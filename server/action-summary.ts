// §2.2 — AI-generated Action List summary (server side).
//
// Single source of truth for:
//   1. The canonical prompt-input string passed to OpenAI.
//   2. The SHA-256 hash stored on the defect.
//
// INVARIANT: the inputs that feed the prompt MUST equal the inputs that feed
// the hash. If you change one, change the other in the same edit. A drift
// between the two would let real source changes leave a summary silently stale
// (or vice versa). Both are produced by buildActionSummaryInput().
//
// Stale-check: a defect's actionSummary is "stale" when
//   actionSummaryInputHash !== sha256(buildActionSummaryInput(defect).promptInput)
// (or when the cached summary is null).
//
// The fallback util (truncate to ≤25 words at a word boundary, ending in a
// full stop) is exported so the same logic is used by:
//   - the regenerate endpoints when the OpenAI key is missing / the call fails;
//   - the renderers, when a defect still has no cached summary at export time.

import crypto from "crypto";

// ---------- Config ----------

// One constant so swapping models later is a single-line change.
export const ACTION_SUMMARY_MODEL = "gpt-4o-mini";
export const ACTION_SUMMARY_TEMPERATURE = 0.2;
export const ACTION_SUMMARY_MAX_TOKENS = 80;
export const ACTION_SUMMARY_TIMEOUT_MS = 10_000;
export const ACTION_SUMMARY_MAX_WORDS = 25;

// Hold-point / safety-critical phrases the model must preserve verbatim.
// Locked list per §2.2 decision: hold point, do not, must not, stop work, WHS.
export const HOLD_POINT_PHRASES = ["hold point", "do not", "must not", "stop work", "WHS"] as const;

// ---------- Canonicalisation ----------

/** Trim and collapse internal whitespace to a single space. */
function normaliseField(value: string | null | undefined): string {
  return String(value ?? "").trim().replace(/\s+/g, " ");
}

export interface ActionSummaryInputs {
  observation: string;     // raw defect.comment
  actionRequired: string;  // raw defect.actionRequired
  category: string;        // resolved category label (or code, or "")
  workType: string;        // resolved work-type label (or code, or "")
}

export interface BuiltActionSummaryInput {
  /** The exact normalised values, in the order they appear in the prompt. */
  fields: { observation: string; actionRequired: string; category: string; workType: string };
  /** Canonical hash key — joined with NUL separators in a fixed field order. */
  hashKey: string;
  /** SHA-256 hex of hashKey. Stored on the defect. */
  hash: string;
}

/**
 * Canonicalise the four prompt fields and produce the hash. Prompt inputs and
 * hash inputs are guaranteed identical because both come out of this one
 * function. Do NOT bypass it.
 *
 * Field order is fixed: observation, actionRequired, category, workType.
 * Hash uses a NUL separator (\u0000) which cannot appear in user text.
 */
export function buildActionSummaryInput(raw: ActionSummaryInputs): BuiltActionSummaryInput {
  const fields = {
    observation: normaliseField(raw.observation),
    actionRequired: normaliseField(raw.actionRequired),
    category: normaliseField(raw.category),
    workType: normaliseField(raw.workType),
  };
  const hashKey = [fields.observation, fields.actionRequired, fields.category, fields.workType].join("\u0000");
  const hash = crypto.createHash("sha256").update(hashKey).digest("hex");
  return { fields, hashKey, hash };
}

// ---------- Fallback truncation ----------

/**
 * Word-boundary truncation. Never cuts mid-word. Returns a single sentence
 * ending in a full stop, ≤ maxWords words.
 *
 * Used when:
 *   - The OpenAI key is missing, the call fails, or returns empty.
 *   - The renderer encounters a defect with no cached summary.
 *
 * Strategy:
 *   1. Prefer actionRequired (it already describes the required action).
 *   2. Fall back to observation if actionRequired is empty.
 *   3. Strip newlines, collapse whitespace.
 *   4. If text has a sentence boundary within the first maxWords words, cut
 *      there. Otherwise take the first maxWords words.
 *   5. Ensure trailing full stop.
 */
export function fallbackActionSummary(
  observation: string | null | undefined,
  actionRequired: string | null | undefined,
  maxWords: number = ACTION_SUMMARY_MAX_WORDS,
): string {
  const source = normaliseField(actionRequired) || normaliseField(observation);
  if (!source) return "";

  // Prefer a sentence boundary within the budget.
  const words = source.split(" ");
  if (words.length <= maxWords) {
    return ensureTerminalPeriod(source);
  }

  // Try to find a sentence end (. ! ?) inside the first maxWords words.
  let charsConsumed = 0;
  let bestEnd = -1;
  for (let i = 0; i < maxWords && i < words.length; i++) {
    const w = words[i];
    charsConsumed += w.length + (i > 0 ? 1 : 0); // +1 for the space we joined on
    if (/[.!?]$/.test(w)) bestEnd = charsConsumed;
  }
  if (bestEnd > 0) {
    return source.slice(0, bestEnd).trim();
  }

  // No sentence boundary in budget — hard cut at word boundary.
  const truncated = words.slice(0, maxWords).join(" ");
  return ensureTerminalPeriod(truncated);
}

function ensureTerminalPeriod(s: string): string {
  const trimmed = s.trim();
  if (!trimmed) return "";
  return /[.!?]$/.test(trimmed) ? trimmed : `${trimmed}.`;
}

// ---------- OpenAI call ----------

const SYSTEM_PROMPT = `You write one-sentence action items for a façade defect remediation report. \
Output a single imperative sentence, ≤25 words, plain English, no preamble, no bullet points, no markdown, no quotes. \
Preserve any hold points and safety-critical phrases ("hold point", "do not", "must not", "stop work", "WHS") verbatim. \
If the input does not describe a required action, infer one from the observation. \
Never invent specifications, materials, or trade names not present in the input.`;

function buildUserPrompt(fields: BuiltActionSummaryInput["fields"]): string {
  return [
    `Observation: ${fields.observation || "(none)"}`,
    `Action required: ${fields.actionRequired || "(none)"}`,
    `Category: ${fields.category || "(none)"}`,
    `Work type: ${fields.workType || "(none)"}`,
  ].join("\n");
}

export interface GenerateResult {
  summary: string;
  source: "ai" | "fallback";
  hash: string;
  /** Set when source === "fallback" to explain why. */
  fallbackReason?: string;
}

/**
 * Generate a one-sentence action summary. Tries OpenAI first; falls back to a
 * word-boundary truncation on any error / missing key / empty response.
 *
 * Pure function: does NOT touch the database. Callers persist the result.
 */
export async function generateActionSummary(raw: ActionSummaryInputs): Promise<GenerateResult> {
  const built = buildActionSummaryInput(raw);
  const apiKey = process.env.OPENAI_API_KEY;

  // Fast path: nothing to summarise at all.
  if (!built.fields.observation && !built.fields.actionRequired) {
    return { summary: "", source: "fallback", hash: built.hash, fallbackReason: "empty-input" };
  }

  if (!apiKey) {
    return {
      summary: fallbackActionSummary(raw.observation, raw.actionRequired),
      source: "fallback",
      hash: built.hash,
      fallbackReason: "missing-api-key",
    };
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), ACTION_SUMMARY_TIMEOUT_MS);
  try {
    const resp = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: ACTION_SUMMARY_MODEL,
        temperature: ACTION_SUMMARY_TEMPERATURE,
        max_tokens: ACTION_SUMMARY_MAX_TOKENS,
        messages: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: buildUserPrompt(built.fields) },
        ],
      }),
      signal: controller.signal,
    });
    if (!resp.ok) {
      const errText = await resp.text().catch(() => "");
      console.warn(`[action-summary] OpenAI HTTP ${resp.status}: ${errText.slice(0, 200)}`);
      return {
        summary: fallbackActionSummary(raw.observation, raw.actionRequired),
        source: "fallback",
        hash: built.hash,
        fallbackReason: `openai-http-${resp.status}`,
      };
    }
    const json: any = await resp.json();
    const content: string = json?.choices?.[0]?.message?.content ?? "";
    const cleaned = sanitiseModelOutput(content);
    if (!cleaned) {
      return {
        summary: fallbackActionSummary(raw.observation, raw.actionRequired),
        source: "fallback",
        hash: built.hash,
        fallbackReason: "empty-completion",
      };
    }
    return { summary: cleaned, source: "ai", hash: built.hash };
  } catch (err: any) {
    const reason = err?.name === "AbortError" ? "openai-timeout" : `openai-error:${err?.message ?? "unknown"}`;
    console.warn(`[action-summary] ${reason}`);
    return {
      summary: fallbackActionSummary(raw.observation, raw.actionRequired),
      source: "fallback",
      hash: built.hash,
      fallbackReason: reason,
    };
  } finally {
    clearTimeout(timer);
  }
}

/**
 * Tidy model output: strip surrounding quotes / backticks, collapse whitespace,
 * drop a trailing newline, and ensure a terminal full stop. Does NOT enforce
 * the ≤25-word limit (the prompt does that) but does cap at a hard word limit
 * as a defensive measure.
 */
function sanitiseModelOutput(raw: string): string {
  let s = (raw ?? "").trim();
  // Strip wrapping quotes / backticks if the model added them despite the prompt.
  s = s.replace(/^['"`]+|['"`]+$/g, "").trim();
  // Collapse internal whitespace / newlines.
  s = s.replace(/\s+/g, " ").trim();
  if (!s) return "";
  // Defensive word cap (the prompt asks for ≤25 but bad outputs happen).
  const words = s.split(" ");
  if (words.length > ACTION_SUMMARY_MAX_WORDS) {
    s = words.slice(0, ACTION_SUMMARY_MAX_WORDS).join(" ");
  }
  return ensureTerminalPeriod(s);
}

// ---------- Staleness helper ----------

/**
 * True when the cached summary is missing or its stored hash doesn't match
 * the live prompt-input hash. Used by the UI badge, the bulk-regenerate count,
 * and the export-time confirm dialog.
 */
export function isActionSummaryStale(
  cached: { actionSummary: string | null; actionSummaryInputHash: string | null },
  raw: ActionSummaryInputs,
): boolean {
  if (!cached.actionSummary) return true;
  const liveHash = buildActionSummaryInput(raw).hash;
  return cached.actionSummaryInputHash !== liveHash;
}
