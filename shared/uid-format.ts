// Project-configurable record UID format.
//
// A project may define its own UID as an ordered set of named parts joined by a
// pattern, e.g. GRC Fin Rectification:
//   parts:   Level, Fin, Bracket
//   pattern: "{level}.{fin}-{bracket}"   →   Bracket B, Fin 14, Level 9 = "9.14-B"
//
// Projects with no format (empty projects.uid_format) keep the standard
// Elevation-Drop-Level-WorkType-Number UID; nothing in this module changes them.
//
// In a formatted project:
//   - the part values are the record's location and live in defects.location_structured
//     (keyed by part key), so the register, cards and filters show them as the location;
//   - the UID is exactly the rendered pattern, with no sequence number, so a UID names
//     one place: a second record at the same place opens the existing one instead;
//   - work type is still recorded for the category label, but is not part of the UID.

export interface UidFormatPart {
  key: string;      // lowercase identifier used in the pattern, e.g. "fin"
  label: string;    // shown in the form and report, e.g. "Fin"
  options: string[]; // pick list; free typing is always allowed as well
}

export interface UidFormat {
  pattern: string;
  parts: UidFormatPart[];
}

export type UidPartValues = Record<string, string | null | undefined>;

// Reserved placeholder for the record's work type code (e.g. GL), so a project can
// end its UID with the work type the way the standard UID does: "{window}-{worktype}".
export const WORK_TYPE_TOKEN = "worktype";

// Pattern used when parts are set but no pattern is typed: the parts joined with "-",
// then the work type, matching the standard UID's order (location, then work type).
export function defaultUidPattern(keys: string[]): string {
  const ks = keys.filter((k) => k && k !== WORK_TYPE_TOKEN);
  if (ks.length === 0) return "";
  return [...ks, WORK_TYPE_TOKEN].map((k) => `{${k}}`).join("-");
}

// Turn a label into a pattern key: "Fin no." → "finno".
export function partKeyFromLabel(label: string): string {
  return String(label || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

// Parse the stored projects.uid_format value. Returns null for "no custom format"
// (empty, malformed, no parts, or a pattern that uses none of the parts).
export function parseUidFormat(raw: unknown): UidFormat | null {
  if (!raw) return null;
  let obj: any = raw;
  if (typeof raw === "string") {
    const t = raw.trim();
    if (!t) return null;
    try { obj = JSON.parse(t); } catch { return null; }
  }
  if (!obj || typeof obj !== "object") return null;
  const parts: UidFormatPart[] = (Array.isArray(obj.parts) ? obj.parts : [])
    .map((p: any) => {
      const label = String(p?.label ?? "").trim();
      const key = partKeyFromLabel(p?.key || label);
      const options = (Array.isArray(p?.options) ? p.options : [])
        .map((o: any) => String(o ?? "").trim())
        .filter((o: string) => o !== "");
      return { key, label: label || key, options };
    })
    .filter((p: UidFormatPart) => p.key !== "");
  if (parts.length === 0) return null;
  // An empty pattern means "use the default", not "no format": someone who has set up
  // parts expects them to be used.
  const pattern = String(obj.pattern ?? "").trim() || defaultUidPattern(parts.map((p) => p.key));
  if (!pattern || !parts.some((p) => pattern.includes(`{${p.key}}`))) return null;
  return { pattern, parts };
}

type Token = { kind: "lit"; text: string } | { kind: "part"; key: string };

function tokenize(pattern: string): Token[] {
  const out: Token[] = [];
  const re = /\{([a-z0-9]+)\}/g;
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(pattern))) {
    if (m.index > last) out.push({ kind: "lit", text: pattern.slice(last, m.index) });
    out.push({ kind: "part", key: m[1] });
    last = m.index + m[0].length;
  }
  if (last < pattern.length) out.push({ kind: "lit", text: pattern.slice(last) });
  return out;
}

const clean = (v: unknown) => String(v ?? "").trim();

// Values for renderUid: the record's location parts plus its work type code. A part
// that happens to be called "worktype" wins over the work type column.
export function uidValues(parts: UidPartValues, workTypeCode: unknown): UidPartValues {
  return { [WORK_TYPE_TOKEN]: String(workTypeCode ?? "").trim(), ...parts };
}

// Build the UID from part values. A blank part is skipped together with the
// separator in front of it, so Level 9 / Fin 14 with no bracket gives "9.14".
export function renderUid(fmt: UidFormat, values: UidPartValues): string {
  const tokens = tokenize(fmt.pattern);
  let out = "";
  let pendingLit = "";
  let seenPart = false; // text before the first placeholder is a prefix (e.g. "W"), not a separator
  for (const t of tokens) {
    if (t.kind === "lit") { pendingLit += t.text; continue; }
    const v = clean(values?.[t.key]);
    const lead = out || !seenPart ? pendingLit : "";
    seenPart = true;
    if (!v) { pendingLit = ""; continue; }
    out += lead + v;
    pendingLit = "";
  }
  return out;
}

// True when every part in the format has a value.
export function isUidComplete(fmt: UidFormat, values: UidPartValues): boolean {
  return fmt.parts.every((p) => clean(values?.[p.key]) !== "");
}

// The pattern with labels in place of keys, for the report and setup screen:
// "{level}.{fin}-{bracket}" → "Level.Fin-Bracket".
export function describeUidFormat(fmt: UidFormat): string {
  const labels = new Map<string, string>([[WORK_TYPE_TOKEN, "Work type"], ...fmt.parts.map((p) => [p.key, p.label] as [string, string])]);
  return tokenize(fmt.pattern)
    .map((t) => (t.kind === "lit" ? t.text : labels.get(t.key) || t.key))
    .join("");
}

// Example UID using the first pick-list value of each part (or a placeholder).
export function exampleUid(fmt: UidFormat): string {
  const sample: Record<string, string> = {};
  sample[WORK_TYPE_TOKEN] = "GL";
  fmt.parts.forEach((p, i) => { sample[p.key] = p.options[0] || String(i + 1); });
  return renderUid(fmt, sample);
}

function escapeRe(s: string) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// Read part values back out of a UID string. Used when a formatted record has no
// stored location (e.g. older rows) and to pick a location out of note text.
// Returns null when the text doesn't match the full pattern.
export function parseUidWithFormat(fmt: UidFormat, uid: string): Record<string, string> | null {
  const tokens = tokenize(fmt.pattern);
  const keys: string[] = [];
  let src = "^";
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t.kind === "lit") { src += escapeRe(t.text); continue; }
    keys.push(t.key);
    // A part can't contain the separator that follows it.
    const next = tokens[i + 1];
    const stop = next && next.kind === "lit" ? escapeRe(next.text.charAt(0)) : "";
    src += stop ? `([^${stop}\\s]+?)` : "([A-Za-z0-9]+)";
  }
  src += "$";
  const m = new RegExp(src).exec(String(uid || "").trim());
  if (!m) return null;
  const out: Record<string, string> = {};
  const partKeys = new Set(fmt.parts.map((p) => p.key));
  // The work type is a column of its own, not a location part.
  keys.forEach((k, i) => { if (partKeys.has(k)) out[k] = m[i + 1]; });
  return out;
}
