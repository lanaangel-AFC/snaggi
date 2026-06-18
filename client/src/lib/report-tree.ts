// ============================================================================
// report-tree.ts — single source of truth for export routing AND filtering.
//
// Pass 2 (per snaggi-export-profiles-spec.md "Split into two passes"):
//   buildReportTree(data, profile) returns a fully RESOLVED tree:
//     - audience filtering (section-level hardcoded + record-level)
//     - category ordering + treatment (itemise / summarise / hide)
//     - photo trimming (contractor = all; client = NEW + 1 most-recent OLD)
//     - carried-forward register (open items NOT in locationsCovered)
//     - client-only Progress Summary block (auto-computed counts)
//     - appendix mode (contractor = full; client = reference line)
//
//   The DOCX/PDF renderers (render-docx.ts / render-pdf.ts) consume this tree
//   and walk it in master section order. They DO NOT re-filter — the tree IS
//   the contract. This enforces "one dataset, two exports".
// ============================================================================

import { formatLocation, getLocationDimensions } from "@shared/location";
import { sortByUid } from "./render-helpers";

export type ProfileKey = "contractor" | "client";

export type Treatment = "itemise" | "summarise" | "hide";

export type CategoryTreatment = { code: string; treatment: string };

export type ExportProfile = {
  filenameSuffix: string;
  categoryTreatments: CategoryTreatment[];
};

export type ExportProfiles = {
  contractor: ExportProfile;
  client: ExportProfile;
};

// Defaults mirror the schema/backfill D1 defaults so the tree is robust even if
// a project predates the export_profiles backfill.
const DEFAULT_PROFILES: ExportProfiles = {
  contractor: {
    filenameSuffix: "Contractor",
    categoryTreatments: [
      { code: "RR", treatment: "itemise" },
      { code: "PI", treatment: "itemise" },
      { code: "RD", treatment: "itemise" },
      { code: "PN", treatment: "summarise" },
    ],
  },
  client: {
    filenameSuffix: "Client",
    categoryTreatments: [
      { code: "RD", treatment: "itemise" },
      { code: "PN", treatment: "itemise" },
      { code: "PI", treatment: "itemise" },
      { code: "RR", treatment: "summarise" },
    ],
  },
};

const UNCATEGORISED = "__uncat__";
// Single fallback label for rows whose category cannot be resolved. Used for the
// Action List "Category" column.
const UNCATEGORISED_LABEL = "(uncategorised)";

function safeFilenamePart(v: unknown): string {
  return String(v ?? "").replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

// Parse the project's stored exportProfiles JSON, falling back to defaults.
export function parseExportProfiles(raw: unknown): ExportProfiles {
  if (raw && typeof raw === "object") {
    const p = raw as any;
    return {
      contractor: p.contractor || DEFAULT_PROFILES.contractor,
      client: p.client || DEFAULT_PROFILES.client,
    };
  }
  if (typeof raw === "string" && raw.trim()) {
    try {
      const p = JSON.parse(raw);
      return {
        contractor: p.contractor || DEFAULT_PROFILES.contractor,
        client: p.client || DEFAULT_PROFILES.client,
      };
    } catch {
      /* fall through */
    }
  }
  return DEFAULT_PROFILES;
}

// Parse the project's category list (label lookup by code).
export type CategoryDef = { code: string; label: string; isDefault?: boolean };
export function parseCategories(raw: unknown): CategoryDef[] {
  if (Array.isArray(raw)) return raw as CategoryDef[];
  if (typeof raw === "string" && raw.trim()) {
    try {
      const p = JSON.parse(raw);
      if (Array.isArray(p)) return p as CategoryDef[];
    } catch {
      /* fall through */
    }
  }
  return [];
}

// ---------------------------------------------------------------------------
// Tree node types — the renderer contract.
// ---------------------------------------------------------------------------

// A group of defects sharing a category. `treatment` tells the renderer how to
// emit it; the tree-builder has ALREADY dropped hidden groups, so a group in the
// tree is either itemised (rows present) or summarised (count + note only).
export type CategoryGroup =
  | {
      kind: "itemise";
      categoryCode: string; // raw code, or UNCATEGORISED sentinel
      label: string; // human label, e.g. "Rectify" or "(uncategorised)"
      defects: any[]; // filtered defect objects, photos already trimmed
    }
  | {
      kind: "summary";
      categoryCode: string;
      label: string;
      count: number;
      note: string | null; // client: "Itemised in the contractor report"; contractor: null
    };

export type ThreeBucket = {
  new: CategoryGroup[];
  amended: CategoryGroup[];
  completed: CategoryGroup[];
};

export type ProgressSummary = {
  open: number;
  closedThisPeriod: number;
  overdue: number;
  total: number;
};

export type AppendixMode = "full" | "reference";

export type ReportTree = {
  profile: ProfileKey;
  filenameSuffix: string;
  // Resolved download filename WITHOUT extension. Renderers append .docx/.pdf.
  filenameBase: string;
  // Category codes in this profile's configured order (excludes uncategorised).
  categoryOrder: string[];

  // The raw payload, kept for cover/intro fields the renderers still read
  // directly (project/report metadata). Renderers must NOT pull defects/status
  // out of here — those live in the filtered tree nodes below.
  project: any;
  report: any;

  // Section 2: Action List (the summary register), grouped + treated.
  actionList: { groups: CategoryGroup[] };

  // Section 3: Project Status (audience-filtered).
  projectStatus: {
    narratives: any[];
    program: any | null;
    stageMap: any | null;
    empty: boolean; // true => render header + "Nothing for this report"
  };

  // Client only: auto-computed Progress Summary (between Project Status and This
  // Inspection). null for contractor.
  progressSummary: ProgressSummary | null;

  // Section 4: This Inspection (new/amended/completed), each grouped + treated.
  thisInspection: ThreeBucket & { empty: boolean };

  // Section 5: Carried-forward register (open items NOT in locationsCovered).
  // Always itemised (ignores treatment) but uses the profile's category order.
  carriedForward: { groups: CategoryGroup[]; empty: boolean };

  // Section 6: appendix behaviour.
  appendixMode: AppendixMode;
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

// Record-level audience inclusion. `both` always in; otherwise must equal profile.
function audienceIncludes(itemAudience: unknown, profile: ProfileKey): boolean {
  const a = (itemAudience == null ? "both" : String(itemAudience)).toLowerCase();
  return a === "both" || a === profile;
}

// Case-insensitive substring check between the assembled structured-location
// string and the report's locationsCovered text. Returns true if the location is
// considered "covered" this round (so the defect does NOT carry forward).
// Empty/blank locationsCovered => treat as "everything covered" => true.
export function isLocationCovered(structuredLocation: string, locationsCoveredText: unknown): boolean {
  const covered = String(locationsCoveredText ?? "").trim().toLowerCase();
  if (!covered) return true; // empty => everything covered => carried-forward empty
  const loc = String(structuredLocation ?? "").trim().toLowerCase();
  if (!loc) return false; // no location string => can't be matched => carries forward
  return covered.includes(loc);
}

// Computed display status (mirrors the app's stored status -> display mapping).
// stored: open | complete | archived. "Amended" is an open defect touched this
// inspection (observation/action/location amended, but not status->complete).
function displayStatusOf(d: any): "Open" | "Amended" | "Closed" | "Archived" {
  if (d.status === "archived") return "Archived";
  if (d.status === "complete") return "Closed";
  const ev = d.events;
  if (ev && !ev.isNew && ev.amendedFields) {
    const af = ev.amendedFields;
    const amended = af.observation || af.action || af.photos > 0 ||
      af.locationsAdded > 0 || af.locationsAmended > 0;
    if (amended) return "Amended";
  }
  return "Open";
}

// Trim photos for the profile.
//   contractor: ALL photos (OLD + NEW), preserving order.
//   client: all photos with age "current" (originReportId === current report id)
//           PLUS the single most-recent OLD photo (latest createdAt among "prior"
//           photos). The most-recent-OLD rule REPLACES the entire prior set; if a
//           defect has neither current nor prior photos, the array is empty.
function trimPhotos(defect: any, profile: ProfileKey, currentReportId: number): any[] {
  const photos: any[] = Array.isArray(defect.photos) ? defect.photos : [];
  if (profile === "contractor") return photos;

  const ageOf = (p: any): "current" | "prior" => {
    // A photo is "current" if it first appeared on this report. Respect the
    // server's isThisInspection flag (which already honours newOverride) when
    // present, otherwise fall back to originReportId/reportId comparison.
    if (typeof p.isThisInspection === "boolean") return p.isThisInspection ? "current" : "prior";
    const origin = p.originReportId ?? p.reportId;
    return origin === currentReportId ? "current" : "prior";
  };

  const current = photos.filter((p) => ageOf(p) === "current");
  const prior = photos.filter((p) => ageOf(p) === "prior");

  let mostRecentOld: any | null = null;
  for (const p of prior) {
    if (!mostRecentOld) { mostRecentOld = p; continue; }
    // latest createdAt wins
    if (String(p.createdAt ?? "").localeCompare(String(mostRecentOld.createdAt ?? "")) > 0) {
      mostRecentOld = p;
    }
  }

  // Keep the natural slot order: current photos in their existing order, then the
  // single most-recent OLD photo appended (if any).
  const result = [...current];
  if (mostRecentOld) result.push(mostRecentOld);
  return result;
}

// Apply a defect's profile-specific photo trim in-place (returns a shallow clone).
function withTrimmedPhotos(defect: any, profile: ProfileKey, currentReportId: number): any {
  return { ...defect, photos: trimPhotos(defect, profile, currentReportId) };
}

// Build category-grouped output for a set of defects.
//   - groups follow `order` (profile's category order); uncategorised always last.
//   - treatment from `treatmentMap`; uncategorised is always "itemise" and cannot
//     be hidden.
//   - `ignoreTreatment` (carried-forward) forces every group to itemise but keeps
//     the category order.
function buildGroups(
  defects: any[],
  order: string[],
  treatmentMap: Map<string, Treatment>,
  labelMap: Map<string, string>,
  profile: ProfileKey,
  ignoreTreatment: boolean,
  // Action List only: resolve a per-row category label and sort completed
  // (displayStatus === "Closed") rows to the bottom of each itemised group.
  actionListMode: boolean = false,
): CategoryGroup[] {
  // Bucket defects by category code (uncategorised under sentinel).
  const buckets = new Map<string, any[]>();
  for (const d of defects) {
    const code = d.categoryCode && String(d.categoryCode).trim() ? String(d.categoryCode) : UNCATEGORISED;
    if (!buckets.has(code)) buckets.set(code, []);
    buckets.get(code)!.push(d);
  }

  const groups: CategoryGroup[] = [];

  const emit = (code: string, treatment: Treatment) => {
    let rows = (buckets.get(code) || []).slice().sort(sortByUid);
    if (rows.length === 0) return; // nothing in this category this section
    const label = code === UNCATEGORISED ? UNCATEGORISED_LABEL : (labelMap.get(code) || code);
    const effective: Treatment = ignoreTreatment ? "itemise" : treatment;
    if (effective === "hide") return; // omit entirely
    if (effective === "summarise") {
      groups.push({
        kind: "summary",
        categoryCode: code,
        label,
        count: rows.length,
        note: profile === "client" ? "Itemised in the contractor report" : null,
      });
    } else {
      if (actionListMode) {
        // Resolve the category label once per row (looked up by the row's own
        // categoryCode; single fallback constant for unresolved codes).
        rows = rows.map((d) => {
          const rowCode = d.categoryCode && String(d.categoryCode).trim() ? String(d.categoryCode) : "";
          const categoryLabel = rowCode ? (labelMap.get(rowCode) || rowCode) : UNCATEGORISED_LABEL;
          return { ...d, categoryLabel };
        });
        // Stable sort: completed (displayStatus === "Closed") rows to the bottom,
        // preserving existing order within each partition.
        const open = rows.filter((d) => displayStatusOf(d) !== "Closed");
        const closed = rows.filter((d) => displayStatusOf(d) === "Closed");
        rows = [...open, ...closed];
      }
      groups.push({ kind: "itemise", categoryCode: code, label, defects: rows });
    }
  };

  // Ordered known categories first.
  for (const code of order) {
    if (code === UNCATEGORISED) continue;
    emit(code, treatmentMap.get(code) || "itemise");
  }
  // Any category present on defects but NOT in the profile order (e.g. a brand
  // new category that hasn't been synced yet) — itemise, after the ordered ones.
  for (const code of buckets.keys()) {
    if (code === UNCATEGORISED) continue;
    if (!order.includes(code)) emit(code, treatmentMap.get(code) || "itemise");
  }
  // Uncategorised always last, always itemise.
  emit(UNCATEGORISED, "itemise");

  return groups;
}

// ---------------------------------------------------------------------------
// Build the Pass 2 report tree. `data` is the /report-data response.
// ---------------------------------------------------------------------------

export function buildReportTree(data: any, profile: ProfileKey): ReportTree {
  const profiles = parseExportProfiles(data?.project?.exportProfiles);
  const chosen = profiles[profile] || DEFAULT_PROFILES[profile];
  const filenameSuffix = (chosen?.filenameSuffix || (profile === "client" ? "Client" : "Contractor")).trim();

  const afcRef = safeFilenamePart(data?.project?.afcReference);
  const inspectionNumber = safeFilenamePart(data?.report?.inspectionNumber);
  const suffixPart = safeFilenamePart(filenameSuffix);
  const filenameBase = `${afcRef}_SVR${inspectionNumber}_${suffixPart}`;

  const treatments = chosen?.categoryTreatments || [];
  const categoryOrder = treatments.map((t) => t.code);
  const treatmentMap = new Map<string, Treatment>();
  for (const t of treatments) {
    const tr = (t.treatment as Treatment);
    treatmentMap.set(t.code, tr === "summarise" || tr === "hide" ? tr : "itemise");
  }

  const cats = parseCategories(data?.project?.categories);
  const labelMap = new Map<string, string>();
  for (const c of cats) labelMap.set(c.code, c.label);

  const dims = getLocationDimensions(data?.project?.locationDimensions);
  const currentReportId = data?.report?.id;
  const locationsCovered = data?.report?.locationsCovered;

  // ---- Record-level audience filter on the report's defects ----
  const allDefectsRaw: any[] = Array.isArray(data?.defects) ? data.defects : [];
  const includedDefects = allDefectsRaw
    .filter((d) => audienceIncludes(d.audience, profile))
    .map((d) => withTrimmedPhotos(d, profile, currentReportId));

  const defectsOnlyAndObs = includedDefects; // we group defects & observations together by category

  // ---- Classify into new / amended / completed (mirrors renderer logic) ----
  const allHaveEvents = includedDefects.some((d) => d.events);
  const isCompletedThisInspection = (d: any): boolean =>
    !d.events?.isNew && !!d.events?.amendedFields?.statusChange && d.events.amendedFields.statusChange.to === "complete";
  const hasAnyAmendedField = (d: any): boolean => {
    if (!d.events) return false;
    const af = d.events.amendedFields;
    return af.observation || af.action || af.photos > 0 || af.locationsAdded > 0 || af.locationsAmended > 0 || !!af.statusChange;
  };

  const newItems = includedDefects.filter((d) => allHaveEvents ? d.events?.isNew : false);
  const amendedItems = includedDefects.filter((d) =>
    allHaveEvents ? (!d.events?.isNew && !isCompletedThisInspection(d) && hasAnyAmendedField(d)) : false);
  const completedItems = includedDefects.filter((d) =>
    allHaveEvents ? isCompletedThisInspection(d) : false);

  // ---- Action List: every defect on the report, grouped + treated ----
  const actionGroups = buildGroups(defectsOnlyAndObs, categoryOrder, treatmentMap, labelMap, profile, false, true);

  // ---- This Inspection: three buckets, each grouped + treated ----
  const tiNew = buildGroups(newItems, categoryOrder, treatmentMap, labelMap, profile, false);
  const tiAmended = buildGroups(amendedItems, categoryOrder, treatmentMap, labelMap, profile, false);
  const tiCompleted = buildGroups(completedItems, categoryOrder, treatmentMap, labelMap, profile, false);
  const tiEmpty = tiNew.length === 0 && tiAmended.length === 0 && tiCompleted.length === 0;

  // ---- Carried-forward register: open items NOT covered this round ----
  // "open" = not complete/archived. Filter by locationStructured vs locationsCovered.
  const carriedDefects = includedDefects.filter((d) => {
    if (d.status === "complete" || d.status === "archived") return false;
    const locStr = formatLocation(d.locationStructured, dims);
    return !isLocationCovered(locStr, locationsCovered);
  });
  const carriedGroups = buildGroups(carriedDefects, categoryOrder, treatmentMap, labelMap, profile, true);

  // ---- Project Status: audience-filtered ----
  const ps = data?.projectStatus || {};
  const narratives = (Array.isArray(ps.narrativeHolds) ? ps.narrativeHolds : [])
    .filter((n: any) => audienceIncludes(n.audience, profile))
    .slice()
    .sort((a: any, b: any) => {
      const so = (a.sortOrder ?? 0) - (b.sortOrder ?? 0);
      if (so !== 0) return so;
      return String(a.dateRaised ?? "").localeCompare(String(b.dateRaised ?? ""));
    });
  const program = ps.programSchedule && audienceIncludes(ps.programSchedule.audience, profile)
    ? ps.programSchedule : null;
  const stageMap = ps.stageProgressMap && audienceIncludes(ps.stageProgressMap.audience, profile)
    ? ps.stageProgressMap : null;
  const projectStatusEmpty = narratives.length === 0 && !program && !stageMap;

  // ---- Progress Summary (client only) ----
  // "Closed this period" rule (three-tier):
  //   A defect counts as closed this period iff its STORED status is "complete"
  //   AND it was closed ON the currently-rendered report. "Closed on this report"
  //   is detected with the following priority:
  //
  //   1. PRIMARY (authoritative when present): statusHistory contains a row with
  //      newStatus === "complete" AND reportId === currentReport.id. This is the
  //      milestone-8 status_history signal and is trusted whenever it exists.
  //
  //   2. FALLBACK (closure never recorded in history): the defect has NO
  //      statusHistory row with newStatus === "complete" anywhere (so tier 1 can
  //      never fire for it), AND defect.dateClosed is set, AND dateClosed >=
  //      currentReport.inspectionDate. Many defects were closed before the
  //      status_history table shipped (or via a path that didn't write a row),
  //      so statusHistory is unreliable retroactively; dateClosed >= this
  //      report's open date is the heuristic that this report is the closing one.
  //
  //   3. SAFETY NET (same-inspection create+close, no history at all): the defect
  //      has NO statusHistory entries whatsoever, AND defect.reportId ===
  //      currentReport.id (it lives on this report). Covers a defect created and
  //      closed within one inspection before any history row was written.
  //
  //   If none of the three match, the defect is NOT counted — it was most likely
  //   closed in an earlier inspection.
  //
  //   dateClosed may be a YYYY-MM-DD string (live close path) or a full ISO
  //   timestamp (server date_closed backfill from status_history.created_at), and
  //   inspectionDate is YYYY-MM-DD, so both are normalised to their date portion
  //   (first 10 chars) before the lexical compare.
  const dateOnly = (v: unknown): string => String(v ?? "").slice(0, 10);
  const closedOnThisReport = (d: any): boolean => {
    if (d.status !== "complete") return false;
    const rows: any[] = Array.isArray(d.statusHistory) ? d.statusHistory : [];
    const completeRows = rows.filter((r) => r && r.newStatus === "complete");

    // Tier 1 — authoritative status_history signal for this report.
    if (completeRows.some((r) => r.reportId === currentReportId)) return true;

    // Tier 2 — no recorded closure in history at all; fall back to dateClosed.
    const inspectionDate = dateOnly(data?.report?.inspectionDate);
    if (completeRows.length === 0 && d.dateClosed && inspectionDate) {
      if (dateOnly(d.dateClosed) >= inspectionDate) return true;
    }

    // Tier 3 — safety net: no history whatsoever, but the defect lives on this
    // report (created + closed within this inspection before history existed).
    if (rows.length === 0 && d.reportId === currentReportId) return true;

    return false;
  };

  let progressSummary: ProgressSummary | null = null;
  if (profile === "client") {
    const today = new Date();
    let open = 0, closedThisPeriod = 0, overdue = 0;
    for (const d of includedDefects) {
      const ds = displayStatusOf(d);
      if (ds === "Open" || ds === "Amended") {
        open++;
        if (d.dueDate && new Date(d.dueDate) < today) overdue++;
      }
      // Closed this period := any defect on this report whose stored status is
      // 'complete' AND whose closing event (newStatus='complete' row in
      // statusHistory) occurred on the currently-rendered report.
      if (ds === "Closed" && closedOnThisReport(d)) closedThisPeriod++;
    }
    progressSummary = { open, closedThisPeriod, overdue, total: includedDefects.length };
  }

  // ---- Appendix mode ----
  const appendixMode: AppendixMode = profile === "client" ? "reference" : "full";

  return {
    profile,
    filenameSuffix,
    filenameBase,
    categoryOrder,
    project: data?.project,
    report: data?.report,
    actionList: { groups: actionGroups },
    projectStatus: { narratives, program, stageMap, empty: projectStatusEmpty },
    progressSummary,
    thisInspection: { new: tiNew, amended: tiAmended, completed: tiCompleted, empty: tiEmpty },
    carriedForward: { groups: carriedGroups, empty: carriedGroups.length === 0 },
    appendixMode,
  };
}
