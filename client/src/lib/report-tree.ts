// ============================================================================
// report-tree.ts — Pass 1 single source of truth for export routing.
//
// Pass 1 scope (per snaggi-export-profiles-spec.md, lines 304/309):
//   - buildReportTree(data, profile) computes the RESOLVED FILENAME and the
//     per-profile CATEGORY ORDER metadata.
//   - Both profiles produce the SAME report body in Pass 1; the only observable
//     differences are (a) the filename suffix and (b) the category ordering that
//     the renderers may use for grouping. No summarise/hide behaviour, no
//     audience filtering, no photo trimming — those are Pass 2.
//   - The renderers (render-pdf / render-docx, currently inline in
//     report-detail.tsx) consume `tree.filename` so the profile→filename routing
//     is verifiable end-to-end.
// ============================================================================

export type ProfileKey = "contractor" | "client";

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

function safeFilenamePart(v: unknown): string {
  return String(v ?? "").replace(/[^a-zA-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
}

// Parse the project's stored exportProfiles JSON, falling back to defaults.
export function parseExportProfiles(raw: unknown): ExportProfiles {
  if (raw && typeof raw === "object") return raw as ExportProfiles;
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

export type ReportTree = {
  profile: ProfileKey;
  filenameSuffix: string;
  // Resolved download filename WITHOUT extension. Renderers append .docx/.pdf.
  filenameBase: string;
  // Category codes in this profile's configured order (Pass 1: order only).
  categoryOrder: string[];
};

// Build the Pass 1 report tree. `data` is the /report-data response.
export function buildReportTree(data: any, profile: ProfileKey): ReportTree {
  const profiles = parseExportProfiles(data?.project?.exportProfiles);
  const chosen = profiles[profile] || DEFAULT_PROFILES[profile];
  const filenameSuffix = (chosen?.filenameSuffix || (profile === "client" ? "Client" : "Contractor")).trim();

  const afcRef = safeFilenamePart(data?.project?.afcReference);
  const inspectionNumber = safeFilenamePart(data?.report?.inspectionNumber);
  const suffixPart = safeFilenamePart(filenameSuffix);

  // {afcReference}_SVR{inspectionNumber}_{suffix}
  const filenameBase = `${afcRef}_SVR${inspectionNumber}_${suffixPart}`;

  const categoryOrder = (chosen?.categoryTreatments || []).map((t) => t.code);

  return { profile, filenameSuffix, filenameBase, categoryOrder };
}
