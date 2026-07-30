// Synthetic-fixture harness for the §2.2 Action List table.
//
// Renders the real renderDocx() / renderPdf() against a hand-built ReportTree so
// the §2.2 column headers and cells can be eyeballed without a live project.
// Emits:
//   fixtures/action-list-section-2.2.docx
//   fixtures/action-list-section-2.2.pdf
//   fixtures/action-list-section-2.2.txt   (§2.2 tables extracted from the DOCX)
//
// Run with: npx tsx scripts/action-list-fixture.ts

import { mkdirSync, writeFileSync } from "fs";
import { dirname, join } from "path";
import { fileURLToPath } from "url";

const here = dirname(fileURLToPath(import.meta.url));
const outDir = join(here, "..", "fixtures");

// The renderers are browser modules: they read Blob/FileReader and fetch the AFC
// logo. Node 18+ supplies Blob; FileReader is stubbed just enough for
// blobToDataUrl, and fetch is forced to fail so loadAfcLogo() returns null.
(globalThis as any).fetch = async () => { throw new Error("no network in fixture"); };
(globalThis as any).FileReader = class {
  onload: any = null;
  onerror: any = null;
  result: any = null;
  readAsDataURL(blob: any) {
    blob.arrayBuffer().then((buf: ArrayBuffer) => {
      this.result = `data:${blob.type || "application/octet-stream"};base64,${Buffer.from(buf).toString("base64")}`;
      this.onload?.({ target: this });
    });
  }
};

const defect = (
  uid: string,
  comment: string,
  actionSummary: string,
  assignedTo: string,
  dueDate: string,
  status: string,
  categoryLabel: string,
) => ({
  uid,
  comment,
  action: actionSummary,
  actionSummary,
  assignedTo,
  dueDate,
  status,
  categoryLabel,
  recordType: "defect",
  locationStructured: null,
  locations: [],
  photos: [],
  statusHistory: [],
  events: { isNew: true, amendedFields: { photos: 0, locationsAdded: 0, locationsAmended: 0 }, photosAddedThisInspection: [] },
});

// One group per status grouping the Action List renders, to prove the fix covers
// all of them (including the uncategorised bucket).
const groups = [
  {
    kind: "itemise" as const,
    categoryCode: "RR",
    label: "Rectify",
    defects: [
      defect("E-04-10-LR-01", "Cracked lintel above window head, spalling to soffit.", "Rectify cracked lintel and make good soffit.", "ABC Remedial", "2026-08-14", "open", "Rectify"),
      defect("N-02-03-RR-02", "Render debonded at parapet return.", "Remove debonded render and reinstate to spec.", "ABC Remedial", "2026-07-01", "open", "Rectify"),
    ],
  },
  {
    kind: "itemise" as const,
    categoryCode: "WIP",
    label: "Work in Progress",
    defects: [
      defect("S-01-05-GW-03", "General making good ongoing to south elevation.", "Continue general works and confirm completion.", "Site Team", "2026-09-02", "open", "Work in Progress"),
    ],
  },
  {
    kind: "itemise" as const,
    categoryCode: "CMP",
    label: "Complete",
    defects: [
      defect("W-03-07-CK-04", "Perimeter caulking renewed to glazing units.", "Caulking renewal complete — no further action.", "Seal Co", "2026-06-20", "complete", "Complete"),
    ],
  },
  {
    kind: "itemise" as const,
    categoryCode: "__uncat__",
    label: "(uncategorised)",
    defects: [
      // Blank assignedTo / dueDate to show the em-dash placeholders.
      defect("E-05-11-PT-05", "Paint finish patchy to east elevation drop 5.", "Rework paint finish to achieve uniform coverage.", "", "", "open", "(uncategorised)"),
    ],
  },
];

const tree: any = {
  profile: "contractor",
  filenameSuffix: "Contractor",
  filenameBase: "260730-AFC-24001-1TestSt-TestProject-SVR-03",
  categoryOrder: ["RR", "WIP", "CMP"],
  project: {
    id: 1,
    name: "Test Project",
    address: "1 Test Street, Sydney",
    client: "Test Owners Corporation",
    inspector: "A. Inspector",
    afcReference: "AFC-24001",
    reportTitle: "Facade Remediation Works",
    locationDimensions: null,
    categories: [],
    roles: "[]",
    areaRefTemplate: "",
    exportProfiles: null,
  },
  report: {
    id: 10,
    inspectionNumber: "3",
    revision: "01",
    inspectionDate: "2026-07-30",
    createdAt: "2026-07-30T00:00:00Z",
    locationsCovered: "",
    projectSnapshot: null,
  },
  actionList: { groups },
  projectStatus: { narratives: [], program: null, stageMap: null, empty: true },
  progressSummary: { open: 4, closedThisPeriod: 1, overdue: 1, total: 5 },
  thisInspection: { new: [], amended: [], completed: [], empty: true },
  carriedForward: { groups: [], empty: true },
  appendixMode: "full",
};

// Extract every w:tbl whose header row starts with "ID" — i.e. the §2.2 tables —
// and print them as aligned text grids.
function extractActionTables(xml: string): string {
  const out: string[] = [];
  const tables = xml.match(/<w:tbl>[\s\S]*?<\/w:tbl>/g) || [];
  for (const tbl of tables) {
    const rows = tbl.match(/<w:tr[\s>][\s\S]*?<\/w:tr>/g) || [];
    const grid = rows.map((tr) => {
      const cells = tr.match(/<w:tc>[\s\S]*?<\/w:tc>/g) || [];
      return cells.map((tc) => (tc.match(/<w:t[^>]*>([^<]*)<\/w:t>/g) || [])
        .map((t) => t.replace(/<[^>]+>/g, "")).join(""));
    });
    if (grid.length === 0 || grid[0][0] !== "ID") continue;
    const cols = Math.max(...grid.map((r) => r.length));
    const widths: number[] = [];
    for (let c = 0; c < cols; c++) widths.push(Math.max(...grid.map((r) => (r[c] || "").length)));
    out.push(grid.map((r, ri) => {
      const line = r.map((v, c) => (v || "").padEnd(widths[c])).join(" | ");
      return ri === 0 ? `${line}\n${widths.map((w) => "-".repeat(w)).join("-+-")}` : line;
    }).join("\n"));
  }
  return out.join("\n\n");
}

async function main() {
  mkdirSync(outDir, { recursive: true });

  const { renderDocx } = await import("../client/src/lib/render-docx");
  const docxBlob = await renderDocx(tree, { profile: "contractor" });
  const docxBuf = Buffer.from(await docxBlob.arrayBuffer());
  writeFileSync(join(outDir, "action-list-section-2.2.docx"), docxBuf);

  // Pull document.xml straight out of the .docx to render the table as text.
  const JSZip = (await import("jszip")).default;
  const zip = await JSZip.loadAsync(docxBuf);
  const xml = await zip.file("word/document.xml")!.async("string");
  const text = extractActionTables(xml);
  writeFileSync(
    join(outDir, "action-list-section-2.2.txt"),
    `Section 2.2 Action List — synthetic fixture (extracted from the generated DOCX)\n\n${text}\n`,
  );

  const { renderPdf } = await import("../client/src/lib/render-pdf");
  const pdfBlob = await renderPdf(tree, { profile: "contractor" });
  writeFileSync(join(outDir, "action-list-section-2.2.pdf"), Buffer.from(await pdfBlob.arrayBuffer()));

  console.log(text);
  console.log("\nWrote fixtures/action-list-section-2.2.{docx,pdf,txt}");
}

main().catch((e) => { console.error(e); process.exit(1); });
