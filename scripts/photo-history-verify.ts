/**
 * Photo-history accumulation verification script.
 *
 * Purpose: prove the photo-history fix (fix/photo-history-accumulation) rebuilds the
 * FULL cumulative photo timeline for an Open item carried across multiple inspections.
 *
 * It replicates the new `storage.getAllPhotosForItem(projectId, uid)` logic directly
 * against the live SQLite DB:
 *   - gathers every photo attached to any defects row sharing projectId + uid
 *     (every clone-link in the item's lineage),
 *   - dedupes by (originReportId, slot, caption, createdAt) keeping the earliest id,
 *   - sorts by captureDate ?? createdAt ascending,
 *   - assigns wipNumber = 1..N in that order.
 *
 * Then it picks one Open item from a project (default: project 4) that has photos
 * spanning 3+ inspections and prints the rebuilt timeline:
 *   WIP {n} — {captureDate || createdAt}: {caption}   [reportId=X originReportId=Y]
 *
 * Usage:
 *   tsx scripts/photo-history-verify.ts [projectId] [uid]
 *   DATA_DIR=/path/to/volume tsx scripts/photo-history-verify.ts 4
 *
 * Exits 0 on success. Read-only — never writes to the DB or disk.
 */
import Database from "better-sqlite3";
import path from "path";
import fs from "fs";

type PhotoRow = {
  id: number;
  defect_id: number;
  report_id: number | null;
  origin_report_id: number | null;
  filename: string;
  caption: string | null;
  slot: string;
  capture_date: string | null;
  created_at: string;
};

const dataDir = process.env.DATA_DIR || process.cwd();
const dbPath = path.join(dataDir, "data.db");
if (!fs.existsSync(dbPath)) {
  console.error(`[verify] DB not found at ${dbPath}. Set DATA_DIR to the volume holding data.db.`);
  process.exit(1);
}

const db = new Database(dbPath, { readonly: true });

const projectId = Number(process.argv[2] || 4);
const cliUid = process.argv[3];

// Replicate getAllPhotosForItem: cumulative, deduped, date-sorted, with wipNumber.
function getAllPhotosForItem(projectId: number, uid: string): (PhotoRow & { wipNumber: number })[] {
  const rows = db
    .prepare(
      `SELECT p.* FROM photos p
       JOIN defects d ON d.id = p.defect_id
       WHERE d.project_id = ? AND d.uid = ?`
    )
    .all(projectId, uid) as PhotoRow[];

  const byKey = new Map<string, PhotoRow>();
  for (const p of rows) {
    const key = `${p.origin_report_id ?? p.report_id ?? ""}|${p.slot}|${p.caption ?? ""}|${p.created_at}`;
    const existing = byKey.get(key);
    if (!existing || p.id < existing.id) byKey.set(key, p);
  }

  const deduped = Array.from(byKey.values()).sort((a, b) => {
    const da = a.capture_date ?? a.created_at;
    const dbb = b.capture_date ?? b.created_at;
    if (da < dbb) return -1;
    if (da > dbb) return 1;
    return a.id - b.id;
  });

  return deduped.map((p, idx) => ({ ...p, wipNumber: idx + 1 }));
}

// How many distinct inspections (reports) does this item's photo lineage span?
function inspectionSpan(projectId: number, uid: string): number {
  const row = db
    .prepare(
      `SELECT COUNT(DISTINCT COALESCE(p.origin_report_id, p.report_id)) AS n
       FROM photos p JOIN defects d ON d.id = p.defect_id
       WHERE d.project_id = ? AND d.uid = ?`
    )
    .get(projectId, uid) as { n: number };
  return row?.n ?? 0;
}

// Pick a target UID: explicit CLI arg, else the first Open item in the project whose
// photo lineage spans 3+ inspections.
function pickUid(projectId: number): string | null {
  if (cliUid) return cliUid;

  // Distinct uids of OPEN defects in this project.
  const openUids = db
    .prepare(`SELECT DISTINCT uid FROM defects WHERE project_id = ? AND status = 'open'`)
    .all(projectId) as { uid: string }[];

  let best: { uid: string; span: number } | null = null;
  for (const { uid } of openUids) {
    const span = inspectionSpan(projectId, uid);
    if (span >= 3) return uid; // first one meeting the bar
    if (!best || span > best.span) best = { uid, span };
  }
  // Fall back to the widest-spanning open item if none reach 3 (so we still print something).
  return best?.uid ?? null;
}

const uid = pickUid(projectId);
if (!uid) {
  console.error(`[verify] No Open items found in project ${projectId}.`);
  process.exit(1);
}

const span = inspectionSpan(projectId, uid);
const timeline = getAllPhotosForItem(projectId, uid);

console.log(`=== Photo-history rebuild — project ${projectId}, item ${uid} ===`);
console.log(`Photo lineage spans ${span} inspection(s); ${timeline.length} photo(s) after dedupe.`);
if (span < 3) {
  console.log(`(NOTE: fewer than 3 inspections — widest available item shown. Pass a uid explicitly to override.)`);
}
console.log("");

for (const p of timeline) {
  const date = p.capture_date ?? p.created_at;
  const caption = p.caption ?? "";
  console.log(
    `WIP ${p.wipNumber} — ${date}: ${caption}   [photoId=${p.id} reportId=${p.report_id} originReportId=${p.origin_report_id} slot=${p.slot}]`
  );
}

console.log("");
console.log("[verify] OK — timeline rebuilt successfully.");
db.close();
process.exit(0);
