import {
  type User, type InsertUser, users,
  type Project, type InsertProject, projects,
  type Report, type InsertReport, reports,
  type Defect, type InsertDefect, defects,
  type Photo, type InsertPhoto, photos,
  type Elevation, type InsertElevation, elevations,
  type Marker, type InsertMarker, markers,
  type ObservationHistory, type InsertObservationHistory, observationHistory,
  type ActionHistory, type InsertActionHistory, actionHistory,
  type DefectLocation, type InsertDefectLocation, defectLocations,
  type ShareLink, type InsertShareLink, shareLinks,
  type StatusHistory, type InsertStatusHistory, statusHistory,
} from "@shared/schema";
import { drizzle } from "drizzle-orm/better-sqlite3";
import Database from "better-sqlite3";
import { eq, and, desc } from "drizzle-orm";
import path from "path";
import fs from "fs";

// Use DATA_DIR env var for persistent storage (Railway volume), fallback to cwd
const dataDir = process.env.DATA_DIR || process.cwd();
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });

const dbPath = path.join(dataDir, "data.db");
const sqlite = new Database(dbPath);
sqlite.pragma("journal_mode = WAL");

// Auto-create tables on startup (no need for drizzle-kit push on deploy)
sqlite.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    address TEXT NOT NULL,
    client TEXT NOT NULL,
    inspector TEXT NOT NULL,
    afc_reference TEXT DEFAULT '',
    revision TEXT DEFAULT '01',

    inspection_number TEXT DEFAULT '',
    inspection_date TEXT DEFAULT '',
    locations_covered TEXT DEFAULT '',
    elevations TEXT DEFAULT '[]',
    attendees TEXT DEFAULT '[]',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    inspection_number TEXT DEFAULT '',
    inspection_date TEXT DEFAULT '',
    revision TEXT DEFAULT '01',
    locations_covered TEXT DEFAULT '',
    elevations TEXT DEFAULT '[]',
    attendees TEXT DEFAULT '[]',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS defects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    report_id INTEGER,
    uid TEXT NOT NULL,
    date_opened TEXT NOT NULL,
    date_closed TEXT,
    comment TEXT NOT NULL,
    action_required TEXT NOT NULL,
    assigned_to TEXT DEFAULT '',
    due_date TEXT DEFAULT '',
    verification_method TEXT NOT NULL,
    verification_person TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'open',
    record_type TEXT NOT NULL DEFAULT 'defect',
    updated_at TEXT
  );
  CREATE TABLE IF NOT EXISTS photos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    defect_id INTEGER NOT NULL,
    filename TEXT NOT NULL,
    caption TEXT,
    slot TEXT NOT NULL DEFAULT 'wip1',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS elevations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    name TEXT NOT NULL,
    filename TEXT NOT NULL,
    file_type TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (project_id) REFERENCES projects(id) ON DELETE CASCADE
  );
  CREATE TABLE IF NOT EXISTS markers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    elevation_id INTEGER NOT NULL,
    defect_id INTEGER,
    defect_uid TEXT NOT NULL,
    status TEXT NOT NULL DEFAULT 'open',
    note TEXT,
    x_percent REAL NOT NULL,
    y_percent REAL NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (elevation_id) REFERENCES elevations(id) ON DELETE CASCADE
  );
  CREATE TABLE IF NOT EXISTS observation_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    defect_id INTEGER NOT NULL,
    report_id INTEGER NOT NULL,
    text TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (defect_id) REFERENCES defects(id) ON DELETE CASCADE,
    FOREIGN KEY (report_id) REFERENCES reports(id)
  );
  CREATE TABLE IF NOT EXISTS action_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    defect_id INTEGER NOT NULL,
    report_id INTEGER NOT NULL,
    text TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (defect_id) REFERENCES defects(id) ON DELETE CASCADE,
    FOREIGN KEY (report_id) REFERENCES reports(id)
  );
  CREATE TABLE IF NOT EXISTS defect_locations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    defect_id INTEGER NOT NULL,
    uid TEXT DEFAULT '',
    "drop" TEXT DEFAULT '',
    elevation TEXT DEFAULT '',
    level TEXT DEFAULT '',
    description TEXT DEFAULT '',
    display_order INTEGER DEFAULT 0,
    created_at TEXT NOT NULL,
    FOREIGN KEY (defect_id) REFERENCES defects(id) ON DELETE CASCADE
  );
  CREATE TABLE IF NOT EXISTS global_settings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    work_types TEXT DEFAULT '[]'
  );
  CREATE TABLE IF NOT EXISTS share_links (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    token TEXT NOT NULL UNIQUE,
    report_id INTEGER NOT NULL,
    recipient_name TEXT DEFAULT '',
    created_at TEXT NOT NULL
  );
  CREATE TABLE IF NOT EXISTS status_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    defect_id INTEGER NOT NULL,
    old_status TEXT,
    new_status TEXT NOT NULL,
    report_id INTEGER,
    created_at TEXT,
    FOREIGN KEY (defect_id) REFERENCES defects(id) ON DELETE CASCADE
  );
`);

// Add new columns to existing tables (safe: ALTER TABLE ADD COLUMN IF NOT EXISTS via try/catch)
const safeAddColumn = (table: string, col: string, colDef: string) => {
  try { sqlite.exec(`ALTER TABLE ${table} ADD COLUMN ${col} ${colDef}`); } catch {}
};
safeAddColumn("projects", "locations_covered", "TEXT DEFAULT ''");
safeAddColumn("projects", "elevations", "TEXT DEFAULT '[]'");
safeAddColumn("reports", "elevations", "TEXT DEFAULT '[]'");
safeAddColumn("defects", "record_type", "TEXT NOT NULL DEFAULT 'defect'");
safeAddColumn("projects", "custom_drops", "TEXT DEFAULT '[]'");
safeAddColumn("projects", "custom_levels", "TEXT DEFAULT '[]'");
safeAddColumn("projects", "custom_work_types", "TEXT DEFAULT '[]'");
safeAddColumn("defects", "report_id", "INTEGER");
safeAddColumn("defects", "updated_at", "TEXT");
safeAddColumn("defects", "created_at", "TEXT");
safeAddColumn("markers", "location_id", "INTEGER");
safeAddColumn("projects", "enabled_uid_parts", `TEXT DEFAULT '{"elevation":true,"drop":true,"level":true,"workType":true}'`);
safeAddColumn("projects", "primary_work_types", "TEXT DEFAULT '[]'");
safeAddColumn("photos", "report_id", "INTEGER");
safeAddColumn("photos", "origin_report_id", "INTEGER"); // report where photo FIRST appeared (survives cloning)
safeAddColumn("photos", "new_override", "TEXT"); // "new" | "not-new" | null (auto-detect)
safeAddColumn("defect_locations", "updated_at", "TEXT");
// Structured UID parts on defects — source of truth for the form (no re-parsing of the uid string)
safeAddColumn("defects", "elevation_code", "TEXT");
safeAddColumn("defects", "drop_code", "TEXT");
safeAddColumn("defects", "level_code", "TEXT");
safeAddColumn("defects", "work_type_code", "TEXT");
safeAddColumn("defects", "seq_number", "TEXT");
// SVR template reformat — Stage A schema additions.
// legacy_id stays NULL in Stage 1 (Stage 2 / apply populates it). DO NOT backfill here.
safeAddColumn("defects", "legacy_id", "TEXT");
safeAddColumn("defects", "location_structured", "TEXT"); // JSON {elevation,drop,level} etc — source of truth for location string
safeAddColumn("defects", "inspection_opened", "INTEGER"); // inspection_number where this record first appeared
// date_opened / date_closed / verification_method / verification_person already exist on defects
// (see CREATE TABLE above). Mapping noted; not re-added to avoid touching live data.
// Project location dimensions — drives the project-general formatLocation() helper.
safeAddColumn("projects", "location_dimensions", `TEXT DEFAULT '["elevation","drop","level"]'`);

// Meta table for one-time migration flags
sqlite.exec(`CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value TEXT)`);

// Photo backfill — runs ONCE, then records completion in meta so it never re-corrupts.
// Original logic (e0ff323) ran on every startup for orphan photos. Now gated.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'photo_backfill_v1'`).get();
  if (!alreadyRan) {
    const orphanPhotos = sqlite.prepare(
      `SELECT p.id, p.defect_id, p.created_at, d.project_id
       FROM photos p JOIN defects d ON d.id = p.defect_id
       WHERE p.report_id IS NULL`
    ).all() as { id: number; defect_id: number; created_at: string; project_id: number }[];

    if (orphanPhotos.length > 0) {
      const projectIds = [...new Set(orphanPhotos.map(p => p.project_id))];
      const projectReports: Record<number, { id: number; created_at: string }[]> = {};
      for (const pid of projectIds) {
        projectReports[pid] = sqlite.prepare(
          `SELECT id, created_at FROM reports WHERE project_id = ? ORDER BY created_at ASC`
        ).all(pid) as { id: number; created_at: string }[];
      }

      const updateStmt = sqlite.prepare(`UPDATE photos SET report_id = ? WHERE id = ?`);
      for (const photo of orphanPhotos) {
        const reports = projectReports[photo.project_id] || [];
        if (reports.length === 0) continue;
        const photoTs = new Date(photo.created_at).getTime();
        let matched: number | null = null;
        for (let i = 0; i < reports.length; i++) {
          const rStart = new Date(reports[i].created_at).getTime();
          const rEnd = i + 1 < reports.length ? new Date(reports[i + 1].created_at).getTime() : Infinity;
          if (photoTs >= rStart && photoTs < rEnd) {
            matched = reports[i].id;
            break;
          }
        }
        if (matched === null) matched = reports[0].id;
        updateStmt.run(matched, photo.id);
      }
    }
    // Mark as completed so it never re-runs
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('photo_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

// Origin-report backfill — runs ONCE, sets origin_report_id via UID+slot matching.
// For each photo, finds the EARLIEST photo with the same (defect_uid, slot) in the same
// project (ordered by report inspection_number). That earliest photo's report_id becomes
// the origin_report_id for all later copies. Photos with no earlier match get their own report_id.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'photo_origin_backfill_v1'`).get();
  if (!alreadyRan) {
    // Gather all photos with their defect UID and project, ordered by inspection_number ASC
    const allPhotos = sqlite.prepare(`
      SELECT p.id, p.defect_id, p.report_id, p.slot, d.uid as defect_uid, d.project_id,
             r.inspection_number
      FROM photos p
      JOIN defects d ON d.id = p.defect_id
      LEFT JOIN reports r ON r.id = p.report_id
      ORDER BY CAST(r.inspection_number AS INTEGER) ASC, p.id ASC
    `).all() as {
      id: number; defect_id: number; report_id: number | null; slot: string;
      defect_uid: string; project_id: number; inspection_number: string | null;
    }[];

    // Map: project_id → defect_uid → slot → earliest report_id
    const originMap = new Map<string, number>(); // key = "projectId:uid:slot"

    const updateStmt = sqlite.prepare(`UPDATE photos SET origin_report_id = ? WHERE id = ?`);
    let backfilled = 0;

    for (const photo of allPhotos) {
      const key = `${photo.project_id}:${photo.defect_uid}:${photo.slot}`;
      const existingOrigin = originMap.get(key);
      if (existingOrigin !== undefined) {
        // This is a carry-over copy — use the earliest origin
        updateStmt.run(existingOrigin, photo.id);
      } else {
        // First occurrence of this (uid, slot) in this project — origin is its own report_id
        const origin = photo.report_id;
        if (origin !== null) {
          originMap.set(key, origin);
        }
        updateStmt.run(origin, photo.id);
      }
      backfilled++;
    }

    console.log(`[origin-backfill] Set origin_report_id for ${backfilled} photos`);
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('photo_origin_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

// Seed global_settings with default work types if empty
{
  const existing = sqlite.prepare("SELECT id FROM global_settings LIMIT 1").get();
  if (!existing) {
    const defaultWorkTypes = [
      {code:"CR",label:"Concrete Repair"},{code:"CK",label:"Caulking"},{code:"PT",label:"Painting"},
      {code:"WP",label:"Waterproofing"},{code:"GL",label:"Glazing"},{code:"CL",label:"Cleaning"},
      {code:"ST",label:"Steelwork"},{code:"SE",label:"Sealant"},{code:"FL",label:"Flashing"},
      {code:"RR",label:"Render Repair"},{code:"GK",label:"Gasket"},{code:"SR",label:"Steel Repair"},
      {code:"SM",label:"Steel Removal"},{code:"BR",label:"Brick Repair"},{code:"LR",label:"Lintel Repair"},
      {code:"SS",label:"Sill Stabilisation"},{code:"GW",label:"General Works"},{code:"OT",label:"Other"},
    ];
    sqlite.prepare("INSERT INTO global_settings (work_types) VALUES (?)").run(JSON.stringify(defaultWorkTypes));
  }
}

// Built-in work type codes that are always valid (mirrors client WORK_TYPES + global defaults).
export const BUILTIN_WORK_TYPE_CODES = [
  "CR","CK","PT","WP","GL","CL","ST","SE","FL","RR","GK","SR","SM","BR","LR","SS","GW","OT",
];

// Codes that must NEVER be a custom work type (collide with the level field / are too ambiguous).
const FORBIDDEN_WORK_TYPE_CODES = new Set(["LEV", "LEVEL", "L"]);

// Decide whether a custom work type code should be stripped.
function isBadWorkTypeCode(code: string): boolean {
  if (!code) return true;
  const c = code.trim().toUpperCase();
  if (FORBIDDEN_WORK_TYPE_CODES.has(c)) return true;
  if (c.length < 2) return true; // single letters collide with elevation codes (N/S/E/W) and levels
  if (BUILTIN_WORK_TYPE_CODES.includes(c)) return true; // duplicates a built-in
  return false;
}

// Parse an assembled UID into structured parts using known elevation + work-type code sets.
// Format: [Elevation]-[Drop]-[Level]-[WorkType]-[Number]; any segment may be omitted; Number is last.
export function parseUidParts(
  uid: string,
  knownElevCodes: Set<string>,
  knownWtCodes: Set<string>,
): { elevation: string; drop: string; level: string; workType: string; seq: string } {
  const empty = { elevation: "", drop: "", level: "", workType: "", seq: "" };
  if (!uid) return empty;
  const parts = uid.split("-");
  if (parts.length === 0) return empty;

  const lastIdx = parts.length - 1;
  const seq = parts[lastIdx];

  // Find work type — prefer the second-to-last position, else any matching known code.
  let wtIdx = -1;
  if (lastIdx >= 1 && knownWtCodes.has(parts[lastIdx - 1])) {
    wtIdx = lastIdx - 1;
  } else {
    for (let i = lastIdx - 1; i >= 0; i--) {
      if (knownWtCodes.has(parts[i])) { wtIdx = i; break; }
    }
  }

  if (wtIdx < 0) return { ...empty, seq };

  const workType = parts[wtIdx];
  const before = parts.slice(0, wtIdx);
  let elevation = "";
  for (let i = 0; i < before.length; i++) {
    if (knownElevCodes.has(before[i])) { elevation = before[i]; before.splice(i, 1); break; }
  }
  const drop = before.length >= 1 ? before[0] : "";
  const level = before.length >= 2 ? before[1] : "";
  return { elevation, drop, level, workType, seq };
}

// One-time cleanup of bad custom work types across all projects (gated by meta flag).
// Strips entries like "LEV"/"LEVEL", single letters, or codes that duplicate a built-in.
export function cleanupCustomWorkTypes(): { projectsChanged: number; removed: string[] } {
  const rows = sqlite.prepare(`SELECT id, custom_work_types FROM projects`).all() as
    { id: number; custom_work_types: string | null }[];
  const removed: string[] = [];
  let projectsChanged = 0;
  const updateStmt = sqlite.prepare(`UPDATE projects SET custom_work_types = ? WHERE id = ?`);
  for (const row of rows) {
    let list: { code: string; label: string }[];
    try { list = JSON.parse(row.custom_work_types || "[]"); } catch { continue; }
    if (!Array.isArray(list)) continue;
    const cleaned = list.filter((wt) => {
      const bad = !wt || typeof wt.code !== "string" || isBadWorkTypeCode(wt.code);
      if (bad && wt?.code) removed.push(wt.code);
      return !bad;
    });
    if (cleaned.length !== list.length) {
      updateStmt.run(JSON.stringify(cleaned), row.id);
      projectsChanged++;
    }
  }
  return { projectsChanged, removed };
}

// Gated startup cleanup of bad custom work types (e.g. the stuck "LEV" entry).
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'custom_work_types_cleanup_v1'`).get();
  if (!alreadyRan) {
    const { projectsChanged, removed } = cleanupCustomWorkTypes();
    if (removed.length > 0) {
      console.log(`[work-type-cleanup] Removed ${removed.length} bad custom work types (${removed.join(", ")}) from ${projectsChanged} project(s)`);
    }
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('custom_work_types_cleanup_v1', ?)`).run(new Date().toISOString());
  }
}

// Gated backfill of structured UID parts from existing assembled uid strings.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'uid_parts_backfill_v1'`).get();
  if (!alreadyRan) {
    // Build the global work-type code set once.
    const globalWtCodes = new Set<string>(BUILTIN_WORK_TYPE_CODES);
    try {
      const gs = sqlite.prepare(`SELECT work_types FROM global_settings LIMIT 1`).get() as { work_types: string } | undefined;
      if (gs?.work_types) {
        (JSON.parse(gs.work_types) as { code: string }[]).forEach((wt) => wt?.code && globalWtCodes.add(wt.code));
      }
    } catch {}

    // Per-project elevation/work-type code maps.
    const codeMap: Record<string, string> = {
      "North": "N", "South": "S", "East": "E", "West": "W",
      "North East": "NE", "North West": "NW", "South East": "SE", "South West": "SW",
    };
    const projects = sqlite.prepare(`SELECT id, elevations, custom_work_types FROM projects`).all() as
      { id: number; elevations: string | null; custom_work_types: string | null }[];
    const projElev = new Map<number, Set<string>>();
    const projWt = new Map<number, Set<string>>();
    for (const p of projects) {
      const elevSet = new Set<string>();
      try {
        (JSON.parse(p.elevations || "[]") as string[]).forEach((label) => {
          elevSet.add(codeMap[label] || label.substring(0, 3).toUpperCase());
          elevSet.add(label);
        });
      } catch {}
      // Always accept the standard compass codes too.
      ["N","S","E","W","NE","NW","SE","SW"].forEach((c) => elevSet.add(c));
      projElev.set(p.id, elevSet);

      const wtSet = new Set<string>(globalWtCodes);
      try {
        (JSON.parse(p.custom_work_types || "[]") as { code: string }[]).forEach((wt) => wt?.code && wtSet.add(wt.code));
      } catch {}
      projWt.set(p.id, wtSet);
    }

    const defectRows = sqlite.prepare(`SELECT id, project_id, uid FROM defects`).all() as
      { id: number; project_id: number; uid: string }[];
    const updateStmt = sqlite.prepare(
      `UPDATE defects SET elevation_code = ?, drop_code = ?, level_code = ?, work_type_code = ?, seq_number = ? WHERE id = ?`
    );
    let backfilled = 0;
    for (const d of defectRows) {
      const elevSet = projElev.get(d.project_id) || new Set<string>(["N","S","E","W","NE","NW","SE","SW"]);
      const wtSet = projWt.get(d.project_id) || globalWtCodes;
      const parsed = parseUidParts(d.uid, elevSet, wtSet);
      updateStmt.run(
        parsed.elevation || null,
        parsed.drop || null,
        parsed.level || null,
        parsed.workType || null,
        parsed.seq || null,
        d.id,
      );
      backfilled++;
    }
    console.log(`[uid-parts-backfill] Backfilled structured UID parts for ${backfilled} defects`);
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('uid_parts_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

// ── SVR reformat Stage A backfills (all gated; run ONCE) ──────────────────────
// NOTE: legacy_id is deliberately NOT backfilled here. It stays NULL until Stage 2
// (apply) is approved. Stage 1 is preview-only.

// Helper: read a project's ordered location dimensions (JSON array). Falls back to
// the East Elevation default. Project-general — no hard-coded project IDs.
function readLocationDimensions(rawConfig: string | null | undefined): string[] {
  if (rawConfig) {
    try {
      const parsed = JSON.parse(rawConfig);
      if (Array.isArray(parsed) && parsed.length > 0) return parsed as string[];
    } catch {}
  }
  return ["elevation", "drop", "level"];
}

// Backfill location_structured from the existing structured UID columns, keyed by
// each project's configured location dimensions. Source of truth for the location
// string in BOTH the register and the card.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'svr_location_structured_backfill_v1'`).get();
  if (!alreadyRan) {
    const projDims = new Map<number, string[]>();
    const projRows = sqlite.prepare(`SELECT id, location_dimensions FROM projects`).all() as
      { id: number; location_dimensions: string | null }[];
    for (const p of projRows) projDims.set(p.id, readLocationDimensions(p.location_dimensions));

    const rows = sqlite.prepare(
      `SELECT id, project_id, elevation_code, drop_code, level_code FROM defects WHERE location_structured IS NULL`
    ).all() as { id: number; project_id: number; elevation_code: string | null; drop_code: string | null; level_code: string | null }[];
    const upd = sqlite.prepare(`UPDATE defects SET location_structured = ? WHERE id = ?`);
    let filled = 0;
    for (const d of rows) {
      const dims = projDims.get(d.project_id) || ["elevation", "drop", "level"];
      // Map the project's known structured columns onto its declared dimensions.
      // Available source values from the b166a50 structured columns:
      const source: Record<string, string | null> = {
        elevation: d.elevation_code,
        drop: d.drop_code,
        level: d.level_code,
        // For projects whose first dimension is "stage", reuse the level-like slot.
        stage: d.drop_code ?? d.elevation_code,
      };
      const obj: Record<string, string> = {};
      for (const dim of dims) {
        const v = source[dim];
        if (v != null && String(v).trim() !== "") obj[dim] = String(v);
      }
      upd.run(JSON.stringify(obj), d.id);
      filled++;
    }
    console.log(`[svr-location-structured-backfill] Populated location_structured for ${filled} defects`);
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('svr_location_structured_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

// Backfill date_opened from createdAt for any row where date_opened is null/empty.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'svr_date_opened_backfill_v1'`).get();
  if (!alreadyRan) {
    const res = sqlite.prepare(
      `UPDATE defects SET date_opened = created_at
       WHERE (date_opened IS NULL OR date_opened = '') AND created_at IS NOT NULL`
    ).run();
    console.log(`[svr-date-opened-backfill] Set date_opened from created_at for ${res.changes} defects`);
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('svr_date_opened_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

// Backfill date_closed from status_history (when status last changed to "complete")
// for any complete defect that has no date_closed recorded.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'svr_date_closed_backfill_v1'`).get();
  if (!alreadyRan) {
    const rows = sqlite.prepare(
      `SELECT d.id,
              (SELECT sh.created_at FROM status_history sh
               WHERE sh.defect_id = d.id AND sh.new_status = 'complete'
               ORDER BY sh.created_at DESC LIMIT 1) AS closed_at
       FROM defects d
       WHERE d.status = 'complete' AND (d.date_closed IS NULL OR d.date_closed = '')`
    ).all() as { id: number; closed_at: string | null }[];
    const upd = sqlite.prepare(`UPDATE defects SET date_closed = ? WHERE id = ?`);
    let filled = 0;
    for (const r of rows) {
      if (r.closed_at) { upd.run(r.closed_at, r.id); filled++; }
    }
    console.log(`[svr-date-closed-backfill] Set date_closed from status_history for ${filled} defects`);
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('svr_date_closed_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

// (inspection_opened backfill runs AFTER the orphan-report migration below, since
//  it depends on every defect having a report_id assigned.)

// Migration: for existing defects without reportId, create a default "Report 1" for each project
{
  const orphanRows = sqlite.prepare(
    `SELECT DISTINCT project_id FROM defects WHERE report_id IS NULL`
  ).all() as { project_id: number }[];
  for (const row of orphanRows) {
    // Check if this project already has a report
    const existing = sqlite.prepare(
      `SELECT id FROM reports WHERE project_id = ?`
    ).get(row.project_id) as { id: number } | undefined;
    let reportId: number;
    if (existing) {
      reportId = existing.id;
    } else {
      // Pull per-visit fields from old project row to populate the default report
      const proj = sqlite.prepare(`SELECT * FROM projects WHERE id = ?`).get(row.project_id) as any;
      const result = sqlite.prepare(
        `INSERT INTO reports (project_id, inspection_number, inspection_date, revision, locations_covered, attendees, created_at)
         VALUES (?, ?, ?, ?, ?, ?, ?)`
      ).run(
        row.project_id,
        proj?.inspection_number || "",
        proj?.inspection_date || "",
        proj?.revision || "01",
        proj?.locations_covered || "",
        proj?.attendees || "[]",
        new Date().toISOString(),
      );
      reportId = Number(result.lastInsertRowid);
    }
    sqlite.prepare(
      `UPDATE defects SET report_id = ? WHERE project_id = ? AND report_id IS NULL`
    ).run(reportId, row.project_id);
  }
}

// Backfill inspection_opened: the inspection_number where each UID FIRST appeared.
// Runs after the orphan-report migration so every defect has a report_id. Defects
// are cloned per inspection sharing the same UID, so the earliest report (by
// createdAt) containing a given UID within a project defines the open inspection.
{
  const alreadyRan = sqlite.prepare(`SELECT value FROM meta WHERE key = 'svr_inspection_opened_backfill_v1'`).get();
  if (!alreadyRan) {
    const rows = sqlite.prepare(
      `SELECT d.id, d.project_id, d.uid,
              (SELECT r.inspection_number FROM defects d2
                 JOIN reports r ON r.id = d2.report_id
               WHERE d2.project_id = d.project_id AND d2.uid = d.uid AND r.id IS NOT NULL
               ORDER BY r.created_at ASC LIMIT 1) AS first_insp
       FROM defects d`
    ).all() as { id: number; project_id: number; uid: string; first_insp: string | null }[];
    const upd = sqlite.prepare(`UPDATE defects SET inspection_opened = ? WHERE id = ?`);
    let filled = 0;
    for (const r of rows) {
      // inspection_number is stored as text (e.g. "05"); store its integer form.
      const n = r.first_insp != null ? parseInt(String(r.first_insp), 10) : NaN;
      upd.run(Number.isFinite(n) ? n : null, r.id);
      if (Number.isFinite(n)) filled++;
    }
    console.log(`[svr-inspection-opened-backfill] Set inspection_opened for ${filled} defects`);
    sqlite.prepare(`INSERT OR REPLACE INTO meta (key, value) VALUES ('svr_inspection_opened_backfill_v1', ?)`).run(new Date().toISOString());
  }
}

export const db = drizzle(sqlite);
export { dataDir, sqlite };

const uploadDir = path.join(dataDir, "uploads");

export interface IStorage {
  // Users
  getUser(id: number): Promise<User | undefined>;
  getUserByUsername(username: string): Promise<User | undefined>;
  createUser(user: InsertUser): Promise<User>;
  // Projects
  getProjects(): Promise<Project[]>;
  getProject(id: number): Promise<Project | undefined>;
  createProject(project: InsertProject): Promise<Project>;
  updateProject(id: number, project: Partial<InsertProject>): Promise<Project | undefined>;
  deleteProject(id: number): Promise<void>;
  // Reports
  getReportsByProject(projectId: number): Promise<Report[]>;
  getReport(id: number): Promise<Report | undefined>;
  createReport(report: InsertReport): Promise<Report>;
  updateReport(id: number, report: Partial<InsertReport>): Promise<Report | undefined>;
  deleteReport(id: number): Promise<void>;
  copyReport(reportId: number): Promise<Report>;
  // Defects
  getDefectsByProject(projectId: number): Promise<Defect[]>;
  getDefectsByReport(reportId: number): Promise<Defect[]>;
  getDefect(id: number): Promise<Defect | undefined>;
  createDefect(defect: InsertDefect): Promise<Defect>;
  updateDefect(id: number, defect: Partial<InsertDefect>): Promise<Defect | undefined>;
  deleteDefect(id: number): Promise<void>;
  getNextDefectUid(projectId: number, prefix?: string): Promise<string>;
  getDefectByUid(projectId: number, uid: string): Promise<Defect | undefined>;
  // Photos
  getPhoto(id: number): Promise<Photo | undefined>;
  getPhotosByDefect(defectId: number): Promise<Photo[]>;
  createPhoto(photo: InsertPhoto): Promise<Photo>;
  updatePhotoCaption(id: number, caption: string): Promise<Photo | undefined>;
  updatePhotoReportId(id: number, reportId: number): Promise<Photo | undefined>;
  updatePhotoNewOverride(id: number, newOverride: string | null): Promise<Photo | undefined>;
  deletePhoto(id: number): Promise<Photo | undefined>;
  // Elevations
  getElevationsByProject(projectId: number): Promise<Elevation[]>;
  getElevation(id: number): Promise<Elevation | undefined>;
  createElevation(elevation: InsertElevation): Promise<Elevation>;
  deleteElevation(id: number): Promise<void>;
  // Markers
  getMarkersByElevation(elevationId: number): Promise<Marker[]>;
  getMarker(id: number): Promise<Marker | undefined>;
  createMarker(marker: InsertMarker): Promise<Marker>;
  updateMarker(id: number, marker: Partial<InsertMarker>): Promise<Marker | undefined>;
  deleteMarker(id: number): Promise<void>;
  updateMarkersByDefectId(defectId: number, updates: Partial<InsertMarker>): Promise<void>;
  // Observation/Action History
  getObservationHistory(defectId: number): Promise<ObservationHistory[]>;
  getActionHistory(defectId: number): Promise<ActionHistory[]>;
  createObservationHistory(entry: InsertObservationHistory): Promise<ObservationHistory>;
  createActionHistory(entry: InsertActionHistory): Promise<ActionHistory>;
  // Defect Locations
  getDefectLocations(defectId: number): Promise<DefectLocation[]>;
  getDefectLocation(id: number): Promise<DefectLocation | undefined>;
  createDefectLocation(input: InsertDefectLocation): Promise<DefectLocation>;
  updateDefectLocation(id: number, patch: Partial<InsertDefectLocation>): Promise<DefectLocation | undefined>;
  deleteDefectLocation(id: number): Promise<void>;
  // Status History
  getStatusHistory(defectId: number): Promise<StatusHistory[]>;
  createStatusHistory(entry: InsertStatusHistory): Promise<StatusHistory>;
}

export class DatabaseStorage implements IStorage {
  // Users
  async getUser(id: number): Promise<User | undefined> {
    return db.select().from(users).where(eq(users.id, id)).get();
  }
  async getUserByUsername(username: string): Promise<User | undefined> {
    return db.select().from(users).where(eq(users.username, username)).get();
  }
  async createUser(insertUser: InsertUser): Promise<User> {
    return db.insert(users).values(insertUser).returning().get();
  }

  // Projects
  async getProjects(): Promise<Project[]> {
    return db.select().from(projects).orderBy(desc(projects.id)).all();
  }
  async getProject(id: number): Promise<Project | undefined> {
    return db.select().from(projects).where(eq(projects.id, id)).get();
  }
  async createProject(project: InsertProject): Promise<Project> {
    return db.insert(projects).values(project).returning().get();
  }
  async updateProject(id: number, project: Partial<InsertProject>): Promise<Project | undefined> {
    return db.update(projects).set(project).where(eq(projects.id, id)).returning().get();
  }
  async deleteProject(id: number): Promise<void> {
    // Delete all photos for all defects in this project
    const projectDefects = db.select().from(defects).where(eq(defects.projectId, id)).all();
    for (const d of projectDefects) {
      db.delete(photos).where(eq(photos.defectId, d.id)).run();
    }
    db.delete(defects).where(eq(defects.projectId, id)).run();
    db.delete(reports).where(eq(reports.projectId, id)).run();
    // Delete markers and elevations for this project
    const projectElevations = db.select().from(elevations).where(eq(elevations.projectId, id)).all();
    for (const e of projectElevations) {
      db.delete(markers).where(eq(markers.elevationId, e.id)).run();
    }
    db.delete(elevations).where(eq(elevations.projectId, id)).run();
    db.delete(projects).where(eq(projects.id, id)).run();
  }

  // Reports
  async getReportsByProject(projectId: number): Promise<Report[]> {
    return db.select().from(reports).where(eq(reports.projectId, projectId)).orderBy(desc(reports.id)).all();
  }
  async getReport(id: number): Promise<Report | undefined> {
    return db.select().from(reports).where(eq(reports.id, id)).get();
  }
  async createReport(report: InsertReport): Promise<Report> {
    return db.insert(reports).values(report).returning().get();
  }
  async updateReport(id: number, report: Partial<InsertReport>): Promise<Report | undefined> {
    return db.update(reports).set(report).where(eq(reports.id, id)).returning().get();
  }
  async deleteReport(id: number): Promise<void> {
    // Delete all photos for all defects in this report
    const reportDefects = db.select().from(defects).where(eq(defects.reportId, id)).all();
    for (const d of reportDefects) {
      db.delete(photos).where(eq(photos.defectId, d.id)).run();
    }
    db.delete(defects).where(eq(defects.reportId, id)).run();
    db.delete(reports).where(eq(reports.id, id)).run();
  }
  async copyReport(reportId: number): Promise<Report> {
    const source = db.select().from(reports).where(eq(reports.id, reportId)).get();
    if (!source) throw new Error("Source report not found");

    // Increment inspection number
    const currentNum = parseInt(source.inspectionNumber || "0", 10);
    const newInspectionNumber = String(currentNum + 1).padStart(2, "0");

    const newReport = db.insert(reports).values({
      projectId: source.projectId,
      inspectionNumber: newInspectionNumber,
      inspectionDate: new Date().toISOString().split("T")[0],
      revision: "01",
      locationsCovered: source.locationsCovered,
      elevations: source.elevations,
      attendees: source.attendees,
      createdAt: new Date().toISOString(),
    }).returning().get();

    // Copy all defects from source report
    const sourceDefects = db.select().from(defects).where(eq(defects.reportId, reportId)).all();
    for (const d of sourceDefects) {
      const newDefect = db.insert(defects).values({
        projectId: d.projectId,
        reportId: newReport.id,
        uid: d.uid,
        dateOpened: d.dateOpened,
        dateClosed: d.dateClosed,
        comment: d.comment,
        actionRequired: d.actionRequired,
        assignedTo: d.assignedTo,
        dueDate: d.dueDate,
        verificationMethod: d.verificationMethod,
        verificationPerson: d.verificationPerson,
        status: d.status,
        recordType: d.recordType,
      }).returning().get();

      // Seed history rows for carried-forward text (so user sees "this came from Insp-N")
      const now = new Date().toISOString();
      if (d.comment?.trim()) {
        db.insert(observationHistory).values({
          defectId: newDefect.id,
          reportId: source.id,
          text: d.comment,
          createdAt: now,
        }).run();
      }
      if (d.actionRequired?.trim()) {
        db.insert(actionHistory).values({
          defectId: newDefect.id,
          reportId: source.id,
          text: d.actionRequired,
          createdAt: now,
        }).run();
      }

      // Copy photos (preserve origin_report_id so carry-over photos keep their original source)
      const defectPhotos = db.select().from(photos).where(eq(photos.defectId, d.id)).all();
      for (const p of defectPhotos) {
        // Copy file on disk
        const srcPath = path.join(uploadDir, p.filename);
        if (fs.existsSync(srcPath)) {
          const ext = path.extname(p.filename);
          const newFilename = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}${ext}`;
          const destPath = path.join(uploadDir, newFilename);
          fs.copyFileSync(srcPath, destPath);
          db.insert(photos).values({
            defectId: newDefect.id,
            filename: newFilename,
            caption: p.caption,
            slot: p.slot,
            reportId: newReport.id,
            originReportId: p.originReportId ?? p.reportId, // preserve original source report
            createdAt: new Date().toISOString(),
          }).run();
        }
      }
    }

    return newReport;
  }

  // Defects
  async getDefectsByProject(projectId: number): Promise<Defect[]> {
    return db.select().from(defects).where(eq(defects.projectId, projectId)).all();
  }
  async getDefectsByReport(reportId: number): Promise<Defect[]> {
    return db.select().from(defects).where(eq(defects.reportId, reportId)).all();
  }
  async getDefect(id: number): Promise<Defect | undefined> {
    return db.select().from(defects).where(eq(defects.id, id)).get();
  }
  async createDefect(defect: InsertDefect): Promise<Defect> {
    const now = new Date().toISOString();
    return db.insert(defects).values({ ...defect, updatedAt: now, createdAt: defect.createdAt || now }).returning().get();
  }
  async updateDefect(id: number, defect: Partial<InsertDefect>): Promise<Defect | undefined> {
    return db.update(defects).set({ ...defect, updatedAt: new Date().toISOString() }).where(eq(defects.id, id)).returning().get();
  }
  async deleteDefect(id: number): Promise<void> {
    db.delete(photos).where(eq(photos.defectId, id)).run();
    db.delete(defects).where(eq(defects.id, id)).run();
  }
  async getNextDefectUid(projectId: number, prefix?: string): Promise<string> {
    if (!prefix) {
      return "01-01-CR-01";
    }
    // prefix is like "01-13-CR" — count existing UIDs that start with this prefix
    const existing = db.select().from(defects).where(eq(defects.projectId, projectId)).all();
    const matching = existing.filter((d) => d.uid.startsWith(prefix + "-"));
    const num = matching.length + 1;
    return `${prefix}-${String(num).padStart(2, "0")}`;
  }

  async getDefectByUid(projectId: number, uid: string): Promise<Defect | undefined> {
    // UIDs may appear in multiple reports (copies); return the most recent (highest reportId)
    const results = db
      .select()
      .from(defects)
      .where(and(eq(defects.projectId, projectId), eq(defects.uid, uid)))
      .orderBy(desc(defects.reportId))
      .all();
    return results[0];
  }

  // Photos
  async getPhoto(id: number): Promise<Photo | undefined> {
    return db.select().from(photos).where(eq(photos.id, id)).get();
  }
  async getPhotosByDefect(defectId: number): Promise<Photo[]> {
    return db.select().from(photos).where(eq(photos.defectId, defectId)).all();
  }
  async createPhoto(photo: InsertPhoto): Promise<Photo> {
    return db.insert(photos).values(photo).returning().get();
  }
  async updatePhotoCaption(id: number, caption: string): Promise<Photo | undefined> {
    return db.update(photos).set({ caption }).where(eq(photos.id, id)).returning().get();
  }
  async updatePhotoReportId(id: number, reportId: number): Promise<Photo | undefined> {
    return db.update(photos).set({ reportId }).where(eq(photos.id, id)).returning().get();
  }
  async updatePhotoNewOverride(id: number, newOverride: string | null): Promise<Photo | undefined> {
    return db.update(photos).set({ newOverride }).where(eq(photos.id, id)).returning().get();
  }
  async deletePhoto(id: number): Promise<Photo | undefined> {
    const photo = db.select().from(photos).where(eq(photos.id, id)).get();
    if (photo) {
      db.delete(photos).where(eq(photos.id, id)).run();
    }
    return photo;
  }

  // Elevations
  async getElevationsByProject(projectId: number): Promise<Elevation[]> {
    return db.select().from(elevations).where(eq(elevations.projectId, projectId)).orderBy(desc(elevations.id)).all();
  }
  async getElevation(id: number): Promise<Elevation | undefined> {
    return db.select().from(elevations).where(eq(elevations.id, id)).get();
  }
  async createElevation(elevation: InsertElevation): Promise<Elevation> {
    return db.insert(elevations).values(elevation).returning().get();
  }
  async deleteElevation(id: number): Promise<void> {
    db.delete(markers).where(eq(markers.elevationId, id)).run();
    db.delete(elevations).where(eq(elevations.id, id)).run();
  }

  // Markers
  async getMarkersByElevation(elevationId: number): Promise<Marker[]> {
    return db.select().from(markers).where(eq(markers.elevationId, elevationId)).all();
  }
  async getMarker(id: number): Promise<Marker | undefined> {
    return db.select().from(markers).where(eq(markers.id, id)).get();
  }
  async createMarker(marker: InsertMarker): Promise<Marker> {
    return db.insert(markers).values(marker).returning().get();
  }
  async updateMarker(id: number, marker: Partial<InsertMarker>): Promise<Marker | undefined> {
    return db.update(markers).set(marker).where(eq(markers.id, id)).returning().get();
  }
  async deleteMarker(id: number): Promise<void> {
    db.delete(markers).where(eq(markers.id, id)).run();
  }
  async updateMarkersByDefectId(defectId: number, updates: Partial<InsertMarker>): Promise<void> {
    db.update(markers).set(updates).where(eq(markers.defectId, defectId)).run();
  }

  // Observation/Action History
  async getObservationHistory(defectId: number): Promise<ObservationHistory[]> {
    return db.select().from(observationHistory).where(eq(observationHistory.defectId, defectId)).orderBy(desc(observationHistory.id)).all();
  }
  async getActionHistory(defectId: number): Promise<ActionHistory[]> {
    return db.select().from(actionHistory).where(eq(actionHistory.defectId, defectId)).orderBy(desc(actionHistory.id)).all();
  }
  async createObservationHistory(entry: InsertObservationHistory): Promise<ObservationHistory> {
    return db.insert(observationHistory).values(entry).returning().get();
  }
  async createActionHistory(entry: InsertActionHistory): Promise<ActionHistory> {
    return db.insert(actionHistory).values(entry).returning().get();
  }

  // Defect Locations
  async getDefectLocations(defectId: number): Promise<DefectLocation[]> {
    return db.select().from(defectLocations).where(eq(defectLocations.defectId, defectId)).orderBy(defectLocations.displayOrder).all();
  }
  async getDefectLocation(id: number): Promise<DefectLocation | undefined> {
    return db.select().from(defectLocations).where(eq(defectLocations.id, id)).get();
  }
  async createDefectLocation(input: InsertDefectLocation): Promise<DefectLocation> {
    return db.insert(defectLocations).values(input).returning().get();
  }
  async updateDefectLocation(id: number, patch: Partial<InsertDefectLocation>): Promise<DefectLocation | undefined> {
    return db.update(defectLocations).set({ ...patch, updatedAt: new Date().toISOString() }).where(eq(defectLocations.id, id)).returning().get();
  }
  async deleteDefectLocation(id: number): Promise<void> {
    // Unlink any markers referencing this location
    sqlite.prepare(`UPDATE markers SET location_id = NULL WHERE location_id = ?`).run(id);
    db.delete(defectLocations).where(eq(defectLocations.id, id)).run();
  }

  // Status History
  async getStatusHistory(defectId: number): Promise<StatusHistory[]> {
    return db.select().from(statusHistory).where(eq(statusHistory.defectId, defectId)).orderBy(desc(statusHistory.id)).all();
  }
  async createStatusHistory(entry: InsertStatusHistory): Promise<StatusHistory> {
    return db.insert(statusHistory).values(entry).returning().get();
  }

  // Share Links
  async createShareLink(data: InsertShareLink): Promise<ShareLink> {
    return db.insert(shareLinks).values(data).returning().get();
  }
  async getShareLinkByToken(token: string): Promise<ShareLink | undefined> {
    return db.select().from(shareLinks).where(eq(shareLinks.token, token)).get();
  }
  async getShareLinksByReport(reportId: number): Promise<ShareLink[]> {
    return db.select().from(shareLinks).where(eq(shareLinks.reportId, reportId)).all();
  }
  async deleteShareLink(id: number): Promise<void> {
    db.delete(shareLinks).where(eq(shareLinks.id, id)).run();
  }
}

// Global settings helpers (raw SQL — single-row table)
export function getGlobalSettings(): { workTypes: { code: string; label: string }[] } {
  const row = sqlite.prepare("SELECT work_types FROM global_settings LIMIT 1").get() as any;
  return { workTypes: row ? JSON.parse(row.work_types) : [] };
}

export function addGlobalWorkType(wt: { code: string; label: string }): void {
  const current = getGlobalSettings().workTypes;
  if (!current.find((w: any) => w.code === wt.code)) {
    current.push(wt);
    sqlite.prepare("UPDATE global_settings SET work_types = ? WHERE id = (SELECT id FROM global_settings LIMIT 1)").run(JSON.stringify(current));
  }
}

export const storage = new DatabaseStorage();
