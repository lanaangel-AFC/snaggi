// Regression harness: the clone body was extracted out of startNextInspection into a shared
// helper. This proves the bulk path still behaves identically.
import fs from "fs";
import path from "path";
import { storage, sqlite } from "../server/storage";

const now = new Date().toISOString();
const checks: Array<[string, boolean, string]> = [];
const check = (n: string, ok: boolean, d = "") => checks.push([n, ok, d]);

async function main() {
  const project = await storage.createProject({
    name: "SNI", inspector: "I", address: "1 Test St", client: "C", createdAt: now,
    categories: JSON.stringify([{ code: "WIP", label: "Work in progress" }]),
  } as any);
  const pid = project.id;
  const r1 = await storage.createReport({ projectId: pid, inspectionNumber: "01", inspectionDate: "2026-01-01", createdAt: now } as any);

  const base = {
    projectId: pid, reportId: r1.id, dateOpened: now, comment: "obs", actionRequired: "act",
    verificationMethod: "v", verificationPerson: "p", createdAt: now,
  };
  const openA = await storage.createDefect({
    ...base, uid: "E-01-01-RR-01", status: "open", categoryCode: "WIP", audience: "builder",
    legacyId: "E-01-01-OLD-01", elevationCode: "E", dropCode: "01", levelCode: "01",
    workTypeCode: "RR", seqNumber: 1, inspectionOpened: "01",
  } as any);
  await storage.createDefect({ ...base, uid: "E-01-02-CR-01", status: "open" } as any);
  const closed = await storage.createDefect({ ...base, uid: "E-01-03-CR-01", status: "complete" } as any);

  const uploadsDir = path.join(process.env.DATA_DIR || ".", "uploads");
  fs.mkdirSync(uploadsDir, { recursive: true });
  fs.writeFileSync(path.join(uploadsDir, "a.jpg"), "x");
  sqlite.prepare(`INSERT INTO photos (defect_id, filename, slot, report_id, created_at) VALUES (?,?,?,?,?)`)
    .run(openA.id, "a.jpg", "defect1", r1.id, now);

  const elev = await storage.createElevation({ projectId: pid, name: "East", filename: "f.png", fileType: "image", createdAt: now } as any);
  const pin = await storage.createMarker({
    elevationId: elev.id, defectId: openA.id, defectUid: "E-01-01-RR-01",
    status: "open", xPercent: 5, yPercent: 5, createdAt: now,
  } as any);
  const pinClosed = await storage.createMarker({
    elevationId: elev.id, defectId: closed.id, defectUid: "E-01-03-CR-01",
    status: "complete", xPercent: 8, yPercent: 8, createdAt: now,
  } as any);

  const r2 = await storage.startNextInspection(r1.id, { inspectionNumber: "02", inspectionDate: "2026-02-01" } as any);

  const cloned = sqlite.prepare(`SELECT * FROM defects WHERE report_id = ? ORDER BY uid`).all(r2.id) as any[];
  check("only open records cloned", cloned.length === 2, `got ${cloned.length}`);
  check("closed record not cloned", !cloned.some((d) => d.uid === "E-01-03-CR-01"), "");

  const a = cloned.find((d) => d.uid === "E-01-01-RR-01")!;
  check("categoryCode carried", a.category_code === "WIP", String(a.category_code));
  check("audience carried", a.audience === "builder", String(a.audience));
  check("legacyId carried", a.legacy_id === "E-01-01-OLD-01", String(a.legacy_id));
  check("inspectionOpened preserved", Number(a.inspection_opened) === 1, String(a.inspection_opened));
  check("structured identity carried", a.work_type_code === "RR" && Number(a.seq_number) === 1, `${a.work_type_code}/${a.seq_number}`);
  check("clone is open", a.status === "open", a.status);

  const obs = sqlite.prepare(`SELECT * FROM observation_history WHERE defect_id = ?`).all(a.id) as any[];
  const act = sqlite.prepare(`SELECT * FROM action_history WHERE defect_id = ?`).all(a.id) as any[];
  check("prior observation history seeded against the source report", obs.length === 1 && obs[0].report_id === r1.id,
    JSON.stringify(obs.map((o) => o.report_id)));
  check("prior action history seeded against the source report", act.length === 1 && act[0].report_id === r1.id,
    JSON.stringify(act.map((o) => o.report_id)));

  const ph = sqlite.prepare(`SELECT * FROM photos WHERE defect_id = ?`).all(a.id) as any[];
  check("photo cloned", ph.length === 1, `got ${ph.length}`);
  check("photo reportId is the new report", ph[0]?.report_id === r2.id, String(ph[0]?.report_id));
  check("photo originReportId preserved", ph[0]?.origin_report_id === r1.id, String(ph[0]?.origin_report_id));
  check("photo file actually copied", !!ph[0] && fs.existsSync(path.join(uploadsDir, ph[0].filename)), String(ph[0]?.filename));

  const meta = sqlite.prepare(`SELECT value FROM meta WHERE key = ?`).get(`start_next_inspection_for_report_${r1.id}`) as any;
  check("completion meta key written", !!meta, "missing");

  // A pin must follow the record into the new inspection without a manual resync.
  const pinAfter = sqlite.prepare(`SELECT defect_id, status FROM markers WHERE id = ?`).get(pin.id) as any;
  check("pin relinked to the new inspection's copy", pinAfter.defect_id === a.id, `pin points at ${pinAfter.defect_id}, new copy is ${a.id}`);
  check("pin status matches the new copy", pinAfter.status === "open", pinAfter.status);
  const closedPin = sqlite.prepare(`SELECT defect_id, status FROM markers WHERE id = ?`).get(pinClosed.id) as any;
  check("pin for a closed record stays on the record that closed it", closedPin.defect_id === closed.id, String(closedPin.defect_id));
  check("pin for a closed record stays complete", closedPin.status === "complete", closedPin.status);

  let failed = 0;
  for (const [n, ok, d] of checks) {
    console.log(`${ok ? "PASS" : "FAIL"}  ${n}${ok ? "" : "  -- " + d}`);
    if (!ok) failed++;
  }
  console.log(failed === 0 ? `\nALL PASS (${checks.length})` : `\nFAILED(${failed}/${checks.length})`);
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((e) => { console.error(e); process.exit(1); });
