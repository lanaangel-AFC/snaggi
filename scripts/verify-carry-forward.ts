// Harness: amending a record from an earlier inspection must create it in the inspection the
// user is working in, leave the issued report untouched, and never duplicate or move backwards.
import express from "express";
import { createServer } from "http";
import fs from "fs";
import path from "path";
import { storage, sqlite } from "../server/storage";
import { registerRoutes } from "../server/routes";

const now = new Date().toISOString();
const checks: Array<[string, boolean, string]> = [];
const check = (name: string, ok: boolean, detail = "") => checks.push([name, ok, detail]);

async function main() {
  const project = await storage.createProject({
    name: "CF", inspector: "I", address: "1 Test St", client: "C", createdAt: now,
    categories: JSON.stringify([{ code: "RR", label: "Rectify" }]),
  } as any);
  const pid = project.id;

  const mkReport = (n: string) =>
    storage.createReport({ projectId: pid, inspectionNumber: n, inspectionDate: "2026-01-0" + n, createdAt: now } as any);
  const r7 = await mkReport("7");
  const r8 = await mkReport("8");
  const r9 = await mkReport("9");

  const base = {
    projectId: pid, dateOpened: now, comment: "original observation",
    actionRequired: "original action", verificationMethod: "v", verificationPerson: "p",
    createdAt: now,
  };

  // Closed at inspection 7, so Start Next Inspection never carried it into 8 or 9.
  const old = await storage.createDefect({
    ...base, reportId: r7.id, uid: "E-10-09-CR-01", status: "complete", categoryCode: "RR",
  } as any);
  // A real file must exist on disk or the clone correctly refuses to copy it.
  const uploadsDir = path.join(process.env.DATA_DIR || ".", "uploads");
  fs.mkdirSync(uploadsDir, { recursive: true });
  fs.writeFileSync(path.join(uploadsDir, "old.jpg"), "not-really-a-jpeg");
  sqlite.prepare(`INSERT INTO photos (defect_id, filename, slot, report_id, created_at) VALUES (?,?,?,?,?)`)
    .run(old.id, "old.jpg", "defect1", r7.id, now);

  const elev = await storage.createElevation({ projectId: pid, name: "East", filename: "f.png", fileType: "image", createdAt: now } as any);
  const pin = await storage.createMarker({
    elevationId: elev.id, defectId: old.id, defectUid: "E-10-09-CR-01",
    status: "complete", xPercent: 5, yPercent: 5, createdAt: now,
  } as any);

  const app = express();
  app.use(express.json());
  const http = createServer(app);
  await registerRoutes(http, app);
  await new Promise<void>((r) => http.listen(0, r));
  const port = (http.address() as any).port;

  const patch = async (id: number, body: any) => {
    const res = await fetch(`http://127.0.0.1:${port}/api/defects/${id}`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    return res.json() as any;
  };
  const rowsFor = (uid: string) =>
    sqlite.prepare(`SELECT id, report_id, uid, status, comment, action_required, category_code FROM defects WHERE project_id = ? AND uid = ? ORDER BY id`)
      .all(pid, uid) as any[];

  // --- 1. amend from the current inspection -> carried forward
  const res1 = await patch(old.id, {
    comment: "water egress now visible",
    actionRequired: "break out and redo",
    status: "open",
    contextReportId: r9.id,
  });
  check("carriedForward reported", res1.carriedForward === true, JSON.stringify(res1.carriedForward));
  check("amendment landed on a new row", res1.id !== old.id, `got ${res1.id} vs ${old.id}`);
  check("new row belongs to the current inspection", res1.reportId === r9.id, `got ${res1.reportId} want ${r9.id}`);
  check("new row carries the amendment", res1.comment === "water egress now visible", res1.comment);
  check("new row is open", res1.status === "open", res1.status);
  check("category carried onto the new row", res1.categoryCode === "RR", String(res1.categoryCode));

  const source = sqlite.prepare(`SELECT * FROM defects WHERE id = ?`).get(old.id) as any;
  check("issued report left untouched (text)", source.comment === "original observation", source.comment);
  check("issued report left untouched (action)", source.action_required === "original action", source.action_required);
  check("issued report left untouched (status)", source.status === "complete", source.status);

  const clonePhotos = sqlite.prepare(`SELECT COUNT(*) c FROM photos WHERE defect_id = ?`).get(res1.id) as any;
  check("photos cloned onto the new row", clonePhotos.c === 1, `got ${clonePhotos.c}`);

  const pinAfter = sqlite.prepare(`SELECT defect_id, status FROM markers WHERE id = ?`).get(pin.id) as any;
  check("pin repointed to the live row", pinAfter.defect_id === res1.id, `got ${pinAfter.defect_id} want ${res1.id}`);
  check("pin status follows the record", pinAfter.status === "open", pinAfter.status);

  // --- 2. amend again -> must not create a second row
  const res2 = await patch(res1.id, { comment: "second edit", contextReportId: r9.id });
  check("second amendment stays on the same row", res2.id === res1.id, `got ${res2.id}`);
  check("no duplicate row in the current inspection", rowsFor("E-10-09-CR-01").filter((r) => r.report_id === r9.id).length === 1,
    JSON.stringify(rowsFor("E-10-09-CR-01").map((r) => r.report_id)));

  // --- 3. amending the historical row directly from its own inspection stays in place
  const res3 = await patch(old.id, { comment: "typo fix in the issued report", contextReportId: r7.id });
  check("editing within its own inspection stays in place", res3.id === old.id, `got ${res3.id}`);
  check("no carry-forward flag when context matches", !res3.carriedForward, String(res3.carriedForward));

  // --- 4. never move backwards: row in 9, context 8
  const res4 = await patch(res1.id, { comment: "backwards attempt", contextReportId: r8.id });
  check("does not move a record backwards", res4.id === res1.id, `got ${res4.id}`);
  check("no row created in the earlier inspection",
    rowsFor("E-10-09-CR-01").filter((r) => r.report_id === r8.id).length === 0,
    JSON.stringify(rowsFor("E-10-09-CR-01").map((r) => r.report_id)));

  // --- 5. renamed lineage already present in the target must not duplicate
  const oldB = await storage.createDefect({ ...base, reportId: r7.id, uid: "E-06-04-LR-01", status: "complete" } as any);
  await storage.createDefect({ ...base, reportId: r9.id, uid: "E-06-04-STI-01", status: "open", legacyId: "E-06-04-LR-01" } as any);
  const res5 = await patch(oldB.id, { comment: "amend via rename lineage", contextReportId: r9.id });
  check("rename lineage resolves to the existing row", res5.uid === "E-06-04-STI-01", res5.uid);
  check("no duplicate created for renamed lineage",
    (sqlite.prepare(`SELECT COUNT(*) c FROM defects WHERE project_id = ? AND report_id = ? AND (uid = ? OR uid = ?)`)
      .get(pid, r9.id, "E-06-04-LR-01", "E-06-04-STI-01") as any).c === 1,
    "expected exactly 1");

  let failed = 0;
  for (const [name, ok, detail] of checks) {
    console.log(`${ok ? "PASS" : "FAIL"}  ${name}${ok ? "" : "  -- " + detail}`);
    if (!ok) failed++;
  }
  console.log(failed === 0 ? `\nALL PASS (${checks.length})` : `\nFAILED(${failed}/${checks.length})`);
  http.close();
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((e) => { console.error(e); process.exit(1); });
