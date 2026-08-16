// Harness: a pin's colour must always be the live record's status, in both directions, and
// editing a superseded copy of a record must never drag the pin onto that copy.
import express from "express";
import { createServer } from "http";
import { storage, sqlite } from "../server/storage";
import { registerRoutes } from "../server/routes";

const now = new Date().toISOString();
const checks: Array<[string, boolean, string]> = [];
const check = (n: string, ok: boolean, d = "") => checks.push([n, ok, d]);

async function main() {
  const project = await storage.createProject({
    name: "MS", inspector: "I", address: "a", client: "c", createdAt: now,
  } as any);
  const pid = project.id;
  const mk = (n: string) =>
    storage.createReport({ projectId: pid, inspectionNumber: n, inspectionDate: "2026-01-01", createdAt: now } as any);
  const r7 = await mk("07");
  const r8 = await mk("08");
  const r10 = await mk("10");

  const base = {
    projectId: pid, dateOpened: now, comment: "c", actionRequired: "a",
    verificationMethod: "v", verificationPerson: "p", createdAt: now,
  };
  // The shape of the user's record: open at 07, signed off at 08, never carried to 10.
  const rowOld = await storage.createDefect({ ...base, reportId: r7.id, uid: "E-04-09-LR-01", status: "open" } as any);
  const rowLive = await storage.createDefect({ ...base, reportId: r8.id, uid: "E-04-09-LR-01", status: "complete" } as any);
  const elev = await storage.createElevation({ projectId: pid, name: "East", filename: "f.png", fileType: "image", createdAt: now } as any);
  const pin = await storage.createMarker({
    elevationId: elev.id, defectId: rowOld.id, defectUid: "E-04-09-LR-01",
    status: "open", xPercent: 5, yPercent: 5, createdAt: now,
  } as any);

  const app = express();
  app.use(express.json());
  const http = createServer(app);
  await registerRoutes(http, app);
  await new Promise<void>((r) => http.listen(0, r));
  const port = (http.address() as any).port;

  const patch = async (id: number, body: any) =>
    (await fetch(`http://127.0.0.1:${port}/api/defects/${id}`, {
      method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
    })).json() as any;
  const resolved = async (reportId?: number) => {
    const q = reportId != null ? `?reportId=${reportId}` : "";
    const all = (await (await fetch(`http://127.0.0.1:${port}/api/elevations/${elev.id}/markers/resolved${q}`)).json()) as any[];
    return all.find((m) => m.id === pin.id);
  };
  const pinRow = () => sqlite.prepare(`SELECT defect_id, status FROM markers WHERE id = ?`).get(pin.id) as any;

  // --- the reported symptom: editing the superseded 07 copy must not touch the pin
  await patch(rowOld.id, { status: "open", comment: "amended on the old copy", contextReportId: r7.id });
  let p = pinRow();
  check("pin not dragged onto the superseded copy", p.defect_id === rowLive.id, `pin points at ${p.defect_id}, live row is ${rowLive.id}`);
  check("pin status not stamped from the superseded copy", p.status === "complete", p.status);
  let m = await resolved(r10.id);
  check("drawing shows the live record's status", m.resolved.status === "complete", m.resolved.status);
  check("drawing names the inspection the status came from", m.resolved.inspectionNumber === "08", String(m.resolved.inspectionNumber));
  check("drawing flags it as not in this inspection", m.resolved.inReport === false, String(m.resolved.inReport));

  // --- direction 1: complete -> pin green, on the live record
  await patch(rowLive.id, { status: "open", contextReportId: r8.id });
  check("reopening the live record turns the pin open", pinRow().status === "open", pinRow().status);
  check("resolved status follows too", (await resolved(r8.id)).resolved.status === "open", "");
  await patch(rowLive.id, { status: "complete", dateClosed: "2026-08-01", contextReportId: r8.id });
  check("completing the live record turns the pin complete", pinRow().status === "complete", pinRow().status);
  check("resolved status follows back", (await resolved(r8.id)).resolved.status === "complete", "");

  // --- direction 2: amending from the current inspection carries forward, and the pin moves
  const res = await patch(rowLive.id, { status: "open", comment: "touch ups outstanding", contextReportId: r10.id });
  check("carried forward into the current inspection", res.carriedForward === true && res.reportId === r10.id, `${res.carriedForward}/${res.reportId}`);
  p = pinRow();
  check("pin now points at the current inspection's row", p.defect_id === res.id, `${p.defect_id} vs ${res.id}`);
  check("pin is open again", p.status === "open", p.status);
  m = await resolved(r10.id);
  check("pin is now in this inspection", m.resolved.inReport === true, String(m.resolved.inReport));
  check("no stale inspection warning once carried forward", m.resolved.inspectionNumber === "10", String(m.resolved.inspectionNumber));
  const oldRow = sqlite.prepare(`SELECT status, comment FROM defects WHERE id = ?`).get(rowLive.id) as any;
  check("the signed-off 08 copy is left as issued", oldRow.status === "complete", oldRow.status);

  // --- a pin with no lineage at all still keeps its own status
  const free = await storage.createMarker({
    elevationId: elev.id, defectId: null, defectUid: "NOT-A-RECORD",
    status: "open", xPercent: 9, yPercent: 9, createdAt: now,
  } as any);
  const all = (await (await fetch(`http://127.0.0.1:${port}/api/elevations/${elev.id}/markers/resolved?reportId=${r10.id}`)).json()) as any[];
  const freeM = all.find((x) => x.id === free.id);
  check("unlinked pin survives resolution", !!freeM && freeM.resolved === null && freeM.status === "open", JSON.stringify(freeM?.resolved));

  let failed = 0;
  for (const [n, ok, d] of checks) {
    console.log(`${ok ? "PASS" : "FAIL"}  ${n}${ok ? "" : "  -- " + d}`);
    if (!ok) failed++;
  }
  console.log(failed === 0 ? `\nALL PASS (${checks.length})` : `\nFAILED(${failed}/${checks.length})`);
  http.close();
  process.exit(failed === 0 ? 0 : 1);
}

main().catch((e) => { console.error(e); process.exit(1); });
