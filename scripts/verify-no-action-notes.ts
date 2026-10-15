// Harness (WP5): observations with no action needed become "notes" in new projects only.
// Notes take a Note-{inspection}-{nn} reference, never carry forward, never reach the to-do
// list or the action summaries, and switching No -> Yes gives them a full UID.
import express from "express";
import { createServer } from "http";
import { storage, sqlite } from "../server/storage";
import { registerRoutes } from "../server/routes";

const now = new Date().toISOString();
const checks: Array<[string, boolean, string]> = [];
const check = (n: string, ok: boolean, d = "") => checks.push([n, ok, d]);

async function main() {
  const app = express();
  app.use(express.json());
  const http = createServer(app);
  await registerRoutes(http, app);
  await new Promise<void>((r) => http.listen(0, r));
  const B = `http://127.0.0.1:${(http.address() as any).port}`;
  const call = async (m: string, p: string, body?: any) => {
    const r = await fetch(`${B}${p}`, { method: m, headers: { "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
    return { status: r.status, body: r.status === 204 ? null : await r.json() };
  };

  const newP = (await call("POST", "/api/projects", { name: "New", address: "a", client: "c", inspector: "i", createdAt: now })).body;
  const oldP = (await call("POST", "/api/projects", { name: "Old", address: "a", client: "c", inspector: "i", createdAt: now, featureVersion: 1 })).body;
  const r12 = (await call("POST", `/api/projects/${newP.id}/reports`, { inspectionNumber: "12", inspectionDate: "2026-10-15", createdAt: now })).body;
  const rOld = (await call("POST", `/api/projects/${oldP.id}/reports`, { inspectionNumber: "03", inspectionDate: "2026-10-15", createdAt: now })).body;

  const rec = (projectId: number, reportId: number, extra: any) => call("POST", `/api/projects/${projectId}/defects`, {
    uidPrefix: "E-04-07-GW", uidOverride: extra.uidOverride, reportId, dateOpened: "2026-10-15", comment: "c", actionRequired: "",
    verificationMethod: "", verificationPerson: "", status: "open", elevationCode: "E", dropCode: "04", levelCode: "07",
    workTypeCode: "GW", seqNumber: "01", createdAt: now, ...extra,
  });

  // ---------------------------------------------------------------- creating notes
  const n1 = (await rec(newP.id, r12.id, { recordType: "observation", actionNeeded: false, uidOverride: "E-04-07-GW-01" })).body;
  const n2 = (await rec(newP.id, r12.id, { recordType: "observation", actionNeeded: false, uidOverride: "E-04-07-GW-01" })).body;
  check("first note gets Note-12-01", n1.uid === "Note-12-01", n1.uid);
  check("second note gets Note-12-02 (no clash with the form's UID)", n2.uid === "Note-12-02", n2.uid);
  check("note stored with actionNeeded false", n1.actionNeeded === false, String(n1.actionNeeded));
  const obs = (await rec(newP.id, r12.id, { recordType: "observation", actionNeeded: true, uidOverride: "E-04-07-GW-01", actionRequired: "Seal" })).body;
  check("observation with action keeps its UID", obs.uid === "E-04-07-GW-01", obs.uid);
  const def = (await rec(newP.id, r12.id, { recordType: "defect", actionNeeded: false, uidOverride: "E-04-07-GW-02", actionRequired: "Fix" })).body;
  check("a defect can never be a note", def.uid === "E-04-07-GW-02" && def.actionNeeded === true, `${def.uid} ${def.actionNeeded}`);

  const oldTry = await rec(oldP.id, rOld.id, { recordType: "observation", actionNeeded: false, uidOverride: "E-04-07-GW-01" });
  check("old project: creating a note is refused", oldTry.status === 400, String(oldTry.status));
  const oldObs = (await rec(oldP.id, rOld.id, { recordType: "observation", uidOverride: "E-04-07-GW-01", actionRequired: "x" })).body;
  const oldPatch = await call("PATCH", `/api/defects/${oldObs.id}`, { actionNeeded: false });
  check("old project: switching an observation to No is refused", oldPatch.status === 400, String(oldPatch.status));
  check("old project: record unchanged", (await call("GET", `/api/defects/${oldObs.id}`)).body.uid === "E-04-07-GW-01");

  // ---------------------------------------------------------------- editing
  const e1 = (await call("PATCH", `/api/defects/${n1.id}`, { comment: "edited", uid: "E-04-07-GW-09" })).body;
  check("editing a note keeps its reference", e1.uid === "Note-12-01" && e1.comment === "edited", e1.uid);

  // ---------------------------------------------------------------- to-do + summaries
  const stale = (await call("GET", `/api/reports/${r12.id}/action-summaries/stale`)).body;
  check("notes never listed as needing an action summary", !(stale.stale ?? []).some((s: any) => s.uid.startsWith("Note-")), JSON.stringify(stale.stale?.map((s: any) => s.uid)));

  // ---------------------------------------------------------------- start next inspection
  const r13 = await storage.startNextInspection(r12.id, { inspectionNumber: "13", inspectionDate: "2026-10-29" });
  const in13 = await storage.getDefectsByReport(r13.id);
  check("notes not carried into the next inspection", !in13.some((d) => d.uid.startsWith("Note-")), in13.map((d) => d.uid).join(","));
  check("actionable records still carried", in13.some((d) => d.uid === "E-04-07-GW-01") && in13.some((d) => d.uid === "E-04-07-GW-02"));
  const todos13 = storage.getInspectionTodos(r13.id);
  check("notes not on the to-do list", !todos13.some((t: any) => String(t.uid).startsWith("Note-")), todos13.map((t: any) => t.uid).join(","));
  const fresh13 = storage.previewInspectionTodos(r13.id);
  check("notes not on a rebuilt to-do list either", !fresh13.some((t: any) => String(t.uid).startsWith("Note-")));
  const e2 = await call("PATCH", `/api/defects/${n2.id}`, { comment: "from 13", contextReportId: r13.id });
  check("editing an old note from the next inspection edits it in place", e2.body.id === n2.id && e2.body.reportId === r12.id && !e2.body.carriedForward, JSON.stringify({ id: e2.body.id, cf: e2.body.carriedForward }));
  check("no copy of the note made in 13", !(await storage.getDefectsByReport(r13.id)).some((d) => d.uid.startsWith("Note-")));
  const n3 = (await rec(newP.id, r13.id, { recordType: "observation", actionNeeded: false })).body;
  check("numbering restarts per inspection: Note-13-01", n3.uid === "Note-13-01", n3.uid);

  // ---------------------------------------------------------------- No -> Yes
  const y1 = (await call("PATCH", `/api/defects/${n1.id}`, { actionNeeded: true, actionRequired: "Now needs sealing", uid: "Note-12-01" })).body;
  check("No -> Yes assigns a full UID", /^E-04-07-GW-\d\d$/.test(y1.uid) && y1.actionNeeded === true, y1.uid);
  check("the new UID is unique in the project", y1.uid !== "E-04-07-GW-01" && y1.uid !== "E-04-07-GW-02", y1.uid);
  const y2 = (await call("PATCH", `/api/defects/${n3.id}`, { actionNeeded: true, uid: "E-04-07-GW-07" })).body;
  check("No -> Yes uses the form's UID when free", y2.uid === "E-04-07-GW-07", y2.uid);
  const d2 = (await call("PATCH", `/api/defects/${n2.id}`, { recordType: "defect" })).body;
  check("note switched to defect gets a UID", !d2.uid.startsWith("Note-") && d2.actionNeeded === true, d2.uid);

  // ---------------------------------------------------------------- Yes -> No on a carried record
  const carried = in13.find((d) => d.uid === "E-04-07-GW-01")!;
  const refuse = await call("PATCH", `/api/defects/${carried.id}`, { actionNeeded: false });
  check("record with earlier history can't become a note", refuse.status === 400, String(refuse.status));
  const fresh = (await rec(newP.id, r13.id, { recordType: "observation", uidOverride: "E-04-07-GW-20", actionRequired: "x", categoryCode: "RR" })).body;
  const toNote = (await call("PATCH", `/api/defects/${fresh.id}`, { actionNeeded: false, uid: "E-04-07-GW-20" })).body;
  check("new observation Yes -> No becomes Note-13-02 (numbers never reused)", toNote.uid === "Note-13-02" && toNote.categoryCode === null, `${toNote.uid} ${toNote.categoryCode}`);

  // ---------------------------------------------------------------- pins
  const elevId = sqlite.prepare("INSERT INTO elevations (project_id, name, filename, file_type, created_at) VALUES (?, 'N', 'x.png', 'image', ?)").run(newP.id, now).lastInsertRowid as number;
  await call("POST", `/api/elevations/${elevId}/markers`, { defectId: toNote.id, defectUid: toNote.uid, status: "open", xPercent: 10, yPercent: 10 });
  await call("POST", `/api/elevations/${elevId}/markers`, { defectId: fresh.id === toNote.id ? y2.id : y2.id, defectUid: y2.uid, status: "open", xPercent: 20, yPercent: 20 });
  const pins = (await call("GET", `/api/elevations/${elevId}/markers/resolved?reportId=${r13.id}`)).body;
  const notePin = pins.find((m: any) => m.defectId === toNote.id);
  const otherPin = pins.find((m: any) => m.defectId === y2.id);
  check("note pin flagged as a note (grey)", notePin?.resolved?.isNote === true, JSON.stringify(notePin?.resolved));
  check("ordinary pin not flagged", otherPin?.resolved?.isNote === false);

  const failed = checks.filter(([, ok]) => !ok);
  for (const [n, ok, d] of checks) console.log(`${ok ? "PASS" : "FAIL"}  ${n}${ok ? "" : "  << " + d}`);
  console.log(`\n${checks.length - failed.length}/${checks.length} passed`);
  sqlite.close();
  process.exit(failed.length ? 1 : 0);
}
main().catch((e) => { console.error(e); process.exit(1); });
