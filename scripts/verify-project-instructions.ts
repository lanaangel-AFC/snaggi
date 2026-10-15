// Harness (WP6): Project Instructions. New projects only; numbered PI-01… per project and
// never reused; need work type(s), action and responsibility; edits are logged; withdrawn
// ones stay listed; exports get them as they stood for each report, newest first.
import express from "express";
import { createServer } from "http";
import { storage, sqlite, getGlobalSettings } from "../server/storage";
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

  const P = (await call("POST", "/api/projects", { name: "New", address: "a", client: "c", inspector: "i", createdAt: now })).body;
  const O = (await call("POST", "/api/projects", { name: "Old", address: "a", client: "c", inspector: "i", createdAt: now, featureVersion: 1 })).body;
  const r1 = (await call("POST", `/api/projects/${P.id}/reports`, { inspectionNumber: "01", inspectionDate: "2026-01-10", createdAt: "2026-01-10T00:00:00.000Z" })).body;
  const r2 = (await call("POST", `/api/projects/${P.id}/reports`, { inspectionNumber: "02", inspectionDate: "2026-10-16", createdAt: "2026-02-01T00:00:00.000Z" })).body;
  const r3 = (await call("POST", `/api/projects/${P.id}/reports`, { inspectionNumber: "03", inspectionDate: "2099-01-01", createdAt: "2099-01-01T00:00:00.000Z" })).body;
  const ro = (await call("POST", `/api/projects/${O.id}/reports`, { inspectionNumber: "01", inspectionDate: "2026-10-16", createdAt: now })).body;

  // The server stamps createdAt itself, so set the inspection timeline directly.
  for (const [r, t] of [[r1, "2026-01-10"], [r2, "2026-02-01"], [r3, "2099-01-01"]] as const)
    sqlite.prepare("UPDATE reports SET created_at = ? WHERE id = ?").run(`${t}T00:00:00.000Z`, r.id);
  const base = { reportId: r2.id, workTypes: ["SE"], observation: "Sealant applied without primer", actionRequired: "Prime all sealant joints before application", responsibility: "Contractor", dateIssued: "2026-10-01" };

  // ---------------------------------------------------------------- create + validate
  const a = await call("POST", `/api/projects/${P.id}/instructions`, base);
  check("first instruction is PI-01", a.status === 201 && a.body.ref === "PI-01", JSON.stringify(a.body).slice(0, 120));
  check("work type label resolved", a.body.workTypeLabels?.[0] === "SE Sealant", JSON.stringify(a.body.workTypeLabels));
  const b = await call("POST", `/api/projects/${P.id}/instructions`, { ...base, workTypes: ["SE", "WP", "se"], responsibility: "Waterproofing subcontractor", dateIssued: "2026-10-16" });
  check("second is PI-02 with several work types (deduplicated)", b.body.ref === "PI-02" && JSON.stringify(b.body.workTypes) === '["SE","WP"]', JSON.stringify(b.body.workTypes));
  check("a new party can be the responsibility", b.body.responsibility === "Waterproofing subcontractor");
  check("no work type refused", (await call("POST", `/api/projects/${P.id}/instructions`, { ...base, workTypes: [] })).status === 400);
  check("no action refused", (await call("POST", `/api/projects/${P.id}/instructions`, { ...base, actionRequired: " " })).status === 400);
  check("no responsibility refused", (await call("POST", `/api/projects/${P.id}/instructions`, { ...base, responsibility: "" })).status === 400);
  check("old project refused", (await call("POST", `/api/projects/${O.id}/instructions`, { ...base, reportId: ro.id })).status === 400);
  check("old project lists none", JSON.stringify((await call("GET", `/api/projects/${O.id}/instructions`)).body) === "[]");
  check("inspection from another project refused", (await call("POST", `/api/projects/${P.id}/instructions`, { ...base, reportId: ro.id })).status === 400);

  // ---------------------------------------------------------------- order
  const list = (await call("GET", `/api/projects/${P.id}/instructions`)).body;
  check("listed newest first", list.map((i: any) => i.ref).join(",") === "PI-02,PI-01", list.map((i: any) => i.ref).join(","));

  // ---------------------------------------------------------------- edit
  const e = (await call("PATCH", `/api/instructions/${a.body.id}`, { actionRequired: "Prime and tool all sealant joints", observation: "Primer missing at level 3", ref: "PI-99" })).body;
  check("wording can be changed", e.actionRequired === "Prime and tool all sealant joints" && e.observation === "Primer missing at level 3");
  check("ID never changes", e.ref === "PI-01", e.ref);
  check("changes logged in history", e.history.length === 2 && e.history.some((h: any) => h.field === "actionRequired" && h.oldValue === base.actionRequired), JSON.stringify(e.history));
  check("blank action refused on edit", (await call("PATCH", `/api/instructions/${a.body.id}`, { actionRequired: "" })).status === 400);
  const same = (await call("PATCH", `/api/instructions/${a.body.id}`, { actionRequired: "Prime and tool all sealant joints" })).body;
  check("unchanged save logs nothing", same.history.length === 2, String(same.history.length));

  // ---------------------------------------------------------------- withdraw / reinstate
  const w = (await call("PATCH", `/api/instructions/${a.body.id}`, { status: "withdrawn" })).body;
  check("withdraw sets status and date", w.status === "withdrawn" && !!w.dateWithdrawn);
  check("withdrawn stays listed", (await call("GET", `/api/projects/${P.id}/instructions`)).body.some((i: any) => i.id === a.body.id));
  const rs = (await call("PATCH", `/api/instructions/${a.body.id}`, { status: "active" })).body;
  check("reinstate clears the withdrawal", rs.status === "active" && rs.dateWithdrawn === null);
  check("bad status refused", (await call("PATCH", `/api/instructions/${a.body.id}`, { status: "closed" })).status === 400);
  await call("PATCH", `/api/instructions/${a.body.id}`, { status: "withdrawn" });

  // ---------------------------------------------------------------- numbering
  const c = (await call("POST", `/api/projects/${P.id}/instructions`, base)).body;
  await call("DELETE", `/api/instructions/${c.id}`);
  const d = (await call("POST", `/api/projects/${P.id}/instructions`, base)).body;
  check("deleted number never reused (PI-03 deleted, next PI-04)", c.ref === "PI-03" && d.ref === "PI-04", `${c.ref} ${d.ref}`);
  await call("DELETE", `/api/instructions/${d.id}`);

  // ---------------------------------------------------------------- export data
  const rd = async (id: number) => (await call("GET", `/api/reports/${id}/report-data`)).body.projectInstructions;
  const at1 = await rd(r1.id), at2 = await rd(r2.id), at3 = await rd(r3.id);
  check("older report (before they were issued) has none", at1.length === 0, JSON.stringify(at1));
  check("report lists them newest first", at3.map((i: any) => i.ref).join(",") === "PI-02,PI-01", at3.map((i: any) => i.ref).join(","));
  check("withdrawn one shown as withdrawn", at3.find((i: any) => i.ref === "PI-01")?.status === "withdrawn");
  check("current wording used", at3.find((i: any) => i.ref === "PI-01")?.actionRequired === "Prime and tool all sealant joints");
  sqlite.prepare("UPDATE project_instructions SET date_withdrawn = '2099-06-01T00:00:00.000Z' WHERE id = ?").run(a.body.id);
  const at2b = await rd(r2.id);
  check("withdrawn after a report: that report still shows it ongoing", at2b.find((i: any) => i.ref === "PI-01")?.status === "active", JSON.stringify(at2b.map((i: any) => i.status)));
  check("report data carries work type labels", JSON.stringify(at2[0]?.workTypeLabels) === '["SE Sealant","WP Waterproofing"]', JSON.stringify(at2[0]?.workTypeLabels));
  check("old project report has none", JSON.stringify(await rd(ro.id)) === "[]");

  // ---------------------------------------------------------------- isolation
  const next = await storage.startNextInspection(r3.id, { inspectionNumber: "04", inspectionDate: "2099-02-01" });
  check("starting the next inspection copies nothing", (await storage.getDefectsByReport(next.id)).length === 0);
  check("not on the to-do list", storage.getInspectionTodos(next.id).length === 0);
  check("still two instructions after the next inspection", (await call("GET", `/api/projects/${P.id}/instructions`)).body.length === 2);

  // ---------------------------------------------------------------- library + delete
  check("Project Note retired from the generic library", !getGlobalSettings().categoryLibrary.some((x) => x.code === "PN"), JSON.stringify(getGlobalSettings().categoryLibrary));
  await call("DELETE", `/api/projects/${P.id}`);
  const left = (sqlite.prepare("SELECT count(*) n FROM project_instructions WHERE project_id = ?").get(P.id) as any).n;
  check("deleting the project removes its instructions", left === 0, String(left));

  const failed = checks.filter(([, ok]) => !ok);
  for (const [n, ok, dd] of checks) console.log(`${ok ? "PASS" : "FAIL"}  ${n}${ok ? "" : "  << " + dd}`);
  console.log(`\n${checks.length - failed.length}/${checks.length} passed`);
  sqlite.close();
  process.exit(failed.length ? 1 : 0);
}
main().catch((e) => { console.error(e); process.exit(1); });
