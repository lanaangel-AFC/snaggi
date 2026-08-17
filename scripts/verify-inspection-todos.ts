// Harness: the inspection to-do list must be a stable snapshot taken when the inspection
// starts, must group items by due date and by whether they were stranded in an earlier
// inspection, and must tick itself off when the linked form is saved.
import express from "express";
import { createServer } from "http";
import { storage, sqlite } from "../server/storage";
import { registerRoutes } from "../server/routes";

const now = new Date().toISOString();
const checks: Array<[string, boolean, string]> = [];
const check = (n: string, ok: boolean, d = "") => checks.push([n, ok, d]);

async function main() {
  const project = await storage.createProject({
    name: "TD", inspector: "I", address: "a", client: "c", createdAt: now,
  } as any);
  const pid = project.id;
  const mk = (n: string, date: string) =>
    storage.createReport({ projectId: pid, inspectionNumber: n, inspectionDate: date, createdAt: now } as any);
  const r7 = await mk("07", "2026-06-17");
  const r8 = await mk("08", "2026-07-27");
  const r11 = await mk("11", "2026-08-12");

  const base = {
    projectId: pid, dateOpened: now, comment: "c", actionRequired: "a",
    verificationMethod: "v", verificationPerson: "p", createdAt: now,
  };

  // --- records in the current inspection (11)
  const overdue = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-01",
    status: "open", dueDate: "2026-07-01", assignedTo: "Contractor",
    comment: "Cracking to soffit", actionRequired: "Grind out and patch" } as any);
  const dueOnDay = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-02",
    status: "open", dueDate: "2026-08-12", assignedTo: "Client" } as any);
  const notYetDue = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-03",
    status: "open", dueDate: "2026-08-28", assignedTo: "Consultant" } as any);
  const noDueDate = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-04",
    status: "open", dueDate: "" } as any);
  const alreadyClosed = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-05",
    status: "complete", dueDate: "2026-07-01" } as any);
  // summary preferred over the long action text
  const withSummary = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-06",
    status: "open", dueDate: "2026-07-01", actionRequired: "A long paragraph of action text.",
    actionSummary: "Reseal the head joint." } as any);
  // "None" is written by the summariser when it declines; must fall back, not display "None"
  const summaryNone = await storage.createDefect({ ...base, reportId: r11.id, uid: "E-01-09-CR-07",
    status: "open", dueDate: "2026-07-01", actionRequired: "Replace the gasket.",
    actionSummary: "None" } as any);

  // --- stranded: open at 07, never carried into 08 or 11
  const stranded = await storage.createDefect({ ...base, reportId: r7.id, uid: "E-02-09-LR-01",
    status: "open", dueDate: "2026-06-01", assignedTo: "Contractor",
    comment: "Flashing lifted", actionRequired: "Refix flashing" } as any);
  // --- NOT stranded: closed in an earlier inspection
  await storage.createDefect({ ...base, reportId: r8.id, uid: "E-02-09-LR-02", status: "complete" } as any);
  // --- NOT stranded: open at 07 but its lineage IS present in 11
  await storage.createDefect({ ...base, reportId: r7.id, uid: "E-02-09-LR-03", status: "open" } as any);
  await storage.createDefect({ ...base, reportId: r11.id, uid: "E-02-09-LR-03", status: "open", dueDate: "2026-09-30" } as any);

  const app = express();
  app.use(express.json());
  const http = createServer(app);
  await registerRoutes(http, app);
  await new Promise<void>((r) => http.listen(0, r));
  const port = (http.address() as any).port;
  const B = `http://127.0.0.1:${port}`;

  const getTodos = async (reportId: number) => (await (await fetch(`${B}/api/reports/${reportId}/todos`)).json()) as any;
  const patch = async (id: number, body: any) =>
    (await fetch(`${B}/api/defects/${id}`, {
      method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
    })).json() as any;
  const post = async (p: string, body: any) =>
    (await fetch(`${B}/api/reports${p}`, {
      method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
    })).json() as any;

  // ---------------------------------------------------------------- generation
  let list = await getTodos(r11.id);
  check("no list before generation", list.total === 0 && list.generated === false, `total=${list.total}`);

  const dry = await post(`/${r11.id}/todos/generate`, {});
  check("dry run is flagged as such", dry.dryRun === true, JSON.stringify(dry).slice(0, 120));
  check("dry run previews what it would create", dry.wouldCreate > 0 && dry.sample.length > 0, `wouldCreate=${dry.wouldCreate}`);
  check("dry run writes nothing", (await getTodos(r11.id)).total === 0, `total=${(await getTodos(r11.id)).total}`);

  const gen = await post(`/${r11.id}/todos/generate`, { dryRun: false });
  check("generation is not skipped on a fresh report", gen.skipped === false || gen.total > 0, JSON.stringify(gen).slice(0, 140));

  list = await getTodos(r11.id);
  const byUid = new Map<string, any>(list.items.map((i: any) => [i.uid, i]));
  const grp = (uid: string) => byUid.get(uid)?.groupKey;

  // ---------------------------------------------------------------- grouping
  check("overdue item is grouped due", grp("E-01-09-CR-01") === "due", String(grp("E-01-09-CR-01")));
  check("due exactly on the inspection date is grouped due", grp("E-01-09-CR-02") === "due", String(grp("E-01-09-CR-02")));
  check("due after the inspection date is grouped open", grp("E-01-09-CR-03") === "open", String(grp("E-01-09-CR-03")));
  check("no due date is grouped open", grp("E-01-09-CR-04") === "open", String(grp("E-01-09-CR-04")));
  check("already-complete record is not listed", !byUid.has("E-01-09-CR-05"), "");
  check("stranded item is listed", grp("E-02-09-LR-01") === "stranded", String(grp("E-02-09-LR-01")));
  check("closed earlier record is not stranded", !byUid.has("E-02-09-LR-02"), "");
  check("lineage present in this inspection is not stranded",
    byUid.get("E-02-09-LR-03")?.groupKey === "open", String(grp("E-02-09-LR-03")));
  check("no duplicate rows per uid", list.items.length === new Set(list.items.map((i: any) => i.uid)).size,
    `${list.items.length} rows`);

  // ---------------------------------------------------------------- payload fields
  const od = byUid.get("E-01-09-CR-01");
  check("carries the observation", od.comment === "Cracking to soffit", od.comment);
  check("carries the action", od.actionText === "Grind out and patch", od.actionText);
  check("carries the party", od.assignedTo === "Contractor", od.assignedTo);
  check("carries the due date", od.dueDate === "2026-07-01", od.dueDate);
  check("summary preferred over long action text",
    byUid.get("E-01-09-CR-06")?.actionText === "Reseal the head joint.", byUid.get("E-01-09-CR-06")?.actionText);
  check('summary of "None" falls back to the action text',
    byUid.get("E-01-09-CR-07")?.actionText === "Replace the gasket.", byUid.get("E-01-09-CR-07")?.actionText);

  // ---------------------------------------------------------------- link targets
  check("in-inspection item links to this inspection", od.linkReportId === r11.id, String(od.linkReportId));
  check("stranded item links to its own inspection",
    byUid.get("E-02-09-LR-01")?.linkReportId === r7.id, String(byUid.get("E-02-09-LR-01")?.linkReportId));
  check("stranded item names its inspection",
    byUid.get("E-02-09-LR-01")?.sourceInspectionNumber === "07", String(byUid.get("E-02-09-LR-01")?.sourceInspectionNumber));

  // ---------------------------------------------------------------- snapshot stability
  const before = list.total;
  await storage.createDefect({ ...base, reportId: r11.id, uid: "E-03-09-CR-99", status: "open", dueDate: "2026-01-01" } as any);
  list = await getTodos(r11.id);
  check("a record added mid-inspection does not appear (snapshot, not live)",
    list.total === before, `${before} -> ${list.total}`);

  const again = await post(`/${r11.id}/todos/generate`, { dryRun: false });
  check("re-generating an existing list is skipped", again.skipped === true, JSON.stringify(again).slice(0, 120));

  // ---------------------------------------------------------------- tick-off on save
  await patch(overdue.id, { comment: "Cracking has been patched", contextReportId: r11.id });
  list = await getTodos(r11.id);
  check("saving the form ticks the to-do off",
    !!byUidOf(list, "E-01-09-CR-01").doneAt, "");
  check('tick reason is "updated" for a content edit',
    byUidOf(list, "E-01-09-CR-01").doneReason === "updated", byUidOf(list, "E-01-09-CR-01").doneReason);
  check("done count is reported", list.done === 1, `done=${list.done}`);
  check("a ticked item stays on the list", list.total === before, `${list.total}`);
  check("other items are untouched", !byUidOf(list, "E-01-09-CR-02").doneAt, "");

  await patch(dueOnDay.id, { status: "complete", dateClosed: "2026-08-12", contextReportId: r11.id });
  list = await getTodos(r11.id);
  check("completing the record ticks the to-do off", !!byUidOf(list, "E-01-09-CR-02").doneAt, "");
  check('tick reason is "completed" when the record closes',
    byUidOf(list, "E-01-09-CR-02").doneReason === "completed", byUidOf(list, "E-01-09-CR-02").doneReason);
  check("live status is surfaced for display",
    byUidOf(list, "E-01-09-CR-02").liveStatus === "complete", String(byUidOf(list, "E-01-09-CR-02").liveStatus));

  // a summary-only PATCH is a background job and must not tick anything
  const doneBefore = list.done;
  await patch(notYetDue.id, { actionSummary: "Regenerated by the summariser", actionSummarySource: "ai" });
  list = await getTodos(r11.id);
  check("summary-only update does not tick an item off",
    list.done === doneBefore && !byUidOf(list, "E-01-09-CR-03").doneAt, `done=${list.done}`);

  // a stranded item carried forward from the drawing must still tick off, even though the
  // save lands on a NEW row in a different report
  const cf = await patch(stranded.id, { comment: "Flashing refixed", contextReportId: r11.id });
  check("stranded item carried forward on save", cf.carriedForward === true || cf.reportId === r11.id,
    `id=${cf.id} report=${cf.reportId}`);
  list = await getTodos(r11.id);
  check("carried-forward save ticks the stranded to-do off",
    !!byUidOf(list, "E-02-09-LR-01").doneAt, "");

  // ---------------------------------------------------------------- new inspection
  const r12 = await storage.startNextInspection(r11.id, { inspectionNumber: "12", inspectionDate: "2026-09-20" } as any);
  const l12 = await getTodos(r12.id);
  check("starting an inspection generates its list automatically", l12.total > 0, `total=${l12.total}`);
  check("the new list starts with nothing ticked", l12.done === 0, `done=${l12.done}`);
  check("the new list is scoped to the new inspection",
    l12.items.every((i: any) => i.reportId === r12.id), "");
  check("items open in the new inspection are grouped, not stranded",
    l12.items.filter((i: any) => i.groupKey === "stranded").length === 0,
    JSON.stringify(l12.items.filter((i: any) => i.groupKey === "stranded").map((i: any) => i.uid)));
  check("previous inspection's list is unaffected", (await getTodos(r11.id)).total === before, "");
  check("mid-inspection record now appears in the next list",
    l12.items.some((i: any) => i.uid === "E-03-09-CR-99"), "");
  const dueIn12 = l12.items.filter((i: any) => i.groupKey === "due").map((i: any) => i.uid);
  check("due grouping recomputed against the NEW inspection date",
    dueIn12.includes("E-01-09-CR-03"), JSON.stringify(dueIn12));

  // ---------------------------------------------------------------- report
  const failed = checks.filter(([, ok]) => !ok);
  for (const [n, ok, d] of checks) console.log(`${ok ? "PASS" : "FAIL"}  ${n}${ok ? "" : "  << " + d}`);
  console.log(`\n${checks.length - failed.length}/${checks.length} passed`);
  sqlite.close();
  process.exit(failed.length ? 1 : 0);
}

function byUidOf(list: any, uid: string) {
  return list.items.find((i: any) => i.uid === uid) ?? {};
}

main().catch((e) => { console.error(e); process.exit(1); });
