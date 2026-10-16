// Harness (WP8): the new report layout for new projects. Checks how records are sorted into
// sections 2.2 / 4 / 5 / 6, the Action List tick-box, the server's carried-in flag, the
// numbering (Appendix A last, unnumbered) and that old projects keep the previous layout.
import express from "express";
import { createServer } from "http";
import { registerRoutes } from "../server/routes";
import { buildReportTree } from "../client/src/lib/report-tree";
import { isInActionList } from "../client/src/lib/action-list";
import { buildOutlineV2 } from "../client/src/lib/report-outline-v2";

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
    const text = await r.text();
    return { status: r.status, body: text ? JSON.parse(text) : null };
  };

  // ------------------------------------------------------------- Action List defaults
  check("RR defaults into the Action List", isInActionList({ code: "RR" }));
  check("WIP defaults out", !isInActionList({ code: "WIP" }));
  check("Provide Information (PI) defaults in", isInActionList({ code: "PI", label: "Provide Information" }));
  check("a category labelled Project Instruction defaults out", !isInActionList({ code: "PI", label: "Project Instruction" }) && !isInActionList({ code: "PJ", label: "Project instructions" }));
  check("PN defaults out", !isInActionList({ code: "PN" }));
  check("explicit tick wins over the default", isInActionList({ code: "WIP", inActionList: true }));
  check("explicit untick wins over the default", !isInActionList({ code: "RR", inActionList: false }));

  const setup = async (fv?: number) => {
    const p = (await call("POST", "/api/projects", { name: fv ? "Old" : "New", address: "1 Layout St", client: "C", inspector: "I", createdAt: now, ...(fv ? { featureVersion: fv } : {}) })).body;
    const put = await call("PUT", `/api/projects/${p.id}/categories`, [
      { code: "RR", label: "Rectify", isDefault: true },
      { code: "RD", label: "Redesign" },
      { code: "RFI", label: "Request Information", inActionList: true },
      { code: "WIP", label: "Work in progress" },
      { code: "MS", label: "Monitor", inActionList: false },
      { code: "UN", label: "Unused category" },
      { code: "PI", label: "Project Instruction" },
    ]);
    const r1 = (await call("POST", `/api/projects/${p.id}/reports`, { inspectionNumber: "01", inspectionDate: "2026-10-08", createdAt: now })).body.id;
    const rec = (rid: number, uid: string, extra: any = {}) => {
      const [e, dr, lv, wt, sq] = uid.split("-");
      return call("POST", `/api/projects/${p.id}/defects`, {
        reportId: rid, dateOpened: "2026-10-15", comment: "c", actionRequired: "a", verificationMethod: "", verificationPerson: "",
        status: "open", createdAt: now, elevationCode: e, dropCode: dr, levelCode: lv, workTypeCode: wt, seqNumber: sq,
        uidPrefix: `${e}-${dr}-${lv}-${wt}`, uidOverride: uid, recordType: "defect", ...extra,
      });
    };
    await rec(r1, "N-01-05-CR-01", { categoryCode: "RR" });
    await rec(r1, "N-01-04-CR-02", { categoryCode: "RR" });
    await rec(r1, "S-02-05-GL-01", { categoryCode: "RD" });
    await rec(r1, "E-01-01-WP-01", { categoryCode: "WIP" });
    await rec(r1, "N-01-02-CR-03", { categoryCode: "RR" });
    await rec(r1, "N-01-03-SE-01", { categoryCode: "RFI", recordType: "observation" });
    const r2 = (await call("POST", `/api/reports/${r1}/start-next-inspection`, { inspectionNumber: "02", inspectionDate: "2026-10-15" })).body.id;
    const rows = Object.fromEntries(((await call("GET", `/api/reports/${r2}/defects`)).body as any[]).map((d) => [d.uid, d]));
    await call("PATCH", `/api/defects/${rows["N-01-02-CR-03"].id}`, { status: "complete", dateClosed: "2026-10-15", contextReportId: r2 });
    await call("PATCH", `/api/defects/${rows["N-01-04-CR-02"].id}`, { comment: "amended", contextReportId: r2 });
    await rec(r2, "N-02-05-CR-04", { categoryCode: "RR" });
    await rec(r2, "S-03-02-GL-02", { categoryCode: "RD" });
    await rec(r2, "N-02-01-SE-02", { recordType: "observation" }); // uncategorised observation
    await rec(r2, "N-02-03-GW-05"); // uncategorised defect
    await rec(r2, "E-02-02-GW-01", { categoryCode: "MS" });
    await rec(r2, "W-01-01-WP-02", { categoryCode: "WIP" });
    const closedNew = (await rec(r2, "E-03-03-GW-02", { categoryCode: "RR" })).body;
    await call("PATCH", `/api/defects/${closedNew.id}`, { status: "complete", dateClosed: "2026-10-15", contextReportId: r2 });
    if (!fv) await rec(r2, "N-02-02-GW-03", { recordType: "observation", actionNeeded: false });
    return { p, r2, put };
  };

  // ------------------------------------------------------------- new project
  const N = await setup();
  check("PUT categories accepted", N.put.status === 200, String(N.put.status));
  const cats = JSON.parse((await call("GET", `/api/projects/${N.p.id}`)).body.categories || "[]");
  check("PUT keeps an explicit tick", cats.find((c: any) => c.code === "RFI")?.inActionList === true, JSON.stringify(cats.find((c: any) => c.code === "RFI")));
  check("PUT keeps an explicit untick", cats.find((c: any) => c.code === "MS")?.inActionList === false, JSON.stringify(cats.find((c: any) => c.code === "MS")));
  check("PUT leaves unset categories unset", cats.find((c: any) => c.code === "RR")?.inActionList === undefined, JSON.stringify(cats.find((c: any) => c.code === "RR")));

  const data = (await call("GET", `/api/reports/${N.r2}/report-data`)).body;
  const byUid = (u: string) => (data.defects as any[]).find((d) => d.uid === u);
  check("carried record flagged carriedIn", byUid("N-01-05-CR-01")?.events?.carriedIn === true, JSON.stringify(byUid("N-01-05-CR-01")?.events?.carriedIn));
  check("new record not carriedIn", byUid("N-02-05-CR-04")?.events?.carriedIn === false, JSON.stringify(byUid("N-02-05-CR-04")?.events?.carriedIn));

  const tree = buildReportTree(data, "contractor");
  const v2 = tree.v2;
  check("new project gets the new layout", !!v2);
  if (v2) {
    const labels = (gs: any[]) => gs.map((g) => g.label);
    const uids = (gs: any[]) => gs.flatMap((g) => (g.kind === "itemise" ? g.defects.map((d: any) => d.uid) : [])).sort();
    const al = labels(v2.actionList);
    check("Action List holds Rectify, Redesign, Request Information", ["Rectify", "Redesign", "Request Information"].every((l) => al.includes(l)), al.join("|"));
    check("Action List leaves out WIP, Monitor and Project Instruction", !al.some((l) => ["Work in progress", "Monitor", "Project Instruction"].includes(l)), al.join("|"));
    check("Action List keeps uncategorised records", al.includes("Uncategorised") && al.includes("Observations"), al.join("|"));
    check("Action List order: ticked categories, then Uncategorised, then Observations",
      al.join("|") === "Rectify|Redesign|Request Information|Uncategorised|Observations", al.join("|"));
    const all = [v2.actionList, v2.newThisInspection, v2.closedThisInspection, v2.runningOpen].flatMap(labels);
    check("unused category never appears", !all.includes("Unused category"), all.join("|"));
    check("section 4 = open records raised this inspection",
      uids(v2.newThisInspection).join(",") === ["E-02-02-GW-01", "N-02-01-SE-02", "N-02-03-GW-05", "N-02-05-CR-04", "S-03-02-GL-02", "W-01-01-WP-02"].sort().join(","), uids(v2.newThisInspection).join(","));
    check("section 4 includes categories outside the Action List", labels(v2.newThisInspection).includes("Monitor") && labels(v2.newThisInspection).includes("Work in progress"), labels(v2.newThisInspection).join("|"));
    check("section 4 has the uncategorised observation under Observations", labels(v2.newThisInspection).at(-1) === "Observations", labels(v2.newThisInspection).join("|"));
    check("notes listed for section 4", v2.notes.length === 1 && String(v2.notes[0].uid).startsWith("Note-"), v2.notes.map((n: any) => n.uid).join(","));
    check("section 5 = closed this inspection (new and carried)", uids(v2.closedThisInspection).join(",") === "E-03-03-GW-02,N-01-02-CR-03", uids(v2.closedThisInspection).join(","));
    check("section 6 = open items from previous inspections only",
      uids(v2.runningOpen).join(",") === ["E-01-01-WP-01", "N-01-03-SE-01", "N-01-04-CR-02", "N-01-05-CR-01", "S-02-05-GL-01"].sort().join(","), uids(v2.runningOpen).join(","));
    check("amended earlier item stays in its category in section 6", v2.runningOpen.find((g: any) => g.label === "Rectify")?.defects.some((d: any) => d.uid === "N-01-04-CR-02") === true);
    check("section 6 is always itemised", v2.runningOpen.every((g: any) => g.kind === "itemise"));
    check("2.2 links only to records with a detail entry", !v2.detailIds.includes(byUid("N-01-05-CR-01").id) && v2.detailIds.includes(byUid("N-02-05-CR-04").id));
    check("elevation pins are limited to this report's records", v2.elevationRecordIds.length > 0 && v2.elevationRecordIds.every((id: number) => (data.defects as any[]).some((d) => d.id === id)));

    const outline = buildOutlineV2(tree, [{ id: 7, name: "North" }, { id: 8, name: "South" }]);
    const l1 = outline.filter((e) => e.level === 1);
    check("level 1 numbered 1-7, then Appendix A", l1.map((e) => e.no).join(",") === "1,2,3,4,5,6,7,Appendix A", l1.map((e) => `${e.no}:${e.title}`).join(" | "));
    check("Appendix A is last, with no section number", outline.at(-1)?.no === "Appendix A" && !/^\d/.test(outline.at(-1)!.no) && outline.at(-1)!.level === 1, JSON.stringify(outline.at(-1)));
    check("no numbered heading after Appendix A", outline.findIndex((e) => e.key === "appendixA") === outline.length - 1);
    const sec2 = outline.filter((e) => e.level === 2 && e.no.startsWith("2.")).map((e) => e.no);
    check("level 2 numbered consecutively", sec2.every((n, i) => n === `2.${i + 1}`), sec2.join(","));
    const sec7 = outline.filter((e) => e.level === 2 && e.no.startsWith("7.")).map((e) => e.title);
    check("one section 7 entry per elevation", sec7.join("|") === "North|South", sec7.join("|"));
  }

  // ------------------------------------------------------------- old project
  const O = await setup(1);
  const oldData = (await call("GET", `/api/reports/${O.r2}/report-data`)).body;
  const oldTree = buildReportTree(oldData, "contractor");
  check("old project keeps the previous layout (v2 null)", oldTree.v2 === null);
  check("old project still builds its previous sections", Array.isArray(oldTree.actionList.groups) && oldTree.actionList.groups.length > 0);

  http.close();
  let fail = 0;
  for (const [n, ok, d] of checks) { if (!ok) fail++; console.log(`${ok ? "PASS" : "FAIL"}  ${n}${ok || !d ? "" : `  -> ${d}`}`); }
  console.log(`\n${checks.length - fail}/${checks.length} passed`);
  process.exit(fail ? 1 : 0);
}
main().catch((e) => { console.error(e); process.exit(1); });
