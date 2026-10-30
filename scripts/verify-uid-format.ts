// Checks for project UID formats (shared/uid-format.ts) and their report wording.
// Run: npx tsx scripts/verify-uid-format.ts
import { parseUidFormat, renderUid, describeUidFormat, parseUidWithFormat, isUidComplete, exampleUid } from "../shared/uid-format";
import { uidNomenclatureSentence, sortByUid, workTypeLabelFor, areaRefFor } from "../client/src/lib/render-helpers";
import { formatLocation } from "../shared/location";

let pass = 0, fail = 0;
const check = (name: string, ok: boolean, info?: unknown) => {
  if (ok) { pass++; console.log("PASS ", name); } else { fail++; console.log("FAIL ", name, info ?? ""); }
};

const grcRaw = JSON.stringify({
  pattern: "{level}.{fin}-{bracket}",
  parts: [
    { key: "level", label: "Level", options: [] },
    { key: "fin", label: "Fin", options: [] },
    { key: "bracket", label: "Bracket", options: ["A", "B", "C"] },
  ],
});
const grc = parseUidFormat(grcRaw)!;
check("GRC format parses", !!grc && grc.parts.length === 3);
check("empty format = standard UID", parseUidFormat("") === null && parseUidFormat(null) === null);
check("malformed format = standard UID", parseUidFormat("{bad") === null);
check("pattern without parts = standard UID", parseUidFormat(JSON.stringify({ pattern: "-", parts: [{ label: "Fin" }] })) === null);
check("9.14-B", renderUid(grc, { level: "9", fin: "14", bracket: "B" }) === "9.14-B");
check("missing bracket drops its separator", renderUid(grc, { level: "9", fin: "14" }) === "9.14");
check("missing fin", renderUid(grc, { level: "9", bracket: "B" }) === "9-B");
check("nothing entered", renderUid(grc, {}) === "");
check("complete check", isUidComplete(grc, { level: "9", fin: "14", bracket: "B" }) && !isUidComplete(grc, { level: "9" }));
check("description", describeUidFormat(grc) === "Level.Fin-Bracket", describeUidFormat(grc));
check("read back 9.14-B", JSON.stringify(parseUidWithFormat(grc, "9.14-B")) === JSON.stringify({ level: "9", fin: "14", bracket: "B" }), parseUidWithFormat(grc, "9.14-B"));
check("read back 10.2-AB", JSON.stringify(parseUidWithFormat(grc, "10.2-AB")) === JSON.stringify({ level: "10", fin: "2", bracket: "AB" }));
check("non-matching text", parseUidWithFormat(grc, "W-01-09-SLT-02") === null);
check("example uid", exampleUid(grc) === "1.2-A", exampleUid(grc));
check("keys from labels", parseUidFormat(JSON.stringify({ pattern: "{finno}", parts: [{ label: "Fin No." }] }))!.parts[0].key === "finno");

const project = { uidFormat: grcRaw };
const rec = (uid: string, loc: any, wt = "SLT") => ({ uid, workTypeCode: wt, locationStructured: JSON.stringify(loc) });
const s1 = uidNomenclatureSentence(project, [rec("9.14-B", { level: "9", fin: "14", bracket: "B" })]);
check("1.5 wording with example", s1 === "The UID comprises the following components: Level.Fin-Bracket. For example, 9.14-B refers to Level 9, Fin 14, Bracket B.", s1);
const s2 = uidNomenclatureSentence(project, []);
check("1.5 wording without records", s2 === "The UID comprises the following components: Level.Fin-Bracket.", s2);
check("1.5 standard wording unchanged", uidNomenclatureSentence({ uidFormat: "" }, []).startsWith("The UID comprises the following components: Area Ref"));

const dims = ["level", "fin", "bracket"];
check("location text", formatLocation(JSON.stringify({ level: "9", fin: "14", bracket: "B" }), dims) === "Level 9, Fin 14, Bracket B");
check("action list area cell", areaRefFor(rec("9.14-B", { level: "9", fin: "14", bracket: "B" }), "", project, dims) === "Level 9, Fin 14, Bracket B");
check("work type label from column", workTypeLabelFor(rec("9.14-B", {}, "SLT")) !== "" && workTypeLabelFor(rec("9.14-AB", {}, "")) !== undefined);

const sorted = [rec("10.2-A", {}), rec("9.14-B", {}), rec("9.2-C", {}), rec("9.14-A", {})].sort(sortByUid).map((r) => r.uid);
check("natural sort", sorted.join(",") === "9.2-C,9.14-A,9.14-B,10.2-A", sorted);
const legacy = [{ uid: "E-05-07-OT-02", workTypeCode: "OT" }, { uid: "E-05-07-OT-01", workTypeCode: "OT" }].sort(sortByUid).map((r) => r.uid);
check("standard sort unchanged", legacy.join(",") === "E-05-07-OT-01,E-05-07-OT-02", legacy);

console.log(`\n${pass}/${pass + fail} passed`);
if (fail) process.exit(1);
