// WP2: record wording — feature version on projects and the wording helpers.
import { storage } from "../server/storage";
import { usesRecordWording, recordWords, typeWords } from "../client/src/lib/wording";

let pass = 0, fail = 0;
const check = (name: string, ok: boolean, detail = "") => { if (ok) pass++; else { fail++; console.log(`FAIL ${name} ${detail}`); } };
const base = { address: "x", client: "x", inspector: "x", createdAt: new Date().toISOString() } as any;

const fresh = await storage.createProject({ ...base, name: "New" });
check("new project gets feature version 2", fresh.featureVersion === 2, String(fresh.featureVersion));
const legacy = await storage.createProject({ ...base, name: "Old", featureVersion: 1 });
check("explicit version 1 respected (test seeding)", legacy.featureVersion === 1);
const edited = await storage.updateProject(fresh.id, { name: "Renamed" } as any);
check("editing a project keeps its version", edited?.featureVersion === 2);

check("new project uses record wording", usesRecordWording(fresh));
check("old project keeps defect wording", !usesRecordWording(legacy));
check("missing version counts as old", !usesRecordWording({}) && !usesRecordWording(null));
check("general wording, new: record", recordWords(fresh).One === "Record" && recordWords(fresh).many === "records");
check("general wording, old: defect", recordWords(legacy).One === "Defect");
check("form wording, new observation", typeWords(fresh, "observation").One === "Observation");
check("form wording, new defect", typeWords(fresh, "defect").One === "Defect");
check("form wording, old observation stays defect", typeWords(legacy, "observation").One === "Defect");

console.log(`${pass}/${pass + fail} passed`);
process.exit(fail ? 1 : 0);
