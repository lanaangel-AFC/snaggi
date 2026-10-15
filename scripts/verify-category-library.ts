// WP4: generic category library shared across projects.
import { storage, sqlite, getGlobalSettings, addToCategoryLibrary, removeFromCategoryLibrary } from "../server/storage";

let pass = 0, fail = 0;
const check = (name: string, ok: boolean, detail = "") => {
  if (ok) pass++; else { fail++; console.log(`FAIL ${name} ${detail}`); }
};

const lib0 = getGlobalSettings().categoryLibrary;
// WP6 retired "Project Note" (PN) from the library: Project Instructions have their own button.
check("fresh install seeds the standard four (no Project Note)", JSON.stringify(lib0.map((c) => c.code)) === JSON.stringify(["RR", "WIP", "PI", "RD"]), JSON.stringify(lib0));
check("seed labels", lib0.find((c) => c.code === "PI")?.label === "Provide Information");
check("seed is gated by a meta key", !!sqlite.prepare(`SELECT value FROM meta WHERE key='category_library_seed_v1'`).get());
check("work types still returned alongside", getGlobalSettings().workTypes.length > 0);

const a = addToCategoryLibrary({ code: "var", label: "Variation" });
check("add: new code added and upper-cased", a.added && a.library.some((c) => c.code === "VAR" && c.label === "Variation"));
const dup = addToCategoryLibrary({ code: "VAR", label: "Something else" });
check("add: same code is not overwritten", !dup.added && dup.existing?.label === "Variation");
check("add: library still has one VAR", getGlobalSettings().categoryLibrary.filter((c) => c.code === "VAR").length === 1);
const bad = addToCategoryLibrary({ code: " ", label: "x" });
check("add: blank code rejected", !bad.added);

// A project using a library category keeps it when the library entry is removed.
const p = await storage.createProject({ name: "Lib test", address: "x", client: "x", inspector: "x", createdAt: new Date().toISOString() } as any);
await storage.updateProject(p.id, { categories: JSON.stringify([{ code: "VAR", label: "Variation" }, { code: "RR", label: "Rectify" }]) } as any);
const r = removeFromCategoryLibrary("VAR");
check("remove: entry removed from library", r.removed && !getGlobalSettings().categoryLibrary.some((c) => c.code === "VAR"));
const after = JSON.parse((await storage.getProject(p.id))!.categories || "[]");
check("remove: project keeps its copy", after.some((c: any) => c.code === "VAR"));
check("remove: unknown code is a no-op", !removeFromCategoryLibrary("NOPE").removed);

// Editing a project's label never changes the library.
await storage.updateProject(p.id, { categories: JSON.stringify([{ code: "RR", label: "Rectify and repair" }]) } as any);
check("project label edit leaves library label alone", getGlobalSettings().categoryLibrary.find((c) => c.code === "RR")?.label === "Rectify");

// Existing projects' categories were not touched by the library seed.
check("library seed does not write project categories", JSON.parse((await storage.getProject(p.id))!.categories || "[]").length === 1);

console.log(`${pass}/${pass + fail} passed`);
process.exit(fail ? 1 : 0);
