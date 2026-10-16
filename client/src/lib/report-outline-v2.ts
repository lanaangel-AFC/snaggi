// WP8: section outline for the new report layout (new projects only). Both the PDF and
// the Word renderer number their headings from this one list, so the two exports and
// the PDF contents page always agree.
//
//   1. Introduction            1.1 General, then Scope of works (if any), Inspection
//                              particulars, Background information (if any), Record UID
//                              nomenclature, Limitations — numbered consecutively.
//   2. Project Status          2.1 Progress Summary, 2.2 Action List, then each
//                              narrative, Program and Stage Map that is present.
//   3. Project Instructions
//   4. New This Inspection
//   5. Closed This Inspection
//   6. Running List of Open Items
//   7. Elevations              7.1, 7.2 … one per elevation drawing with pins.
//   Appendix A                 Technical Guidance & Method Statements (unnumbered, last).

import type { ReportTree } from "./report-tree";
import { parseProjectSnapshot, resolveProjectField } from "./render-helpers";

export type OutlineEntry = { key: string; level: 1 | 2; no: string; title: string };

function jsonArray(raw: unknown): any[] {
  try {
    const v = JSON.parse(String(raw || "[]"));
    return Array.isArray(v) ? v : [];
  } catch {
    return [];
  }
}

export function scopeRowsV2(tree: ReportTree): any[] {
  const snap = parseProjectSnapshot((tree.report as any)?.projectSnapshot);
  return jsonArray(resolveProjectField(snap, tree.project, "scopeOfWorks"));
}

export function backgroundDocsV2(tree: ReportTree): any[] {
  const snap = parseProjectSnapshot((tree.report as any)?.projectSnapshot);
  return jsonArray(resolveProjectField(snap, tree.project, "backgroundDocs"));
}

export function buildOutlineV2(tree: ReportTree, elevations: Array<{ id: number; name: string }>): OutlineEntry[] {
  const out: OutlineEntry[] = [];
  const section = (no: number, key: string, title: string, subs: Array<[string, string]>) => {
    out.push({ key, level: 1, no: String(no), title });
    subs.forEach(([k, t], i) => out.push({ key: k, level: 2, no: `${no}.${i + 1}`, title: t }));
  };

  const s1: Array<[string, string]> = [["1.general", "General"]];
  if (scopeRowsV2(tree).length > 0) s1.push(["1.scope", "Scope of works"]);
  s1.push(["1.particulars", "Inspection particulars"]);
  if (backgroundDocsV2(tree).length > 0) s1.push(["1.background", "Background information"]);
  s1.push(["1.uid", "Record UID nomenclature"]);
  s1.push(["1.limitations", "Limitations"]);
  section(1, "1", "Introduction", s1);

  const s2: Array<[string, string]> = [];
  if (tree.progressSummary) s2.push(["2.progress", "Progress Summary"]);
  s2.push(["2.action", "Action List"]);
  tree.projectStatus.narratives.forEach((n: any, i: number) => s2.push([`2.narrative.${i}`, String(n.title || "Narrative")]));
  if (tree.projectStatus.program) s2.push(["2.program", "Program"]);
  if (tree.projectStatus.stageMap) s2.push(["2.stagemap", "Stage Map"]);
  section(2, "2", "Project Status", s2);

  section(3, "3", "Project Instructions", []);
  section(4, "4", "New This Inspection", []);
  section(5, "5", "Closed This Inspection", []);
  section(6, "6", "Running List of Open Items", []);
  section(7, "7", "Elevations", elevations.map((e) => [`7.elevation.${e.id}`, e.name]));

  out.push({ key: "appendixA", level: 1, no: "Appendix A", title: "Technical Guidance & Method Statements" });
  return out;
}

/** Heading text for an outline key: "2.2 Action List", "4. New This Inspection",
 *  "Appendix A – Technical Guidance & Method Statements". */
export function outlineHeading(outline: OutlineEntry[], key: string): { no: string; title: string; text: string } {
  const e = outline.find((x) => x.key === key);
  if (!e) return { no: "", title: key, text: key };
  if (e.no.startsWith("Appendix")) return { no: e.no, title: e.title, text: `${e.no} \u2013 ${e.title}` };
  const no = e.level === 1 ? `${e.no}.` : e.no;
  return { no, title: e.title, text: `${no} ${e.title}` };
}
