// WP8: section outline for the new report layout (new projects only). Both the PDF and
// the Word renderer number their headings from this one list, so the two exports and
// the PDF contents page always agree.
//
//   1. Introduction            1.1 General, then Scope of works (if any), Inspection
//                              particulars, Background information (if any), Record UID
//                              nomenclature, Limitations — numbered consecutively.
//   2. Project Status          2.1 Progress Summary, 2.2 Action List, then Project
//                              Narratives and Holds (if any) with each narrative/hold as
//                              a level 3 heading (2.3.1, 2.3.2 …), then Program and Stage
//                              Map if present.
//   3. Project Instructions
//   4. New This Inspection
//   5. Closed This Inspection
//   6. Running List of Open Items
//   7. Elevations              7.1, 7.2 … one per elevation drawing with pins.
//   Appendix A                 Technical Guidance & Method Statements (unnumbered, last).

import type { ReportTree } from "./report-tree";
import { parseProjectSnapshot, resolveProjectField } from "./render-helpers";

export const NARRATIVES_HEADING = "Project Narratives and Holds";

export type OutlineEntry = { key: string; level: 1 | 2 | 3; no: string; title: string };

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
  // A level 2 entry may carry level 3 children: [key, title, [[childKey, childTitle], …]].
  type Sub = [string, string] | [string, string, Array<[string, string]>];
  const section = (no: number, key: string, title: string, subs: Sub[]) => {
    out.push({ key, level: 1, no: String(no), title });
    subs.forEach(([k, t, kids], i) => {
      const n2 = `${no}.${i + 1}`;
      out.push({ key: k, level: 2, no: n2, title: t });
      (kids || []).forEach(([k3, t3], j) => out.push({ key: k3, level: 3, no: `${n2}.${j + 1}`, title: t3 }));
    });
  };

  const s1: Sub[] = [["1.general", "General"]];
  if (scopeRowsV2(tree).length > 0) s1.push(["1.scope", "Scope of works"]);
  s1.push(["1.particulars", "Inspection particulars"]);
  if (backgroundDocsV2(tree).length > 0) s1.push(["1.background", "Background information"]);
  s1.push(["1.uid", "Record UID nomenclature"]);
  s1.push(["1.limitations", "Limitations"]);
  section(1, "1", "Introduction", s1);

  const s2: Sub[] = [];
  if (tree.progressSummary) s2.push(["2.progress", "Progress Summary"]);
  s2.push(["2.action", "Action List"]);
  const narratives = tree.projectStatus.narratives;
  if (narratives.length > 0) {
    s2.push(["2.narratives", NARRATIVES_HEADING,
      narratives.map((n: any, i: number): [string, string] => [`2.narrative.${i}`, String(n.title || "Narrative").trim() || "Narrative"])]);
  }
  if (tree.projectStatus.program) s2.push(["2.program", "Program"]);
  if (tree.projectStatus.stageMap) s2.push(["2.stagemap", "Stage Map"]);
  section(2, "2", "Project Status", s2);

  section(3, "3", "Project Instructions", []);
  section(4, "4", "New This Inspection", []);
  section(5, "5", "Closed This Inspection", []);
  section(6, "6", "Running List of Open Items", []);
  section(7, "7", "Elevations", elevations.map((e): Sub => [`7.elevation.${e.id}`, e.name]));

  out.push({ key: "appendixA", level: 1, no: "Appendix A", title: "Technical Guidance & Method Statements" });
  return out;
}

// Title page lines below the address, top to bottom: project name, report title, then
// "Inspection AA | Revision 01". The report title falls back to "Site Visit Report"; a
// report title identical to the project name is shown once.
export function titlePageLinesV2(o: { projectName?: string; reportTitle?: string; inspectionNumber?: string; revision?: string }): { project: string; reportTitle: string; revisionLine: string } {
  const t = (v: any) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim());
  const project = t(o.projectName);
  let reportTitle = t(o.reportTitle) || "Site Visit Report";
  if (project && reportTitle.toLowerCase() === project.toLowerCase()) reportTitle = "Site Visit Report";
  const rev = `Revision ${t(o.revision) || "01"}`;
  const insp = t(o.inspectionNumber);
  return { project, reportTitle, revisionLine: insp ? `Inspection ${insp} | ${rev}` : rev };
}

// Harvard-style reference for a background document:
//   Originator (Year) Title. Doc. no. X, Rev. Y.
// The title is the italic part. "n.d." stands in for a missing year. The document type
// is not shown.
export type HarvardParts = { lead: string; title: string; tail: string };
export function harvardReference(b: any): HarvardParts | null {
  const s = (v: any) => (v == null ? "" : String(v).replace(/\s+/g, " ").trim());
  const stripEnd = (v: string) => v.replace(/[.\s]+$/, "");
  const originator = stripEnd(s(b?.originator));
  const m = s(b?.date).match(/\b(19|20)\d{2}\b/);
  const year = m ? m[0] : "";
  const title = stripEnd(s(b?.title));
  const details = [s(b?.docNumbers) ? `Doc. no. ${s(b.docNumbers)}` : "", s(b?.revision) ? `Rev. ${s(b.revision)}` : ""].filter(Boolean);
  if (!originator && !title && details.length === 0) return null;
  const lead = originator ? `${originator} (${year || "n.d."}) ` : (year ? `(${year}) ` : "");
  const tail = details.length ? ` ${stripEnd(details.join(", "))}.` : "";
  return { lead, title: title ? `${title}.` : "", tail };
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
