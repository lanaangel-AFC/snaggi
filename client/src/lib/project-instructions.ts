// WP6: Project Instructions in exports. New projects (featureVersion >= 2) get a
// "3. Project Instructions" section straight after "2. Project Status"; the later
// sections move down one. Older projects keep their original section numbers.
import { usesRecordWording } from "./wording";

export const INSTRUCTION_HEADERS = ["ID", "Date Issued", "Work Type(s)", "Instruction", "Action Required", "Responsibility", "By When"];

export function hasInstructionSection(data: any): boolean {
  return usesRecordWording(data?.project);
}

/** Section number for a section that was numbered `base` before Project Instructions existed. */
export function sectionNo(data: any, base: number): number {
  return hasInstructionSection(data) && base >= 3 ? base + 1 : base;
}

/** Local calendar date (YYYY-MM-DD) for a stored timestamp, so a withdrawal made on the
 *  morning of the 16th in Sydney doesn't print as the 15th (UTC). */
export function localDate(iso: string | null | undefined): string {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return String(iso).slice(0, 10);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

/** One row per instruction, newest first (the server already sorts them). */
export function instructionRows(data: any): string[][] {
  const list: any[] = Array.isArray(data?.projectInstructions) ? data.projectInstructions : [];
  return list.map((i) => [
    String(i.ref || ""),
    String(i.dateIssued || ""),
    (Array.isArray(i.workTypeLabels) && i.workTypeLabels.length ? i.workTypeLabels : i.workTypes || []).join(", "),
    String(i.observation || "").trim() || "\u2014",
    String(i.actionRequired || "").trim() || "\u2014",
    String(i.responsibility || "").trim() || "\u2014",
    i.status === "withdrawn" ? `Withdrawn ${localDate(i.dateWithdrawn)}`.trim() : "Ongoing",
  ]);
}

export const INSTRUCTION_SUBTITLE =
  "Instructions to the contractor that apply to all similar works across the whole project. Newest first.";
