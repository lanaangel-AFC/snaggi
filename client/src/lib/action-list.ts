// WP8: which categories appear in the report's 2.2 Action List (new projects only).
// A category's explicit "Action List" tick in project setup always wins. Until it has
// been ticked or unticked the default is: in the list, except Work in progress (WIP),
// Progress notes (PN) and any category labelled "Project Instruction", which ask nothing
// of the contractor within a timeframe. "Provide Information" (PI) is in the list.
export const ACTION_LIST_DEFAULT_OFF = new Set(["WIP", "PN"]);
const PROJECT_INSTRUCTION_LABEL = /project instruction/i;

export function isInActionList(cat: { code: string; label?: string; inActionList?: boolean } | null | undefined): boolean {
  if (!cat) return false;
  if (typeof cat.inActionList === "boolean") return cat.inActionList;
  if (PROJECT_INSTRUCTION_LABEL.test(String(cat.label || ""))) return false;
  return !ACTION_LIST_DEFAULT_OFF.has(String(cat.code).trim().toUpperCase());
}
