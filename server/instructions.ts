// WP6: Project Instructions — project-wide instructions to the contractor tied to one or
// more work types rather than a location. New projects only (featureVersion >= 2).
// Kept in their own table so they never enter per-inspection copying, to-do lists or pins.
import { sqlite } from "./storage";

export type Instruction = {
  id: number;
  projectId: number;
  reportId: number | null;
  ref: string;
  workTypes: string[];
  observation: string;
  actionRequired: string;
  responsibility: string;
  status: "active" | "withdrawn";
  dateIssued: string;
  dateWithdrawn: string | null;
  photos: { filename: string; caption?: string | null }[];
  createdAt: string;
  updatedAt: string | null;
};

const parseArr = (raw: unknown): any[] => {
  try { const v = JSON.parse(String(raw ?? "[]")); return Array.isArray(v) ? v : []; } catch { return []; }
};

const fromRow = (r: any): Instruction => ({
  id: r.id,
  projectId: r.project_id,
  reportId: r.report_id ?? null,
  ref: r.ref,
  workTypes: parseArr(r.work_types).map(String),
  observation: r.observation ?? "",
  actionRequired: r.action_required ?? "",
  responsibility: r.responsibility ?? "",
  status: r.status === "withdrawn" ? "withdrawn" : "active",
  dateIssued: r.date_issued,
  dateWithdrawn: r.date_withdrawn ?? null,
  photos: parseArr(r.photos),
  createdAt: r.created_at,
  updatedAt: r.updated_at ?? null,
});

// Newest first: by issue date, then by number.
const sortNewestFirst = (list: Instruction[]) =>
  list.sort((a, b) => (b.dateIssued || "").localeCompare(a.dateIssued || "") || b.id - a.id);

export function listInstructions(projectId: number): Instruction[] {
  const rows = sqlite.prepare("SELECT * FROM project_instructions WHERE project_id = ?").all(projectId);
  return sortNewestFirst(rows.map(fromRow));
}

export function getInstruction(id: number): Instruction | undefined {
  const r = sqlite.prepare("SELECT * FROM project_instructions WHERE id = ?").get(id);
  return r ? fromRow(r) : undefined;
}

// PI-01, PI-02… per project. The highest number issued is kept in meta so a deleted
// instruction's number is never handed out again.
function nextInstructionRef(projectId: number): string {
  const key = `pi_seq_project_${projectId}`;
  const issued = sqlite.prepare("SELECT value FROM meta WHERE key = ?").get(key) as { value: string } | undefined;
  let max = issued ? parseInt(issued.value, 10) || 0 : 0;
  const rows = sqlite.prepare("SELECT ref FROM project_instructions WHERE project_id = ?").all(projectId) as { ref: string }[];
  for (const { ref } of rows) {
    const m = /^PI-(\d+)$/.exec(ref);
    if (m && Number(m[1]) > max) max = Number(m[1]);
  }
  sqlite.prepare("INSERT OR REPLACE INTO meta (key, value) VALUES (?, ?)").run(key, String(max + 1));
  return `PI-${String(max + 1).padStart(2, "0")}`;
}

export type InstructionInput = {
  reportId?: number | null;
  workTypes?: string[];
  observation?: string;
  actionRequired?: string;
  responsibility?: string;
  dateIssued?: string;
  status?: "active" | "withdrawn";
};

const cleanWorkTypes = (v: unknown): string[] =>
  Array.from(new Set((Array.isArray(v) ? v : []).map((x) => String(x).trim().toUpperCase()).filter(Boolean)));

export function validateInstruction(input: InstructionInput, partial: boolean): string | null {
  if (!partial || input.workTypes !== undefined) {
    if (cleanWorkTypes(input.workTypes).length === 0) return "Choose at least one work type.";
  }
  if (!partial || input.actionRequired !== undefined) {
    if (!String(input.actionRequired ?? "").trim()) return "Action required is needed.";
  }
  if (!partial || input.responsibility !== undefined) {
    if (!String(input.responsibility ?? "").trim()) return "Choose who is responsible.";
  }
  if (input.dateIssued !== undefined && !/^\d{4}-\d{2}-\d{2}$/.test(String(input.dateIssued))) return "Date issued must be a date.";
  if (input.status !== undefined && input.status !== "active" && input.status !== "withdrawn") return "Status must be active or withdrawn.";
  return null;
}

export function createInstruction(projectId: number, input: InstructionInput): Instruction {
  const now = new Date().toISOString();
  const info = sqlite.prepare(`INSERT INTO project_instructions
    (project_id, report_id, ref, work_types, observation, action_required, responsibility, status, date_issued, photos, created_at, updated_at)
    VALUES (?, ?, ?, ?, ?, ?, ?, 'active', ?, '[]', ?, ?)`).run(
    projectId, input.reportId ?? null, nextInstructionRef(projectId), JSON.stringify(cleanWorkTypes(input.workTypes)),
    String(input.observation ?? "").trim(), String(input.actionRequired ?? "").trim(), String(input.responsibility ?? "").trim(),
    input.dateIssued || now.slice(0, 10), now, now,
  );
  return getInstruction(Number(info.lastInsertRowid))!;
}

// Applies a patch and logs every changed field to instruction_history.
export function updateInstruction(id: number, input: InstructionInput): Instruction | undefined {
  const cur = getInstruction(id);
  if (!cur) return undefined;
  const now = new Date().toISOString();
  const sets: string[] = [];
  const vals: any[] = [];
  const log = (field: string, oldV: string | null, newV: string | null) =>
    sqlite.prepare("INSERT INTO instruction_history (instruction_id, field, old_value, new_value, changed_at) VALUES (?, ?, ?, ?, ?)").run(id, field, oldV, newV, now);

  const text = (key: "observation" | "actionRequired" | "responsibility", col: string) => {
    if (input[key] === undefined) return;
    const v = String(input[key] ?? "").trim();
    if (v === cur[key]) return;
    sets.push(`${col} = ?`); vals.push(v); log(key, cur[key], v);
  };
  text("observation", "observation");
  text("actionRequired", "action_required");
  text("responsibility", "responsibility");
  if (input.workTypes !== undefined) {
    const v = cleanWorkTypes(input.workTypes);
    if (JSON.stringify(v) !== JSON.stringify(cur.workTypes)) {
      sets.push("work_types = ?"); vals.push(JSON.stringify(v)); log("workTypes", cur.workTypes.join(", "), v.join(", "));
    }
  }
  if (input.dateIssued !== undefined && input.dateIssued !== cur.dateIssued) {
    sets.push("date_issued = ?"); vals.push(input.dateIssued); log("dateIssued", cur.dateIssued, input.dateIssued);
  }
  if (input.status !== undefined && input.status !== cur.status) {
    sets.push("status = ?", "date_withdrawn = ?");
    vals.push(input.status, input.status === "withdrawn" ? now : null);
    log("status", cur.status, input.status);
  }
  if (sets.length === 0) return cur;
  sets.push("updated_at = ?"); vals.push(now);
  sqlite.prepare(`UPDATE project_instructions SET ${sets.join(", ")} WHERE id = ?`).run(...vals, id);
  return getInstruction(id);
}

export function setInstructionPhotos(id: number, photos: Instruction["photos"]): Instruction | undefined {
  sqlite.prepare("UPDATE project_instructions SET photos = ?, updated_at = ? WHERE id = ?").run(JSON.stringify(photos), new Date().toISOString(), id);
  return getInstruction(id);
}

export function deleteInstruction(id: number): void {
  sqlite.prepare("DELETE FROM instruction_history WHERE instruction_id = ?").run(id);
  sqlite.prepare("DELETE FROM project_instructions WHERE id = ?").run(id);
}


export function getInstructionHistory(id: number) {
  return (sqlite.prepare("SELECT * FROM instruction_history WHERE instruction_id = ? ORDER BY id DESC").all(id) as any[]).map((r) => ({
    id: r.id, field: r.field, oldValue: r.old_value, newValue: r.new_value, changedAt: r.changed_at,
  }));
}

// The instructions as they stood for one report: issued before the next inspection began,
// and shown as withdrawn only if they were withdrawn before then. Wording is the current
// wording (the report shows current wording only).
export function instructionsForReport(projectId: number, windowEnd: string): Instruction[] {
  return listInstructions(projectId)
    .filter((i) => i.createdAt < windowEnd)
    .map((i) => (i.status === "withdrawn" && i.dateWithdrawn && i.dateWithdrawn >= windowEnd
      ? { ...i, status: "active" as const, dateWithdrawn: null }
      : i));
}
