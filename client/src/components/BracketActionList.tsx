// 2.2 Bracket Action List editor (projects with bracketActions on; GRC Fin Rectification).
// Project-level rows carried across inspections. ID follows the project's UID format
// (GRC: Level.Fin-Bracket → 9.14-B); work type is always SLT – Slot Lengthening.

import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { Plus, Trash2 } from "lucide-react";
import { useState } from "react";
import { parseUidFormat, renderUid } from "@shared/uid-format";

export const BRACKET_CONDITIONS = ["Incomplete routing depth", "Required clearance not achieved", "Damage to bracket"];
const STATUSES = ["Open", "Pending", "Closed"];
const STATUS_CLASS: Record<string, string> = {
  Open: "text-amber-600", Pending: "text-blue-600", Closed: "text-green-700",
};

type Row = {
  id: number; level: string; fin: string; bracket: string; workType: string;
  condition: string; consultantInspection: boolean; status: string; dateBy: string;
};

async function errMessage(e: any): Promise<string> {
  const m = String(e?.message || "");
  const json = m.slice(m.indexOf("{"));
  try { return JSON.parse(json).message || m; } catch { return m || "Save failed"; }
}

export default function BracketActionList({ projectId, project }: { projectId: string; project: any }) {
  const { toast } = useToast();
  const key = [`/api/projects/${projectId}/bracket-actions`];
  const { data: rows = [] } = useQuery<Row[]>({ queryKey: key });
  const fmt = parseUidFormat(project?.uidFormat);
  const uidOf = (r: { level: string; fin: string; bracket: string }) =>
    fmt ? renderUid(fmt, r) : [r.level, r.fin, r.bracket].filter(Boolean).join(".");
  const bracketOptions = fmt?.parts.find((p) => p.key === "bracket")?.options?.length
    ? fmt!.parts.find((p) => p.key === "bracket")!.options : ["A", "B", "C"];

  const [draft, setDraft] = useState({ level: "", fin: "", bracket: "" });
  const refresh = () => queryClient.invalidateQueries({ queryKey: key });

  const addMutation = useMutation({
    mutationFn: async () => (await apiRequest("POST", `/api/projects/${projectId}/bracket-actions`, { ...draft, status: "Open" })).json(),
    onSuccess: () => { refresh(); setDraft((d) => ({ level: d.level, fin: "", bracket: "" })); },
    onError: async (e) => toast({ title: await errMessage(e), variant: "destructive" }),
  });
  const patchMutation = useMutation({
    mutationFn: async ({ id, patch }: { id: number; patch: Partial<Row> }) => (await apiRequest("PATCH", `/api/bracket-actions/${id}`, patch)).json(),
    onSuccess: refresh,
    onError: async (e) => { toast({ title: await errMessage(e), variant: "destructive" }); refresh(); },
  });
  const deleteMutation = useMutation({
    mutationFn: async (id: number) => apiRequest("DELETE", `/api/bracket-actions/${id}`),
    onSuccess: refresh,
  });
  const patch = (id: number, p: Partial<Row>) => patchMutation.mutate({ id, patch: p });

  const preview = uidOf(draft);
  const canAdd = !!draft.level.trim() && !!draft.fin.trim() && !!draft.bracket.trim();

  return (
    <div className="mb-8" data-testid="bracket-action-list">
      <div className="flex items-center justify-between mb-1">
        <h3 className="text-base font-medium">Bracket Action List</h3>
        <span className="text-xs text-muted-foreground">{rows.length} bracket{rows.length === 1 ? "" : "s"}</span>
      </div>
      <p className="text-xs text-muted-foreground mb-3">
        Shown in the report under 2.2 Action List. Work type is SLT – Slot Lengthening for every row.
      </p>

      <Card className="p-3 mb-3">
        <div className="grid grid-cols-[1fr_1fr_1fr_auto] gap-2 items-end">
          <div>
            <Label className="text-xs">Level</Label>
            <Input value={draft.level} onChange={(e) => setDraft({ ...draft, level: e.target.value.replace(/\s+/g, "") })} className="font-mono text-center" data-testid="bracket-new-level" />
          </div>
          <div>
            <Label className="text-xs">Fin</Label>
            <Input value={draft.fin} onChange={(e) => setDraft({ ...draft, fin: e.target.value.replace(/\s+/g, "") })} className="font-mono text-center" data-testid="bracket-new-fin" />
          </div>
          <div>
            <Label className="text-xs">Bracket</Label>
            <Input list="bracket-options" value={draft.bracket} onChange={(e) => setDraft({ ...draft, bracket: e.target.value.replace(/\s+/g, "").toUpperCase() })} className="font-mono text-center" autoComplete="off" data-testid="bracket-new-bracket" />
            <datalist id="bracket-options">{bracketOptions.map((o) => <option key={o} value={o} />)}</datalist>
          </div>
          <Button size="sm" className="h-9" disabled={!canAdd || addMutation.isPending} onClick={() => addMutation.mutate()} data-testid="bracket-add">
            <Plus className="w-4 h-4 mr-1" /> Add
          </Button>
        </div>
        {preview && <p className="text-xs text-muted-foreground mt-2">ID: <span className="font-mono font-semibold text-foreground">{preview}</span></p>}
      </Card>

      {rows.length === 0 ? (
        <p className="text-sm text-muted-foreground">No brackets yet.</p>
      ) : (
        <div className="space-y-2">
          {rows.map((r) => (
            <Card key={r.id} className="p-3" data-testid={`bracket-row-${uidOf(r)}`}>
              <div className="flex items-center justify-between gap-2 mb-2">
                <span className="font-mono font-semibold">{uidOf(r)}</span>
                <div className="flex items-center gap-2">
                  <span className="text-xs text-muted-foreground">SLT</span>
                  <Button size="icon" variant="ghost" className="h-8 w-8" aria-label="Delete row"
                    onClick={() => { if (confirm(`Remove ${uidOf(r)} from the Action List?`)) deleteMutation.mutate(r.id); }}>
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
              </div>
              <div className="grid grid-cols-[auto_minmax(0,1fr)_minmax(0,1.4fr)] sm:grid-cols-[minmax(0,2fr)_auto_minmax(0,1fr)_minmax(0,1fr)] gap-2 items-end">
                <div className="col-span-3 sm:col-span-1">
                  <Label className="text-xs">Observed condition</Label>
                  <Select value={r.condition || undefined} onValueChange={(v) => patch(r.id, { condition: v })}>
                    <SelectTrigger className="h-9" data-testid={`bracket-condition-${r.id}`}><SelectValue placeholder="Select…" /></SelectTrigger>
                    <SelectContent>{BRACKET_CONDITIONS.map((c) => <SelectItem key={c} value={c}>{c}</SelectItem>)}</SelectContent>
                  </Select>
                </div>
                <div>
                  <Label className="text-xs whitespace-nowrap">Consultant insp.</Label>
                  <div className="flex rounded-md border overflow-hidden h-9">
                    {[true, false].map((v) => (
                      <button key={String(v)} type="button"
                        className={`px-3 text-sm font-medium ${r.consultantInspection === v ? "bg-primary text-primary-foreground" : "bg-background text-muted-foreground hover:bg-accent/60"}`}
                        onClick={() => patch(r.id, { consultantInspection: v })}
                        data-testid={`bracket-consultant-${r.id}-${v ? "y" : "n"}`}>
                        {v ? "Y" : "N"}
                      </button>
                    ))}
                  </div>
                </div>
                <div>
                  <Label className="text-xs">Status</Label>
                  <Select value={r.status} onValueChange={(v) => patch(r.id, { status: v })}>
                    <SelectTrigger className={`h-9 font-medium ${STATUS_CLASS[r.status] || ""}`} data-testid={`bracket-status-${r.id}`}><SelectValue /></SelectTrigger>
                    <SelectContent>{STATUSES.map((s) => <SelectItem key={s} value={s}>{s}</SelectItem>)}</SelectContent>
                  </Select>
                </div>
                <div>
                  <Label className="text-xs">Date by</Label>
                  <Input type="date" className="h-9 min-w-0 w-full px-2" value={r.dateBy || ""} onChange={(e) => patch(r.id, { dateBy: e.target.value })} data-testid={`bracket-date-${r.id}`} />
                </div>
              </div>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}
