// Bracket Record Review editor (projects with bracketActions on; GRC Fin Rectification).
// A log of the contractor bracket records reviewed each week and the outcome. The report
// links each review to the latest inspection when it is entered, and that inspection's
// report shows it in 2.2. Acceptable reviews are then archived. "Requires rectification"
// reviews stay current (and in 2.2) until their Bracket Action List item is closed.

import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { Plus, Trash2, CheckCircle2, ListPlus } from "lucide-react";
import { useState } from "react";
import { parseUidFormat, renderUid } from "@shared/uid-format";

export const REVIEW_OUTCOMES = ["Acceptable", "Requires rectification"] as const;
const OUTCOME_ON: Record<string, string> = {
  Acceptable: "bg-green-700 text-white",
  "Requires rectification": "bg-red-600 text-white",
};

type Review = { id: number; level: string; bracket: string; dateSubmitted: string; outcome: string; reportId: number | null; createdAt: string };
type ReportRow = { id: number; inspectionNumber: string | null; createdAt: string };
type ActionRow = { id: number; level: string; fin: string; bracket: string; status: string };
export type BracketPrefill = { level: string; fin: string; bracket: string; nonce: number };

async function errMessage(e: any): Promise<string> {
  const m = String(e?.message || "");
  const json = m.slice(m.indexOf("{"));
  try { return JSON.parse(json).message || m; } catch { return m || "Save failed"; }
}

// "14-B", "14.B", "14 B", "14B" → fin 14, bracket B. A bare letter ("B") gives no fin.
export function parseReviewBracket(text: string): { fin: string; bracket: string } {
  const t = String(text || "").trim().toUpperCase();
  const m = /^(\d+)\s*[-.\s]?\s*([A-Z]+)$/.exec(t);
  if (m) return { fin: m[1], bracket: m[2] };
  if (/^[A-Z]+$/.test(t)) return { fin: "", bracket: t };
  return { fin: "", bracket: "" };
}

const todayIso = () => {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
};
const recordedOn = (iso: string) => {
  const d = new Date(iso);
  return Number.isNaN(d.getTime()) ? "" : d.toLocaleDateString("en-AU", { day: "numeric", month: "short", year: "numeric" });
};

export default function BracketRecordReview({ projectId, project, onPrefillAction }: {
  projectId: string; project: any; onPrefillAction: (p: Omit<BracketPrefill, "nonce">) => void;
}) {
  const { toast } = useToast();
  const key = [`/api/projects/${projectId}/bracket-reviews`];
  const actionKey = [`/api/projects/${projectId}/bracket-actions`];
  const { data: rows = [] } = useQuery<Review[]>({ queryKey: key });
  const { data: actions = [] } = useQuery<ActionRow[]>({ queryKey: actionKey });
  const { data: reports = [] } = useQuery<ReportRow[]>({ queryKey: [`/api/projects/${projectId}/reports`] });
  const latest = [...reports].sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)))[0];
  const inspectionLabel = (reportId: number | null) => {
    const r = reports.find((x) => x.id === reportId);
    return r?.inspectionNumber ? `Inspection ${String(r.inspectionNumber).padStart(2, "0")}` : "";
  };
  const fmt = parseUidFormat(project?.uidFormat);
  const uidOf = (r: { level: string; fin: string; bracket: string }) =>
    fmt ? renderUid(fmt, r) : [r.level, r.fin, r.bracket].filter(Boolean).join(".");

  const [draft, setDraft] = useState({ level: "", bracket: "", dateSubmitted: todayIso(), outcome: "" });
  const [showArchived, setShowArchived] = useState(false);
  const refresh = () => queryClient.invalidateQueries({ queryKey: key });

  const addMutation = useMutation({
    mutationFn: async () => (await apiRequest("POST", `/api/projects/${projectId}/bracket-reviews`, draft)).json(),
    onSuccess: () => { refresh(); setDraft((d) => ({ ...d, bracket: "", outcome: "" })); },
    onError: async (e) => toast({ title: await errMessage(e), variant: "destructive" }),
  });
  const patchMutation = useMutation({
    mutationFn: async ({ id, patch }: { id: number; patch: Partial<Review> }) => (await apiRequest("PATCH", `/api/bracket-reviews/${id}`, patch)).json(),
    onSuccess: refresh,
    onError: async (e) => { toast({ title: await errMessage(e), variant: "destructive" }); refresh(); },
  });
  const deleteMutation = useMutation({
    mutationFn: async (id: number) => apiRequest("DELETE", `/api/bracket-reviews/${id}`),
    onSuccess: refresh,
  });
  const addActionMutation = useMutation({
    mutationFn: async (b: { level: string; fin: string; bracket: string }) =>
      (await apiRequest("POST", `/api/projects/${projectId}/bracket-actions`, { ...b, status: "Open" })).json(),
    onSuccess: (_d, b) => { queryClient.invalidateQueries({ queryKey: actionKey }); toast({ title: `${uidOf(b)} added to the Bracket Action List`, description: "Set its observed condition and date by there." }); },
    onError: async (e) => toast({ title: await errMessage(e), variant: "destructive" }),
  });
  const patch = (id: number, p: Partial<Review>) => patchMutation.mutate({ id, patch: p });

  // The Action List row this review points at, if it is already there.
  const actionFor = (r: Review) => {
    const { fin, bracket } = parseReviewBracket(r.bracket);
    if (!fin || !bracket) return null;
    return actions.find((a) => a.level === r.level && a.fin === fin && a.bracket === bracket) || null;
  };
  const addToActionList = (r: Review) => {
    const { fin, bracket } = parseReviewBracket(r.bracket);
    if (r.level && fin && bracket) addActionMutation.mutate({ level: r.level, fin, bracket });
    else {
      onPrefillAction({ level: r.level, fin, bracket });
      toast({ title: "Finish the bracket ID in the Bracket Action List", description: "Enter the Fin and Bracket, then Add." });
    }
  };

  const canAdd = !!draft.level.trim() && !!draft.bracket.trim() && !!draft.outcome;
  // Current = recorded for the latest inspection, or a rectification still awaiting closure.
  // Everything else is archived: it was reported in its own inspection and won't repeat.
  const isCurrent = (r: Review) =>
    r.reportId == null || r.reportId === latest?.id ||
    (r.outcome === "Requires rectification" && actionFor(r)?.status !== "Closed");
  const current = rows.filter(isCurrent);
  const archived = rows.filter((r) => !isCurrent(r));
  const visible = showArchived ? [...current, ...archived] : current;
  const awaiting = rows.filter((r) => r.outcome === "Requires rectification" && !actionFor(r)).length;

  return (
    <div className="mb-8" data-testid="bracket-record-review">
      <div className="flex items-center justify-between mb-1">
        <h3 className="text-base font-medium">Bracket Record Review</h3>
        <span className="text-xs text-muted-foreground">{current.length} current · {archived.length} archived</span>
      </div>
      <p className="text-xs text-muted-foreground mb-3">
        Contractor bracket records reviewed and the outcome. New reviews belong to the latest inspection{latest?.inspectionNumber ? ` (${inspectionLabel(latest.id)})` : ""} and appear in its report as 2.2 Bracket Record Review. Acceptable reviews are then archived; reviews requiring rectification stay in 2.2 until their Bracket Action List item is closed.
      </p>

      <Card className="p-3 mb-3">
        <div className="grid grid-cols-2 sm:grid-cols-[1fr_1.2fr_1.3fr] gap-2 items-end">
          <div>
            <Label className="text-xs">Level</Label>
            <Input value={draft.level} onChange={(e) => setDraft({ ...draft, level: e.target.value.replace(/\s+/g, "") })} className="font-mono text-center" data-testid="review-new-level" />
          </div>
          <div>
            <Label className="text-xs">Bracket</Label>
            <Input value={draft.bracket} placeholder="e.g. 14-B" onChange={(e) => setDraft({ ...draft, bracket: e.target.value.toUpperCase() })} className="font-mono text-center" autoComplete="off" data-testid="review-new-bracket" />
          </div>
          <div className="col-span-2 sm:col-span-1">
            <Label className="text-xs">Date submitted</Label>
            <Input type="date" className="h-9 min-w-0 w-full px-2" value={draft.dateSubmitted} onChange={(e) => setDraft({ ...draft, dateSubmitted: e.target.value })} data-testid="review-new-date" />
          </div>
        </div>
        <div className="flex flex-wrap items-end gap-2 mt-2">
          <div className="flex-1 min-w-[16rem]">
            <Label className="text-xs">Outcome</Label>
            <div className="flex rounded-md border overflow-hidden h-9">
              {REVIEW_OUTCOMES.map((o) => (
                <button key={o} type="button"
                  className={`flex-1 px-3 text-sm font-medium whitespace-nowrap ${draft.outcome === o ? OUTCOME_ON[o] : "bg-background text-muted-foreground hover:bg-accent/60"}`}
                  onClick={() => setDraft({ ...draft, outcome: o })}
                  data-testid={`review-new-outcome-${o === "Acceptable" ? "acceptable" : "rectify"}`}>
                  {o}
                </button>
              ))}
            </div>
          </div>
          <Button size="sm" className="h-9" disabled={!canAdd || addMutation.isPending} onClick={() => addMutation.mutate()} data-testid="review-add">
            <Plus className="w-4 h-4 mr-1" /> Add
          </Button>
        </div>
      </Card>

      {awaiting > 0 && (
        <p className="text-xs text-red-600 mb-2" data-testid="review-awaiting">
          {awaiting} bracket{awaiting === 1 ? "" : "s"} requiring rectification not yet in the Bracket Action List.
        </p>
      )}

      {rows.length === 0 || (visible.length === 0 && archived.length === 0) ? (
        <p className="text-sm text-muted-foreground">No reviews yet.</p>
      ) : (
        <div className="space-y-2">
          {visible.map((r) => {
            const linked = r.outcome === "Requires rectification" ? actionFor(r) : null;
            return (
              <Card key={r.id} className="p-3" data-testid={`review-row-${r.id}`}>
                <div className="flex items-center justify-between gap-2 mb-2">
                  <span className="text-xs text-muted-foreground" data-testid={`review-meta-${r.id}`}>
                    Recorded {recordedOn(r.createdAt)}{inspectionLabel(r.reportId) ? ` · ${inspectionLabel(r.reportId)}` : ""}
                    {!isCurrent(r) ? " · Archived" : r.reportId !== latest?.id && r.reportId != null ? " · Carried until closed" : ""}
                  </span>
                  <Button size="icon" variant="ghost" className="h-8 w-8" aria-label="Delete review"
                    onClick={() => { if (confirm(`Delete the review of Level ${r.level}, Bracket ${r.bracket}?`)) deleteMutation.mutate(r.id); }}>
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-[1fr_1.2fr_1.3fr] gap-2 items-end">
                  <div>
                    <Label className="text-xs">Level</Label>
                    <Input defaultValue={r.level} className="h-9 font-mono text-center" key={`l-${r.id}-${r.level}`}
                      onBlur={(e) => { const v = e.target.value.replace(/\s+/g, ""); if (v !== r.level) patch(r.id, { level: v }); }} data-testid={`review-level-${r.id}`} />
                  </div>
                  <div>
                    <Label className="text-xs">Bracket</Label>
                    <Input defaultValue={r.bracket} className="h-9 font-mono text-center" key={`b-${r.id}-${r.bracket}`}
                      onBlur={(e) => { const v = e.target.value.toUpperCase().trim(); if (v !== r.bracket) patch(r.id, { bracket: v }); }} data-testid={`review-bracket-${r.id}`} />
                  </div>
                  <div className="col-span-2 sm:col-span-1">
                    <Label className="text-xs">Date submitted</Label>
                    <Input type="date" className="h-9 min-w-0 w-full px-2" value={r.dateSubmitted || ""} onChange={(e) => patch(r.id, { dateSubmitted: e.target.value })} data-testid={`review-date-${r.id}`} />
                  </div>
                  <div className="col-span-2 sm:col-span-3">
                    <Label className="text-xs">Outcome</Label>
                    <div className="flex rounded-md border overflow-hidden h-9">
                      {REVIEW_OUTCOMES.map((o) => (
                        <button key={o} type="button"
                          className={`flex-1 px-2 text-sm font-medium whitespace-nowrap ${r.outcome === o ? OUTCOME_ON[o] : "bg-background text-muted-foreground hover:bg-accent/60"}`}
                          onClick={() => { if (r.outcome !== o) patch(r.id, { outcome: o }); }}
                          data-testid={`review-outcome-${r.id}-${o === "Acceptable" ? "acceptable" : "rectify"}`}>
                          {o}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
                {r.outcome === "Requires rectification" && (
                  <div className="mt-2 flex items-center justify-end">
                    {linked ? (
                      <span className="text-xs text-green-700 flex items-center gap-1" data-testid={`review-linked-${r.id}`}>
                        <CheckCircle2 className="w-3.5 h-3.5" /> In Bracket Action List as {uidOf(linked)} ({linked.status})
                      </span>
                    ) : (
                      <Button size="sm" variant="outline" className="h-8" disabled={addActionMutation.isPending} onClick={() => addToActionList(r)} data-testid={`review-to-action-${r.id}`}>
                        <ListPlus className="w-4 h-4 mr-1" /> Add to Bracket Action List
                      </Button>
                    )}
                  </div>
                )}
              </Card>
            );
          })}
          {current.length === 0 && !showArchived && (
            <p className="text-sm text-muted-foreground">No current reviews. Earlier reviews are archived.</p>
          )}
          {archived.length > 0 && (
            <Button variant="ghost" size="sm" className="w-full" onClick={() => setShowArchived(!showArchived)} data-testid="review-show-archived">
              {showArchived ? "Hide archived reviews" : `Show ${archived.length} archived review${archived.length === 1 ? "" : "s"}`}
            </Button>
          )}
        </div>
      )}
    </div>
  );
}
