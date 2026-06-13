import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import {
  Dialog, DialogContent, DialogHeader, DialogTitle,
} from "@/components/ui/dialog";
import {
  Select, SelectContent, SelectItem, SelectTrigger, SelectValue,
} from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { Plus, Trash2, Upload, ArrowUp, ArrowDown } from "lucide-react";
import { useState } from "react";
import type { NarrativeHold, ProgramSchedule, StageProgressMap } from "@shared/schema";

const AUDIENCES = [
  { value: "both", label: "Both" },
  { value: "contractor", label: "Contractor only" },
  { value: "client", label: "Client only" },
];
const HOLD_STATUSES = ["Active", "Lifted", "For information"];
const STAGE_STATUSES = ["Not started", "Underway", "Complete"];

type Figure = { filename: string; caption: string };
type Stage = { stageName: string; status: string };

// Upload a status image (program/stage/figure), returns the stored filename.
async function uploadStatusImage(projectId: string, file: File): Promise<string> {
  const fd = new FormData();
  fd.append("image", file);
  const res = await fetch(`/api/projects/${projectId}/status-images`, { method: "POST", body: fd });
  if (!res.ok) throw new Error("Image upload failed");
  const data = await res.json();
  return data.filename as string;
}

function AudienceSelect({ value, onChange }: { value: string; onChange: (v: string) => void }) {
  return (
    <Select value={value} onValueChange={onChange}>
      <SelectTrigger className="h-9"><SelectValue /></SelectTrigger>
      <SelectContent>
        {AUDIENCES.map((a) => <SelectItem key={a.value} value={a.value}>{a.label}</SelectItem>)}
      </SelectContent>
    </Select>
  );
}

export default function ProjectStatusSection({ projectId }: { projectId: string }) {
  const { toast } = useToast();

  const { data: holds = [] } = useQuery<NarrativeHold[]>({ queryKey: [`/api/projects/${projectId}/narrative-holds`] });
  const { data: program } = useQuery<ProgramSchedule | null>({ queryKey: [`/api/projects/${projectId}/program-schedule`] });
  const { data: stageMap } = useQuery<StageProgressMap | null>({ queryKey: [`/api/projects/${projectId}/stage-progress-map`] });

  // ---- Narrative/Hold modal ----
  const [holdOpen, setHoldOpen] = useState(false);
  const [editingHold, setEditingHold] = useState<NarrativeHold | null>(null);
  const blankHold = { title: "", body: "", status: "Active", dateRaised: "", dateLifted: "", audience: "both", sortOrder: 0 };
  const [holdForm, setHoldForm] = useState<Record<string, any>>(blankHold);
  const [holdFigures, setHoldFigures] = useState<Figure[]>([]);

  const openHoldModal = (h?: NarrativeHold) => {
    if (h) {
      setEditingHold(h);
      setHoldForm({
        title: h.title || "", body: h.body || "", status: h.status || "Active",
        dateRaised: h.dateRaised || "", dateLifted: h.dateLifted || "",
        audience: h.audience || "both", sortOrder: h.sortOrder ?? 0,
      });
      let figs: Figure[] = [];
      try { figs = JSON.parse(h.figures || "[]"); } catch {}
      setHoldFigures(figs);
    } else {
      setEditingHold(null);
      setHoldForm({ ...blankHold, sortOrder: holds.length });
      setHoldFigures([]);
    }
    setHoldOpen(true);
  };

  const saveHoldMutation = useMutation({
    mutationFn: async () => {
      const payload = { ...holdForm, figures: JSON.stringify(holdFigures) };
      if (editingHold) {
        const res = await apiRequest("PATCH", `/api/narrative-holds/${editingHold.id}`, payload);
        return res.json();
      }
      const res = await apiRequest("POST", `/api/projects/${projectId}/narrative-holds`, payload);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/narrative-holds`] });
      setHoldOpen(false);
      toast({ title: editingHold ? "Narrative updated" : "Narrative added" });
    },
    onError: (e: any) => toast({ title: e?.message || "Save failed", variant: "destructive" }),
  });

  const deleteHoldMutation = useMutation({
    mutationFn: async (hid: number) => { await apiRequest("DELETE", `/api/narrative-holds/${hid}`); },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/narrative-holds`] });
      toast({ title: "Narrative deleted" });
    },
  });

  const addHoldFigure = async (file: File) => {
    try {
      const filename = await uploadStatusImage(projectId, file);
      setHoldFigures((prev) => [...prev, { filename, caption: "" }]);
    } catch (e: any) { toast({ title: e?.message || "Upload failed", variant: "destructive" }); }
  };

  // ---- Program form ----
  const [progForm, setProgForm] = useState<Record<string, any> | null>(null);
  const progState = progForm ?? {
    programImageFilename: program?.programImageFilename || "",
    asAtDate: program?.asAtDate || "",
    varianceText: program?.varianceText || "",
    projectedCompletion: program?.projectedCompletion || "",
    statusNarrative: program?.statusNarrative || "",
    audience: program?.audience || "both",
  };
  const setProg = (patch: Record<string, any>) => setProgForm({ ...progState, ...patch });

  const saveProgramMutation = useMutation({
    mutationFn: async () => {
      const res = await apiRequest("PUT", `/api/projects/${projectId}/program-schedule`, progState);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/program-schedule`] });
      setProgForm(null);
      toast({ title: "Program saved" });
    },
    onError: (e: any) => toast({ title: e?.message || "Save failed", variant: "destructive" }),
  });

  // ---- Stage map form ----
  const [stageForm, setStageForm] = useState<{ planImageFilename: string; audience: string; stages: Stage[] } | null>(null);
  const stageState = stageForm ?? {
    planImageFilename: stageMap?.planImageFilename || "",
    audience: stageMap?.audience || "both",
    stages: (() => { try { return JSON.parse(stageMap?.stages || "[]"); } catch { return []; } })() as Stage[],
  };
  const setStage = (patch: Partial<typeof stageState>) => setStageForm({ ...stageState, ...patch });

  const saveStageMutation = useMutation({
    mutationFn: async () => {
      const res = await apiRequest("PUT", `/api/projects/${projectId}/stage-progress-map`, {
        planImageFilename: stageState.planImageFilename,
        audience: stageState.audience,
        stages: JSON.stringify(stageState.stages),
      });
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/stage-progress-map`] });
      setStageForm(null);
      toast({ title: "Stage map saved" });
    },
    onError: (e: any) => toast({ title: e?.message || "Save failed", variant: "destructive" }),
  });

  const moveStage = (idx: number, dir: -1 | 1) => {
    const next = [...stageState.stages];
    const j = idx + dir;
    if (j < 0 || j >= next.length) return;
    [next[idx], next[j]] = [next[j], next[idx]];
    setStage({ stages: next });
  };

  return (
    <div className="mt-10">
      <h2 className="text-lg font-semibold mb-4">Project Status</h2>

      {/* ===== Narratives / Holds ===== */}
      <div className="mb-8">
        <div className="flex items-center justify-between mb-3">
          <h3 className="text-base font-medium">Narratives / Holds</h3>
          <Button size="sm" onClick={() => openHoldModal()} data-testid="add-narrative">
            <Plus className="w-4 h-4 mr-2" /> Add Narrative
          </Button>
        </div>
        {holds.length === 0 ? (
          <p className="text-sm text-muted-foreground">No narratives or holds yet.</p>
        ) : (
          <div className="space-y-2">
            {holds.map((h) => (
              <Card key={h.id} className="p-3 flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="font-medium">{h.title || "(untitled)"}</span>
                    <span className="text-xs px-1.5 py-0.5 rounded bg-muted">{h.status}</span>
                    <span className="text-xs px-1.5 py-0.5 rounded bg-accent">{h.audience}</span>
                  </div>
                  {h.body && <p className="text-sm text-muted-foreground line-clamp-2 mt-1">{h.body}</p>}
                </div>
                <div className="flex gap-1 shrink-0">
                  <Button size="sm" variant="ghost" onClick={() => openHoldModal(h)}>Edit</Button>
                  <Button size="sm" variant="ghost" onClick={() => { if (confirm("Delete this narrative?")) deleteHoldMutation.mutate(h.id); }}>
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
              </Card>
            ))}
          </div>
        )}
      </div>

      {/* ===== Program / Schedule ===== */}
      <div className="mb-8">
        <h3 className="text-base font-medium mb-3">Program / Schedule</h3>
        <Card className="p-4 space-y-3">
          <div>
            <Label>Program image (Gantt)</Label>
            <div className="flex items-center gap-2 mt-1">
              <Input type="file" accept="image/*" onChange={async (e) => {
                const f = e.target.files?.[0]; if (!f) return;
                try { const fn = await uploadStatusImage(projectId, f); setProg({ programImageFilename: fn }); }
                catch (err: any) { toast({ title: err?.message || "Upload failed", variant: "destructive" }); }
              }} />
              {progState.programImageFilename && (
                <img src={`/api/uploads/${progState.programImageFilename}`} alt="program" className="h-12 rounded border" />
              )}
            </div>
          </div>
          <div className="grid grid-cols-2 gap-3">
            <div>
              <Label>As-at date</Label>
              <Input value={progState.asAtDate} onChange={(e) => setProg({ asAtDate: e.target.value })} placeholder="e.g. 2026-06-01" />
            </div>
            <div>
              <Label>Projected completion</Label>
              <Input value={progState.projectedCompletion} onChange={(e) => setProg({ projectedCompletion: e.target.value })} />
            </div>
          </div>
          <div>
            <Label>Variance</Label>
            <Input value={progState.varianceText} onChange={(e) => setProg({ varianceText: e.target.value })} placeholder="e.g. behind 2 weeks" />
          </div>
          <div>
            <Label>Status narrative</Label>
            <Textarea value={progState.statusNarrative} onChange={(e) => setProg({ statusNarrative: e.target.value })} rows={2} />
          </div>
          <div>
            <Label>Audience</Label>
            <AudienceSelect value={progState.audience} onChange={(v) => setProg({ audience: v })} />
          </div>
          <Button onClick={() => saveProgramMutation.mutate()} disabled={saveProgramMutation.isPending}>
            {saveProgramMutation.isPending ? "Saving..." : "Save Program"}
          </Button>
        </Card>
      </div>

      {/* ===== Stage Progress Map ===== */}
      <div className="mb-8">
        <h3 className="text-base font-medium mb-3">Stage Progress Map</h3>
        <Card className="p-4 space-y-3">
          <div>
            <Label>Plan image</Label>
            <div className="flex items-center gap-2 mt-1">
              <Input type="file" accept="image/*" onChange={async (e) => {
                const f = e.target.files?.[0]; if (!f) return;
                try { const fn = await uploadStatusImage(projectId, f); setStage({ planImageFilename: fn }); }
                catch (err: any) { toast({ title: err?.message || "Upload failed", variant: "destructive" }); }
              }} />
              {stageState.planImageFilename && (
                <img src={`/api/uploads/${stageState.planImageFilename}`} alt="plan" className="h-12 rounded border" />
              )}
            </div>
          </div>
          <div>
            <div className="flex items-center justify-between mb-1">
              <Label>Stages</Label>
              <Button size="sm" variant="outline" onClick={() => setStage({ stages: [...stageState.stages, { stageName: "", status: "Not started" }] })}>
                <Plus className="w-4 h-4 mr-1" /> Add Stage
              </Button>
            </div>
            <div className="space-y-2">
              {stageState.stages.map((s, idx) => (
                <div key={idx} className="flex items-center gap-2">
                  <Input
                    value={s.stageName}
                    placeholder="Stage name"
                    onChange={(e) => {
                      const next = [...stageState.stages]; next[idx] = { ...next[idx], stageName: e.target.value }; setStage({ stages: next });
                    }}
                  />
                  <Select value={s.status} onValueChange={(v) => {
                    const next = [...stageState.stages]; next[idx] = { ...next[idx], status: v }; setStage({ stages: next });
                  }}>
                    <SelectTrigger className="w-40 h-9"><SelectValue /></SelectTrigger>
                    <SelectContent>
                      {STAGE_STATUSES.map((st) => <SelectItem key={st} value={st}>{st}</SelectItem>)}
                    </SelectContent>
                  </Select>
                  <Button size="icon" variant="ghost" onClick={() => moveStage(idx, -1)}><ArrowUp className="w-4 h-4" /></Button>
                  <Button size="icon" variant="ghost" onClick={() => moveStage(idx, 1)}><ArrowDown className="w-4 h-4" /></Button>
                  <Button size="icon" variant="ghost" onClick={() => setStage({ stages: stageState.stages.filter((_, i) => i !== idx) })}>
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
              ))}
            </div>
          </div>
          <div>
            <Label>Audience</Label>
            <AudienceSelect value={stageState.audience} onChange={(v) => setStage({ audience: v })} />
          </div>
          <Button onClick={() => saveStageMutation.mutate()} disabled={saveStageMutation.isPending}>
            {saveStageMutation.isPending ? "Saving..." : "Save Stage Map"}
          </Button>
        </Card>
      </div>

      {/* Narrative/Hold modal */}
      <Dialog open={holdOpen} onOpenChange={setHoldOpen}>
        <DialogContent className="max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>{editingHold ? "Edit Narrative" : "Add Narrative"}</DialogTitle>
          </DialogHeader>
          <form onSubmit={(e) => { e.preventDefault(); saveHoldMutation.mutate(); }} className="space-y-3">
            <div>
              <Label>Title</Label>
              <Input value={holdForm.title} onChange={(e) => setHoldForm({ ...holdForm, title: e.target.value })} />
            </div>
            <div>
              <Label>Body</Label>
              <Textarea value={holdForm.body} onChange={(e) => setHoldForm({ ...holdForm, body: e.target.value })} rows={5} />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Label>Status</Label>
                <Select value={holdForm.status} onValueChange={(v) => setHoldForm({ ...holdForm, status: v })}>
                  <SelectTrigger className="h-9"><SelectValue /></SelectTrigger>
                  <SelectContent>
                    {HOLD_STATUSES.map((s) => <SelectItem key={s} value={s}>{s}</SelectItem>)}
                  </SelectContent>
                </Select>
              </div>
              <div>
                <Label>Audience</Label>
                <AudienceSelect value={holdForm.audience} onChange={(v) => setHoldForm({ ...holdForm, audience: v })} />
              </div>
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <Label>Date raised</Label>
                <Input type="date" value={holdForm.dateRaised} onChange={(e) => setHoldForm({ ...holdForm, dateRaised: e.target.value })} />
              </div>
              <div>
                <Label>Date lifted</Label>
                <Input type="date" value={holdForm.dateLifted} onChange={(e) => setHoldForm({ ...holdForm, dateLifted: e.target.value })} />
              </div>
            </div>
            <div>
              <Label>Sort order</Label>
              <Input type="number" value={holdForm.sortOrder} onChange={(e) => setHoldForm({ ...holdForm, sortOrder: Number(e.target.value) })} />
            </div>
            <div>
              <div className="flex items-center justify-between mb-1">
                <Label>Figures</Label>
                <label className="inline-flex items-center gap-1 text-sm cursor-pointer text-primary">
                  <Upload className="w-4 h-4" /> Add figure
                  <input type="file" accept="image/*" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) addHoldFigure(f); }} />
                </label>
              </div>
              <div className="space-y-2">
                {holdFigures.map((fig, idx) => (
                  <div key={idx} className="flex items-center gap-2">
                    <img src={`/api/uploads/${fig.filename}`} alt="figure" className="h-12 rounded border" />
                    <Input
                      value={fig.caption}
                      placeholder="Caption"
                      onChange={(e) => { const next = [...holdFigures]; next[idx] = { ...next[idx], caption: e.target.value }; setHoldFigures(next); }}
                    />
                    <Button size="icon" variant="ghost" type="button" onClick={() => setHoldFigures(holdFigures.filter((_, i) => i !== idx))}>
                      <Trash2 className="w-4 h-4" />
                    </Button>
                  </div>
                ))}
              </div>
            </div>
            <Button type="submit" className="w-full" disabled={saveHoldMutation.isPending}>
              {saveHoldMutation.isPending ? "Saving..." : "Save Narrative"}
            </Button>
          </form>
        </DialogContent>
      </Dialog>
    </div>
  );
}
