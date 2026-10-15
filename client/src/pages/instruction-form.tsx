// WP6: Project Instruction form. A project-wide instruction to the contractor that applies
// to one or more work types (not a location) for the life of the project. It has a
// project ID (PI-01…), a responsibility, By When fixed as "Ongoing", and no verification.
// Withdrawn instructions stay listed in reports, marked "Withdrawn".
import { useEffect, useMemo, useRef, useState } from "react";
import { useParams, useLocation, Link } from "wouter";
import { useQuery } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "@/components/ui/select";
import { useToast } from "@/hooks/use-toast";
import { ArrowLeft, Camera, Loader2, Trash2, History, Ban, RotateCcw } from "lucide-react";
import { usesRecordWording } from "@/lib/wording";
import { localDate } from "@/lib/project-instructions";
import type { Project, Report } from "@shared/schema";

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

// Same parties as the record form's By Whom list.
const PERSON_ROLES = ["Consultant", "Contractor", "Client", "Project Manager", "Facility Manager"];
const NEW_PARTY = "__new_party__";

type Instruction = {
  id: number;
  projectId: number;
  reportId: number | null;
  ref: string;
  workTypes: string[];
  workTypeLabels?: string[];
  observation: string;
  actionRequired: string;
  responsibility: string;
  status: "active" | "withdrawn";
  dateIssued: string;
  dateWithdrawn: string | null;
  photos: { filename: string; caption?: string | null }[];
  history?: { id: number; field: string; oldValue: string | null; newValue: string | null; changedAt: string }[];
};

const FIELD_NAMES: Record<string, string> = {
  observation: "Observation",
  actionRequired: "Action required",
  responsibility: "Responsibility",
  workTypes: "Work types",
  dateIssued: "Date issued",
  status: "Status",
};

// Keyed by the instruction so moving from one instruction to another (or to "new")
// starts with a fresh form instead of the previous instruction's values.
export default function InstructionFormPage() {
  const { instructionId } = useParams<{ instructionId?: string }>();
  return <InstructionForm key={instructionId || "new"} />;
}

function InstructionForm() {
  const { projectId, reportId, instructionId } = useParams<{ projectId: string; reportId: string; instructionId?: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const isEdit = !!instructionId && instructionId !== "new";
  const backHref = `/projects/${projectId}/reports/${reportId}`;

  // ---- data (all hooks above any early return)
  const { data: project } = useQuery<Project>({ queryKey: [`/api/projects/${projectId}`] });
  const { data: report } = useQuery<Report>({ queryKey: [`/api/reports/${reportId}`] });
  const { data: globalSettings } = useQuery<{ workTypes: { code: string; label: string }[] }>({ queryKey: ["/api/global-settings"] });
  const { data: allInstructions = [] } = useQuery<Instruction[]>({ queryKey: [`/api/projects/${projectId}/instructions`] });
  const { data: existing, isLoading: loadingExisting } = useQuery<Instruction>({
    queryKey: [`/api/instructions/${instructionId}`],
    enabled: isEdit,
  });

  const [workTypes, setWorkTypes] = useState<string[]>([]);
  const [observation, setObservation] = useState("");
  const [actionRequired, setActionRequired] = useState("");
  const [responsibility, setResponsibility] = useState("");
  const [addingParty, setAddingParty] = useState(false);
  const [newParty, setNewParty] = useState("");
  const [dateIssued, setDateIssued] = useState("");
  const [saving, setSaving] = useState(false);
  const [uploading, setUploading] = useState(false);
  const loadedRef = useRef(false);
  const fileRef = useRef<HTMLInputElement>(null);

  // Load the saved instruction once.
  useEffect(() => {
    if (!existing || loadedRef.current) return;
    loadedRef.current = true;
    setWorkTypes(existing.workTypes);
    setObservation(existing.observation);
    setActionRequired(existing.actionRequired);
    setResponsibility(existing.responsibility);
    setDateIssued(existing.dateIssued);
  }, [existing]);

  // A new instruction is dated to the inspection it is issued in.
  useEffect(() => {
    if (isEdit || dateIssued || !report) return;
    const d = (report as any)?.inspectionDate;
    const t = new Date();
    const today = `${t.getFullYear()}-${String(t.getMonth() + 1).padStart(2, "0")}-${String(t.getDate()).padStart(2, "0")}`;
    setDateIssued(/^\d{4}-\d{2}-\d{2}$/.test(d || "") ? d : today);
  }, [isEdit, report, dateIssued]);

  // Work types: the project's primary ones first, then its custom ones, then the rest.
  const workTypeOptions = useMemo(() => {
    const parse = (raw: unknown) => { try { const v = JSON.parse(String(raw || "[]")); return Array.isArray(v) ? v : []; } catch { return []; } };
    const global: { code: string; label: string }[] = globalSettings?.workTypes || [];
    const custom: { code: string; label: string }[] = parse((project as any)?.customWorkTypes);
    const primary: string[] = parse((project as any)?.primaryWorkTypes);
    const all = new Map<string, string>();
    for (const w of [...custom, ...global]) if (w?.code && !all.has(w.code)) all.set(w.code, w.label || w.code);
    for (const c of workTypes) if (!all.has(c)) all.set(c, c);
    const first = [...primary, ...custom.map((w) => w.code)].filter((c, i, a) => all.has(c) && a.indexOf(c) === i);
    const rest = Array.from(all.keys()).filter((c) => !first.includes(c));
    return { first: first.map((c) => ({ code: c, label: all.get(c)! })), rest: rest.map((c) => ({ code: c, label: all.get(c)! })) };
  }, [globalSettings, project, workTypes]);
  const [showAllWorkTypes, setShowAllWorkTypes] = useState(false);

  // Parties: the usual list plus any already used on this project's instructions.
  const parties = useMemo(() => {
    const used = allInstructions.map((i) => i.responsibility).filter(Boolean);
    const list = [...PERSON_ROLES];
    for (const p of [...used, responsibility]) if (p && !list.includes(p)) list.push(p);
    return list;
  }, [allInstructions, responsibility]);

  const enabled = usesRecordWording(project as any);
  const withdrawn = existing?.status === "withdrawn";
  const dirty = isEdit && !!existing && (
    JSON.stringify(workTypes) !== JSON.stringify(existing.workTypes) || observation !== existing.observation ||
    actionRequired !== existing.actionRequired || responsibility !== existing.responsibility || dateIssued !== existing.dateIssued
  );

  // ---- early returns (after every hook)
  if (!project || (isEdit && loadingExisting)) {
    return <div className="flex items-center justify-center py-24"><Loader2 className="w-5 h-5 animate-spin text-muted-foreground" /></div>;
  }
  if (!enabled) {
    return (
      <div className="max-w-2xl mx-auto p-4">
        <p className="text-sm text-muted-foreground">Project Instructions are only available in projects created after this feature was added.</p>
        <Link href={backHref}><Button variant="ghost" className="mt-4"><ArrowLeft className="w-4 h-4 mr-2" />Back to Report</Button></Link>
      </div>
    );
  }
  if (isEdit && !existing) {
    return <div className="max-w-2xl mx-auto p-4 text-sm text-muted-foreground">Instruction not found.</div>;
  }

  const toggleWorkType = (code: string) =>
    setWorkTypes((prev) => (prev.includes(code) ? prev.filter((c) => c !== code) : [...prev, code]));

  const refresh = () => {
    queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/instructions`] });
    if (isEdit) queryClient.invalidateQueries({ queryKey: [`/api/instructions/${instructionId}`] });
  };

  const save = async () => {
    const party = addingParty ? newParty.trim() : responsibility;
    if (workTypes.length === 0) return toast({ title: "Choose at least one work type", variant: "destructive" });
    if (!actionRequired.trim()) return toast({ title: "Action required is needed", variant: "destructive" });
    if (!party) return toast({ title: "Choose who is responsible", variant: "destructive" });
    setSaving(true);
    try {
      const body = { workTypes, observation, actionRequired, responsibility: party, dateIssued };
      const res = isEdit
        ? await apiRequest("PATCH", `/api/instructions/${instructionId}`, body)
        : await apiRequest("POST", `/api/projects/${projectId}/instructions`, { ...body, reportId: Number(reportId) });
      const saved: Instruction = await res.json();
      setResponsibility(party); setAddingParty(false); setNewParty("");
      refresh();
      toast({ title: isEdit ? `${saved.ref} updated` : `Project Instruction ${saved.ref} created` });
      if (!isEdit) navigate(`/projects/${projectId}/reports/${reportId}/instructions/${saved.id}`);
      else loadedRef.current = false;
    } catch (e: any) {
      toast({ title: e?.message || "Could not save", variant: "destructive" });
    } finally {
      setSaving(false);
    }
  };

  const setStatus = async (status: "active" | "withdrawn") => {
    if (status === "withdrawn" && !window.confirm(`Withdraw ${existing!.ref}? It stays listed in reports, marked Withdrawn.`)) return;
    try {
      await apiRequest("PATCH", `/api/instructions/${instructionId}`, { status });
      refresh();
      toast({ title: status === "withdrawn" ? `${existing!.ref} withdrawn` : `${existing!.ref} reinstated` });
    } catch (e: any) {
      toast({ title: e?.message || "Could not update", variant: "destructive" });
    }
  };

  const remove = async () => {
    if (!window.confirm(`Delete ${existing!.ref} permanently? To keep it in reports, withdraw it instead.`)) return;
    try {
      await apiRequest("DELETE", `/api/instructions/${instructionId}`);
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${projectId}/instructions`] });
      toast({ title: `${existing!.ref} deleted` });
      navigate(backHref);
    } catch (e: any) {
      toast({ title: e?.message || "Could not delete", variant: "destructive" });
    }
  };

  const uploadPhoto = async (file: File) => {
    setUploading(true);
    try {
      const fd = new FormData();
      fd.append("photo", file);
      const res = await fetch(`${API_BASE}/api/instructions/${instructionId}/photos`, { method: "POST", body: fd });
      if (!res.ok) throw new Error((await res.json().catch(() => ({})))?.message || "Upload failed");
      refresh();
    } catch (e: any) {
      toast({ title: e?.message || "Upload failed", variant: "destructive" });
    } finally {
      setUploading(false);
      if (fileRef.current) fileRef.current.value = "";
    }
  };

  const removePhoto = async (filename: string) => {
    if (!window.confirm("Remove this photo?")) return;
    try {
      await apiRequest("DELETE", `/api/instructions/${instructionId}/photos/${encodeURIComponent(filename)}`);
      refresh();
    } catch (e: any) {
      toast({ title: e?.message || "Could not remove", variant: "destructive" });
    }
  };

  const chip = (w: { code: string; label: string }) => {
    const on = workTypes.includes(w.code);
    return (
      <button
        key={w.code}
        type="button"
        onClick={() => toggleWorkType(w.code)}
        aria-pressed={on}
        data-testid={`chip-worktype-${w.code}`}
        className={`flex flex-col items-start p-2.5 rounded-lg border text-left transition-colors ${on ? "border-primary bg-primary/10" : "hover:bg-accent/60"}`}
      >
        <span className="font-mono font-semibold text-sm">{w.code}</span>
        <span className="text-xs text-muted-foreground">{w.label}</span>
      </button>
    );
  };

  return (
    <div className="max-w-2xl mx-auto p-4 pb-24 space-y-5">
      <Link href={backHref}>
        <button type="button" className="flex items-center text-sm text-muted-foreground hover:text-foreground">
          <ArrowLeft className="w-4 h-4 mr-1.5" /> Back to Report
        </button>
      </Link>

      <div className="flex flex-wrap items-center gap-2">
        <h1 className="text-xl font-semibold">{isEdit ? "Edit Project Instruction" : "New Project Instruction"}</h1>
        {withdrawn && (
          <Badge variant="secondary" className="bg-slate-100 text-slate-700 border border-slate-300" data-testid="badge-withdrawn">
            Withdrawn {localDate(existing?.dateWithdrawn)}
          </Badge>
        )}
      </div>

      <Card className="p-4 bg-accent/30">
        <div className="flex items-center justify-between gap-3">
          <span className="text-sm font-medium">Instruction ID</span>
          <span className="font-mono font-semibold" data-testid="text-instruction-ref">{existing?.ref || "Assigned on save"}</span>
        </div>
        <p className="text-xs text-muted-foreground mt-2">
          Applies to all similar works across the whole project, not to one location. It is listed in every report
          from now on, in Section 3. It has no verification and no close-out date.
        </p>
      </Card>

      <div>
        <Label>Work types</Label>
        <p className="text-xs text-muted-foreground mb-2">Choose one or more.</p>
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2" data-testid="worktype-chips">
          {(workTypeOptions.first.length ? workTypeOptions.first : workTypeOptions.rest).map(chip)}
          {workTypeOptions.first.length > 0 && (showAllWorkTypes || workTypeOptions.rest.some((w) => workTypes.includes(w.code))) &&
            workTypeOptions.rest.filter((w) => showAllWorkTypes || workTypes.includes(w.code)).map(chip)}
        </div>
        {workTypeOptions.first.length > 0 && !showAllWorkTypes && (
          <Button type="button" variant="ghost" size="sm" className="mt-1" onClick={() => setShowAllWorkTypes(true)} data-testid="button-more-worktypes">
            Show all work types
          </Button>
        )}
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
        <div>
          <Label htmlFor="dateIssued">Date issued</Label>
          <Input id="dateIssued" type="date" value={dateIssued} onChange={(e) => setDateIssued(e.target.value)} data-testid="input-date-issued" />
        </div>
        <div>
          <Label>By When</Label>
          <div className="h-9 px-3 flex items-center rounded-md border bg-muted/40 text-sm" data-testid="text-by-when">Ongoing</div>
        </div>
      </div>

      <div>
        <Label htmlFor="observation">Observation</Label>
        <Textarea id="observation" rows={3} value={observation} onChange={(e) => setObservation(e.target.value)}
          placeholder="What was observed that prompts this instruction" data-testid="input-observation" />
      </div>

      <div>
        <Label htmlFor="actionRequired">Action required</Label>
        <Textarea id="actionRequired" rows={3} value={actionRequired} onChange={(e) => setActionRequired(e.target.value)}
          placeholder="The instruction to the contractor for all similar works" data-testid="input-action-required" />
      </div>

      <div>
        <Label>Responsibility</Label>
        {!addingParty ? (
          <Select value={responsibility} onValueChange={(v) => { if (v === NEW_PARTY) { setAddingParty(true); } else setResponsibility(v); }}>
            <SelectTrigger data-testid="select-responsibility"><SelectValue placeholder="Select..." /></SelectTrigger>
            <SelectContent>
              {parties.map((p) => <SelectItem key={p} value={p}>{p}</SelectItem>)}
              <SelectItem value={NEW_PARTY}>+ Add a new party…</SelectItem>
            </SelectContent>
          </Select>
        ) : (
          <div className="flex gap-2">
            <Input autoFocus value={newParty} onChange={(e) => setNewParty(e.target.value)} placeholder="e.g. Waterproofing subcontractor" data-testid="input-new-party" />
            <Button type="button" variant="ghost" onClick={() => { setAddingParty(false); setNewParty(""); }}>Cancel</Button>
          </div>
        )}
      </div>

      <div className="flex flex-wrap gap-2">
        <Button onClick={save} disabled={saving || (isEdit && !dirty && !addingParty)} data-testid="button-save-instruction">
          {saving ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : null}
          {isEdit ? "Save changes" : "Create instruction"}
        </Button>
        {isEdit && (withdrawn ? (
          <Button variant="outline" onClick={() => setStatus("active")} data-testid="button-reinstate"><RotateCcw className="w-4 h-4 mr-2" />Reinstate</Button>
        ) : (
          <Button variant="outline" onClick={() => setStatus("withdrawn")} data-testid="button-withdraw"><Ban className="w-4 h-4 mr-2" />Withdraw</Button>
        ))}
        {isEdit && (
          <Button variant="ghost" className="text-destructive" onClick={remove} data-testid="button-delete-instruction"><Trash2 className="w-4 h-4 mr-2" />Delete</Button>
        )}
      </div>

      {isEdit && (
        <Card className="p-4 space-y-3">
          <div className="flex items-center justify-between">
            <h3 className="text-sm font-medium">Photos (optional)</h3>
            <Button type="button" variant="outline" size="sm" disabled={uploading} onClick={() => fileRef.current?.click()} data-testid="button-add-instruction-photo">
              {uploading ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : <Camera className="w-4 h-4 mr-2" />} Add photo
            </Button>
            <input ref={fileRef} type="file" accept="image/*" className="hidden" onChange={(e) => { const f = e.target.files?.[0]; if (f) uploadPhoto(f); }} data-testid="input-instruction-photo" />
          </div>
          {existing!.photos.length === 0 ? (
            <p className="text-xs text-muted-foreground">No photos. Photos stay with the instruction; the report table is text only.</p>
          ) : (
            <div className="grid grid-cols-3 gap-2">
              {existing!.photos.map((p) => (
                <div key={p.filename} className="relative">
                  <img src={`${API_BASE}/api/uploads/${p.filename}`} alt="" className="w-full aspect-square object-cover rounded" />
                  <button type="button" onClick={() => removePhoto(p.filename)} className="absolute top-1 right-1 bg-background/90 rounded p-1" aria-label="Remove photo">
                    <Trash2 className="w-3.5 h-3.5" />
                  </button>
                </div>
              ))}
            </div>
          )}
        </Card>
      )}

      {isEdit && (
        <Card className="p-4">
          <h3 className="text-sm font-medium flex items-center gap-1.5 mb-2"><History className="w-4 h-4" /> Edit history</h3>
          {(existing!.history ?? []).length === 0 ? (
            <p className="text-xs text-muted-foreground">No changes since it was issued.</p>
          ) : (
            <ul className="space-y-2" data-testid="instruction-history">
              {existing!.history!.map((h) => (
                <li key={h.id} className="text-xs">
                  <span className="text-muted-foreground">{localDate(h.changedAt)} · {FIELD_NAMES[h.field] || h.field}</span>
                  <div className="break-words"><span className="line-through text-muted-foreground">{h.oldValue || "—"}</span> → {h.newValue || "—"}</div>
                </li>
              ))}
            </ul>
          )}
        </Card>
      )}
    </div>
  );
}
