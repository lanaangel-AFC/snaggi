import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Link, useParams, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuSeparator,
  DropdownMenuSub,
  DropdownMenuSubContent,
  DropdownMenuSubTrigger,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";
import { useToast } from "@/hooks/use-toast";
import {
  ArrowLeft, Plus, FileText, Camera, ChevronRight, Trash2,
  MapPin, User, UserCheck, AlertTriangle, CheckCircle2, Archive,
  ChevronDown, FileDown, Eye, Settings, X, ImageDown, Share2, Copy, Link as LinkIcon
} from "lucide-react";
import {
  Dialog, DialogContent, DialogHeader, DialogTitle,
} from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";
import { Checkbox } from "@/components/ui/checkbox";
import type { Project, Report, Defect } from "@shared/schema";
import { getLocationDimensions } from "@shared/location";
import { buildReportTree, type ProfileKey } from "@/lib/report-tree";
import { renderDocx } from "@/lib/render-docx";
import { renderPdf } from "@/lib/render-pdf";
import { formatDefectLocation, formatReportDate } from "@/lib/render-helpers";
import { useState, useMemo } from "react";

const STANDARD_ELEVATIONS = [
  "North", "North East", "East", "South East",
  "South", "South West", "West", "North West",
];

// Trigger a browser download for an export Blob.
function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

export default function ReportDetail() {
  const { projectId, reportId } = useParams<{ projectId: string; reportId: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const [generating, setGenerating] = useState<string | null>(null);
  const [editOpen, setEditOpen] = useState(false);
  const [editForm, setEditForm] = useState<Record<string, string>>({});
  const [customElevation, setCustomElevation] = useState("");
  const [shareOpen, setShareOpen] = useState(false);
  const [shareRecipient, setShareRecipient] = useState("");
  const [generatedLink, setGeneratedLink] = useState("");
  // afcReference export guard: when afcReference is blank we block export and
  // prompt for it. pendingExport stores the {profile, format} to resume after save.
  const [afcGuardOpen, setAfcGuardOpen] = useState(false);
  const [afcGuardValue, setAfcGuardValue] = useState("");
  const [afcGuardSaving, setAfcGuardSaving] = useState(false);
  const [pendingExport, setPendingExport] = useState<{ profile: ProfileKey; format: "word" | "pdf"; audienceSuffix: string } | null>(null);

  const { data: project } = useQuery<Project>({
    queryKey: ["/api/projects", projectId],
  });

  const { data: report, isLoading: reportLoading } = useQuery<Report>({
    queryKey: ["/api/reports", reportId],
  });

  // Project's location dimensions — drives the single formatLocation() helper for the card.
  const cardDims = useMemo(() => getLocationDimensions((project as any)?.locationDimensions), [(project as any)?.locationDimensions]);
  // SVR Stage 2 — whether to hide the "(prev. {legacy_id})" alias on cards/register.
  const hideLegacyAliases = Boolean((project as any)?.hideLegacyAliases);

  const { data: defects, isLoading: defectsLoading } = useQuery<Defect[]>({
    queryKey: [`/api/reports/${reportId}/defects`],
  });

  const openEditDialog = () => {
    if (!report) return;
    setEditForm({
      inspectionNumber: report.inspectionNumber || "",
      inspectionDate: report.inspectionDate || "",
      revision: report.revision || "01",
      locationsCovered: report.locationsCovered || "",
      elevations: (report as any).elevations || project?.elevations || "[]",
      attendees: report.attendees || "[]",
    });
    setCustomElevation("");
    setEditOpen(true);
  };

  const updateReportMutation = useMutation({
    mutationFn: async (data: Record<string, string>) => {
      const res = await apiRequest("PATCH", `/api/reports/${reportId}`, data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/reports", reportId] });
      setEditOpen(false);
      toast({ title: "Report updated" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: async (defectId: number) => {
      await apiRequest("DELETE", `/api/defects/${defectId}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/defects`] });
      toast({ title: "Defect deleted" });
    },
  });

  // Fetch enriched report data for change-indicator badges
  const { data: reportData } = useQuery<any>({
    queryKey: [`/api/reports/${reportId}/report-data`],
    enabled: !!reportId,
  });

  // Build a map of defect id → events for badges
  const defectEventsMap = useMemo(() => {
    const map = new Map<number, { tag: "NEW" | "AMENDED" | "COMPLETED" | null; summary: string }>();
    if (!reportData?.defects) return map;
    for (const d of reportData.defects) {
      if (!d.events) continue;
      const { isNew, amendedFields } = d.events;
      if (isNew) {
        map.set(d.id, { tag: "NEW", summary: "" });
      } else {
        const parts: string[] = [];
        if (amendedFields.observation) parts.push("Observation amended");
        if (amendedFields.action) parts.push("Action amended");
        if (amendedFields.photos > 0) parts.push(`${amendedFields.photos} new photo${amendedFields.photos > 1 ? "s" : ""}`);
        if (amendedFields.locationsAdded > 0) parts.push("Location added");
        if (amendedFields.locationsAmended > 0) parts.push("Location amended");
        if (amendedFields.statusChange) parts.push(`Status: ${amendedFields.statusChange.from} \u2192 ${amendedFields.statusChange.to}`);
        if (parts.length > 0) {
          const completedThisInspection = !!amendedFields.statusChange && amendedFields.statusChange.to === "complete";
          map.set(d.id, { tag: completedThisInspection ? "COMPLETED" : "AMENDED", summary: parts.map(p => `\u2022 ${p}`).join("  ") });
        } else {
          map.set(d.id, { tag: null, summary: "" });
        }
      }
    }
    return map;
  }, [reportData]);

  const { data: shareLinks } = useQuery<any[]>({
    queryKey: [`/api/reports/${reportId}/share-links`],
  });

  const createShareLinkMutation = useMutation({
    mutationFn: async (recipientName: string) => {
      const res = await apiRequest("POST", `/api/reports/${reportId}/share-links`, { recipientName });
      return res.json();
    },
    onSuccess: (link: any) => {
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/share-links`] });
      const url = `${window.location.origin}${window.location.pathname}#/share/${link.token}`;
      setGeneratedLink(url);
    },
  });

  const deleteShareLinkMutation = useMutation({
    mutationFn: async (id: number) => {
      await apiRequest("DELETE", `/api/share-links/${id}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/share-links`] });
      toast({ title: "Share link revoked" });
    },
  });

  // Parse variable-length UID into sortable parts
  const parseUidParts = (uid: string) => {
    const parts = uid.split("-");
    const wtIdx = parts.findIndex((p) => /^[A-Z]{2,3}$/i.test(p));
    if (wtIdx < 0) return { elev: "", drop: 0, level: 0, work: "", num: 0 };
    const before = parts.slice(0, wtIdx);
    const alphaIdx = before.findIndex((p) => /^[A-Z]+$/i.test(p));
    const elev = alphaIdx >= 0 ? before[alphaIdx] : "";
    const numerics = before.filter((_, i) => i !== alphaIdx);
    return {
      elev,
      drop: parseInt(numerics[0] || "0", 10),
      level: parseInt(numerics[1] || "0", 10),
      work: parts[wtIdx] || "",
      num: parseInt(parts[wtIdx + 1] || "0", 10),
    };
  };

  // Sort by: Elevation, Drop (asc), Level (desc/highest first), WorkType (grouped), Number (asc)
  const sortByUid = (a: Defect, b: Defect): number => {
    const ap = parseUidParts(a.uid);
    const bp = parseUidParts(b.uid);
    if (ap.elev !== bp.elev) return ap.elev.localeCompare(bp.elev);
    if (ap.drop !== bp.drop) return ap.drop - bp.drop;
    if (ap.level !== bp.level) return bp.level - ap.level;
    if (ap.work !== bp.work) return ap.work.localeCompare(bp.work);
    return ap.num - bp.num;
  };

  const activeDefects = useMemo(() =>
    (defects?.filter((d) => d.status !== "complete" && d.recordType !== "observation") ?? []).sort(sortByUid), [defects]);
  const activeObservations = useMemo(() =>
    (defects?.filter((d) => d.status !== "complete" && d.recordType === "observation") ?? []).sort(sortByUid), [defects]);
  const completedAll = useMemo(() =>
    (defects?.filter((d) => d.status === "complete") ?? []).sort((a, b) => {
      const dateCompare = (b.dateClosed ?? "").localeCompare(a.dateClosed ?? "");
      if (dateCompare !== 0) return dateCompare;
      return sortByUid(a, b);
    }),
    [defects]);

  // ==================== EXPORT ORCHESTRATION ====================
  // Single entry point for the Export Report menu. Enforces the afcReference
  // guard (Pass 1): if the project has no afcReference, block and prompt; the
  // chosen export resumes after the reference is saved.
  const runExport = (profile: ProfileKey, format: "word" | "pdf", audienceSuffix: string) => {
    if (format === "word") handleGenerateWord(profile, audienceSuffix);
    else handleGeneratePdf(profile, audienceSuffix);
  };

  // `audienceSuffix` is the filename audience word: "" for the default unified
  // Report, "Contractor"/"Client" for the split-by-audience exports.
  const handleGenerate = (profile: ProfileKey, format: "word" | "pdf", audienceSuffix: string = "") => {
    const afcRef = ((project as any)?.afcReference || "").trim();
    if (!afcRef) {
      setPendingExport({ profile, format, audienceSuffix });
      setAfcGuardValue("");
      setAfcGuardOpen(true);
      return;
    }
    runExport(profile, format, audienceSuffix);
  };

  const saveAfcAndContinue = async () => {
    const value = afcGuardValue.trim();
    if (!value) {
      toast({ title: "AFC reference required", variant: "destructive" });
      return;
    }
    setAfcGuardSaving(true);
    try {
      await apiRequest("PATCH", `/api/projects/${projectId}`, { afcReference: value });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId] });
      setAfcGuardOpen(false);
      const next = pendingExport;
      setPendingExport(null);
      if (next) runExport(next.profile, next.format, next.audienceSuffix);
    } catch (e: any) {
      toast({ title: e?.message || "Save failed", variant: "destructive" });
    } finally {
      setAfcGuardSaving(false);
    }
  };

  // ==================== EXPORT (Pass 2: tree + renderers) ====================
  // The export handlers are now thin orchestrators: fetch the report-data,
  // build the RESOLVED report tree (report-tree.ts does ALL audience filtering,
  // category treatments, photo trimming, carried-forward and progress-summary
  // computation) then hand the tree to a renderer. The renderers walk the tree
  // in master section order and DO NOT re-filter. See render-docx.ts / render-pdf.ts.
  const handleGeneratePdf = async (profile: ProfileKey = "contractor", audienceSuffix: string = "") => {
    setGenerating("pdf");
    try {
      const res = await apiRequest("GET", `/api/reports/${reportId}/report-data`);
      const data = await res.json();
      const tree = buildReportTree(data, profile, audienceSuffix);
      const blob = await renderPdf(tree, { profile });
      downloadBlob(blob, `${tree.filenameBase}.pdf`);
      toast({ title: "PDF report downloaded" });
    } catch (err) {
      console.error(err);
      toast({ title: "Error generating PDF", variant: "destructive" });
    } finally {
      setGenerating(null);
    }
  };

  const handleGenerateWord = async (profile: ProfileKey = "contractor", audienceSuffix: string = "") => {
    setGenerating("word");
    try {
      const res = await apiRequest("GET", `/api/reports/${reportId}/report-data`);
      const data = await res.json();
      const tree = buildReportTree(data, profile, audienceSuffix);
      const blob = await renderDocx(tree, { profile });
      downloadBlob(blob, `${tree.filenameBase}.docx`);
      toast({ title: "Word report downloaded" });
    } catch (err) {
      console.error(err);
      toast({ title: "Error generating Word document", variant: "destructive" });
    } finally {
      setGenerating(null);
    }
  };


  if (reportLoading) {
    return (
      <div className="max-w-3xl mx-auto px-4 py-8">
        <div className="h-8 w-48 bg-muted animate-pulse rounded mb-4" />
        <div className="h-4 w-64 bg-muted animate-pulse rounded mb-8" />
        <div className="space-y-3">
          {[1, 2, 3].map((i) => (
            <div key={i} className="h-20 rounded-lg bg-muted animate-pulse" />
          ))}
        </div>
      </div>
    );
  }

  if (!report || !project) return null;

  return (
    <div className="max-w-3xl mx-auto px-4 py-8">
      {/* Header */}
      <div className="mb-8">
        <Link href={`/projects/${projectId}`}>
          <button className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors mb-4">
            <ArrowLeft className="w-4 h-4" />
            Back to Project
          </button>
        </Link>
        <div className="flex items-start justify-between gap-4">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">
              {project.name}
            </h1>
            <div className="flex flex-wrap gap-x-4 gap-y-1 mt-1 text-sm text-muted-foreground">
              {report.inspectionNumber && (
                <span>Inspection #{report.inspectionNumber}</span>
              )}
              {report.inspectionDate && (
                <span>{formatReportDate(report.inspectionDate)}</span>
              )}
              {report.revision && (
                <span>Rev {report.revision}</span>
              )}
            </div>
            {report.locationsCovered && (
              <p className="text-sm text-muted-foreground mt-1">
                <MapPin className="w-3.5 h-3.5 inline mr-1" />
                {report.locationsCovered}
              </p>
            )}
          </div>
          <Button variant="ghost" size="icon" onClick={openEditDialog}>
            <Settings className="w-4 h-4" />
          </Button>
        </div>

        {/* Edit Report Dialog */}
        <Dialog open={editOpen} onOpenChange={setEditOpen}>
          <DialogContent className="max-h-[90vh] overflow-y-auto">
            <DialogHeader>
              <DialogTitle>Edit Report</DialogTitle>
            </DialogHeader>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                updateReportMutation.mutate(editForm);
              }}
              className="space-y-4"
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label>Inspection Number</Label>
                  <Input value={editForm.inspectionNumber || ""} onChange={(e) => setEditForm({ ...editForm, inspectionNumber: e.target.value })} />
                </div>
                <div>
                  <Label>Inspection Date</Label>
                  <Input type="date" value={editForm.inspectionDate || ""} onChange={(e) => setEditForm({ ...editForm, inspectionDate: e.target.value })} />
                </div>
              </div>
              <div>
                <Label>Revision</Label>
                <Input value={editForm.revision || ""} onChange={(e) => setEditForm({ ...editForm, revision: e.target.value })} />
              </div>
              <div>
                <Label>Locations Covered</Label>
                <Textarea value={editForm.locationsCovered || ""} onChange={(e) => setEditForm({ ...editForm, locationsCovered: e.target.value })} rows={2} />
              </div>
              {/* Elevations picker */}
              <div>
                <Label className="mb-2 block">Elevations</Label>
                <div className="grid grid-cols-2 gap-2">
                  {STANDARD_ELEVATIONS.map((elev) => {
                    const selected: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                    return (
                      <label key={elev} className="flex items-center gap-2 text-sm cursor-pointer">
                        <Checkbox
                          checked={selected.includes(elev)}
                          onCheckedChange={(checked) => {
                            const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                            if (checked) { sel.push(elev); } else { const idx = sel.indexOf(elev); if (idx !== -1) sel.splice(idx, 1); }
                            setEditForm({ ...editForm, elevations: JSON.stringify(sel) });
                          }}
                        />
                        {elev}
                      </label>
                    );
                  })}
                </div>
                {(() => {
                  const selected: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                  const custom = selected.filter((e) => !STANDARD_ELEVATIONS.includes(e));
                  return custom.length > 0 ? (
                    <div className="flex flex-wrap gap-1.5 mt-2">
                      {custom.map((c) => (
                        <span key={c} className="inline-flex items-center gap-1 px-2 py-0.5 bg-accent rounded text-xs">
                          {c}
                          <button type="button" onClick={() => {
                            const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                            setEditForm({ ...editForm, elevations: JSON.stringify(sel.filter((e) => e !== c)) });
                          }} className="text-muted-foreground hover:text-destructive"><X className="w-3 h-3" /></button>
                        </span>
                      ))}
                    </div>
                  ) : null;
                })()}
                <div className="flex gap-2 mt-2">
                  <Input placeholder="Add custom elevation..." value={customElevation} onChange={(e) => setCustomElevation(e.target.value)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" && customElevation.trim()) {
                        e.preventDefault();
                        const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                        if (!sel.includes(customElevation.trim())) { sel.push(customElevation.trim()); setEditForm({ ...editForm, elevations: JSON.stringify(sel) }); }
                        setCustomElevation("");
                      }
                    }}
                  />
                  <Button type="button" variant="outline" size="sm" className="shrink-0" onClick={() => {
                    if (customElevation.trim()) {
                      const sel: string[] = (() => { try { return JSON.parse(editForm.elevations || "[]"); } catch { return []; } })();
                      if (!sel.includes(customElevation.trim())) { sel.push(customElevation.trim()); setEditForm({ ...editForm, elevations: JSON.stringify(sel) }); }
                      setCustomElevation("");
                    }
                  }}><Plus className="w-4 h-4" /></Button>
                </div>
              </div>
              <div>
                <Label>Attendees</Label>
                <div className="space-y-2 mt-1">
                  {(() => {
                    try {
                      return (JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[]).map((attendee, index) => (
                        <div key={index} className="flex items-center gap-2">
                          <Input
                            placeholder="Name"
                            value={attendee.name}
                            onChange={(e) => {
                              const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                              attendees[index].name = e.target.value;
                              setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                            }}
                          />
                          <Input
                            placeholder="Company / Role"
                            value={attendee.company}
                            onChange={(e) => {
                              const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                              attendees[index].company = e.target.value;
                              setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                            }}
                          />
                          <Button type="button" variant="ghost" size="icon" className="shrink-0" onClick={() => {
                            const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                            attendees.splice(index, 1);
                            setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                          }}>
                            <X className="w-4 h-4" />
                          </Button>
                        </div>
                      ));
                    } catch { return null; }
                  })()}
                  <Button type="button" variant="outline" size="sm" onClick={() => {
                    try {
                      const attendees = JSON.parse(editForm.attendees || "[]") as { name: string; company: string }[];
                      attendees.push({ name: "", company: "" });
                      setEditForm({ ...editForm, attendees: JSON.stringify(attendees) });
                    } catch {}
                  }}>
                    <Plus className="w-4 h-4 mr-1" />
                    Add Attendee
                  </Button>
                </div>
              </div>
              <Button type="submit" className="w-full" disabled={updateReportMutation.isPending}>
                {updateReportMutation.isPending ? "Saving..." : "Save Changes"}
              </Button>
            </form>
          </DialogContent>
        </Dialog>

        {/* AFC reference export guard */}
        <Dialog open={afcGuardOpen} onOpenChange={setAfcGuardOpen}>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>AFC Reference Required</DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              <p className="text-sm text-muted-foreground">
                This project has no AFC reference. It's used in the report filename and cover, so enter it before exporting.
              </p>
              <div>
                <Label>AFC reference</Label>
                <Input
                  placeholder="e.g. AFC-24123"
                  value={afcGuardValue}
                  onChange={(e) => setAfcGuardValue(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter") saveAfcAndContinue(); }}
                  autoFocus
                />
              </div>
              <Button className="w-full" disabled={afcGuardSaving} onClick={saveAfcAndContinue}>
                {afcGuardSaving ? "Saving..." : "Save & Continue"}
              </Button>
            </div>
          </DialogContent>
        </Dialog>

        <Dialog open={shareOpen} onOpenChange={setShareOpen}>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>Share Report</DialogTitle>
            </DialogHeader>
            <div className="space-y-4">
              {!generatedLink ? (
                <>
                  <div>
                    <Label>Recipient Name</Label>
                    <Input
                      placeholder="e.g. Client Name"
                      value={shareRecipient}
                      onChange={(e) => setShareRecipient(e.target.value)}
                    />
                    <p className="text-xs text-muted-foreground mt-1">Shown in the watermark banner on the shared report.</p>
                  </div>
                  <Button
                    className="w-full"
                    disabled={createShareLinkMutation.isPending}
                    onClick={() => createShareLinkMutation.mutate(shareRecipient)}
                  >
                    <LinkIcon className="w-4 h-4 mr-2" />
                    {createShareLinkMutation.isPending ? "Generating..." : "Generate Share Link"}
                  </Button>
                </>
              ) : (
                <div>
                  <Label>Share Link</Label>
                  <div className="flex gap-2 mt-1">
                    <Input readOnly value={generatedLink} className="text-xs" />
                    <Button variant="outline" size="icon" className="shrink-0" onClick={() => {
                      navigator.clipboard.writeText(generatedLink);
                      toast({ title: "Link copied to clipboard" });
                    }}>
                      <Copy className="w-4 h-4" />
                    </Button>
                  </div>
                  <p className="text-xs text-muted-foreground mt-2">Anyone with this link can view a read-only version of this report.</p>
                  <Button variant="outline" className="w-full mt-3" onClick={() => setGeneratedLink("")}>
                    Generate Another Link
                  </Button>
                </div>
              )}

              {shareLinks && shareLinks.length > 0 && (
                <div className="border-t pt-3">
                  <Label className="text-xs text-muted-foreground">Active Share Links</Label>
                  <div className="space-y-2 mt-2">
                    {shareLinks.map((link: any) => (
                      <div key={link.id} className="flex items-center justify-between text-sm">
                        <span>{link.recipientName || "No name"} <span className="text-xs text-muted-foreground">— {new Date(link.createdAt).toLocaleDateString()}</span></span>
                        <Button variant="ghost" size="sm" className="text-destructive h-7" onClick={() => deleteShareLinkMutation.mutate(link.id)}>
                          Revoke
                        </Button>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </DialogContent>
        </Dialog>
      </div>

      {/* Stats */}
      <div className="grid grid-cols-4 gap-3 mb-6">
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold">{defects?.length ?? 0}</div>
          <div className="text-xs text-muted-foreground">Total</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-amber-600">{activeDefects.length}</div>
          <div className="text-xs text-muted-foreground">Defects</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-blue-600">{activeObservations.length}</div>
          <div className="text-xs text-muted-foreground">Observations</div>
        </Card>
        <Card className="p-3 text-center">
          <div className="text-2xl font-semibold text-green-600">{completedAll.length}</div>
          <div className="text-xs text-muted-foreground">Completed</div>
        </Card>
      </div>

      {/* Actions */}
      <div className="flex gap-3 mb-6">
        <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-defect`}>
          <Button>
            <Plus className="w-4 h-4 mr-2" />
            Add Defect
          </Button>
        </Link>
        <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-observation`}>
          <Button variant="secondary">
            <Plus className="w-4 h-4 mr-2" />
            Add Observation
          </Button>
        </Link>

        <DropdownMenu>
          <DropdownMenuTrigger asChild>
            <Button
              variant="secondary"
              disabled={!!generating || !defects?.length}
            >
              <FileDown className="w-4 h-4 mr-2" />
              {generating ? "Generating..." : "Export Report"}
              <ChevronDown className="w-3.5 h-3.5 ml-1.5" />
            </Button>
          </DropdownMenuTrigger>
          <DropdownMenuContent align="start">
            <DropdownMenuLabel>Generate Report</DropdownMenuLabel>
            {/* Primary unified report: full content (all photos + full appendix,
                the contractor content path) with a neutral "Report" label and no
                audience word in the filename. */}
            <DropdownMenuItem onClick={() => handleGenerate("contractor", "word", "")}>
              <FileText className="w-4 h-4 mr-2" />
              Report — Word Document (.docx)
            </DropdownMenuItem>
            <DropdownMenuItem onClick={() => handleGenerate("contractor", "pdf", "")}>
              <FileText className="w-4 h-4 mr-2" />
              Report — PDF (.pdf)
            </DropdownMenuItem>
            <DropdownMenuSeparator />
            {/* Advanced: keep the original contractor/client audience split. These
                paths are unchanged in content; the filename gains a -Contractor /
                -Client audience suffix. */}
            <DropdownMenuSub>
              <DropdownMenuSubTrigger>
                <Settings className="w-4 h-4 mr-2" />
                Export split by audience
              </DropdownMenuSubTrigger>
              <DropdownMenuSubContent>
                <DropdownMenuLabel>Contractor</DropdownMenuLabel>
                <DropdownMenuItem onClick={() => handleGenerate("contractor", "word", "Contractor")}>
                  Word Document (.docx)
                </DropdownMenuItem>
                <DropdownMenuItem onClick={() => handleGenerate("contractor", "pdf", "Contractor")}>
                  PDF (.pdf)
                </DropdownMenuItem>
                <DropdownMenuSeparator />
                <DropdownMenuLabel>Client</DropdownMenuLabel>
                <DropdownMenuItem onClick={() => handleGenerate("client", "word", "Client")}>
                  Word Document (.docx)
                </DropdownMenuItem>
                <DropdownMenuItem onClick={() => handleGenerate("client", "pdf", "Client")}>
                  PDF (.pdf)
                </DropdownMenuItem>
              </DropdownMenuSubContent>
            </DropdownMenuSub>
            <DropdownMenuSeparator />
            <DropdownMenuItem onClick={() => {
              window.open(`${API_BASE}/api/reports/${reportId}/photos-zip?scope=current`, "_blank");
            }}>
              <ImageDown className="w-4 h-4 mr-2" />
              Export Images — {report?.inspectionNumber ? `Insp-${report.inspectionNumber} only` : "this inspection"} (.zip)
            </DropdownMenuItem>
            <DropdownMenuItem onClick={() => {
              window.open(`${API_BASE}/api/reports/${reportId}/photos-zip?scope=all`, "_blank");
            }}>
              <ImageDown className="w-4 h-4 mr-2" />
              Export Images — all photos (.zip)
            </DropdownMenuItem>
          </DropdownMenuContent>
        </DropdownMenu>

        <Button variant="outline" onClick={() => { setShareRecipient(""); setGeneratedLink(""); setShareOpen(true); }}>
          <Share2 className="w-4 h-4 mr-2" />
          Share
        </Button>
      </div>

      {defectsLoading ? (
        <div className="space-y-3">
          {[1, 2].map((i) => (
            <div key={i} className="h-20 rounded-lg bg-muted animate-pulse" />
          ))}
        </div>
      ) : !defects?.length ? (
        <div className="flex flex-col items-center justify-center py-16 text-center">
          <Camera className="w-10 h-10 text-muted-foreground/40 mb-4" />
          <h2 className="text-base font-medium mb-1">No entries recorded</h2>
          <p className="text-sm text-muted-foreground mb-6 max-w-xs">
            Start your inspection by adding defects or observations with photos and details.
          </p>
          <div className="flex gap-3">
            <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-defect`}>
              <Button>
                <Plus className="w-4 h-4 mr-2" />
                Add Defect
              </Button>
            </Link>
            <Link href={`/projects/${projectId}/reports/${reportId}/defects/new-observation`}>
              <Button variant="secondary">
                <Plus className="w-4 h-4 mr-2" />
                Add Observation
              </Button>
            </Link>
          </div>
        </div>
      ) : (
        <div className="space-y-8">
          {activeDefects.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <AlertTriangle className="w-4 h-4 text-amber-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Defects ({activeDefects.length})
                </h2>
              </div>
              <div className="space-y-2">
                {activeDefects.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={projectId!}
                    reportId={reportId!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                    changeTag={defectEventsMap.get(defect.id)?.tag ?? undefined}
                    changeSummary={defectEventsMap.get(defect.id)?.summary}
                    dims={cardDims}
                    hideLegacyAliases={hideLegacyAliases}
                  />
                ))}
              </div>
            </div>
          )}

          {activeObservations.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Eye className="w-4 h-4 text-blue-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Observations ({activeObservations.length})
                </h2>
              </div>
              <div className="space-y-2">
                {activeObservations.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={projectId!}
                    reportId={reportId!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                    changeTag={defectEventsMap.get(defect.id)?.tag ?? undefined}
                    changeSummary={defectEventsMap.get(defect.id)?.summary}
                    dims={cardDims}
                    hideLegacyAliases={hideLegacyAliases}
                  />
                ))}
              </div>
            </div>
          )}

          {completedAll.length > 0 && (
            <div>
              <div className="flex items-center gap-2 mb-3">
                <Archive className="w-4 h-4 text-green-600" />
                <h2 className="text-sm font-semibold uppercase tracking-wide text-muted-foreground">
                  Completed ({completedAll.length})
                </h2>
              </div>
              <div className="space-y-2">
                {completedAll.map((defect) => (
                  <DefectCard
                    key={defect.id}
                    defect={defect}
                    projectId={projectId!}
                    reportId={reportId!}
                    onDelete={() => deleteMutation.mutate(defect.id)}
                    changeTag={defectEventsMap.get(defect.id)?.tag ?? undefined}
                    changeSummary={defectEventsMap.get(defect.id)?.summary}
                    dims={cardDims}
                    hideLegacyAliases={hideLegacyAliases}
                  />
                ))}
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function DefectCard({ defect, projectId, reportId, onDelete, changeTag, changeSummary, dims, hideLegacyAliases }: { defect: Defect; projectId: string; reportId: string; onDelete: () => void; changeTag?: "NEW" | "AMENDED" | "COMPLETED"; changeSummary?: string; dims?: string[]; hideLegacyAliases?: boolean }) {
  const isComplete = defect.status === "complete";
  const locationText = formatDefectLocation(defect, dims || getLocationDimensions(undefined));
  // SVR Stage 2 — show "(prev. {legacy_id})" alias when the UID was migrated and aliases
  // aren't hidden for this project.
  const legacyAlias = defect.legacyId;
  const showLegacyAlias = !hideLegacyAliases && legacyAlias != null && String(legacyAlias).trim() !== "" && legacyAlias !== defect.uid;

  return (
    <Card className="group relative">
      <Link href={`/projects/${projectId}/reports/${reportId}/defects/${defect.id}`}>
        <div className={`flex items-center justify-between p-4 cursor-pointer hover:bg-accent/50 rounded-lg transition-colors ${isComplete ? "opacity-80" : ""}`}>
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2 mb-1">
              <span className="font-mono text-sm font-semibold">{defect.uid}</span>
              {showLegacyAlias && (
                <span className="font-mono text-xs text-muted-foreground">(prev. {legacyAlias})</span>
              )}
              <Badge
                variant="secondary"
                className={isComplete
                  ? "bg-green-100 text-green-800 dark:bg-green-900/30 dark:text-green-400"
                  : "bg-amber-100 text-amber-800 dark:bg-amber-900/30 dark:text-amber-400"
                }
              >
                {isComplete ? (
                  <><CheckCircle2 className="w-3 h-3 mr-1" />Complete</>
                ) : (
                  <><AlertTriangle className="w-3 h-3 mr-1" />Open</>
                )}
              </Badge>
              {changeTag === "NEW" && (
                <Badge variant="secondary" className="bg-emerald-100 text-emerald-800 dark:bg-emerald-900/30 dark:text-emerald-400 text-[10px] px-1.5 py-0">
                  NEW
                </Badge>
              )}
              {changeTag === "AMENDED" && (
                <Badge variant="secondary" className="bg-amber-50 text-amber-700 dark:bg-amber-900/20 dark:text-amber-300 text-[10px] px-1.5 py-0 border border-amber-300">
                  AMENDED
                </Badge>
              )}
              {changeTag === "COMPLETED" && (
                <Badge variant="secondary" className="bg-slate-100 text-slate-700 dark:bg-slate-800/40 dark:text-slate-300 text-[10px] px-1.5 py-0 border border-slate-300">
                  COMPLETED
                </Badge>
              )}
            </div>
            <p className="text-sm text-muted-foreground truncate">{defect.comment}</p>
            {locationText && (
              <p className="flex items-center gap-1 text-xs text-muted-foreground mt-0.5">
                <MapPin className="w-3 h-3 shrink-0" />{locationText}
              </p>
            )}
            {changeSummary && (
              <p className={`text-[11px] mt-0.5 truncate ${changeTag === "COMPLETED" ? "text-slate-600 dark:text-slate-400" : "text-amber-600 dark:text-amber-400"}`}>{changeSummary}</p>
            )}
            <div className="flex gap-3 mt-1 text-xs text-muted-foreground">
              <span>Assigned: {defect.assignedTo}</span>
              {isComplete && defect.dateClosed ? (
                <span className="text-green-600 dark:text-green-400">Completed: {defect.dateClosed}</span>
              ) : (
                <span>Due: {defect.dueDate}</span>
              )}
            </div>
          </div>
          <ChevronRight className="w-5 h-5 text-muted-foreground shrink-0 ml-3" />
        </div>
      </Link>
      <button
        onClick={(e) => {
          e.stopPropagation();
          if (confirm("Delete this defect and all its photos?")) {
            onDelete();
          }
        }}
        className="absolute top-3 right-12 p-1.5 rounded-md opacity-0 group-hover:opacity-100 hover:bg-destructive/10 hover:text-destructive transition-all"
      >
        <Trash2 className="w-4 h-4" />
      </button>
    </Card>
  );
}
