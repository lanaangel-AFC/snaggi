import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Link, useParams, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import {
  Dialog, DialogContent, DialogHeader, DialogTitle,
} from "@/components/ui/dialog";
import { useToast } from "@/hooks/use-toast";
import {
  ArrowLeft, Plus, MapPin, User, UserCheck, ChevronRight, Trash2,
  Settings, X, FileText, Copy, Calendar
} from "lucide-react";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import type { Project, Report, Defect } from "@shared/schema";
import { useState } from "react";

const STANDARD_ELEVATIONS = [
  "North", "North East", "East", "South East",
  "South", "South West", "West", "North West",
];

export default function ProjectDetail() {
  const { id } = useParams<{ id: string }>();
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const [editOpen, setEditOpen] = useState(false);
  const [editForm, setEditForm] = useState<Record<string, string>>({});
  const [customElevation, setCustomElevation] = useState("");
  const [newReportOpen, setNewReportOpen] = useState(false);
  const [newReportForm, setNewReportForm] = useState({
    inspectionNumber: "",
    inspectionDate: new Date().toISOString().split("T")[0],
    revision: "01",
    locationsCovered: "",
    attendees: "[]",
  });

  const { data: project, isLoading: projectLoading } = useQuery<Project>({
    queryKey: ["/api/projects", id],
  });

  const { data: reports, isLoading: reportsLoading } = useQuery<Report[]>({
    queryKey: [`/api/projects/${id}/reports`],
  });

  const openEditDialog = () => {
    if (!project) return;
    setEditForm({
      name: project.name,
      address: project.address,
      client: project.client,
      inspector: project.inspector,
      afcReference: (project as any).afcReference || "",
      elevations: (project as any).elevations || "[]",
    });
    setCustomElevation("");
    setEditOpen(true);
  };

  const updateProjectMutation = useMutation({
    mutationFn: async (data: Record<string, string>) => {
      const res = await apiRequest("PATCH", `/api/projects/${id}`, data);
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/projects", id] });
      setEditOpen(false);
      toast({ title: "Project updated" });
    },
  });

  const createReportMutation = useMutation({
    mutationFn: async (data: typeof newReportForm) => {
      const res = await apiRequest("POST", `/api/projects/${id}/reports`, data);
      return res.json();
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${id}/reports`] });
      setNewReportOpen(false);
      setNewReportForm({ inspectionNumber: "", inspectionDate: new Date().toISOString().split("T")[0], revision: "01", locationsCovered: "", attendees: "[]" });
      toast({ title: "Report created" });
      navigate(`/projects/${id}/reports/${data.id}`);
    },
  });

  const copyReportMutation = useMutation({
    mutationFn: async (reportId: number) => {
      const res = await apiRequest("POST", `/api/reports/${reportId}/copy`);
      return res.json();
    },
    onSuccess: (data) => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${id}/reports`] });
      toast({ title: "Report copied" });
      navigate(`/projects/${id}/reports/${data.id}`);
    },
  });

  const deleteReportMutation = useMutation({
    mutationFn: async (reportId: number) => {
      await apiRequest("DELETE", `/api/reports/${reportId}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: [`/api/projects/${id}/reports`] });
      toast({ title: "Report deleted" });
    },
  });

  if (projectLoading) {
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

  if (!project) return null;

  const mostRecentReport = reports?.[0];

  return (
    <div className="max-w-3xl mx-auto px-4 py-8">
      {/* Header */}
      <div className="mb-8">
        <Link href="/">
          <button className="flex items-center gap-1.5 text-sm text-muted-foreground hover:text-foreground transition-colors mb-4">
            <ArrowLeft className="w-4 h-4" />
            All Projects
          </button>
        </Link>
        <div className="flex items-start justify-between gap-4">
          <div>
            <h1 className="text-xl font-semibold tracking-tight">
              {project.name}
            </h1>
            <div className="flex flex-wrap gap-x-4 gap-y-1 mt-2 text-sm text-muted-foreground">
              <span className="flex items-center gap-1.5">
                <MapPin className="w-3.5 h-3.5" />
                {project.address}
              </span>
              <span className="flex items-center gap-1.5">
                <User className="w-3.5 h-3.5" />
                {project.client}
              </span>
              <span className="flex items-center gap-1.5">
                <UserCheck className="w-3.5 h-3.5" />
                {project.inspector}
              </span>
            </div>
          </div>
          <Button variant="ghost" size="icon" onClick={openEditDialog}>
            <Settings className="w-4 h-4" />
          </Button>
        </div>

        {/* Edit Project Dialog */}
        <Dialog open={editOpen} onOpenChange={setEditOpen}>
          <DialogContent className="max-h-[90vh] overflow-y-auto">
            <DialogHeader>
              <DialogTitle>Edit Project</DialogTitle>
            </DialogHeader>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                updateProjectMutation.mutate(editForm);
              }}
              className="space-y-4"
            >
              <div>
                <Label>Project Name</Label>
                <Input value={editForm.name || ""} onChange={(e) => setEditForm({ ...editForm, name: e.target.value })} required />
              </div>
              <div>
                <Label>Site Address</Label>
                <Input value={editForm.address || ""} onChange={(e) => setEditForm({ ...editForm, address: e.target.value })} required />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label>Client</Label>
                  <Input value={editForm.client || ""} onChange={(e) => setEditForm({ ...editForm, client: e.target.value })} required />
                </div>
                <div>
                  <Label>Inspector</Label>
                  <Input value={editForm.inspector || ""} onChange={(e) => setEditForm({ ...editForm, inspector: e.target.value })} required />
                </div>
              </div>
              <div>
                <Label>AFC Reference</Label>
                <Input value={editForm.afcReference || ""} onChange={(e) => setEditForm({ ...editForm, afcReference: e.target.value })} />
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
                  <Input
                    placeholder="Add custom elevation..."
                    value={customElevation}
                    onChange={(e) => setCustomElevation(e.target.value)}
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
              <Button type="submit" className="w-full" disabled={updateProjectMutation.isPending}>
                {updateProjectMutation.isPending ? "Saving..." : "Save Changes"}
              </Button>
            </form>
          </DialogContent>
        </Dialog>

        {/* New Report Dialog */}
        <Dialog open={newReportOpen} onOpenChange={setNewReportOpen}>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>New Report</DialogTitle>
            </DialogHeader>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                createReportMutation.mutate(newReportForm);
              }}
              className="space-y-4"
            >
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label>Inspection Number</Label>
                  <Input value={newReportForm.inspectionNumber} onChange={(e) => setNewReportForm({ ...newReportForm, inspectionNumber: e.target.value })} placeholder="e.g. 01" />
                </div>
                <div>
                  <Label>Inspection Date</Label>
                  <Input type="date" value={newReportForm.inspectionDate} onChange={(e) => setNewReportForm({ ...newReportForm, inspectionDate: e.target.value })} />
                </div>
              </div>
              <div>
                <Label>Revision</Label>
                <Input value={newReportForm.revision} onChange={(e) => setNewReportForm({ ...newReportForm, revision: e.target.value })} placeholder="01" />
              </div>
              <div>
                <Label>Locations Covered</Label>
                <Input value={newReportForm.locationsCovered} onChange={(e) => setNewReportForm({ ...newReportForm, locationsCovered: e.target.value })} placeholder="e.g. North elevation levels 1-5" />
              </div>
              <Button type="submit" className="w-full" disabled={createReportMutation.isPending}>
                {createReportMutation.isPending ? "Creating..." : "Create Report"}
              </Button>
            </form>
          </DialogContent>
        </Dialog>
      </div>

      {/* Actions */}
      <div className="flex gap-3 mb-6">
        <Button onClick={() => setNewReportOpen(true)}>
          <Plus className="w-4 h-4 mr-2" />
          New Report
        </Button>
        {mostRecentReport && (
          <Button
            variant="secondary"
            disabled={copyReportMutation.isPending}
            onClick={() => copyReportMutation.mutate(mostRecentReport.id)}
          >
            <Copy className="w-4 h-4 mr-2" />
            {copyReportMutation.isPending ? "Copying..." : "Copy Previous Report"}
          </Button>
        )}
      </div>

      {/* Reports List */}
      {reportsLoading ? (
        <div className="space-y-3">
          {[1, 2].map((i) => (
            <div key={i} className="h-20 rounded-lg bg-muted animate-pulse" />
          ))}
        </div>
      ) : !reports?.length ? (
        <div className="flex flex-col items-center justify-center py-16 text-center">
          <FileText className="w-10 h-10 text-muted-foreground/40 mb-4" />
          <h2 className="text-base font-medium mb-1">No reports yet</h2>
          <p className="text-sm text-muted-foreground mb-6 max-w-xs">
            Create your first site visit report to start recording defects and observations.
          </p>
          <Button onClick={() => setNewReportOpen(true)}>
            <Plus className="w-4 h-4 mr-2" />
            New Report
          </Button>
        </div>
      ) : (
        <div className="space-y-3">
          {reports.map((report) => (
            <ReportCard
              key={report.id}
              report={report}
              projectId={id!}
              onDelete={() => {
                if (confirm("Delete this report and all its entries?")) {
                  deleteReportMutation.mutate(report.id);
                }
              }}
            />
          ))}
        </div>
      )}
    </div>
  );
}

function ReportCard({ report, projectId, onDelete }: { report: Report; projectId: string; onDelete: () => void }) {
  // Fetch defect count for this report
  const { data: defects } = useQuery<Defect[]>({
    queryKey: [`/api/reports/${report.id}/defects`],
  });

  const entryCount = defects?.length ?? 0;
  const formatDate = (dateStr?: string | null) => {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    return d.toLocaleDateString("en-AU", { day: "2-digit", month: "short", year: "numeric" });
  };

  return (
    <Card className="group relative">
      <Link href={`/projects/${projectId}/reports/${report.id}`}>
        <div className="flex items-center justify-between p-4 cursor-pointer hover:bg-accent/50 rounded-lg transition-colors">
          <div className="min-w-0 flex-1">
            <div className="flex items-center gap-2 mb-1">
              <span className="font-medium">
                {report.inspectionNumber ? `Inspection #${report.inspectionNumber}` : `Report #${report.id}`}
              </span>
              {report.revision && (
                <span className="text-xs text-muted-foreground bg-muted px-1.5 py-0.5 rounded">
                  Rev {report.revision}
                </span>
              )}
            </div>
            <div className="flex flex-wrap gap-x-4 gap-y-1 text-sm text-muted-foreground">
              {report.inspectionDate && (
                <span className="flex items-center gap-1.5">
                  <Calendar className="w-3.5 h-3.5" />
                  {formatDate(report.inspectionDate)}
                </span>
              )}
              {report.locationsCovered && (
                <span className="flex items-center gap-1.5">
                  <MapPin className="w-3.5 h-3.5" />
                  <span className="truncate max-w-[200px]">{report.locationsCovered}</span>
                </span>
              )}
              <span>{entryCount} {entryCount === 1 ? "entry" : "entries"}</span>
            </div>
          </div>
          <ChevronRight className="w-5 h-5 text-muted-foreground shrink-0 ml-3" />
        </div>
      </Link>
      <button
        onClick={(e) => {
          e.stopPropagation();
          onDelete();
        }}
        className="absolute top-3 right-12 p-1.5 rounded-md opacity-0 group-hover:opacity-100 hover:bg-destructive/10 hover:text-destructive transition-all"
      >
        <Trash2 className="w-4 h-4" />
      </button>
    </Card>
  );
}
