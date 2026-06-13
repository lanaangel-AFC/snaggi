import { useQuery, useMutation, useQueryClient } from "@tanstack/react-query";
import { apiRequest } from "@/lib/queryClient";
import { useParams, Link } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { ArrowLeft, Download, AlertTriangle, CheckCircle2 } from "lucide-react";
import { formatLocation } from "@shared/location";
import { escapeCsvField } from "@shared/csv";
import { useToast } from "@/hooks/use-toast";

interface MigrationSummary {
  totalRowsMigrated: number;
  canonicalRowsChanged: number;
  canonicalRowsUnchanged: number;
  duplicateRowsArchived: number;
  duplicateRowsClosed: number;
}

interface ProjectMeta {
  uidMigrationAppliedAt: string | null;
  uidMigrationSummary: MigrationSummary | null;
}

interface PreviewRow {
  defectId: number;
  legacyId: string;
  proposedUid: string;
  location: Record<string, string>;
  workType: string | null;
  type: string;
  status: string;
  changed: boolean;
  duplicateLegacyId: boolean;
  notes: string | null;
}

interface PreviewResponse {
  projectId: number;
  projectName: string | null;
  locationDimensions: string[];
  uidProtocol: string | null;
  rows: PreviewRow[];
  summary: {
    totalRows: number;
    uniqueLegacyIds: number;
    rowsWhereProposedDiffersFromLegacy: number;
    duplicateLegacyIdGroups: number;
    duplicateResolutionStrategy: string;
  };
}

// Build a CSV string from the preview rows. Same data as the table so the user can
// share it with their team for review before Stage 2 (apply) is requested.
function buildCsv(data: PreviewResponse): string {
  const header = ["Defect ID", "Legacy ID", "Proposed UID", "Location", "Work Type", "Type", "Status", "Changed", "Duplicate", "Notes"];
  const esc = escapeCsvField;
  const lines = [header.join(",")];
  for (const r of data.rows) {
    lines.push([
      r.defectId,
      r.legacyId,
      r.proposedUid,
      formatLocation(r.location, data.locationDimensions),
      r.workType || "",
      r.type,
      r.status,
      r.changed ? "yes" : "no",
      r.duplicateLegacyId ? "yes" : "no",
      r.notes || "",
    ].map(esc).join(","));
  }
  return lines.join("\n");
}

export default function MigrationPreview() {
  const params = useParams();
  // projectId can come from the route (/admin/migration-preview/:projectId) or default to 4.
  const projectId = Number(params.projectId) || 4;

  const queryClient = useQueryClient();
  const { toast } = useToast();

  const { data, isLoading, error } = useQuery<PreviewResponse>({
    queryKey: ["uid-migration-preview", projectId],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/admin/uid-migration-preview?projectId=${projectId}`);
      return res.json();
    },
  });

  // Migration applied state comes from the project endpoint's meta fields.
  const { data: projectMeta } = useQuery<ProjectMeta>({
    queryKey: ["/api/projects", projectId, "migration-meta"],
    queryFn: async () => {
      const res = await apiRequest("GET", `/api/projects/${projectId}`);
      return res.json();
    },
  });
  const appliedAt = projectMeta?.uidMigrationAppliedAt ?? null;
  const summary = projectMeta?.uidMigrationSummary ?? null;

  const applyMutation = useMutation({
    mutationFn: async () => {
      const res = await apiRequest("POST", `/api/admin/uid-migration-apply`, { projectId, confirm: true });
      return res.json();
    },
    onSuccess: (result: any) => {
      toast({
        title: result?.alreadyApplied ? "Already applied" : "Migration applied",
        description: result?.summary
          ? `${result.summary.totalRowsMigrated} rows migrated, ${result.summary.canonicalRowsChanged} UIDs changed`
          : undefined,
      });
      queryClient.invalidateQueries({ queryKey: ["/api/projects", projectId, "migration-meta"] });
      queryClient.invalidateQueries({ queryKey: ["uid-migration-preview", projectId] });
    },
    onError: (err: any) => {
      toast({ title: "Apply failed", description: err?.message || String(err), variant: "destructive" });
    },
  });

  // Server-generated CSV (matches the Stage 2 export schema). Opens the download endpoint.
  const downloadMappingCsv = () => {
    window.open(`/api/admin/uid-migration-export.csv?projectId=${projectId}`, "_blank");
  };

  const downloadCsv = () => {
    if (!data) return;
    const csv = buildCsv(data);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `uid-migration-preview-project-${projectId}.csv`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="max-w-6xl mx-auto p-4 space-y-4">
      <div className="flex items-center gap-2">
        <Link href="/">
          <Button variant="ghost" size="sm"><ArrowLeft className="w-4 h-4 mr-1" />Home</Button>
        </Link>
        <h1 className="text-lg font-semibold">UID Migration Preview</h1>
      </div>

      {isLoading && <p className="text-sm text-muted-foreground">Loading preview…</p>}
      {error && <p className="text-sm text-red-600">Failed to load preview: {(error as Error).message}</p>}

      {data && (
        <>
          {/* Applied notice / Stage 2 gate banner */}
          {appliedAt ? (
            <Card className="p-4 border-green-300 bg-green-50 dark:bg-green-900/20">
              <div className="flex items-start gap-2">
                <CheckCircle2 className="w-5 h-5 text-green-600 shrink-0 mt-0.5" />
                <div>
                  <p className="font-medium text-green-800 dark:text-green-300">
                    MIGRATION APPLIED on {new Date(appliedAt).toLocaleString()}
                  </p>
                  {summary && (
                    <p className="text-sm text-green-700 dark:text-green-400 mt-1">
                      {summary.totalRowsMigrated} rows migrated • {summary.canonicalRowsChanged} canonical UIDs changed •{" "}
                      {summary.canonicalRowsUnchanged} canonical unchanged • {summary.duplicateRowsArchived} duplicates archived •{" "}
                      {summary.duplicateRowsClosed} duplicates closed
                    </p>
                  )}
                </div>
              </div>
            </Card>
          ) : (
            <Card className="p-4 border-amber-300 bg-amber-50 dark:bg-amber-900/20">
              <div className="flex items-start gap-2">
                <AlertTriangle className="w-5 h-5 text-amber-600 shrink-0 mt-0.5" />
                <div>
                  <p className="font-medium text-amber-800 dark:text-amber-300">This is a preview — applying writes to the database.</p>
                  <p className="text-sm text-amber-700 dark:text-amber-400 mt-1">
                    Review the preview, then click APPLY MIGRATION. The apply is idempotent (safe to re-run).
                  </p>
                </div>
              </div>
            </Card>
          )}

          {/* Summary block */}
          <Card className="p-4">
            <div className="flex items-center justify-between gap-4 flex-wrap">
              <div>
                <h2 className="font-semibold">
                  {data.projectName ? `${data.projectName} (project ${data.projectId})` : `Project ${data.projectId} — not found`}
                </h2>
                <p className="text-xs text-muted-foreground mt-0.5">
                  UID protocol: <code className="font-mono">{data.uidProtocol || "—"}</code>
                  {"  •  "}
                  Dimensions: <code className="font-mono">{data.locationDimensions.join(", ")}</code>
                </p>
              </div>
              <div className="flex gap-2">
                <Button onClick={downloadCsv} variant="outline" size="sm" disabled={data.rows.length === 0}>
                  <Download className="w-4 h-4 mr-1" />Download preview as CSV
                </Button>
                {appliedAt ? (
                  <Button onClick={downloadMappingCsv} variant="outline" size="sm">
                    <Download className="w-4 h-4 mr-1" />Download mapping CSV
                  </Button>
                ) : (
                  <Button
                    onClick={() => applyMutation.mutate()}
                    disabled={applyMutation.isPending || data.rows.length === 0}
                    size="sm"
                    className="bg-green-600 hover:bg-green-700 text-white"
                  >
                    {applyMutation.isPending ? "Applying…" : "APPLY MIGRATION"}
                  </Button>
                )}
              </div>
            </div>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3 mt-4 text-sm">
              <Stat label="Total rows" value={data.summary.totalRows} />
              <Stat label="Unique legacy IDs" value={data.summary.uniqueLegacyIds} />
              <Stat label="Proposed ≠ legacy" value={data.summary.rowsWhereProposedDiffersFromLegacy} />
              <Stat label="Duplicate groups" value={data.summary.duplicateLegacyIdGroups} />
            </div>
            <p className="text-xs text-muted-foreground mt-3">
              Duplicate strategy: {data.summary.duplicateResolutionStrategy}
            </p>
          </Card>

          {/* Migration table */}
          <Card className="p-0 overflow-x-auto">
            <table className="w-full text-sm">
              <thead className="bg-muted/50 text-left">
                <tr>
                  <th className="p-2 font-medium">Legacy ID</th>
                  <th className="p-2 font-medium">Proposed UID</th>
                  <th className="p-2 font-medium">Location</th>
                  <th className="p-2 font-medium">Work Type</th>
                  <th className="p-2 font-medium">Type</th>
                  <th className="p-2 font-medium">Status</th>
                  <th className="p-2 font-medium">Changed?</th>
                  <th className="p-2 font-medium">Duplicate?</th>
                </tr>
              </thead>
              <tbody>
                {data.rows.length === 0 && (
                  <tr><td colSpan={8} className="p-4 text-center text-muted-foreground">No rows for this project.</td></tr>
                )}
                {data.rows.map((r) => (
                  <tr key={r.defectId} className="border-t hover:bg-accent/30" title={r.notes || undefined}>
                    <td className="p-2 font-mono">{r.legacyId}</td>
                    <td className="p-2 font-mono">{r.proposedUid}</td>
                    <td className="p-2">{formatLocation(r.location, data.locationDimensions)}</td>
                    <td className="p-2">{r.workType || "—"}</td>
                    <td className="p-2">{r.type}</td>
                    <td className="p-2">{r.status}</td>
                    <td className="p-2">{r.changed ? <span className="text-amber-600 font-medium">yes</span> : "no"}</td>
                    <td className="p-2">{r.duplicateLegacyId ? <span className="text-red-600 font-medium">yes</span> : "no"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>
        </>
      )}
    </div>
  );
}

function Stat({ label, value }: { label: string; value: number }) {
  return (
    <div className="rounded-md border p-2">
      <div className="text-xs text-muted-foreground">{label}</div>
      <div className="text-lg font-semibold">{value}</div>
    </div>
  );
}
