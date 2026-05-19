import { useParams } from "wouter";
import { useQuery } from "@tanstack/react-query";
import { Card } from "@/components/ui/card";

const API_BASE = "__PORT_5000__".startsWith("__") ? "" : "__PORT_5000__";

export default function SharedReport() {
  const { token } = useParams<{ token: string }>();

  const { data, isLoading, error } = useQuery<any>({
    queryKey: [`/api/share/${token}`],
    retry: false,
  });

  if (isLoading) {
    return (
      <div className="max-w-4xl mx-auto px-4 py-8">
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

  if (error || !data) {
    return (
      <div className="max-w-4xl mx-auto px-4 py-16 text-center">
        <h1 className="text-2xl font-semibold mb-2">Link Not Found</h1>
        <p className="text-muted-foreground">This share link is invalid or has been revoked.</p>
      </div>
    );
  }

  return (
    <div className="relative max-w-4xl mx-auto px-4 py-8">
      {/* Watermark overlay */}
      <div className="fixed inset-0 pointer-events-none z-50 flex items-center justify-center opacity-[0.04]">
        <div className="text-9xl font-bold text-black rotate-[-30deg] select-none whitespace-nowrap">
          CONFIDENTIAL
        </div>
      </div>

      {/* Top watermark bar */}
      <div className="border-b border-amber-200 bg-amber-50 px-4 py-2 mb-6 text-xs text-amber-900 flex justify-between items-center rounded">
        <span>Shared with: <strong>{data.recipientName || "Recipient"}</strong></span>
        <span>Confidential — Angel Facade Consulting</span>
      </div>

      {/* Header */}
      <h1 className="text-2xl font-semibold">{data.project.name}</h1>
      <p className="text-sm text-muted-foreground mb-6">
        {data.project.address} &middot; {data.project.client} &middot; Inspector: {data.project.inspector}
      </p>

      <h2 className="text-lg font-semibold mb-1">
        Inspection {data.report.inspectionNumber || "\u2014"}
        {data.report.revision && ` (Rev ${data.report.revision})`}
      </h2>
      <p className="text-sm text-muted-foreground mb-6">
        {data.report.inspectionDate && `Inspected ${data.report.inspectionDate}`}
        {data.report.locationsCovered && ` \u2014 ${data.report.locationsCovered}`}
      </p>

      {/* Defects/Observations list */}
      {data.defects.map((d: any) => (
        <DefectCardReadOnly key={d.id} defect={d} token={token!} />
      ))}

      {/* Bottom watermark */}
      <div className="border-t mt-12 pt-4 text-xs text-center text-muted-foreground">
        This is a read-only shared report for {data.recipientName || "the recipient"} only.
        Confidential — do not forward.
      </div>
    </div>
  );
}

function DefectCardReadOnly({ defect, token }: { defect: any; token: string }) {
  const slotOrder = ["wip1", "wip2", "wip3", "wip4", "wip5", "complete"];
  const sortedPhotos = slotOrder
    .map((s) => (defect.photos || []).find((p: any) => p.slot === s))
    .filter(Boolean);

  return (
    <Card className="p-4 mb-4">
      <div className="flex items-center gap-2 mb-2 flex-wrap">
        <span className="font-mono font-semibold">{defect.uid}</span>
        <span
          className={`px-2 py-0.5 text-xs rounded ${
            defect.recordType === "observation"
              ? "bg-blue-100 text-blue-800"
              : "bg-amber-100 text-amber-800"
          }`}
        >
          {defect.recordType === "observation" ? "Observation" : "Defect"}
        </span>
        <span
          className={`px-2 py-0.5 text-xs rounded ${
            defect.status === "complete"
              ? "bg-green-100 text-green-800"
              : "bg-gray-100 text-gray-800"
          }`}
        >
          {defect.status === "complete" ? "Complete" : "Open"}
        </span>
      </div>

      <div className="space-y-1.5 text-sm">
        {defect.comment && (
          <div>
            <strong>Observation:</strong> {defect.comment}
          </div>
        )}
        {defect.actionRequired && (
          <div>
            <strong>Action Required:</strong> {defect.actionRequired}
          </div>
        )}
        {defect.assignedTo && (
          <div>
            <strong>By Whom:</strong> {defect.assignedTo}
          </div>
        )}
        {defect.dateClosed && (
          <div>
            <strong>Completed:</strong> {defect.dateClosed}
          </div>
        )}
      </div>

      {sortedPhotos.length > 0 && (
        <div className="grid grid-cols-2 gap-3 mt-4">
          {sortedPhotos.map((p: any) => (
            <div key={p.id}>
              <img
                src={`${API_BASE}/api/share/${token}/photo/${p.filename}`}
                alt={p.slot}
                className="w-full rounded"
                loading="lazy"
              />
              <p className="text-xs text-muted-foreground mt-1">
                {p.slot === "complete"
                  ? "Complete"
                  : p.slot.toUpperCase().replace("WIP", "WIP ")}
                {p.caption ? ` \u2014 ${p.caption}` : ""}
              </p>
            </div>
          ))}
        </div>
      )}
    </Card>
  );
}
