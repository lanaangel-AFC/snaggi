import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { useParams, useLocation } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { useToast } from "@/hooks/use-toast";
import {
  ArrowLeft, CheckCircle2, Circle, AlertTriangle, ChevronRight,
  ListChecks, Loader2, History, CalendarClock,
} from "lucide-react";
import { useState, useMemo } from "react";
import type { Report } from "@shared/schema";

type TodoItem = {
  id: number;
  reportId: number;
  defectId: number;
  uid: string;
  groupKey: "due" | "open" | "stranded";
  recordType: string;
  comment: string;
  actionText: string;
  dueDate: string | null;
  assignedTo: string | null;
  sourceReportId: number | null;
  sourceInspectionNumber: string | null;
  doneAt: string | null;
  doneReason: string | null;
  liveStatus: string | null;
  linkReportId: number;
};

type TodoResponse = {
  reportId: number;
  inspectionNumber: string;
  inspectionDate: string;
  generated: boolean;
  total: number;
  done: number;
  items: TodoItem[];
};

const GROUPS: Array<{ key: TodoItem["groupKey"]; title: string; blurb: string; icon: typeof CalendarClock }> = [
  {
    key: "due",
    title: "Due or overdue",
    blurb: "Due on or before this inspection date, for any party.",
    icon: CalendarClock,
  },
  {
    key: "open",
    title: "Other open items",
    blurb: "Open in this inspection but not yet due.",
    icon: Circle,
  },
  {
    key: "stranded",
    title: "Carried over from earlier inspections",
    blurb: "Still open in an earlier inspection and never carried into this one.",
    icon: History,
  },
];

export default function InspectionTodos() {
  const params = useParams();
  const projectId = Number(params.projectId);
  const reportId = Number(params.reportId);
  const [, navigate] = useLocation();
  const { toast } = useToast();
  const [filter, setFilter] = useState<"outstanding" | "done" | "all">("outstanding");

  const { data: report } = useQuery<Report>({
    queryKey: [`/api/reports/${reportId}`],
    enabled: Number.isFinite(reportId),
  });

  const { data, isLoading } = useQuery<TodoResponse>({
    queryKey: [`/api/reports/${reportId}/todos`],
    enabled: Number.isFinite(reportId),
  });

  const generate = useMutation({
    mutationFn: async () => {
      const res = await apiRequest("POST", `/api/reports/${reportId}/todos/generate`, { dryRun: false });
      return res.json();
    },
    onSuccess: (result: { created?: number }) => {
      queryClient.invalidateQueries({ queryKey: [`/api/reports/${reportId}/todos`] });
      toast({ title: "To-do list created", description: `${result?.created ?? 0} items to work through.` });
    },
    onError: (err: Error) => {
      toast({ title: "Could not create the list", description: err.message, variant: "destructive" });
    },
  });

  const items = data?.items ?? [];
  const visible = useMemo(() => {
    if (filter === "outstanding") return items.filter((i) => !i.doneAt);
    if (filter === "done") return items.filter((i) => i.doneAt);
    return items;
  }, [items, filter]);

  const grouped = useMemo(() => {
    const sortItems = (a: TodoItem, b: TodoItem) => {
      const da = a.dueDate || "9999-12-31";
      const db = b.dueDate || "9999-12-31";
      if (da !== db) return da < db ? -1 : 1;
      return a.uid.localeCompare(b.uid);
    };
    return GROUPS.map((g) => ({
      ...g,
      rows: visible.filter((i) => i.groupKey === g.key).slice().sort(sortItems),
      totalInGroup: items.filter((i) => i.groupKey === g.key).length,
    }));
  }, [visible, items]);

  const openForm = (item: TodoItem) => {
    navigate(`/projects/${projectId}/reports/${item.linkReportId}/defects/${item.defectId}`);
  };

  const done = data?.done ?? 0;
  const total = data?.total ?? 0;
  const pct = total > 0 ? Math.round((done / total) * 100) : 0;
  const todayIso = new Date().toISOString().slice(0, 10);

  return (
    <div className="min-h-screen bg-background pb-16">
      <div className="sticky top-0 z-10 bg-background/95 backdrop-blur border-b">
        <div className="max-w-3xl mx-auto px-4 py-3">
          <Button
            variant="ghost"
            size="sm"
            className="gap-1.5 -ml-2 mb-1"
            onClick={() => navigate(`/projects/${projectId}/reports/${reportId}`)}
            data-testid="button-back-to-report"
          >
            <ArrowLeft className="w-4 h-4" />
            Back to Report
          </Button>

          <div className="flex items-start justify-between gap-3">
            <div className="min-w-0">
              <h1 className="text-lg font-semibold flex items-center gap-2">
                <ListChecks className="w-5 h-5 text-primary shrink-0" />
                To do
              </h1>
              <p className="text-sm text-muted-foreground">
                Inspection {report?.inspectionNumber ?? data?.inspectionNumber ?? "--"}
                {(report?.inspectionDate || data?.inspectionDate) && (
                  <> · {report?.inspectionDate ?? data?.inspectionDate}</>
                )}
              </p>
            </div>
            {total > 0 && (
              <div className="text-right shrink-0">
                <div className="text-sm font-semibold tabular-nums" data-testid="text-todo-progress">
                  {done} of {total} done
                </div>
                <div className="w-24 h-1.5 bg-muted rounded-full mt-1.5 overflow-hidden">
                  <div className="h-full bg-primary rounded-full transition-all" style={{ width: `${pct}%` }} />
                </div>
              </div>
            )}
          </div>

          {total > 0 && (
            <div className="flex gap-1.5 mt-3">
              {([
                ["outstanding", `Outstanding (${total - done})`],
                ["done", `Done (${done})`],
                ["all", `All (${total})`],
              ] as const).map(([key, label]) => (
                <Button
                  key={key}
                  size="sm"
                  variant={filter === key ? "default" : "outline"}
                  className="h-7 text-xs"
                  onClick={() => setFilter(key)}
                  data-testid={`button-filter-${key}`}
                >
                  {label}
                </Button>
              ))}
            </div>
          )}
        </div>
      </div>

      <div className="max-w-3xl mx-auto px-4 py-4 space-y-6">
        {isLoading && (
          <div className="flex items-center gap-2 text-sm text-muted-foreground py-8 justify-center">
            <Loader2 className="w-4 h-4 animate-spin" />
            Loading the list
          </div>
        )}

        {!isLoading && total === 0 && (
          <Card className="p-6 text-center">
            <ListChecks className="w-8 h-8 text-muted-foreground mx-auto mb-3" />
            <p className="font-medium">No to-do list for this inspection yet</p>
            <p className="text-sm text-muted-foreground mt-1 mb-4">
              Lists are created automatically when you start an inspection. This one began before that,
              so it can be created now from the records as they stand.
            </p>
            <Button
              onClick={() => generate.mutate()}
              disabled={generate.isPending}
              data-testid="button-generate-todos"
            >
              {generate.isPending ? <Loader2 className="w-4 h-4 mr-2 animate-spin" /> : null}
              Create the list
            </Button>
          </Card>
        )}

        {!isLoading && total > 0 && visible.length === 0 && (
          <Card className="p-6 text-center">
            <CheckCircle2 className="w-8 h-8 text-green-600 mx-auto mb-3" />
            <p className="font-medium">
              {filter === "outstanding" ? "Everything on the list is done" : "Nothing to show"}
            </p>
            {filter === "outstanding" && (
              <p className="text-sm text-muted-foreground mt-1">
                All {total} items have been updated in this inspection.
              </p>
            )}
          </Card>
        )}

        {grouped.map((group) =>
          group.rows.length === 0 ? null : (
            <section key={group.key} data-testid={`section-todo-${group.key}`}>
              <div className="flex items-baseline justify-between gap-2 mb-1">
                <h2 className="text-sm font-semibold flex items-center gap-1.5">
                  <group.icon className="w-4 h-4 text-muted-foreground shrink-0" />
                  {group.title}
                </h2>
                <span className="text-xs text-muted-foreground tabular-nums shrink-0">
                  {group.rows.length} of {group.totalInGroup}
                </span>
              </div>
              <p className="text-xs text-muted-foreground mb-2">{group.blurb}</p>

              <div className="space-y-2">
                {group.rows.map((item) => {
                  const isDone = !!item.doneAt;
                  const overdue = !!item.dueDate && item.dueDate < todayIso && !isDone;
                  return (
                    <Card
                      key={item.id}
                      className={`p-3 cursor-pointer hover-elevate active-elevate-2 ${isDone ? "opacity-60" : ""}`}
                      onClick={() => openForm(item)}
                      data-testid={`card-todo-${item.uid}`}
                    >
                      <div className="flex items-start gap-2.5">
                        {isDone ? (
                          <CheckCircle2 className="w-4 h-4 text-green-600 mt-0.5 shrink-0" />
                        ) : (
                          <Circle className="w-4 h-4 text-muted-foreground mt-0.5 shrink-0" />
                        )}

                        <div className="min-w-0 flex-1">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className={`font-mono text-sm font-semibold ${isDone ? "line-through" : ""}`}>
                              {item.uid}
                            </span>
                            {item.recordType === "observation" && (
                              <Badge variant="secondary" className="text-[10px] px-1.5 py-0">Observation</Badge>
                            )}
                            {item.assignedTo && (
                              <Badge variant="outline" className="text-[10px] px-1.5 py-0">{item.assignedTo}</Badge>
                            )}
                            {item.dueDate && (
                              <Badge
                                variant={overdue ? "destructive" : "outline"}
                                className="text-[10px] px-1.5 py-0"
                              >
                                {overdue ? "Overdue " : "Due "}{item.dueDate}
                              </Badge>
                            )}
                          </div>

                          {item.groupKey === "stranded" && item.sourceInspectionNumber && (
                            <p className="text-xs text-amber-600 dark:text-amber-500 mt-1 flex items-start gap-1">
                              <AlertTriangle className="w-3 h-3 mt-0.5 shrink-0" />
                              Still open in inspection {item.sourceInspectionNumber} and not carried into this one.
                              Opening it will offer to bring it forward.
                            </p>
                          )}

                          {item.comment && (
                            <p className="text-sm text-muted-foreground mt-1 line-clamp-2">{item.comment}</p>
                          )}
                          {item.actionText && (
                            <p className="text-sm mt-1 line-clamp-2">
                              <span className="text-muted-foreground">To do: </span>
                              {item.actionText}
                            </p>
                          )}

                          {isDone && (
                            <p className="text-xs text-green-700 dark:text-green-500 mt-1.5">
                              {item.doneReason === "completed" ? "Closed out" : "Updated"} in this inspection
                              {item.liveStatus === "complete" ? " · record complete" : ""}
                            </p>
                          )}
                        </div>

                        <ChevronRight className="w-4 h-4 text-muted-foreground mt-0.5 shrink-0" />
                      </div>
                    </Card>
                  );
                })}
              </div>
            </section>
          ),
        )}

        {total > 0 && (
          <p className="text-xs text-muted-foreground text-center pt-2">
            This list was fixed when the inspection started. Open an item, update its form and save —
            it ticks off here automatically.
          </p>
        )}
      </div>
    </div>
  );
}
