import { useQuery, useMutation } from "@tanstack/react-query";
import { queryClient, apiRequest } from "@/lib/queryClient";
import { Link } from "wouter";
import { Button } from "@/components/ui/button";
import { Card } from "@/components/ui/card";
import { Dialog, DialogContent, DialogHeader, DialogTitle, DialogTrigger } from "@/components/ui/dialog";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useToast } from "@/hooks/use-toast";
import { Plus, Building2, MapPin, User, Calendar, ChevronRight, Trash2 } from "lucide-react";
import type { Project } from "@shared/schema";
import { useState } from "react";

export default function ProjectList() {
  const { toast } = useToast();
  const [open, setOpen] = useState(false);
  const [form, setForm] = useState({ name: "", address: "", client: "", inspector: "", afcReference: "", revision: "01" });

  const { data: projects, isLoading } = useQuery<Project[]>({
    queryKey: ["/api/projects"],
  });

  const createMutation = useMutation({
    mutationFn: async (data: typeof form) => {
      const res = await apiRequest("POST", "/api/projects", {
        ...data,
        createdAt: new Date().toISOString(),
      });
      return res.json();
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/projects"] });
      setOpen(false);
      setForm({ name: "", address: "", client: "", inspector: "", afcReference: "", revision: "01" });
      toast({ title: "Project created" });
    },
  });

  const deleteMutation = useMutation({
    mutationFn: async (id: number) => {
      await apiRequest("DELETE", `/api/projects/${id}`);
    },
    onSuccess: () => {
      queryClient.invalidateQueries({ queryKey: ["/api/projects"] });
      toast({ title: "Project deleted" });
    },
  });

  return (
    <div className="max-w-3xl mx-auto px-4 py-8">
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-xl font-semibold tracking-tight" data-testid="text-page-title">
            Facade Defect Tracker
          </h1>
          <p className="text-sm text-muted-foreground mt-1">Site inspection projects</p>
        </div>
        <Dialog open={open} onOpenChange={setOpen}>
          <DialogTrigger asChild>
            <Button data-testid="button-new-project">
              <Plus className="w-4 h-4 mr-2" />
              New Project
            </Button>
          </DialogTrigger>
          <DialogContent>
            <DialogHeader>
              <DialogTitle>New Project</DialogTitle>
            </DialogHeader>
            <form
              onSubmit={(e) => {
                e.preventDefault();
                createMutation.mutate(form);
              }}
              className="space-y-4"
            >
              <div>
                <Label htmlFor="name">Project Name</Label>
                <Input
                  id="name"
                  placeholder="e.g. 123 George St — Facade Repair"
                  value={form.name}
                  onChange={(e) => setForm({ ...form, name: e.target.value })}
                  required
                  data-testid="input-project-name"
                />
              </div>
              <div>
                <Label htmlFor="address">Site Address</Label>
                <Input
                  id="address"
                  placeholder="123 George St, Sydney NSW 2000"
                  value={form.address}
                  onChange={(e) => setForm({ ...form, address: e.target.value })}
                  required
                  data-testid="input-project-address"
                />
              </div>
              <div>
                <Label htmlFor="client">Client</Label>
                <Input
                  id="client"
                  placeholder="Client name"
                  value={form.client}
                  onChange={(e) => setForm({ ...form, client: e.target.value })}
                  required
                  data-testid="input-project-client"
                />
              </div>
              <div>
                <Label htmlFor="inspector">Inspector</Label>
                <Input
                  id="inspector"
                  placeholder="Your name"
                  value={form.inspector}
                  onChange={(e) => setForm({ ...form, inspector: e.target.value })}
                  required
                  data-testid="input-project-inspector"
                />
              </div>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <Label htmlFor="afcReference">AFC Reference</Label>
                  <Input
                    id="afcReference"
                    placeholder="AFC-24XXX"
                    value={form.afcReference}
                    onChange={(e) => setForm({ ...form, afcReference: e.target.value })}
                    data-testid="input-project-afc-reference"
                  />
                </div>
                <div>
                  <Label htmlFor="revision">Revision</Label>
                  <Input
                    id="revision"
                    placeholder="01"
                    value={form.revision}
                    onChange={(e) => setForm({ ...form, revision: e.target.value })}
                    data-testid="input-project-revision"
                  />
                </div>
              </div>
              <Button type="submit" className="w-full" disabled={createMutation.isPending} data-testid="button-submit-project">
                {createMutation.isPending ? "Creating..." : "Create Project"}
              </Button>
            </form>
          </DialogContent>
        </Dialog>
      </div>

      {isLoading ? (
        <div className="space-y-3">
          {[1, 2, 3].map((i) => (
            <div key={i} className="h-24 rounded-lg bg-muted animate-pulse" />
          ))}
        </div>
      ) : !projects?.length ? (
        <div className="flex flex-col items-center justify-center py-20 text-center">
          <Building2 className="w-12 h-12 text-muted-foreground/40 mb-4" />
          <h2 className="text-lg font-medium mb-1">No projects yet</h2>
          <p className="text-sm text-muted-foreground mb-6 max-w-xs">
            Create your first inspection project to start tracking facade defects.
          </p>
          <Button onClick={() => setOpen(true)} data-testid="button-empty-new-project">
            <Plus className="w-4 h-4 mr-2" />
            New Project
          </Button>
        </div>
      ) : (
        <div className="space-y-3">
          {projects.map((project) => (
            <Card key={project.id} className="group relative" data-testid={`card-project-${project.id}`}>
              <Link href={`/projects/${project.id}`}>
                <div className="flex items-center justify-between p-4 cursor-pointer hover:bg-accent/50 rounded-lg transition-colors">
                  <div className="min-w-0 flex-1">
                    <h3 className="font-medium truncate">{project.name}</h3>
                    <div className="flex flex-wrap gap-x-4 gap-y-1 mt-1.5 text-sm text-muted-foreground">
                      <span className="flex items-center gap-1.5">
                        <MapPin className="w-3.5 h-3.5 shrink-0" />
                        <span className="truncate">{project.address}</span>
                      </span>
                      <span className="flex items-center gap-1.5">
                        <User className="w-3.5 h-3.5 shrink-0" />
                        {project.client}
                      </span>
                      <span className="flex items-center gap-1.5">
                        <Calendar className="w-3.5 h-3.5 shrink-0" />
                        {new Date(project.createdAt).toLocaleDateString()}
                      </span>
                    </div>
                  </div>
                  <ChevronRight className="w-5 h-5 text-muted-foreground shrink-0 ml-3" />
                </div>
              </Link>
              <button
                onClick={(e) => {
                  e.stopPropagation();
                  if (confirm("Delete this project and all its defects?")) {
                    deleteMutation.mutate(project.id);
                  }
                }}
                className="absolute top-3 right-12 p-1.5 rounded-md opacity-0 group-hover:opacity-100 hover:bg-destructive/10 hover:text-destructive transition-all"
                data-testid={`button-delete-project-${project.id}`}
              >
                <Trash2 className="w-4 h-4" />
              </button>
            </Card>
          ))}
        </div>
      )}
    </div>
  );
}
