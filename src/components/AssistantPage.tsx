import { ReactNode, useState } from 'react';
import { Workflow } from '../types';
import { PageHeader } from './ui';
import { WorkflowList, WorkflowViewer } from './WorkflowViewer';
import { useSearchParams } from 'react-router-dom';

export function AssistantPage({ title, subtitle, icon, workflows }: { title: string; subtitle: string; icon: ReactNode; workflows: Workflow[] }) {
  const [params, setParams] = useSearchParams();
  const fromUrl = params.get('wf');
  const [selectedId, setSelectedId] = useState<string | null>(fromUrl ?? workflows[0]?.id ?? null);
  const selected = workflows.find((w) => w.id === selectedId) ?? null;

  return (
    <div>
      <PageHeader title={title} subtitle={subtitle} icon={icon} />
      <div className="grid gap-6 lg:grid-cols-[280px_1fr]">
        <div className="lg:sticky lg:top-20 lg:self-start lg:max-h-[calc(100vh-6rem)] lg:overflow-y-auto pr-1">
          <WorkflowList
            workflows={workflows}
            selectedId={selectedId}
            onSelect={(id) => { setSelectedId(id); setParams({ wf: id }, { replace: true }); }}
          />
        </div>
        <div>{selected && <WorkflowViewer workflow={selected} />}</div>
      </div>
    </div>
  );
}
