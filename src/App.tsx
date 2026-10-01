import { FormEvent, useEffect, useMemo, useState } from 'react';
import {
  AlertCircle,
  CalendarDays,
  CheckCircle2,
  ChevronLeft,
  ChevronRight,
  CircleDot,
  Clock3,
  FolderKanban,
  GanttChartSquare,
  LayoutDashboard,
  ListTodo,
  Menu,
  Plus,
  Search,
  Settings2,
  Trash2,
  Users,
  Upload,
  Download,
  Database,
  X,
} from 'lucide-react';
import { api, ws } from './lib/client';
import * as XLSX from 'xlsx';

type TaskStatus =
  | 'Não iniciado'
  | 'Em andamento'
  | 'Aguardando'
  | 'Concluído'
  | 'Cancelado';
type TaskPriority = 'Crítica' | 'Alta' | 'Média' | 'Baixa';
type RecurrenceType = 'none' | 'weekly' | 'monthly' | 'quarterly' | 'custom';
type Occurrence = { id: string; dueDate: string; status: TaskStatus; progress: number; completedAt?: string };

type Task = {
  id: string;
  title: string;
  project: string;
  assignee: string;
  support: string;
  type: string;
  startDate: string;
  dueDate: string;
  status: TaskStatus;
  priority: TaskPriority;
  progress: number;
  notes: string;
  completedAt?: string;
  isRecurring: boolean;
  recurrenceType: RecurrenceType;
  recurrenceEndDate: string;
  customDates: string[];
  occurrences: Occurrence[];
  createdAt: number;
  updatedAt: number;
};

type ViewKey =
  | 'dashboard'
  | 'mywork'
  | 'activities'
  | 'projects'
  | 'gantt'
  | 'team'
  | 'calendar'
  | 'data';

type HierarchyMember = { id?: string; name: string; team: string; manager: string; role: string; level: number; active: boolean };

const statuses: TaskStatus[] = [
  'Não iniciado',
  'Em andamento',
  'Aguardando',
  'Concluído',
  'Cancelado',
];
const priorities: TaskPriority[] = ['Crítica', 'Alta', 'Média', 'Baixa'];
const people = ['Ailton', 'Leonardo', 'Monique', 'Isabela', 'Outro'];
const types = [
  'Projeto',
  'Rotina',
  'Análise / Relatório',
  'Reunião / Governança',
  'Processo',
  'Automação / Melhoria',
  'Demanda pontual',
];
const statusClass: Record<string, string> = {
  'Não iniciado': 'status status-not-started',
  'Em andamento': 'status status-progress',
  Aguardando: 'status status-waiting',
  Concluído: 'status status-done',
  Cancelado: 'status status-cancelled',
  Atrasado: 'status status-overdue',
  'Concluído com atraso': 'status status-done-late',
};
const priorityClass: Record<TaskPriority, string> = {
  Crítica: 'priority priority-critical',
  Alta: 'priority priority-high',
  Média: 'priority priority-medium',
  Baixa: 'priority priority-low',
};

const navItems: {
  key: ViewKey;
  label: string;
  icon: typeof LayoutDashboard;
}[] = [
  { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { key: 'mywork', label: 'Meu Trabalho', icon: CircleDot },
  { key: 'activities', label: 'Atividades', icon: ListTodo },
  { key: 'projects', label: 'Projetos', icon: FolderKanban },
  { key: 'gantt', label: 'Gantt', icon: GanttChartSquare },
  { key: 'team', label: 'Equipe', icon: Users },
  { key: 'calendar', label: 'Calendário', icon: CalendarDays },
  { key: 'data', label: 'Dados & Importação', icon: Database },
];

function parseDate(value: string) {
  const [year, month, day] = value.split('-').map(Number);
  return new Date(year, month - 1, day, 12, 0, 0);
}

function toDateInput(date: Date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

function formatDate(value: string) {
  if (!value) return '—';
  return parseDate(value).toLocaleDateString('pt-BR');
}

function taskOccurrences(task: Task): Occurrence[] { return task.isRecurring && task.occurrences?.length ? task.occurrences.slice().sort((a,b)=>a.dueDate.localeCompare(b.dueDate)) : [{id:'single',dueDate:task.dueDate,status:task.status,progress:task.progress,completedAt:task.completedAt}]; }
function nextOccurrence(task: Task) { const all=taskOccurrences(task); return all.find(x=>!['Concluído','Cancelado'].includes(x.status)) || all[all.length-1]; }
function dueDate(task: Task) { return nextOccurrence(task)?.dueDate || task.dueDate; }
function displayStatus(task: Task,today:string){const x=nextOccurrence(task);if(x.status==='Concluído'&&x.completedAt&&x.completedAt.slice(0,10)>x.dueDate)return 'Concluído com atraso';if(!['Concluído','Cancelado'].includes(x.status)&&x.dueDate<today)return 'Atrasado';return x.status}
function isLate(task:Task,today=toDateInput(new Date())){return displayStatus(task,today)==='Atrasado'}
function taskProgress(task:Task){const all=taskOccurrences(task);return task.isRecurring?Math.round(all.reduce((s,x)=>s+x.progress,0)/Math.max(1,all.length)):task.progress}
function recurrenceSummary(task:Task){if(!task.isRecurring)return '';const all=taskOccurrences(task),done=all.filter(x=>x.status==='Concluído').length,label={none:'',weekly:'Semanal',monthly:'Mensal',quarterly:'Trimestral',custom:'Datas específicas'}[task.recurrenceType];return '🔁 '+label+' · '+done+'/'+all.length+' concluídas'}
function daysUntil(value:string,today=toDateInput(new Date())){return Math.ceil((parseDate(value).getTime()-parseDate(today).getTime())/86400000)}

function weekStart(date: Date) {
  const result = new Date(date);
  const day = (result.getDay() + 6) % 7;
  result.setDate(result.getDate() - day);
  result.setHours(12, 0, 0, 0);
  return result;
}

function addDays(date: Date, days: number) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function emptyDraft(): Omit<Task, 'id' | 'createdAt' | 'updatedAt'> {
  return {
    title: '',
    project: '',
    assignee: '',
    support: '',
    type: 'Projeto',
    startDate: toDateInput(new Date()),
    dueDate: toDateInput(addDays(new Date(), 7)),
    status: 'Não iniciado',
    priority: 'Média',
    progress: 0,
    notes: '', completedAt: undefined, isRecurring: false, recurrenceType: 'none', recurrenceEndDate: toDateInput(addDays(new Date(),365)), customDates: [], occurrences: [],
  };
}

function PlanejaSymbol({ className = '' }: { className?: string }) {
  return (
    <svg
      className={'planeja-symbol ' + className}
      viewBox="0 0 220 92"
      role="img"
      aria-label="Símbolo Planeja Aê"
    >
      <line x1="24" y1="82" x2="194" y2="82" className="planeja-logo-base" />
      <circle cx="32" cy="82" r="12" className="planeja-logo-node" />
      <circle cx="82" cy="82" r="12" className="planeja-logo-node" />
      <circle cx="132" cy="82" r="12" className="planeja-logo-node" />
      <circle cx="184" cy="82" r="12" className="planeja-logo-node gold" />
      <circle cx="32" cy="82" r="4" className="planeja-logo-hole" />
      <circle cx="82" cy="82" r="4" className="planeja-logo-hole" />
      <circle cx="132" cy="82" r="4" className="planeja-logo-hole" />
      <circle cx="184" cy="82" r="4" className="planeja-logo-hole" />
      <path d="M26 72V59Q26 55 30 53L38 50V72Z" className="planeja-logo-bar small" />
      <path d="M72 72V42Q72 38 76 36L91 30V72Z" className="planeja-logo-bar medium" />
      <path d="M121 72V24Q121 20 125 18L145 9V72Z" className="planeja-logo-bar large" />
      <path d="M172 72V7Q172 3 176 1L199 -8V72Z" className="planeja-logo-bar gold tall" />
    </svg>
  );
}

function PlanejaMark() {
  return (
    <svg
      className="planeja-mark"
      viewBox="0 0 220 130"
      role="img"
      aria-label="Planeja Aê - Inteligência OFF"
    >
      <g className="planeja-logo-graph">
        <line x1="24" y1="82" x2="194" y2="82" className="planeja-logo-base" />
        <circle cx="32" cy="82" r="12" className="planeja-logo-node" />
        <circle cx="82" cy="82" r="12" className="planeja-logo-node" />
        <circle cx="132" cy="82" r="12" className="planeja-logo-node" />
        <circle cx="184" cy="82" r="12" className="planeja-logo-node gold" />
        <circle cx="32" cy="82" r="4" className="planeja-logo-hole" />
        <circle cx="82" cy="82" r="4" className="planeja-logo-hole" />
        <circle cx="132" cy="82" r="4" className="planeja-logo-hole" />
        <circle cx="184" cy="82" r="4" className="planeja-logo-hole" />
        <path d="M26 72V59Q26 55 30 53L38 50V72Z" className="planeja-logo-bar small" />
        <path d="M72 72V42Q72 38 76 36L91 30V72Z" className="planeja-logo-bar medium" />
        <path d="M121 72V24Q121 20 125 18L145 9V72Z" className="planeja-logo-bar large" />
        <path d="M172 72V7Q172 3 176 1L199 -8V72Z" className="planeja-logo-bar gold tall" />
      </g>
      <text x="18" y="112" className="planeja-logo-word">Planeja</text>
      <text x="137" y="112" className="planeja-logo-word gold-text">Aê</text>
      <text x="50" y="128" className="planeja-logo-sub">INTELIGÊNCIA OFF</text>
    </svg>
  );
}

function PlanejaBrand() {
  return (
    <div className="planeja-brand">
      <PlanejaMark />
    </div>
  );
}

function App() {
  const [view, setView] = useState<ViewKey>('dashboard');
  const [tasks, setTasks] = useState<Task[]>([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');
  const [search, setSearch] = useState('');
  const [statusFilter, setStatusFilter] = useState('Todos');
  const [assigneeFilter, setAssigneeFilter] = useState('Todos');
  const [priorityFilter, setPriorityFilter] = useState('Todos');
  const [currentPerson, setCurrentPerson] = useState('Ailton');
  const [modalOpen, setModalOpen] = useState(false);
  const [editingTask, setEditingTask] = useState<Task | null>(null);
  const [draft, setDraft] = useState(emptyDraft());
  const [formError, setFormError] = useState('');
  const [confirmDelete, setConfirmDelete] = useState(false);
  const [mobileMenu, setMobileMenu] = useState(false);
  const [ganttPeriod, setGanttPeriod] = useState<'q4' | '2027' | 'all'>('q4');
  const [calendarMonth, setCalendarMonth] = useState(new Date(2026, 9, 1, 12));
  const [lastSync, setLastSync] = useState<Date | null>(null);
  const [serverToday,setServerToday]=useState(toDateInput(new Date()));
  const [hierarchy,setHierarchy]=useState<HierarchyMember[]>([]);
  const people = useMemo(() => hierarchy.filter(item=>item.active!==false).map(item=>item.name), [hierarchy]);

  const loadTasks = async () => {
    try {
      const [response, hierarchyResponse] = await Promise.all([api.get('/api/tasks'), api.get('/api/hierarchy')]);
      setTasks(response.data.tasks as Task[]);
      setHierarchy(hierarchyResponse.data.hierarchy as HierarchyMember[]);
      if(response.data.today)setServerToday(response.data.today);
      setLastSync(new Date());
      setError('');
    } catch {
      setError('Não foi possível carregar as atividades. Tente novamente.');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    loadTasks();
    const connection = ws.connect();
    connection.onMessage(message => {
      if (
        message?.type === 'entity.update' &&
        message?.payload?.entity_type === 'workspace'
      ) {
        loadTasks();
      }
    });
    connection.onError(() =>
      setError(
        'A sincronização em tempo real foi interrompida. Os dados continuam disponíveis.'
      )
    );
    connection.ready
      .then(async () => {
        if (!connection.connectionId) return;
        await api.post('/api/subscriptions', {
          entity_type: 'workspace',
          entity_id: 'main',
          connection_id: connection.connectionId,
        });
      })
      .catch(() =>
        setError('Não foi possível iniciar a sincronização em tempo real.')
      );
    return () => {
      const id = connection.connectionId;
      if (id) {
        api
          .post('/api/subscriptions/remove', {
            entity_type: 'workspace',
            entity_id: 'main',
            connection_id: id,
          })
          .catch(() => undefined);
      }
      connection.disconnect();
    };
  }, []);

  const filteredTasks = useMemo(() => {
    const term = search.trim().toLowerCase();
    return tasks.filter(task => {
      const matchesSearch =
        !term ||
        [task.title, task.project, task.assignee, task.type].some(value =>
          value.toLowerCase().includes(term)
        );
      const matchesStatus =
        statusFilter === 'Todos' || displayStatus(task,serverToday) === statusFilter;
      const matchesAssignee =
        assigneeFilter === 'Todos' || task.assignee === assigneeFilter;
      const matchesPriority =
        priorityFilter === 'Todos' || task.priority === priorityFilter;
      return (
        matchesSearch && matchesStatus && matchesAssignee && matchesPriority
      );
    });
  }, [tasks, search, statusFilter, assigneeFilter, priorityFilter, serverToday]);

  const activeTasks=tasks.filter(task=>!['Concluído','Cancelado'].includes(displayStatus(task,serverToday)));
  const lateTasks=activeTasks.filter(task=>isLate(task,serverToday));
  const nextSeven=activeTasks.filter(task=>{const days=daysUntil(dueDate(task),serverToday);return days>=0&&days<=7;});
  const doneThisMonth = tasks.filter(
    task =>
      task.status === 'Concluído' &&
      parseDate(task.dueDate).getMonth() === new Date().getMonth()
  ).length;
  const myTasks = tasks
    .filter(
      task =>
        task.assignee === currentPerson &&
        !['Concluído', 'Cancelado'].includes(task.status)
    )
    .sort(
      (a, b) => parseDate(a.dueDate).getTime() - parseDate(b.dueDate).getTime()
    );

  const openNew = () => {
    setEditingTask(null);
    setDraft({...emptyDraft(), assignee: people[0] || ''});
    setFormError('');
    setConfirmDelete(false);
    setModalOpen(true);
  };

  const openEdit = (task: Task) => {
    setEditingTask(task);
    setDraft({
      title: task.title,
      project: task.project,
      assignee: task.assignee,
      support: task.support || '',
      type: task.type,
      startDate: task.startDate,
      dueDate: task.dueDate,
      status: task.status,
      priority: task.priority,
      progress: task.progress,
      notes: task.notes || '', completedAt:task.completedAt,isRecurring:!!task.isRecurring,recurrenceType:task.recurrenceType||'none',recurrenceEndDate:task.recurrenceEndDate||'',customDates:task.customDates||[],occurrences:task.occurrences||[],
    });
    setFormError('');
    setConfirmDelete(false);
    setModalOpen(true);
  };

  const submitTask = async (event: FormEvent) => {
    event.preventDefault();
    if (!draft.title.trim()) {
      setFormError('Informe o nome da atividade.');
      return;
    }
    if (!draft.project.trim()) {
      setFormError('Informe o projeto ou frente.');
      return;
    }
    if (parseDate(draft.dueDate) < parseDate(draft.startDate)) { setFormError('O prazo não pode ser anterior à data de início.'); return; }
    if(draft.isRecurring&&draft.recurrenceType!=='custom'&&(!draft.recurrenceEndDate||draft.recurrenceEndDate<draft.dueDate)){setFormError('Informe até quando essa recorrência deve acontecer.');return;}
    setSaving(true);
    setFormError('');
    try {
      if (editingTask) {
        const response = await api.put('/api/tasks/' + editingTask.id, draft);
        setTasks(current =>
          current.map(item =>
            item.id === editingTask.id ? (response.data.task as Task) : item
          )
        );
      } else {
        const response = await api.post('/api/tasks', draft);
        setTasks(current => [response.data.task as Task, ...current]);
      }
      setModalOpen(false);
      setError('');
    } catch {
      setFormError(
        'Não foi possível salvar. Confira os dados e tente novamente.'
      );
    } finally {
      setSaving(false);
    }
  };

  const updateOccurrence=async(taskId:string,occurrenceId:string,status:TaskStatus,progress:number)=>{setSaving(true);setFormError('');try{const response=await api.put('/api/tasks/'+taskId+'/occurrences/'+occurrenceId,{status,progress});const updated=response.data.task as Task;setTasks(current=>current.map(x=>x.id===taskId?updated:x));setEditingTask(updated)}catch{setFormError('Não foi possível atualizar esta ocorrência. Tente novamente.')}finally{setSaving(false)}};

  const deleteTask = async () => {
    if (!editingTask) return;
    setSaving(true);
    try {
      await api.delete('/api/tasks/' + editingTask.id);
      setTasks(current => current.filter(item => item.id !== editingTask.id));
      setModalOpen(false);
      setConfirmDelete(false);
    } catch {
      setFormError('Não foi possível excluir a atividade.');
    } finally {
      setSaving(false);
    }
  };

  const projectCards = useMemo(() => {
    const map = new Map<string, Task[]>();
    tasks.forEach(task => {
      const current = map.get(task.project) || [];
      current.push(task);
      map.set(task.project, current);
    });
    return Array.from(map.entries())
      .map(([name, items]) => {
        const progress = Math.round(
          items.reduce((sum, item) => sum + item.progress, 0) / items.length
        );
        return {
          name,
          items,
          progress,
          late: items.filter(isLate).length,
          nextDue: items
            .slice()
            .sort(
              (a, b) =>
                parseDate(a.dueDate).getTime() - parseDate(b.dueDate).getTime()
            )[0]?.dueDate,
        };
      })
      .sort((a, b) => b.items.length - a.items.length);
  }, [tasks]);

  const teamCards = useMemo(
    () =>
      people
        .filter(
          person =>
            person !== 'Outro' || tasks.some(task => task.assignee === 'Outro')
        )
        .map(person => {
          const items = tasks.filter(task => task.assignee === person);
          const open = items.filter(
            task => !['Concluído', 'Cancelado'].includes(task.status)
          );
          const progress = items.length
            ? Math.round(
                items.reduce((sum, item) => sum + item.progress, 0) /
                  items.length
              )
            : 0;
          return {
            person,
            items,
            open,
            progress,
            late: open.filter(isLate).length,
          };
        }),
    [tasks]
  );

  const pageTitle: Record<ViewKey, string> = {
    dashboard: 'Visão Geral',
    mywork: 'Meu Trabalho',
    activities: 'Atividades',
    projects: 'Projetos',
    gantt: 'Gantt de Atividades',
    team: 'Equipe',
    calendar: 'Calendário',
    data: 'Dados & Importação',
  };

  const navigate = (key: ViewKey) => {
    setView(key);
    setMobileMenu(false);
  };

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand">
          <PlanejaBrand />
        </div>
        <nav>
          {navItems.map(item => {
            const Icon = item.icon;
            return (
              <button
                key={item.key}
                className={view === item.key ? 'nav-item active' : 'nav-item'}
                onClick={() => navigate(item.key)}
              >
                <Icon size={19} />
                <span>{item.label}</span>
              </button>
            );
          })}
        </nav>
        <div className="sidebar-footer">
          <Settings2 size={18} />
          <span>Planeja Aê · V1 compartilhada</span>
        </div>
      </aside>

      <main className="main">
        <div className="brand-watermark" aria-hidden="true"><PlanejaSymbol /></div>
        <header className="topbar">
          <button
            className="icon-btn mobile-menu-btn"
            onClick={() => setMobileMenu(!mobileMenu)}
            aria-label="Abrir menu"
          >
            <Menu size={22} />
          </button>
          <div className="mobile-logo"><PlanejaSymbol /></div>
          <div className="page-heading">
            <span className="eyebrow">PLANEJA AÊ · INTELIGÊNCIA OFF</span>
            <h1>{view === 'dashboard' ? 'Olá, Ailton 👋' : pageTitle[view]}</h1>
          </div>
          <div className="top-actions">
            <label className="search-box">
              <Search size={18} />
              <input
                aria-label="Buscar atividades"
                value={search}
                onChange={e => setSearch(e.target.value)}
                placeholder="Buscar atividades, projetos ou pessoas"
              />
            </label>
            <button className="primary-btn" onClick={openNew}>
              <Plus size={18} />
              <span>Nova atividade</span>
            </button>
          </div>
        </header>

        {mobileMenu && (
          <div className="mobile-menu">
            {navItems.map(item => {
              const Icon = item.icon;
              return (
                <button
                  key={item.key}
                  className={view === item.key ? 'active' : ''}
                  onClick={() => navigate(item.key)}
                >
                  <Icon size={18} />
                  {item.label}
                </button>
              );
            })}
          </div>
        )}

        {error && (
          <div className="alert">
            <AlertCircle size={18} />
            <span>{error}</span>
            <button onClick={() => setError('')} aria-label="Fechar aviso">
              <X size={17} />
            </button>
          </div>
        )}

        <section className="content">
          {loading ? (
            <LoadingState />
          ) : (
            <>
              {view === 'dashboard' && (
                <Dashboard
                  tasks={tasks}
                  active={activeTasks.length}
                  late={lateTasks.length}
                  nextSeven={nextSeven.length}
                  doneThisMonth={doneThisMonth}
                  projects={projectCards}
                  team={teamCards}
                  onOpenTask={openEdit}
                />
              )}
              {view === 'mywork' && (
                <MyWork
                  tasks={myTasks}
                  person={currentPerson}
                  setPerson={setCurrentPerson}
                  onOpenTask={openEdit}
                />
              )}
              {view === 'activities' && (
                <Activities
                  tasks={filteredTasks}
                  statusFilter={statusFilter}
                  setStatusFilter={setStatusFilter}
                  assigneeFilter={assigneeFilter}
                  setAssigneeFilter={setAssigneeFilter}
                  priorityFilter={priorityFilter}
                  setPriorityFilter={setPriorityFilter}
                  onOpenTask={openEdit}
                  onNew={openNew}
                  serverToday={serverToday}
                />
              )}
              {view === 'projects' && (
                <Projects projects={projectCards} onOpenTask={openEdit} />
              )}
              {view === 'gantt' && (
                <Gantt
                  tasks={filteredTasks}
                  period={ganttPeriod}
                  setPeriod={setGanttPeriod}
                  onOpenTask={openEdit}
                />
              )}
              {view === 'team' && (
                <Team cards={teamCards} onOpenTask={openEdit} />
              )}
              {view === 'calendar' && (
                <CalendarView
                  tasks={tasks}
                  month={calendarMonth}
                  setMonth={setCalendarMonth}
                  onOpenTask={openEdit}
                />
              )}
              {view === 'data' && (
                <DataManager tasks={tasks} hierarchy={hierarchy} onRefresh={loadTasks} />
              )}
            </>
          )}
        </section>

        <div className="sync-note">
          {lastSync
            ? 'Sincronizado às ' +
              lastSync.toLocaleTimeString('pt-BR', {
                hour: '2-digit',
                minute: '2-digit',
              })
            : 'Sincronizando...'}
        </div>
      </main>

      <nav className="bottom-nav">
        {navItems.slice(0, 5).map(item => {
          const Icon = item.icon;
          return (
            <button
              key={item.key}
              className={view === item.key ? 'active' : ''}
              onClick={() => navigate(item.key)}
            >
              <Icon size={19} />
              <span>{item.label === 'Meu Trabalho' ? 'Meu' : item.label}</span>
            </button>
          );
        })}
      </nav>

      {modalOpen && (
        <TaskModal
          draft={draft}
          setDraft={setDraft}
          people={people}
          editingTask={editingTask}
          formError={formError}
          saving={saving}
          confirmDelete={confirmDelete}
          setConfirmDelete={setConfirmDelete}
          onSubmit={submitTask}
          onDelete={deleteTask}
          onUpdateOccurrence={updateOccurrence}
          serverToday={serverToday}
          onClose={() => setModalOpen(false)}
        />
      )}
    </div>
  );
}

function LoadingState() {
  return (
    <div className="loading-grid">
      <div className="skeleton big" />
      <div className="skeleton big" />
      <div className="skeleton wide" />
    </div>
  );
}

function MetricCard({
  label,
  value,
  icon,
  tone,
}: {
  label: string;
  value: number | string;
  icon: React.ReactNode;
  tone: string;
}) {
  return (
    <div className={'metric-card ' + tone}>
      <div className="metric-icon">{icon}</div>
      <div>
        <span>{label}</span>
        <strong>{value}</strong>
      </div>
    </div>
  );
}

function Dashboard({
  tasks,
  active,
  late,
  nextSeven,
  doneThisMonth,
  projects,
  team,
  onOpenTask,
}: {
  tasks: Task[];
  active: number;
  late: number;
  nextSeven: number;
  doneThisMonth: number;
  projects: { name: string; items: Task[]; progress: number; late: number; nextDue?: string }[];
  team: { person: string; items: Task[]; open: Task[]; progress: number; late: number }[];
  onOpenTask: (task: Task) => void;
}) {
  const upcoming = tasks
    .filter(task => !['Concluído', 'Cancelado'].includes(task.status))
    .sort(
      (a, b) => parseDate(a.dueDate).getTime() - parseDate(b.dueDate).getTime()
    )
    .slice(0, 6);
  return (
    <div className="stack dashboard-home">
      <section className="welcome-hero">
        <div className="welcome-copy">
          <span className="hero-kicker">PLANEJA AÊ · INTELIGÊNCIA OFF</span>
          <h2>Da ideia à conquista, <strong>com mais foco.</strong></h2>
          <p>Organize projetos, alinhe o time e transforme planos em resultados.</p>
        </div>
        <div className="hero-principles">
          <div><b>01</b><span>Mais foco</span><small>menos dispersão</small></div>
          <div><b>02</b><span>Mais resultados</span><small>planos que saem do papel</small></div>
          <div><b>03</b><span>Mais alinhamento</span><small>time evoluindo junto</small></div>
        </div>
      </section>
      <div className="metrics-grid">
        <MetricCard
          label="Abertas"
          value={active}
          icon={<ListTodo />}
          tone="blue"
        />
        <MetricCard
          label="Atrasadas"
          value={late}
          icon={<Clock3 />}
          tone="red"
        />
        <MetricCard
          label="Próximos 7 dias"
          value={nextSeven}
          icon={<CalendarDays />}
          tone="amber"
        />
        <MetricCard
          label="Concluídas no mês"
          value={doneThisMonth}
          icon={<CheckCircle2 />}
          tone="green"
        />
      </div>
      <div className="dashboard-grid">
        <section className="panel">
          <div className="panel-head">
            <div>
              <span className="eyebrow">PRÓXIMAS ENTREGAS</span>
              <h2>O que merece atenção</h2>
            </div>
          </div>
          <div className="task-list compact">
            {upcoming.map(task => (
              <TaskRow
                key={task.id}
                task={task}
                onClick={() => onOpenTask(task)}
              />
            ))}
          </div>
        </section>
        <section className="panel">
          <div className="panel-head">
            <div>
              <span className="eyebrow">PROJETOS</span>
              <h2>Andamento por frente</h2>
            </div>
          </div>
          <div className="progress-list">
            {projects.slice(0, 6).map((project: any) => (
              <div key={project.name} className="progress-item">
                <div>
                  <strong>{project.name}</strong>
                  <span>{project.items.length} atividades</span>
                </div>
                <div className="progress-meta">
                  <span>{project.progress}%</span>
                  <div className="progress-track">
                    <i style={{ width: project.progress + '%' }} />
                  </div>
                </div>
              </div>
            ))}
          </div>
        </section>
        <section className="panel span-2">
          <div className="panel-head">
            <div>
              <span className="eyebrow">CAPACIDADE</span>
              <h2>Carga aberta por executor</h2>
            </div>
          </div>
          <div className="team-strip">
            {team.map((member: any) => (
              <div className="member-card" key={member.person}>
                <div className="avatar">
                  {member.person.slice(0, 2).toUpperCase()}
                </div>
                <div>
                  <strong>{member.person}</strong>
                  <span>{member.open.length} abertas</span>
                </div>
                <b className={member.late ? 'danger-text' : ''}>
                  {member.late ? member.late + ' atrasada(s)' : 'No prazo'}
                </b>
              </div>
            ))}
          </div>
        </section>
      </div>
    </div>
  );
}

function MyWork({
  tasks,
  person,
  setPerson,
  onOpenTask,
}: {
  tasks: Task[];
  person: string;
  setPerson: (value: string) => void;
  onOpenTask: (task: Task) => void;
}) {
  return (
    <div className="stack">
      <div className="section-toolbar">
        <div>
          <span className="eyebrow">FOCO PESSOAL</span>
          <h2>Prioridades de {person}</h2>
        </div>
        <label className="field-inline">
          Visão de
          <select value={person} onChange={e => setPerson(e.target.value)}>
            {people.map(p => (
              <option key={p}>{p}</option>
            ))}
          </select>
        </label>
      </div>
      <div className="mywork-grid">
        {tasks.map(task => (
          <button
            className={'focus-card ' + (isLate(task) ? 'late' : '')}
            key={task.id}
            onClick={() => onOpenTask(task)}
          >
            <div className="focus-top">
              <span className={priorityClass[task.priority]}>
                {task.priority}
              </span>
              <span>{formatDate(task.dueDate)}</span>
            </div>
            <h3>{task.title}</h3>
            <p>{task.project}</p>
            <div className="progress-track">
              <i style={{ width: task.progress + '%' }} />
            </div>
            <div className="focus-bottom">
              <span className={statusClass[task.status]}>{task.status}</span>
              <strong>{task.progress}%</strong>
            </div>
          </button>
        ))}
        {!tasks.length && (
          <Empty
            title="Nada aberto por aqui"
            text="As tarefas atribuídas a esta pessoa aparecerão nesta visão."
          />
        )}
      </div>
    </div>
  );
}

function Activities({
  tasks,
  statusFilter,
  setStatusFilter,
  assigneeFilter,
  setAssigneeFilter,
  priorityFilter,
  setPriorityFilter,
  onOpenTask,
  onNew,
  serverToday,
}: any) {
  return (
    <div className="stack">
      <div className="filters">
        <label>
          Status
          <select
            value={statusFilter}
            onChange={e => setStatusFilter(e.target.value)}
          >
            <option>Todos</option>
            {statuses.map(value => (
              <option key={value}>{value}</option>
            ))}
          </select>
        </label>
        <label>
          Executor
          <select
            value={assigneeFilter}
            onChange={e => setAssigneeFilter(e.target.value)}
          >
            <option>Todos</option>
            {people.map(value => (
              <option key={value}>{value}</option>
            ))}
          </select>
        </label>
        <label>
          Prioridade
          <select
            value={priorityFilter}
            onChange={e => setPriorityFilter(e.target.value)}
          >
            <option>Todos</option>
            {priorities.map(value => (
              <option key={value}>{value}</option>
            ))}
          </select>
        </label>
        <span className="result-count">{tasks.length} resultado(s)</span>
      </div>
      <section className="panel table-panel">
        <div className="desktop-table-wrap">
          <table>
            <thead>
              <tr>
                <th>Atividade</th>
                <th>Projeto / Frente</th>
                <th>Executor</th>
                <th>Prazo</th>
                <th>Status</th>
                <th>Prioridade</th>
                <th>Progresso</th>
              </tr>
            </thead>
            <tbody>
              {tasks.map((task: Task) => (
                <tr key={task.id} onClick={() => onOpenTask(task)}>
                  <td>
                    <strong>{task.title}</strong>
                    <small>{task.type}{task.isRecurring?' · '+recurrenceSummary(task):''}</small>
                  </td>
                  <td>{task.project}</td>
                  <td>{task.assignee}</td>
                  <td className={isLate(task,serverToday)?'danger-text':''}>{formatDate(dueDate(task))}</td>
                  <td><span className={statusClass[displayStatus(task,serverToday)]}>{displayStatus(task,serverToday)}</span>
                  </td>
                  <td>
                    <span className={priorityClass[task.priority]}>
                      {task.priority}
                    </span>
                  </td>
                  <td>
                    <div className="table-progress">
                      <span>{taskProgress(task)}%</span>
                      <div className="progress-track">
                        <i style={{ width: taskProgress(task) + '%' }} />
                      </div>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <div className="mobile-task-list">
          {tasks.map((task: Task) => (
            <TaskRow
              key={task.id}
              task={task}
              onClick={() => onOpenTask(task)}
            />
          ))}
        </div>
        {!tasks.length && (
          <Empty
            title="Nenhuma atividade encontrada"
            text="Ajuste os filtros ou crie uma nova atividade."
            action="Nova atividade"
            onAction={onNew}
          />
        )}
      </section>
    </div>
  );
}

function Projects({ projects, onOpenTask }: any) {
  return (
    <div className="project-grid">
      {projects.map((project: any) => (
        <section className="project-card" key={project.name}>
          <div className="project-card-head">
            <div>
              <span className="eyebrow">PROJETO</span>
              <h2>{project.name}</h2>
            </div>
            <strong>{project.progress}%</strong>
          </div>
          <div className="progress-track large">
            <i style={{ width: project.progress + '%' }} />
          </div>
          <div className="project-stats">
            <span>
              <b>{project.items.length}</b> atividades
            </span>
            <span className={project.late ? 'danger-text' : ''}>
              <b>{project.late}</b> atrasadas
            </span>
            <span>
              <b>{formatDate(project.nextDue)}</b> próxima
            </span>
          </div>
          <div className="task-list mini">
            {project.items.slice(0, 4).map((task: Task) => (
              <TaskRow
                key={task.id}
                task={task}
                onClick={() => onOpenTask(task)}
              />
            ))}
          </div>
        </section>
      ))}
    </div>
  );
}

function Gantt({
  tasks,
  period,
  setPeriod,
  onOpenTask,
}: {
  tasks: Task[];
  period: 'q4' | '2027' | 'all';
  setPeriod: (value: 'q4' | '2027' | 'all') => void;
  onOpenTask: (task: Task) => void;
}) {
  const config =
    period === 'q4'
      ? { start: new Date(2026, 8, 28, 12), weeks: 14 }
      : period === '2027'
        ? { start: new Date(2027, 0, 4, 12), weeks: 52 }
        : { start: new Date(2026, 8, 28, 12), weeks: 66 };
  const weeks = Array.from({ length: config.weeks }, (_, index) =>
    addDays(config.start, index * 7)
  );
  const visibleTasks = tasks.filter(
    task =>
      parseDate(task.startDate) <= addDays(config.start, config.weeks * 7) &&
      parseDate(task.dueDate) >= config.start
  );
  const monthGroups: { label: string; start: number; span: number }[] = [];
  weeks.forEach((week, index) => {
    const label = week
      .toLocaleDateString('pt-BR', { month: 'short', year: 'numeric' })
      .replace('.', '');
    const last = monthGroups[monthGroups.length - 1];
    if (last && last.label === label) last.span += 1;
    else monthGroups.push({ label, start: index + 1, span: 1 });
  });
  const getSpan = (task: Task) => {
    const ganttStart = weekStart(config.start).getTime();
    const start = Math.max(
      0,
      Math.floor(
        (weekStart(parseDate(task.startDate)).getTime() - ganttStart) /
          604800000
      )
    );
    const end = Math.min(
      config.weeks - 1,
      Math.floor(
        (weekStart(parseDate(task.dueDate)).getTime() - ganttStart) / 604800000
      )
    );
    return { start, end };
  };
  return (
    <div className="stack">
      <div className="section-toolbar">
        <div>
          <span className="eyebrow">PLANEJAMENTO</span>
          <h2>Linha do tempo compartilhada</h2>
        </div>
        <div className="segmented">
          <button
            className={period === 'q4' ? 'active' : ''}
            onClick={() => setPeriod('q4')}
          >
            Q4 2026
          </button>
          <button
            className={period === '2027' ? 'active' : ''}
            onClick={() => setPeriod('2027')}
          >
            2027
          </button>
          <button
            className={period === 'all' ? 'active' : ''}
            onClick={() => setPeriod('all')}
          >
            Tudo
          </button>
        </div>
      </div>
      <section className="panel gantt-panel">
        <div className="gantt-scroll">
          <div
            className="gantt-board"
            style={{
              gridTemplateColumns: 'var(--gantt-name-width) repeat(' + config.weeks + ', 46px)',
            }}
          >
            <div className="gantt-fixed gantt-header task-head">Atividade</div>
            <div
              className="month-row"
              style={{
                gridColumn: '2 / span ' + config.weeks,
                gridTemplateColumns: 'repeat(' + config.weeks + ', 46px)',
              }}
            >
              {monthGroups.map(group => (
                <div
                  key={group.label + group.start}
                  className="month-label"
                  style={{ gridColumn: group.start + ' / span ' + group.span }}
                >
                  {group.label}
                </div>
              ))}
            </div>
            <div className="gantt-fixed week-spacer" />
            <div
              className="week-row"
              style={{
                gridColumn: '2 / span ' + config.weeks,
                gridTemplateColumns: 'repeat(' + config.weeks + ', 46px)',
              }}
            >
              {weeks.map(week => (
                <div key={week.toISOString()}>
                  {week.toLocaleDateString('pt-BR', {
                    day: '2-digit',
                    month: '2-digit',
                  })}
                </div>
              ))}
            </div>
            {visibleTasks.map(task => {
              const span = getSpan(task);
              return (
                <div className="gantt-task-fragment" key={task.id}>
                  <button
                    className="gantt-task-name"
                    onClick={() => onOpenTask(task)}
                  >
                    <strong>{task.title}</strong>
                    <span>
                      {task.assignee} · {task.progress}%
                    </span>
                  </button>
                  <div
                    className="gantt-track"
                    style={{
                      gridTemplateColumns: 'repeat(' + config.weeks + ', 46px)',
                    }}
                  >
                    {weeks.map(week => (
                      <i key={week.toISOString()} />
                    ))}
                    <button
                      title={task.title}
                      aria-label={'Abrir ' + task.title}
                      className={
                        'gantt-bar ' +
                        (isLate(task)
                          ? 'late'
                          : task.status === 'Concluído'
                            ? 'done'
                            : task.status === 'Aguardando'
                              ? 'waiting'
                              : task.status === 'Não iniciado'
                                ? 'not-started'
                                : '')
                      }
                      style={{
                        gridColumn: span.start + 1 + ' / ' + (span.end + 2),
                      }}
                      onClick={() => onOpenTask(task)}
                    >
                      <span>{task.progress}%</span>
                    </button>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </section>
    </div>
  );
}

function Team({ cards, onOpenTask }: any) {
  return (
    <div className="team-grid">
      {cards.map((member: any) => (
        <section className="panel person-panel" key={member.person}>
          <div className="person-head">
            <div className="avatar large">
              {member.person.slice(0, 2).toUpperCase()}
            </div>
            <div>
              <h2>{member.person}</h2>
              <span>
                {member.open.length} abertas · {member.progress}% médio
              </span>
            </div>
            <b className={member.late ? 'danger-text' : ''}>
              {member.late} atrasada(s)
            </b>
          </div>
          <div className="progress-track large">
            <i
              style={{ width: Math.min(100, member.open.length * 14) + '%' }}
            />
          </div>
          <div className="task-list mini">
            {member.open.slice(0, 5).map((task: Task) => (
              <TaskRow
                key={task.id}
                task={task}
                onClick={() => onOpenTask(task)}
              />
            ))}
          </div>
        </section>
      ))}
    </div>
  );
}

function CalendarView({
  tasks,
  month,
  setMonth,
  onOpenTask,
}: {
  tasks: Task[];
  month: Date;
  setMonth: (date: Date) => void;
  onOpenTask: (task: Task) => void;
}) {
  const first = new Date(month.getFullYear(), month.getMonth(), 1, 12);
  const startOffset = (first.getDay() + 6) % 7;
  const daysInMonth = new Date(
    month.getFullYear(),
    month.getMonth() + 1,
    0
  ).getDate();
  const cells = Array.from(
    { length: 42 },
    (_, index) => index - startOffset + 1
  );
  const monthLabel = month.toLocaleDateString('pt-BR', {
    month: 'long',
    year: 'numeric',
  });
  return (
    <div className="stack">
      <div className="section-toolbar">
        <div>
          <span className="eyebrow">AGENDA DE ENTREGAS</span>
          <h2 className="capitalize">{monthLabel}</h2>
        </div>
        <div className="calendar-nav">
          <button
            className="icon-btn"
            onClick={() =>
              setMonth(
                new Date(month.getFullYear(), month.getMonth() - 1, 1, 12)
              )
            }
          >
            <ChevronLeft />
          </button>
          <button
            className="icon-btn"
            onClick={() =>
              setMonth(
                new Date(month.getFullYear(), month.getMonth() + 1, 1, 12)
              )
            }
          >
            <ChevronRight />
          </button>
        </div>
      </div>
      <section className="panel calendar-panel">
        <div className="calendar-weekdays">
          {['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SÁB', 'DOM'].map(day => (
            <span key={day}>{day}</span>
          ))}
        </div>
        <div className="calendar-grid">
          {cells.map((day, index) => {
            const valid = day >= 1 && day <= daysInMonth;
            const dateKey = valid
              ? toDateInput(
                  new Date(month.getFullYear(), month.getMonth(), day, 12)
                )
              : '';
            const due=valid?tasks.flatMap(task=>taskOccurrences(task).filter(occ=>occ.dueDate===dateKey).map(occ=>({task,occ}))):[];
            return (
              <div
                className={valid ? 'calendar-cell' : 'calendar-cell muted'}
                key={index}
              >
                {valid && (
                  <>
                    <b>{day}</b>
                    <div className="calendar-events">
                      {due.slice(0,3).map(({task,occ})=>(
                        <button key={task.id+occ.id} className={!['Concluído','Cancelado'].includes(occ.status)&&occ.dueDate<toDateInput(new Date())?'late':''} onClick={()=>onOpenTask(task)}>
                          {task.title}
                        </button>
                      ))}
                      {due.length > 3 && <span>+{due.length - 3}</span>}
                    </div>
                  </>
                )}
              </div>
            );
          })}
        </div>
      </section>
    </div>
  );
}

function TaskRow({ task, onClick }: { task: Task; onClick: () => void }) {
  return (
    <button className="task-row" onClick={onClick}>
      <div className="task-row-main">
        <strong>{task.title}</strong>
        <span>
          {task.project} · {task.assignee}
        </span>
      </div>
      <div className="task-row-side">
        <span className={isLate(task) ? 'due late' : 'due'}>
          {formatDate(task.dueDate)}
        </span>
        <span className={statusClass[task.status]}>{task.status}</span>
      </div>
    </button>
  );
}

function Empty({
  title,
  text,
  action,
  onAction,
}: {
  title: string;
  text: string;
  action?: string;
  onAction?: () => void;
}) {
  return (
    <div className="empty">
      <ListTodo size={28} />
      <h3>{title}</h3>
      <p>{text}</p>
      {action && (
        <button className="secondary-btn" onClick={onAction}>
          {action}
        </button>
      )}
    </div>
  );
}

function TaskModal({
  draft,
  setDraft,
  people,
  editingTask,
  formError,
  saving,
  confirmDelete,
  setConfirmDelete,
  onSubmit,
  onDelete,
  onUpdateOccurrence,
  serverToday,
  onClose,
}: any) {
  const update=(field:string,value:any)=>setDraft((current:any)=>({...current,[field]:value}));
  return (
    <div className="modal-backdrop" role="presentation">
      <div
        className="modal"
        role="dialog"
        aria-modal="true"
        aria-labelledby="task-modal-title"
      >
        <div className="modal-head">
          <div>
            <span className="eyebrow">{editingTask ? 'EDITAR' : 'NOVA'}</span>
            <h2 id="task-modal-title">
              {editingTask ? 'Editar atividade' : 'Nova atividade'}
            </h2>
          </div>
          <button className="icon-btn" onClick={onClose} aria-label="Fechar">
            <X />
          </button>
        </div>
        {confirmDelete ? (
          <div className="confirm-box">
            <AlertCircle />
            <h3>Excluir esta atividade?</h3>
            <p>Essa ação remove a atividade da base compartilhada.</p>
            <div>
              <button
                className="secondary-btn"
                onClick={() => setConfirmDelete(false)}
              >
                Cancelar
              </button>
              <button
                className="danger-btn"
                onClick={onDelete}
                disabled={saving}
              >
                Excluir
              </button>
            </div>
          </div>
        ) : (
          <form onSubmit={onSubmit}>
            {formError && (
              <div className="form-error">
                <AlertCircle size={17} />
                {formError}
              </div>
            )}
            <label className="field field-full">
              Atividade
              <input
                value={draft.title}
                onChange={e => update('title', e.target.value)}
                placeholder="Ex.: Atualizar Scorecard OFF"
              />
            </label>
            <div className="form-grid">
              <label className="field">
                Projeto / Frente
                <input
                  value={draft.project}
                  onChange={e => update('project', e.target.value)}
                  placeholder="Ex.: Scorecard"
                />
              </label>
              <label className="field">
                Tipo
                <select
                  value={draft.type}
                  onChange={e => update('type', e.target.value)}
                >
                  {types.map(value => (
                    <option key={value}>{value}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                Executor
                <select
                  value={draft.assignee}
                  onChange={e => update('assignee', e.target.value)}
                >
                  {people.map(value => (
                    <option key={value}>{value}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                Apoio
                <input
                  value={draft.support}
                  onChange={e => update('support', e.target.value)}
                  placeholder="Opcional"
                />
              </label>
              <label className="field">
                Início
                <input
                  type="date"
                  value={draft.startDate}
                  onChange={e => update('startDate', e.target.value)}
                />
              </label>
              <label className="field">
                {draft.isRecurring?'Primeira entrega':'Prazo'}
                <input
                  type="date"
                  value={draft.dueDate}
                  onChange={e => update('dueDate', e.target.value)}
                />
              </label>
              <label className="field field-full recurrence-toggle"><span>Atividade recorrente</span><input type="checkbox" checked={draft.isRecurring} onChange={e=>setDraft((c:any)=>({...c,isRecurring:e.target.checked,recurrenceType:e.target.checked&&c.recurrenceType==='none'?'monthly':c.recurrenceType}))}/></label>
              {draft.isRecurring&&<><label className="field">Recorrência<select value={draft.recurrenceType} onChange={e=>update('recurrenceType',e.target.value)}><option value="weekly">Semanal</option><option value="monthly">Mensal</option><option value="quarterly">Trimestral</option><option value="custom">Datas específicas</option></select></label>{draft.recurrenceType!=='custom'?<label className="field">Repetir até<input type="date" value={draft.recurrenceEndDate} onChange={e=>update('recurrenceEndDate',e.target.value)}/></label>:<label className="field field-full">Datas adicionais<textarea rows={2} value={(draft.customDates||[]).join(', ')} onChange={e=>update('customDates',e.target.value.split(',').map((x:string)=>x.trim()).filter(Boolean))} placeholder="2026-10-15, 2026-11-20"/></label>}</>}
              <label className="field">
                Status
                <select
                  value={draft.status}
                  onChange={e => update('status', e.target.value)}
                >
                  {statuses.map(value => (
                    <option key={value}>{value}</option>
                  ))}
                </select>
              </label>
              <label className="field">
                Prioridade
                <select
                  value={draft.priority}
                  onChange={e => update('priority', e.target.value)}
                >
                  {priorities.map(value => (
                    <option key={value}>{value}</option>
                  ))}
                </select>
              </label>
              <label className="field field-full">
                Progresso <span>{draft.progress}%</span>
                <input
                  className="range"
                  type="range"
                  min="0"
                  max="100"
                  step="5"
                  value={draft.progress}
                  onChange={e => update('progress', Number(e.target.value))}
                />
              </label>
              {editingTask?.isRecurring&&<div className="field field-full occurrence-history"><span>Ocorrências</span>{taskOccurrences(editingTask).map((o:Occurrence)=><div className="occurrence-row" key={o.id}><b>{formatDate(o.dueDate)}</b><span className={statusClass[(!['Concluído','Cancelado'].includes(o.status)&&o.dueDate<serverToday)?'Atrasado':o.status]}>{(!['Concluído','Cancelado'].includes(o.status)&&o.dueDate<serverToday)?'Atrasado':o.status}</span><select value={o.status} onChange={e=>onUpdateOccurrence(editingTask.id,o.id,e.target.value as TaskStatus,o.progress)}>{statuses.map(v=><option key={v}>{v}</option>)}</select>{o.completedAt&&<small>Concluída em {new Date(o.completedAt).toLocaleDateString('pt-BR')}</small>}</div>)}</div>}
              {editingTask?.completedAt&&!editingTask.isRecurring&&<div className="field field-full completion-note">Conclusão real: <b>{new Date(editingTask.completedAt).toLocaleDateString('pt-BR')}</b></div>}
              <label className="field field-full">
                Observações
                <textarea
                  rows={3}
                  value={draft.notes}
                  onChange={e => update('notes', e.target.value)}
                  placeholder="Dependências, contexto ou próxima ação"
                />
              </label>
            </div>
            <div className="modal-actions">
              <div>
                {editingTask && (
                  <button
                    type="button"
                    className="text-danger"
                    onClick={() => setConfirmDelete(true)}
                  >
                    <Trash2 size={17} />
                    Excluir
                  </button>
                )}
              </div>
              <div>
                <button
                  type="button"
                  className="secondary-btn"
                  onClick={onClose}
                >
                  Cancelar
                </button>
                <button className="primary-btn" disabled={saving}>
                  {saving ? 'Salvando...' : 'Salvar atividade'}
                </button>
              </div>
            </div>
          </form>
        )}
      </div>
    </div>
  );
}

function DataManager({tasks,hierarchy,onRefresh}:{tasks:Task[];hierarchy:HierarchyMember[];onRefresh:()=>Promise<void>}) {
  const [message,setMessage]=useState('');
  const [busy,setBusy]=useState(false);
  const downloadModel=()=>{
    const wb=XLSX.utils.book_new();
    const h=[{Nome:'Ailton Paganucci',Equipe:'Inteligência',Gestor:'Lúcio Conge',Cargo:'Coordenador',Nivel:2},{Nome:'Lúcio Conge',Equipe:'Inteligência',Gestor:'',Cargo:'Dono da área',Nivel:1},{Nome:'Miguel',Equipe:'Execução',Gestor:'Lúcio Conge',Cargo:'Gestor',Nivel:2}];
    const t=[{Atividade:'Exemplo de atividade',Projeto:'Exemplo',Responsavel:'Ailton Paganucci',Apoio:'',Tipo:'Projeto',Inicio:'2026-10-01',Prazo:'2026-10-10',Status:'Não iniciado',Prioridade:'Média',Progresso:0,Observacoes:'',Recorrente:'Não',Recorrencia:'none',RepetirAte:'',DatasAdicionais:''}];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(h),'Hierarquia');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(t),'Tarefas');
    XLSX.writeFile(wb,'Planeja_Ae_Modelo_Importacao.xlsx');
  };
  const exportBackup=()=>{
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(hierarchy.map(({id,...x})=>x)),'Hierarquia');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(tasks),'Tarefas');
    XLSX.writeFile(wb,'Planeja_Ae_Backup.xlsx');
  };
  const importFile=async(event:React.ChangeEvent<HTMLInputElement>)=>{
    const file=event.target.files?.[0]; if(!file)return;
    setBusy(true);setMessage('');
    try{
      const data=await file.arrayBuffer();
      const wb=XLSX.read(data);
      const hs=wb.Sheets['Hierarquia'],ts=wb.Sheets['Tarefas'];
      if(!hs)throw new Error('A planilha precisa ter a aba Hierarquia.');
      const hierarchyRows=XLSX.utils.sheet_to_json(hs,{defval:''});
      await api.post('/api/hierarchy/import',{rows:hierarchyRows});
      let count=0;
      if(ts){const taskRows=XLSX.utils.sheet_to_json(ts,{defval:''}).filter((row:any)=>String(row.Atividade||'').trim());if(taskRows.length){const response=await api.post('/api/tasks/import',{rows:taskRows});count=response.data.count;}}
      await onRefresh();
      setMessage('Importação concluída: hierarquia atualizada e '+count+' tarefa(s) adicionada(s).');
    }catch(e:any){setMessage(e?.response?.data?.error||e?.message||'Não foi possível importar a planilha.');}
    finally{setBusy(false);event.target.value='';}
  };
  const grouped=hierarchy.reduce<Record<string,HierarchyMember[]>>((acc,item)=>{(acc[item.team]??=[]).push(item);return acc;},{});
  return <div className="stack">
    <section className="panel data-hero"><div><span className="eyebrow">BASE MESTRA</span><h2>Importação, hierarquia e backup</h2><p>A hierarquia é definida pela planilha. Depois de importada, fica travada no Planeja Aê. Para mudar, ajuste a planilha e importe novamente.</p></div><div className="data-actions"><button className="secondary-btn" onClick={downloadModel}><Download size={17}/>Baixar planilha padrão</button><label className="primary-btn file-btn"><Upload size={17}/>{busy?'Importando...':'Importar planilha'}<input type="file" accept=".xlsx,.xls" onChange={importFile} disabled={busy}/></label><button className="secondary-btn" onClick={exportBackup}><Database size={17}/>Exportar backup</button></div></section>
    {message&&<div className="import-message">{message}</div>}
    <section className="panel"><div className="panel-head"><div><span className="eyebrow">HIERARQUIA TRAVADA</span><h2>{hierarchy.length} pessoa(s) cadastrada(s)</h2></div><span className="lock-pill">🔒 Alteração somente via planilha</span></div><div className="hierarchy-groups">{Object.entries(grouped).map(([team,members])=><div className="hierarchy-team" key={team}><h3>{team}</h3>{members.sort((a,b)=>a.level-b.level).map(member=><div className="hierarchy-row" key={member.name}><div className="avatar">{member.name.slice(0,2).toUpperCase()}</div><div><strong>{member.name}</strong><span>{member.role}</span></div><div><small>Gestor</small><b>{member.manager||'Topo da estrutura'}</b></div></div>)}</div>)}</div></section>
  </div>;
}

export default App;
