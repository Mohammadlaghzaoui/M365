// ---------- Workflow schema (used by all assistants) ----------
export interface PSCommand {
  label: string;
  command: string;
}

export interface Workflow {
  id: string;
  title: string;
  category?: string;
  explanation: string;
  questions: string[];
  portalPaths: string[];
  steps: string[];
  powershell: PSCommand[];
  requiredRole: string;
  dangers: string[];
  escalation: string[];
  ticketTemplate: string;
}

// ---------- Tickets ----------
export type TicketStatus = 'open' | 'in-progress' | 'waiting' | 'escalated' | 'closed';

export interface Ticket {
  id: string;
  createdAt: string;
  customer: string;
  userName: string;
  userEmail: string;
  service: string;
  category: string;
  errorMessage: string;
  impact: string;
  urgency: 'low' | 'medium' | 'high' | 'critical';
  tried: string;
  screenshots: boolean;
  language: 'en' | 'nl';
  status: TicketStatus;
}

// ---------- Migration projects ----------
export type MigrationType = 'm365-to-m365' | 'google-to-m365' | 'exchange-onprem-to-m365';

export interface StageTask {
  id: string;
  label: string;
  done: boolean;
}

export interface MigrationStage {
  id: string;
  name: string;
  status: 'not-started' | 'in-progress' | 'blocked' | 'done';
  owner: string;
  risk: 'low' | 'medium' | 'high';
  notes: string;
  dueDate: string;
  tasks: StageTask[];
}

export interface MigrationProject {
  id: string;
  createdAt: string;
  name: string;
  customer: string;
  type: MigrationType;
  sourceEnv: string;
  destTenant: string;
  domains: string;
  users: number;
  sharedMailboxes: number;
  groups: number;
  teams: number;
  mailboxSizeNotes: string;
  tool: string;
  cutoverDate: string;
  status: 'planning' | 'active' | 'cutover' | 'completed' | 'on-hold';
  notes: string;
  stages: MigrationStage[];
}

// ---------- Knowledge base ----------
export interface KBArticle {
  id: string;
  title: string;
  service: string;
  tags: string[];
  body: string;
  favorite: boolean;
  custom?: boolean;
}

// ---------- Notes ----------
export interface Note {
  id: string;
  title: string;
  body: string;
  updatedAt: string;
  pinned: boolean;
}

// ---------- Security hardening ----------
export interface SecurityControl {
  id: string;
  name: string;
  priority: 'critical' | 'high' | 'medium' | 'low';
  risk: string;
  recommended: string;
  portalPath: string;
  requiredRole: string;
  quickWin?: boolean;
  plan?: '30' | '90';
}

export interface SecurityControlState {
  status: 'not-reviewed' | 'compliant' | 'partial' | 'non-compliant' | 'n/a';
  evidence: string;
  notes: string;
}

// ---------- Monitoring ----------
export interface MonitoringCheck {
  id: string;
  name: string;
  frequency: string;
  portalPath: string;
  whatToCheck: string;
  normal: string;
  risky: string;
  escalation: string;
  ticketNote: string;
}

// ---------- Syskit ----------
export interface SyskitTask {
  id: string;
  name: string;
  why: string;
  where: string;
  lookFor: string;
  action: string;
  risk: 'low' | 'medium' | 'high';
  ticketNote: string;
  escalationTrigger: string;
}

// ---------- Tenant admin ----------
export interface TenantTask {
  id: string;
  name: string;
  portalPath: string;
  requiredRole: string;
  steps: string[];
  powershell: PSCommand[];
  approvalNeeded: boolean;
  risk: 'low' | 'medium' | 'high';
  ticketNote: string;
}

// ---------- Escalation matrix ----------
export interface EscalationEntry {
  id: string;
  service: string;
  issueType: string;
  whenToEscalate: string;
  evidence: string;
  team: string;
  exampleNote: string;
}

// ---------- PowerShell generator ----------
export interface PSParam {
  key: string;
  label: string;
  placeholder: string;
}

export interface PSTask {
  id: string;
  service: string;
  name: string;
  template: string;
  explanation: string;
  module: string;
  requiredRole: string;
  warning: string;
  params: PSParam[];
}

// ---------- Mail generator ----------
export type MailType =
  | 'ack' | 'more-info' | 'investigating' | 'resolved' | 'escalation'
  | 'migration-announce' | 'migration-reminder' | 'cutover-done' | 'post-migration' | 'security-recommendation';

// ---------- Settings / AI ----------
export type AIProvider = 'openrouter' | 'openai' | 'claude' | 'disabled';

export interface AISettings {
  provider: AIProvider;
  apiKey: string;
  model: string;
  temperature: number;
  systemPrompt: string;
}

export interface ChatMessage {
  role: 'user' | 'assistant';
  content: string;
}

// ---------- Errors helpers ----------
export interface MigrationError {
  id: string;
  name: string;
  meaning: string;
  cause: string;
  fix: string[];
  powershell?: string;
  prevention: string;
  escalationNote: string;
}
