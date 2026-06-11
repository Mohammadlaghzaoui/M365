import { load, save } from '../store/useLocalStorage';

/**
 * Role-Based Access Control.
 *
 * Four roles per the platform spec. The portal enforces these in the UI
 * (hiding/disabling actions); the AGENT enforces them again server-side via the
 * X-Role header so a tampered browser cannot execute privileged actions — the
 * agent is the real boundary, the UI is convenience.
 */
export type Role = 'super_admin' | 'architect' | 'engineer' | 'read_only';

export const ROLES: { id: Role; label: string; desc: string }[] = [
  { id: 'super_admin', label: 'Super Admin', desc: 'Full control — settings, integrations, users, execution, audit' },
  { id: 'architect', label: 'Migration Architect', desc: 'Plan + execute migrations and provisioning; manage knowledge base; no user/settings admin' },
  { id: 'engineer', label: 'Engineer', desc: 'Run validations and migrations; cannot change settings, integrations or knowledge base' },
  { id: 'read_only', label: 'Read Only', desc: 'View dashboards, reports and workflows; no execution, no changes' },
];

/** Capabilities each role grants. */
export type Capability =
  | 'manage_settings'      // AI/SSO/integrations/branding/modules
  | 'manage_users'         // local users + roles
  | 'manage_knowledge'     // GPO knowledge base, KB articles, group catalogue
  | 'execute_migration'    // run real migration/provisioning via agent
  | 'run_validation'       // test runs, live read-only validation
  | 'edit_data'            // tickets, projects, notes
  | 'view_audit';          // view audit log

const MATRIX: Record<Role, Capability[]> = {
  super_admin: ['manage_settings', 'manage_users', 'manage_knowledge', 'execute_migration', 'run_validation', 'edit_data', 'view_audit'],
  architect: ['manage_knowledge', 'execute_migration', 'run_validation', 'edit_data', 'view_audit'],
  engineer: ['execute_migration', 'run_validation', 'edit_data'],
  read_only: [],
};

const ROLE_KEY = 'user-roles'; // { [email]: Role }

function roleMap(): Record<string, Role> {
  return load<Record<string, Role>>(ROLE_KEY, {});
}

export function getRole(email: string): Role {
  const map = roleMap();
  if (map[email]) return map[email];
  // First/seed admin is Super Admin; everyone else defaults to Engineer.
  if (email === 'admin@sorrento.cloud' || Object.keys(map).length === 0) return 'super_admin';
  return 'engineer';
}

export function setRole(email: string, role: Role) {
  const map = roleMap();
  map[email] = role;
  save(ROLE_KEY, map);
}

export function allRoleAssignments(): Record<string, Role> {
  return roleMap();
}

export function can(email: string | undefined, cap: Capability): boolean {
  if (!email) return false;
  return MATRIX[getRole(email)].includes(cap);
}

export function roleLabel(role: Role): string {
  return ROLES.find((r) => r.id === role)?.label ?? role;
}
