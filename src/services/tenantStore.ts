import { DiscoveryResult } from './graphDiscovery';
import { load, save } from '../store/useLocalStorage';

/**
 * Per-tenant separated storage for assessment results.
 * Each tenant's data is stored under its own key (discovery:<tenantId>) so
 * multiple customer tenants never share state. An index tracks known tenants.
 */

export interface TenantIndexEntry {
  tenantId: string;
  displayName: string;
  fetchedAt: string;
  users: number;
  devices: number;
  dataGB: number;
  completeness: number;   // 0-100: exported workloads / total
  desktopUsers: number;
  mobileUsers: number;
  connectedAt: string;    // first time this tenant was connected (kept until removed)
}

export interface ExportAuditEntry {
  t: string;
  operator: string;
  tenantId: string;
  tenantName: string;
  scopes: number;
  objects: number;
  format: string;
}

const INDEX = 'tenant-index';
const AUDIT = 'export-audit';

export function tenantIndex(): TenantIndexEntry[] {
  return load<TenantIndexEntry[]>(INDEX, []);
}

export function saveTenantResult(r: DiscoveryResult): void {
  save(`discovery:${r.org.tenantId}`, r);
  const prev = tenantIndex().find((t) => t.tenantId === r.org.tenantId);
  const idx = tenantIndex().filter((t) => t.tenantId !== r.org.tenantId);
  const exported = r.workloads.filter((w) => w.status === 'Exported').length;
  const completeness = r.workloads.length ? Math.round((exported / r.workloads.length) * 100) : 0;
  const dataGB = Math.round((r.usage.mailboxTotalGB || 0) + (r.usage.oneDriveTotalGB || 0) + (r.usage.spoTotalGB || 0));
  idx.unshift({
    tenantId: r.org.tenantId,
    displayName: r.org.displayName,
    fetchedAt: r.fetchedAt,
    users: r.users.length,
    devices: r.devices?.total ?? 0,
    dataGB,
    completeness,
    desktopUsers: r.appUsage?.desktopApp ?? 0,
    mobileUsers: r.appUsage?.mobileOnly ?? 0,
    connectedAt: prev?.connectedAt ?? r.fetchedAt,
  });
  save(INDEX, idx);
}

export function loadTenantResult(tenantId: string): DiscoveryResult | null {
  return load<DiscoveryResult | null>(`discovery:${tenantId}`, null);
}

export function removeTenantResult(tenantId: string): void {
  localStorage.removeItem(`workpilot:discovery:${tenantId}`);
  save(INDEX, tenantIndex().filter((t) => t.tenantId !== tenantId));
}

export function recordExportAudit(e: Omit<ExportAuditEntry, 't'>): void {
  const log = load<ExportAuditEntry[]>(AUDIT, []);
  log.unshift({ t: new Date().toISOString(), ...e });
  save(AUDIT, log.slice(0, 200));
}

export function exportAudit(): ExportAuditEntry[] {
  return load<ExportAuditEntry[]>(AUDIT, []);
}
