import { load, save } from '../store/useLocalStorage';
import { DeviceRecord } from './graphDiscovery';

/**
 * Workstation naming convention + license proposal logic for migration assessment.
 *
 * Convention (Retrofit example): RTF-01-[WorkerCode][DeviceCode][Last5]
 *   Site/company code: e.g. RTF-01  (Retrofit = RTF, Riverside site = 01)
 *   Worker codes:  O = Office, F = Field, T = Temp, K = Kiosk
 *   Device codes:  L = Laptop, D = Desktop, P = Phone, T = Tablet, M = Mac
 *   Last5: last 5 alphanumerics of the serial number (or device name if no serial).
 */

export const NAMING_FORMAT = '[SITE]-[Worker][Device][Last5]';

export function getSiteCode(tenantId: string): string {
  return load<string>(`naming-site:${tenantId}`, 'RTF-01');
}
export function setSiteCode(tenantId: string, code: string): void {
  save(`naming-site:${tenantId}`, code.trim().toUpperCase());
}

export function getMigrationTarget(tenantId: string): string {
  return load<string>(`migration-target:${tenantId}`, 'Kelso');
}
export function setMigrationTarget(tenantId: string, name: string): void {
  save(`migration-target:${tenantId}`, name.trim());
}

export function getTargetDomain(tenantId: string): string {
  return load<string>(`target-domain:${tenantId}`, 'kelso-industries.com');
}
export function setTargetDomain(tenantId: string, domain: string): void {
  save(`target-domain:${tenantId}`, domain.trim().toLowerCase());
}

/** Map a source UPN/email to the target tenant domain (keeps the local part). */
export function mapUpnToTarget(sourceUpn: string, targetDomain: string): string {
  const local = (sourceUpn || '').split('@')[0];
  return local && targetDomain ? `${local}@${targetDomain}` : '';
}

/** PC (laptop/desktop/mac) vs Mobile (phone/tablet). */
export function endpointClassOf(formFactor: string): 'PC' | 'Mobile' | '' {
  const f = (formFactor || '').toLowerCase();
  if (f === 'laptop' || f === 'desktop' || f === 'mac') return 'PC';
  if (f === 'phone' || f === 'tablet') return 'Mobile';
  return '';
}

/** Single-letter device code per the convention (L/D/P/T/M). */
export function deviceCodeOf(formFactor: string): string {
  const f = (formFactor || '').toLowerCase();
  if (f === 'mac') return 'M';
  if (f === 'laptop') return 'L';
  if (f === 'desktop') return 'D';
  if (f === 'phone') return 'P';
  if (f === 'tablet') return 'T';
  return '';
}

export function deviceTypeLabel(formFactor: string): string {
  const e = endpointClassOf(formFactor);
  return formFactor && e ? `${formFactor} (${e})` : (formFactor || '—');
}

/** Worker code: Office (O) / Field (F); if unknown, infer from the device class. */
export function workerCodeOf(workerType: string, formFactor: string): string {
  if (workerType === 'Office') return 'O';
  if (workerType === 'Field') return 'F';
  return endpointClassOf(formFactor) === 'Mobile' ? 'F' : 'O';
}

export function last5Of(d: Pick<DeviceRecord, 'serialNumber' | 'deviceName'>): string {
  const s = (d.serialNumber || d.deviceName || '').replace(/[^A-Za-z0-9]/g, '');
  return s.slice(-5).toUpperCase();
}

/** Build the suggested workstation name, or '' if the device can't be classified. */
export function suggestedName(site: string, d: DeviceRecord, workerType: string): string {
  const dc = deviceCodeOf(d.formFactor);
  if (!dc) return '';
  const l5 = last5Of(d);
  if (!l5) return '';
  return `${site}-${workerCodeOf(workerType, d.formFactor)}${dc}${l5}`;
}

/** License proposal from worker type. */
export function proposedLicenseOf(workerType: string): string {
  if (workerType === 'Office') return 'Microsoft 365 Business Premium';
  if (workerType === 'Field') return 'Exchange Online (Plan 1)';
  return 'Validate manually';
}

export const LICENSE_RULES = [
  { rule: 'Office / PC users', license: 'Microsoft 365 Business Premium', match: 'Office' },
  { rule: 'Mobile-only field users', license: 'Exchange Online (Plan 1)', match: 'Field' },
  { rule: 'No clear login evidence', license: 'Validate manually', match: '' },
];
