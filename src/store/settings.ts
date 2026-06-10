import { AISettings, SSOSettings, ServiceNowSettings } from '../types';
import { load, save } from './useLocalStorage';

export const DEFAULT_SYSTEM_PROMPT =
  'You are an internal Microsoft 365 Service Provider Assistant. Help the engineer troubleshoot Entra ID, Exchange Online, SharePoint Online, Teams, migrations, BitTitan, Syskit, monitoring and security hardening. Always give safe, practical, step-by-step guidance. Always include portal paths, PowerShell commands when useful, required roles, risks, escalation criteria and ticket notes. Warn before destructive or risky actions.';

export const DEFAULT_AI_SETTINGS: AISettings = {
  provider: 'disabled',
  apiKey: '',
  model: 'anthropic/claude-sonnet-4.5',
  temperature: 0.3,
  systemPrompt: DEFAULT_SYSTEM_PROMPT,
};

export function getAISettings(): AISettings {
  return { ...DEFAULT_AI_SETTINGS, ...load<Partial<AISettings>>('ai-settings', {}) };
}

export function saveAISettings(s: AISettings) {
  save('ai-settings', s);
}

export function aiEnabled(): boolean {
  const s = getAISettings();
  return s.provider !== 'disabled' && !!s.apiKey;
}

// ---------- Microsoft 365 SSO ----------
export const DEFAULT_SSO: SSOSettings = { enabled: false, tenantId: '', clientId: '' };

export function getSSOSettings(): SSOSettings {
  return { ...DEFAULT_SSO, ...load<Partial<SSOSettings>>('sso-settings', {}) };
}

export function saveSSOSettings(s: SSOSettings) {
  save('sso-settings', s);
}

// ---------- ServiceNow ----------
export const DEFAULT_SERVICENOW: ServiceNowSettings = { enabled: false, instanceUrl: '', username: '', password: '' };

export function getServiceNowSettings(): ServiceNowSettings {
  return { ...DEFAULT_SERVICENOW, ...load<Partial<ServiceNowSettings>>('servicenow-settings', {}) };
}

export function saveServiceNowSettings(s: ServiceNowSettings) {
  save('servicenow-settings', s);
}

export function serviceNowEnabled(): boolean {
  const s = getServiceNowSettings();
  return s.enabled && !!s.instanceUrl && !!s.username;
}
