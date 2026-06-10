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

// ---------- Branding ----------
export interface BrandingSettings {
  companyName: string;
  portalName: string;
  supportEmail: string;
  defaultLanguage: 'en' | 'nl';
}

export const DEFAULT_BRANDING: BrandingSettings = {
  companyName: 'Sorrento Cloud',
  portalName: 'M365 WorkPilot',
  supportEmail: 'admin@sorrento.cloud',
  defaultLanguage: 'en',
};

export function getBranding(): BrandingSettings {
  return { ...DEFAULT_BRANDING, ...load<Partial<BrandingSettings>>('branding', {}) };
}

export function saveBranding(b: BrandingSettings) {
  save('branding', b);
}

// ---------- Extra integrations ----------
export interface IntegrationSettings {
  jira: { enabled: boolean; baseUrl: string; email: string; apiToken: string; projectKey: string };
  zendesk: { enabled: boolean; subdomain: string; email: string; apiToken: string };
  teamsWebhook: { enabled: boolean; url: string };
  slackWebhook: { enabled: boolean; url: string };
  bittitan: { enabled: boolean; apiKey: string };
  syskit: { enabled: boolean; baseUrl: string };
}

export const DEFAULT_INTEGRATIONS: IntegrationSettings = {
  jira: { enabled: false, baseUrl: '', email: '', apiToken: '', projectKey: '' },
  zendesk: { enabled: false, subdomain: '', email: '', apiToken: '' },
  teamsWebhook: { enabled: false, url: '' },
  slackWebhook: { enabled: false, url: '' },
  bittitan: { enabled: false, apiKey: '' },
  syskit: { enabled: false, baseUrl: '' },
};

export function getIntegrations(): IntegrationSettings {
  const stored = load<Partial<IntegrationSettings>>('integrations', {});
  return {
    jira: { ...DEFAULT_INTEGRATIONS.jira, ...stored.jira },
    zendesk: { ...DEFAULT_INTEGRATIONS.zendesk, ...stored.zendesk },
    teamsWebhook: { ...DEFAULT_INTEGRATIONS.teamsWebhook, ...stored.teamsWebhook },
    slackWebhook: { ...DEFAULT_INTEGRATIONS.slackWebhook, ...stored.slackWebhook },
    bittitan: { ...DEFAULT_INTEGRATIONS.bittitan, ...stored.bittitan },
    syskit: { ...DEFAULT_INTEGRATIONS.syskit, ...stored.syskit },
  };
}

export function saveIntegrations(s: IntegrationSettings) {
  save('integrations', s);
}
