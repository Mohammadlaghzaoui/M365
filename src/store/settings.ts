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

// ---------- Google SSO ----------
export interface GoogleSSOSettings {
  enabled: boolean;
  clientId: string; // OAuth 2.0 Web client ID from Google Cloud Console
}

export const DEFAULT_GOOGLE_SSO: GoogleSSOSettings = { enabled: false, clientId: '' };

export function getGoogleSSO(): GoogleSSOSettings {
  return { ...DEFAULT_GOOGLE_SSO, ...load<Partial<GoogleSSOSettings>>('google-sso', {}) };
}

export function saveGoogleSSO(s: GoogleSSOSettings) {
  save('google-sso', s);
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
  topdesk: { enabled: boolean; baseUrl: string; username: string; appPassword: string };
  halopsa: { enabled: boolean; baseUrl: string; clientId: string; clientSecret: string };
  connectwise: { enabled: boolean; baseUrl: string; companyId: string; publicKey: string; privateKey: string };
  autotask: { enabled: boolean; apiUser: string; secret: string; integrationCode: string };
  freshservice: { enabled: boolean; domain: string; apiKey: string };
  intune: { enabled: boolean; note: string };
  agent: { enabled: boolean; url: string; apiKey: string };
  onpremAd: {
    enabled: boolean;
    domainFqdn: string;        // ad.customer.local
    netbios: string;           // CUSTOMER
    dcHostname: string;        // dc01.ad.customer.local
    entraConnectServer: string;
    serviceAccount: string;    // delegated provisioning account (no Domain Admin)
    defaultUserOu: string;     // OU=Internal Users,DC=...
    externalUserOu: string;    // OU=External,DC=...
    upnSuffix: string;         // customer.com
  };
}

export const DEFAULT_INTEGRATIONS: IntegrationSettings = {
  jira: { enabled: false, baseUrl: '', email: '', apiToken: '', projectKey: '' },
  zendesk: { enabled: false, subdomain: '', email: '', apiToken: '' },
  teamsWebhook: { enabled: false, url: '' },
  slackWebhook: { enabled: false, url: '' },
  bittitan: { enabled: false, apiKey: '' },
  syskit: { enabled: false, baseUrl: '' },
  topdesk: { enabled: false, baseUrl: '', username: '', appPassword: '' },
  halopsa: { enabled: false, baseUrl: '', clientId: '', clientSecret: '' },
  connectwise: { enabled: false, baseUrl: '', companyId: '', publicKey: '', privateKey: '' },
  autotask: { enabled: false, apiUser: '', secret: '', integrationCode: '' },
  freshservice: { enabled: false, domain: '', apiKey: '' },
  intune: { enabled: false, note: '' },
  agent: { enabled: false, url: 'http://localhost:8787', apiKey: '' },
  onpremAd: { enabled: false, domainFqdn: '', netbios: '', dcHostname: '', entraConnectServer: '', serviceAccount: '', defaultUserOu: '', externalUserOu: '', upnSuffix: '' },
};

export function getIntegrations(): IntegrationSettings {
  const stored = load<Partial<IntegrationSettings>>('integrations', {});
  const merged = {} as IntegrationSettings;
  (Object.keys(DEFAULT_INTEGRATIONS) as (keyof IntegrationSettings)[]).forEach((k) => {
    (merged as unknown as Record<string, unknown>)[k] = { ...DEFAULT_INTEGRATIONS[k], ...(stored[k] as object | undefined) };
  });
  return merged;
}

/** Status overview used by the dashboard integration strip. */
export function integrationStatus(): { name: string; enabled: boolean }[] {
  const i = getIntegrations();
  const sn = getServiceNowSettings();
  const ai = getAISettings();
  const sso = getSSOSettings();
  return [
    { name: 'AI Assistant', enabled: ai.provider !== 'disabled' && !!ai.apiKey },
    { name: 'Migration Agent', enabled: i.agent.enabled && !!i.agent.url && !!i.agent.apiKey },
    { name: 'Microsoft SSO', enabled: sso.enabled && !!sso.clientId },
    { name: 'Google SSO', enabled: getGoogleSSO().enabled && !!getGoogleSSO().clientId },
    { name: 'On-Prem AD', enabled: i.onpremAd.enabled && !!i.onpremAd.domainFqdn },
    { name: 'ServiceNow', enabled: sn.enabled && !!sn.instanceUrl },
    { name: 'TOPdesk', enabled: i.topdesk.enabled && !!i.topdesk.baseUrl },
    { name: 'Jira', enabled: i.jira.enabled && !!i.jira.baseUrl },
    { name: 'Zendesk', enabled: i.zendesk.enabled && !!i.zendesk.subdomain },
    { name: 'HaloPSA', enabled: i.halopsa.enabled && !!i.halopsa.baseUrl },
    { name: 'ConnectWise', enabled: i.connectwise.enabled && !!i.connectwise.baseUrl },
    { name: 'Autotask', enabled: i.autotask.enabled && !!i.autotask.apiUser },
    { name: 'Freshservice', enabled: i.freshservice.enabled && !!i.freshservice.domain },
    { name: 'Teams alerts', enabled: i.teamsWebhook.enabled && !!i.teamsWebhook.url },
    { name: 'Slack alerts', enabled: i.slackWebhook.enabled && !!i.slackWebhook.url },
    { name: 'BitTitan API', enabled: i.bittitan.enabled && !!i.bittitan.apiKey },
    { name: 'Syskit', enabled: i.syskit.enabled && !!i.syskit.baseUrl },
  ];
}

export function saveIntegrations(s: IntegrationSettings) {
  save('integrations', s);
}
