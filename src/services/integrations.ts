import { getIntegrations } from '../store/settings';

/**
 * Extra integrations (Jira, Zendesk, Teams/Slack webhooks).
 * Browser CORS notes:
 *  - Jira Cloud & Zendesk REST APIs do not send CORS headers for browser apps;
 *    tests work via a proxy or when the instance allows the origin. The error
 *    message will say so explicitly.
 *  - Incoming webhooks (Teams/Slack) accept POSTs; responses may be opaque.
 */

export async function testJira(): Promise<string> {
  const { jira } = getIntegrations();
  if (!jira.baseUrl || !jira.email || !jira.apiToken) throw new Error('Fill in base URL, email and API token first.');
  const res = await fetch(`${jira.baseUrl.replace(/\/+$/, '')}/rest/api/3/myself`, {
    headers: { authorization: 'Basic ' + btoa(`${jira.email}:${jira.apiToken}`), accept: 'application/json' },
  }).catch(() => { throw new Error('Network/CORS blocked. Jira Cloud requires a proxy for browser calls — config saved, use it via the backend later.'); });
  if (!res.ok) throw new Error(`Jira ${res.status}: ${await res.text()}`);
  const me = await res.json();
  return `Connected as ${me.displayName ?? me.emailAddress}`;
}

export async function testZendesk(): Promise<string> {
  const { zendesk } = getIntegrations();
  if (!zendesk.subdomain || !zendesk.email || !zendesk.apiToken) throw new Error('Fill in subdomain, email and API token first.');
  const res = await fetch(`https://${zendesk.subdomain}.zendesk.com/api/v2/users/me.json`, {
    headers: { authorization: 'Basic ' + btoa(`${zendesk.email}/token:${zendesk.apiToken}`), accept: 'application/json' },
  }).catch(() => { throw new Error('Network/CORS blocked. Zendesk requires a proxy for browser calls — config saved for backend use.'); });
  if (!res.ok) throw new Error(`Zendesk ${res.status}: ${await res.text()}`);
  const me = await res.json();
  return `Connected as ${me.user?.name ?? zendesk.email}`;
}

async function postWebhook(url: string, payload: unknown): Promise<string> {
  try {
    const res = await fetch(url, {
      method: 'POST',
      headers: { 'content-type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (res.ok || res.type === 'opaque') return 'Test message sent — check the channel.';
    throw new Error(`Webhook ${res.status}: ${await res.text()}`);
  } catch {
    // Retry opaque (no-cors) so the message still goes out when CORS blocks reading the response.
    await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'content-type': 'application/json' }, body: JSON.stringify(payload) });
    return 'Test message sent (response not readable in browser) — check the channel.';
  }
}

export async function testTeamsWebhook(): Promise<string> {
  const { teamsWebhook } = getIntegrations();
  if (!teamsWebhook.url) throw new Error('Fill in the webhook URL first.');
  return postWebhook(teamsWebhook.url, { text: '✅ M365 WorkPilot — Teams webhook connection test succeeded.' });
}

export async function testSlackWebhook(): Promise<string> {
  const { slackWebhook } = getIntegrations();
  if (!slackWebhook.url) throw new Error('Fill in the webhook URL first.');
  return postWebhook(slackWebhook.url, { text: ':white_check_mark: M365 WorkPilot — Slack webhook connection test succeeded.' });
}

/** Send a notification to every enabled chat webhook (used by future alerting). */
export async function notifyChannels(message: string): Promise<void> {
  const s = getIntegrations();
  const jobs: Promise<unknown>[] = [];
  if (s.teamsWebhook.enabled && s.teamsWebhook.url) jobs.push(postWebhook(s.teamsWebhook.url, { text: message }));
  if (s.slackWebhook.enabled && s.slackWebhook.url) jobs.push(postWebhook(s.slackWebhook.url, { text: message }));
  await Promise.allSettled(jobs);
}
