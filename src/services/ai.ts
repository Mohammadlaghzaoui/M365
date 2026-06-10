import { AISettings, ChatMessage } from '../types';
import { getAISettings } from '../store/settings';

/**
 * Provider-agnostic chat completion.
 * Supports OpenRouter (default), OpenAI and Claude (Anthropic).
 * Switching provider only changes endpoint + auth headers + payload mapping.
 */
export async function chat(messages: ChatMessage[], settingsOverride?: AISettings): Promise<string> {
  const s = settingsOverride ?? getAISettings();
  if (s.provider === 'disabled' || !s.apiKey) {
    throw new Error('AI is disabled. Configure a provider and API key in Settings.');
  }

  if (s.provider === 'claude') {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': s.apiKey,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({
        model: s.model || 'claude-sonnet-4-5',
        max_tokens: 2000,
        temperature: s.temperature,
        system: s.systemPrompt,
        messages: messages.map((m) => ({ role: m.role, content: m.content })),
      }),
    });
    if (!res.ok) throw new Error(`Claude API error ${res.status}: ${await res.text()}`);
    const data = await res.json();
    return data.content?.[0]?.text ?? '';
  }

  // OpenRouter and OpenAI share the OpenAI chat-completions schema.
  const endpoint =
    s.provider === 'openrouter'
      ? 'https://openrouter.ai/api/v1/chat/completions'
      : 'https://api.openai.com/v1/chat/completions';

  const headers: Record<string, string> = {
    'content-type': 'application/json',
    authorization: `Bearer ${s.apiKey}`,
  };
  if (s.provider === 'openrouter') {
    headers['HTTP-Referer'] = 'https://m365-workpilot.local';
    headers['X-Title'] = 'M365 WorkPilot';
  }

  const res = await fetch(endpoint, {
    method: 'POST',
    headers,
    body: JSON.stringify({
      model: s.model || (s.provider === 'openrouter' ? 'anthropic/claude-sonnet-4.5' : 'gpt-4o-mini'),
      temperature: s.temperature,
      // Cap the response size: without this, providers reserve the model's full
      // output window (e.g. 64k tokens) which fails on small prepaid balances (402).
      max_tokens: 2000,
      messages: [{ role: 'system', content: s.systemPrompt }, ...messages],
    }),
  });
  if (!res.ok) throw new Error(`${s.provider} API error ${res.status}: ${await res.text()}`);
  const data = await res.json();
  return data.choices?.[0]?.message?.content ?? '';
}

export async function testConnection(settings: AISettings): Promise<string> {
  const reply = await chat(
    [{ role: 'user', content: 'Reply with exactly: Connection OK' }],
    settings,
  );
  return reply.trim();
}
