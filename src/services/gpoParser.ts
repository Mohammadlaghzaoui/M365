import { GpoMapping, GpoSupport, seedGpoMappings } from '../data/gpoKnowledge';
import { load } from '../store/useLocalStorage';

/**
 * GPO XML parser + knowledge-base matcher.
 * Accepts the output of `Get-GPOReport -ReportType XML` (single GPO or -All),
 * extracts individual policy settings from Computer/User configuration, and
 * matches each one against the (seed + admin-edited) knowledge base.
 */

export interface ParsedSetting {
  id: string;
  name: string;
  scope: 'Computer' | 'User' | 'Unknown';
  category: string;
  state?: string;
  rawPath?: string;
  gpoName?: string;
}

export interface MatchedSetting extends ParsedSetting {
  matched: boolean;
  target: string;
  intunePath: string;
  setting: string;
  support: GpoSupport;
  confidence: number;
  notes?: string;
}

let counter = 0;
const nid = () => `s${Date.now().toString(36)}${counter++}`;

function text(el: Element | null | undefined): string {
  return el?.textContent?.trim() ?? '';
}

/** Get the effective knowledge base (seed overlaid with admin edits/additions). */
export function getKnowledgeBase(): GpoMapping[] {
  const custom = load<GpoMapping[]>('gpo-custom-mappings', []);
  const removed = load<string[]>('gpo-removed-mappings', []);
  const overrides = load<Record<string, Partial<GpoMapping>>>('gpo-mapping-overrides', {});
  const seeds = seedGpoMappings.filter((m) => !removed.includes(m.id)).map((m) => ({ ...m, ...overrides[m.id] }));
  return [...custom, ...seeds];
}

export function parseGpoXml(xml: string, fileName: string): ParsedSetting[] {
  const doc = new DOMParser().parseFromString(xml, 'application/xml');
  if (doc.querySelector('parsererror')) {
    throw new Error(`${fileName}: not valid XML.`);
  }
  const settings: ParsedSetting[] = [];

  // Each <GPO> may contain <Computer><ExtensionData>… and <User><ExtensionData>…
  const gpos = Array.from(doc.getElementsByTagName('GPO'));
  const roots = gpos.length ? gpos : [doc.documentElement];

  for (const gpo of roots) {
    const gpoName = text(gpo.getElementsByTagName('Name')[0]) || fileName.replace(/\.xml$/i, '');
    for (const scopeTag of ['Computer', 'User'] as const) {
      const scopeEl = Array.from(gpo.children).find((c) => c.tagName === scopeTag)
        ?? gpo.getElementsByTagName(scopeTag)[0];
      if (!scopeEl) continue;

      // Administrative templates: <Policy><Name>..</Name><State>..</State><Category>..</Category>
      for (const pol of Array.from(scopeEl.getElementsByTagName('Policy'))) {
        const name = text(pol.getElementsByTagName('Name')[0]);
        if (!name) continue;
        settings.push({
          id: nid(), name, scope: scopeTag,
          category: text(pol.getElementsByTagName('Category')[0]) || 'Administrative Templates',
          state: text(pol.getElementsByTagName('State')[0]),
          rawPath: text(pol.getElementsByTagName('Category')[0]),
          gpoName,
        });
      }

      // Security settings / SecurityOptions / Account policies
      for (const tag of ['SecurityOptions', 'Account', 'SettingNumber', 'SettingBoolean']) {
        for (const s of Array.from(scopeEl.getElementsByTagName(tag))) {
          const name = text(s.getElementsByTagName('Name')[0]) || text(s.getElementsByTagName('KeyName')[0]);
          if (!name) continue;
          settings.push({ id: nid(), name, scope: scopeTag, category: 'Security Settings', state: text(s.getElementsByTagName('SettingNumber')[0] ?? s.getElementsByTagName('Display')[0]), gpoName });
        }
      }

      // Group Policy Preferences: Drive maps, Printers, etc. (DriveMapSettings, Printers…)
      for (const pref of ['DriveMapSettings', 'Printers', 'Drive', 'SharedPrinter', 'PortPrinter']) {
        for (const p of Array.from(scopeEl.getElementsByTagName(pref))) {
          const name = pref.replace(/Settings$/, '') + (text(p.getAttribute && p.getElementsByTagName('Properties')[0]?.getAttribute?.('path') ? null : null) || '');
          settings.push({ id: nid(), name: `${pref} (Group Policy Preferences)`, scope: scopeTag, category: `Preferences - ${pref}`, gpoName });
        }
      }

      // Folder Redirection
      for (const fr of Array.from(scopeEl.getElementsByTagName('Folder'))) {
        const id = fr.getAttribute('Id') || 'Folder';
        settings.push({ id: nid(), name: `Folder Redirection (${id})`, scope: scopeTag, category: 'Folder Redirection', gpoName });
      }

      // Scripts (logon/logoff/startup)
      for (const sc of Array.from(scopeEl.getElementsByTagName('Script'))) {
        const cmd = text(sc.getElementsByTagName('Command')[0]);
        settings.push({ id: nid(), name: `Script: ${cmd || 'logon/logoff'}`, scope: scopeTag, category: 'Scripts', gpoName });
      }
    }
  }

  // De-duplicate identical name+scope
  const seen = new Set<string>();
  return settings.filter((s) => {
    const key = `${s.scope}|${s.name}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

export function matchSetting(s: ParsedSetting, kb: GpoMapping[]): MatchedSetting {
  const hay = `${s.name} ${s.category} ${s.rawPath ?? ''}`.toLowerCase();
  let best: GpoMapping | null = null;
  let bestScore = 0;
  for (const m of kb) {
    let score = 0;
    for (const kw of m.match) {
      if (hay.includes(kw.toLowerCase())) score += kw.length; // longer keyword = stronger signal
    }
    if (m.scope !== 'Both' && m.scope === s.scope) score += 2;
    if (score > bestScore) { bestScore = score; best = m; }
  }
  if (best && bestScore > 0) {
    return {
      ...s, matched: true, target: best.target, intunePath: best.intunePath,
      setting: best.setting, support: best.support, confidence: best.confidence, notes: best.notes,
    };
  }
  return {
    ...s, matched: false, target: 'Unsupported', intunePath: '—',
    setting: 'No knowledge-base match (review manually / add mapping)', support: 'Unsupported', confidence: 0,
    notes: 'Not yet in the knowledge base. Add a mapping in the Knowledge Engine to teach WorkPilot this setting.',
  };
}

export interface GpoAnalysis {
  settings: MatchedSetting[];
  total: number;
  supported: number;
  scriptRequired: number;
  manual: number;
  unsupported: number;
  readiness: number; // percentage
  byTarget: { label: string; value: number }[];
}

export function analyzeGpo(parsed: ParsedSetting[]): GpoAnalysis {
  const kb = getKnowledgeBase();
  const settings = parsed.map((p) => matchSetting(p, kb));
  const supported = settings.filter((s) => s.support === 'Supported').length;
  const scriptRequired = settings.filter((s) => s.support === 'ScriptRequired').length;
  const manual = settings.filter((s) => s.support === 'Manual').length;
  const unsupported = settings.filter((s) => s.support === 'Unsupported').length;
  const total = settings.length || 1;
  // Readiness: supported = full credit, script = 0.7, manual = 0.4, unsupported = 0.
  const readiness = Math.round(((supported + scriptRequired * 0.7 + manual * 0.4) / total) * 100);

  const targetCounts: Record<string, number> = {};
  for (const s of settings) targetCounts[s.target] = (targetCounts[s.target] ?? 0) + 1;
  const byTarget = Object.entries(targetCounts).map(([label, value]) => ({ label, value })).sort((a, b) => b.value - a.value);

  return { settings, total: settings.length, supported, scriptRequired, manual, unsupported, readiness, byTarget };
}

export function exportCsv(settings: MatchedSetting[]): string {
  const rows = [
    'Policy Name,Scope,Category,Recommended Target,Intune Path,Setting,Support Status,Confidence,Notes',
    ...settings.map((s) => [s.name, s.scope, s.category, s.target, s.intunePath, s.setting, s.support, s.confidence + '%', s.notes ?? '']
      .map((v) => `"${String(v).replace(/"/g, '""')}"`).join(',')),
  ];
  return rows.join('\n');
}
