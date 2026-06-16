import * as XLSX from 'xlsx';
import { DiscoveryResult } from './graphDiscovery';
import { Assessment } from './migrationAssessment';

/**
 * Template-aware Excel filler.
 *
 * The engineer uploads THEIR OWN .xlsx template (locally, in the browser — it
 * never leaves the machine). We read its sheet names and header rows, map each
 * column header to a discovery field via a synonym dictionary (English + Dutch),
 * fill the data rows under the existing headers, and write the workbook back —
 * preserving the original tabs, headers and any sheets we don't touch.
 */

type Accessor = (ctx: FillContext) => (string | number | boolean)[];
type FieldGetter = (row: Record<string, unknown>) => string | number | boolean;

interface FillContext {
  d: DiscoveryResult;
  a: Assessment | null;
}

// A "dataset" = a list of records + a per-header value resolver.
interface Dataset {
  id: string;
  label: string;
  // synonyms that, when seen in a sheet's headers, hint this dataset
  hints: string[];
  rows: (ctx: FillContext) => Record<string, unknown>[];
  // map a normalized header -> getter
  fields: { synonyms: string[]; get: FieldGetter }[];
}

const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, '');

const DATASETS: Dataset[] = [
  {
    id: 'users', label: 'Users', hints: ['user', 'gebruiker', 'upn', 'mail', 'mailbox', 'medewerker'],
    rows: ({ d }) => d.users as unknown as Record<string, unknown>[],
    fields: [
      { synonyms: ['displayname', 'name', 'naam', 'volledigenaam', 'fullname', 'weergavenaam'], get: (r) => String(r.displayName ?? '') },
      { synonyms: ['userprincipalname', 'upn', 'loginname', 'aanmeldnaam'], get: (r) => String(r.userPrincipalName ?? '') },
      { synonyms: ['mail', 'email', 'emailaddress', 'emailadres', 'primarysmtp', 'smtp'], get: (r) => String(r.mail ?? r.userPrincipalName ?? '') },
      { synonyms: ['usertype', 'type', 'accounttype', 'soort'], get: (r) => String(r.userType ?? '') },
      { synonyms: ['enabled', 'accountenabled', 'actief', 'active', 'status'], get: (r) => (r.accountEnabled !== false ? 'Enabled' : 'Disabled') },
      { synonyms: ['department', 'afdeling', 'dept'], get: (r) => String(r.department ?? '') },
      { synonyms: ['jobtitle', 'title', 'functie', 'functietitel', 'role', 'rol'], get: (r) => String(r.jobTitle ?? '') },
      { synonyms: ['usagelocation', 'location', 'locatie', 'land', 'country'], get: (r) => String(r.usageLocation ?? '') },
      { synonyms: ['licenses', 'license', 'licentie', 'licenties', 'sku', 'licentietype'], get: (r) => String(r.licenses ?? '') },
      { synonyms: ['created', 'createddatetime', 'aangemaakt', 'creatiedatum'], get: (r) => String(r.createdDateTime ?? '') },
      { synonyms: ['lastsignin', 'lastlogon', 'laatsteaanmelding', 'laatstelogin', 'lastlogin'], get: (r) => String(r.lastSignIn ?? '') },
      { synonyms: ['mailboxsize', 'mailboxgrootte', 'grootte', 'sizegb', 'size'], get: () => '' },
    ],
  },
  {
    id: 'groups', label: 'Groups / Teams', hints: ['group', 'groep', 'team', 'distributielijst', 'distribution'],
    rows: ({ d }) => d.groups as unknown as Record<string, unknown>[],
    fields: [
      { synonyms: ['displayname', 'name', 'naam', 'groupname', 'groepsnaam'], get: (r) => String(r.displayName ?? '') },
      { synonyms: ['mail', 'email', 'emailadres', 'adres'], get: (r) => String(r.mail ?? '') },
      { synonyms: ['grouptype', 'type', 'soort', 'groepstype'], get: (r) => String(r.groupType ?? '') },
      { synonyms: ['membership', 'membershiptype', 'lidmaatschap'], get: (r) => String(r.membershipType ?? '') },
      { synonyms: ['visibility', 'zichtbaarheid'], get: (r) => String(r.visibility ?? '') },
      { synonyms: ['isteam', 'team', 'isteams'], get: (r) => (r.isTeam ? 'Yes' : 'No') },
      { synonyms: ['members', 'leden', 'aantalleden', 'membercount'], get: (r) => Number(r.members ?? 0) },
    ],
  },
  {
    id: 'licenses', label: 'Licenses', hints: ['license', 'licentie', 'sku', 'subscription', 'abonnement'],
    rows: ({ d }) => d.licenses as unknown as Record<string, unknown>[],
    fields: [
      { synonyms: ['sku', 'skupartnumber', 'license', 'licentie', 'naam', 'name', 'product'], get: (r) => String(r.skuPartNumber ?? '') },
      { synonyms: ['enabled', 'total', 'totaal', 'aantal', 'purchased', 'gekocht'], get: (r) => Number(r.enabled ?? 0) },
      { synonyms: ['consumed', 'used', 'gebruikt', 'inuse', 'assigned', 'toegewezen'], get: (r) => Number(r.consumed ?? 0) },
      { synonyms: ['available', 'beschikbaar', 'free', 'vrij', 'remaining'], get: (r) => Number(r.available ?? 0) },
    ],
  },
  {
    id: 'domains', label: 'Domains', hints: ['domain', 'domein'],
    rows: ({ d }) => d.domains as unknown as Record<string, unknown>[],
    fields: [
      { synonyms: ['domain', 'domein', 'naam', 'name', 'id'], get: (r) => String(r.id ?? '') },
      { synonyms: ['default', 'standaard'], get: (r) => (r.isDefault ? 'Yes' : 'No') },
      { synonyms: ['verified', 'geverifieerd', 'status'], get: (r) => (r.isVerified ? 'Verified' : 'Unverified') },
      { synonyms: ['services', 'diensten', 'supportedservices'], get: (r) => String(r.supportedServices ?? '') },
    ],
  },
];

export interface SheetPlan {
  sheetName: string;
  headers: string[];
  headerRowIndex: number;     // 0-based
  datasetId: string | null;   // auto-detected, user can override
  matched: number;            // how many headers map
}

export interface TemplateAnalysis {
  workbook: XLSX.WorkBook;
  sheets: SheetPlan[];
}

/** Read an uploaded template file and detect each sheet's headers + best dataset. */
export async function analyzeTemplate(file: File): Promise<TemplateAnalysis> {
  const buf = await file.arrayBuffer();
  const workbook = XLSX.read(buf, { cellStyles: true });
  const sheets: SheetPlan[] = workbook.SheetNames.map((name) => {
    const ws = workbook.Sheets[name];
    const grid: unknown[][] = XLSX.utils.sheet_to_json(ws, { header: 1, blankrows: false });
    // Find the header row = the first row with >=2 non-empty text cells.
    let headerRowIndex = 0;
    for (let i = 0; i < Math.min(grid.length, 10); i++) {
      const nonEmpty = (grid[i] ?? []).filter((c) => String(c ?? '').trim()).length;
      if (nonEmpty >= 2) { headerRowIndex = i; break; }
    }
    const headers = (grid[headerRowIndex] ?? []).map((c) => String(c ?? '').trim());

    // Auto-detect dataset: score by hint + field-synonym matches against headers.
    const normHeaders = headers.map(norm);
    let best: { id: string; matched: number } | null = null;
    for (const ds of DATASETS) {
      let matched = 0;
      for (const h of normHeaders) {
        if (!h) continue;
        if (ds.hints.some((hint) => h.includes(norm(hint)))) matched += 0.5;
        if (ds.fields.some((f) => f.synonyms.includes(h))) matched += 1;
      }
      if (!best || matched > best.matched) best = { id: ds.id, matched };
    }
    return {
      sheetName: name, headers, headerRowIndex,
      datasetId: best && best.matched >= 1 ? best.id : null,
      matched: best ? Math.round(best.matched) : 0,
    };
  });
  return { workbook, sheets };
}

/** Fill the analyzed template with discovery data and trigger a .xlsx download. */
export function fillAndDownload(analysis: TemplateAnalysis, ctx: FillContext, filename: string): { sheet: string; rows: number }[] {
  const summary: { sheet: string; rows: number }[] = [];
  for (const plan of analysis.sheets) {
    if (!plan.datasetId) { summary.push({ sheet: plan.sheetName, rows: 0 }); continue; }
    const ds = DATASETS.find((x) => x.id === plan.datasetId);
    if (!ds) { summary.push({ sheet: plan.sheetName, rows: 0 }); continue; }
    const records = ds.rows(ctx);

    // Resolve a getter per template column header (unmatched columns stay blank).
    const getters: (FieldGetter | null)[] = plan.headers.map((h) => {
      const nh = norm(h);
      const f = ds.fields.find((fld) => fld.synonyms.includes(nh) || fld.synonyms.some((s) => nh.includes(s) || s.includes(nh)));
      return f ? f.get : null;
    });

    const ws = analysis.workbook.Sheets[plan.sheetName];
    // Write data rows starting right after the header row, preserving the header.
    const startRow = plan.headerRowIndex + 1;
    records.forEach((rec, ri) => {
      plan.headers.forEach((_h, ci) => {
        const getter = getters[ci];
        if (!getter) return;
        const addr = XLSX.utils.encode_cell({ r: startRow + ri, c: ci });
        const val = getter(rec);
        ws[addr] = { t: typeof val === 'number' ? 'n' : 's', v: val };
      });
    });
    // Extend the sheet range so the new rows are included.
    const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1');
    range.e.r = Math.max(range.e.r, startRow + records.length - 1);
    range.e.c = Math.max(range.e.c, plan.headers.length - 1);
    ws['!ref'] = XLSX.utils.encode_range(range);
    summary.push({ sheet: plan.sheetName, rows: records.length });
  }
  XLSX.writeFile(analysis.workbook, filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`);
  return summary;
}

export const TEMPLATE_DATASETS = DATASETS.map((d) => ({ id: d.id, label: d.label }));
