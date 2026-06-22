import ExcelJS from 'exceljs';
import { DiscoveryResult } from './graphDiscovery';
import { Assessment } from './migrationAssessment';

/**
 * Template-aware Excel filler — STYLE PRESERVING.
 *
 * The engineer uploads THEIR OWN .xlsx template (locally, in the browser — it
 * never leaves the machine). We read its sheet names and header rows with
 * ExcelJS (which keeps every cell's colours, fonts, borders, column widths,
 * merged cells, tab colours and any sheet we don't touch), map each column
 * header to a discovery field via a synonym dictionary (English + Dutch), and
 * write the data rows under the existing headers using the styling of the
 * template's own data row — so the result looks exactly like the template.
 */

type FieldGetter = (row: Record<string, unknown>) => string | number | boolean;

interface FillContext {
  d: DiscoveryResult;
  a: Assessment | null;
}

interface Dataset {
  id: string;
  label: string;
  hints: string[];
  rows: (ctx: FillContext) => Record<string, unknown>[];
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
      { synonyms: ['usercategory', 'category', 'categorie', 'officeormobile', 'officemobile', 'type2', 'usertype2'], get: (r) => String(r.userCategory ?? '') },
      { synonyms: ['appplatforms', 'platforms', 'platform', 'platformen', 'apps', 'office'], get: (r) => String(r.appPlatforms ?? '') },
      { synonyms: ['lastofficeactivity', 'officeactivity', 'laatsteactiviteit'], get: (r) => String(r.lastOfficeActivity ?? '') },
      { synonyms: ['devices', 'devicecount', 'toestellen', 'aantaltoestellen', 'apparaten'], get: (r) => Number(r.deviceCount ?? 0) },
      { synonyms: ['devicetypes', 'toesteltypes', 'apparaattypes', 'ostypes'], get: (r) => String(r.deviceTypes ?? '') },
      { synonyms: ['mfa', 'mfastatus', 'multifactor'], get: (r) => String(r.mfa ?? '') },
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
  {
    id: 'devices', label: 'Devices', hints: ['device', 'toestel', 'toestellen', 'apparaat', 'computer', 'endpoint', 'intune'],
    rows: ({ d }) => (d.deviceInventory ?? []) as unknown as Record<string, unknown>[],
    fields: [
      { synonyms: ['devicename', 'name', 'naam', 'toestel', 'apparaat', 'hostname', 'computer'], get: (r) => String(r.deviceName ?? '') },
      { synonyms: ['user', 'gebruiker', 'owner', 'eigenaar', 'primaryuser', 'upn'], get: (r) => String(r.user ?? '') },
      { synonyms: ['os', 'operatingsystem', 'besturingssysteem', 'platform'], get: (r) => String(r.os ?? '') },
      { synonyms: ['osversion', 'version', 'versie'], get: (r) => String(r.osVersion ?? '') },
      { synonyms: ['compliance', 'compliancestate', 'compliant', 'conform', 'naleving'], get: (r) => String(r.compliance ?? '') },
      { synonyms: ['ownership', 'owner', 'eigendom', 'beheer', 'managed'], get: (r) => String(r.ownership ?? '') },
      { synonyms: ['manufacturer', 'fabrikant', 'merk', 'make'], get: (r) => String(r.manufacturer ?? '') },
      { synonyms: ['model', 'type'], get: (r) => String(r.model ?? '') },
      { synonyms: ['serial', 'serialnumber', 'serienummer', 'sn'], get: (r) => String(r.serialNumber ?? '') },
      { synonyms: ['formfactor', 'formaat', 'soort'], get: (r) => String(r.formFactor ?? '') },
      { synonyms: ['typecode', 'code', 'devicecode', 'namingcode'], get: (r) => String(r.typeCode ?? '') },
      { synonyms: ['workstationname', 'suggestedname', 'naam', 'name', 'naamgeving', 'computername', 'hostname'], get: (r) => String(r.suggestedName ?? '') },
      { synonyms: ['lastsync', 'laatstesync', 'lastcheckin', 'lastseen'], get: (r) => String(r.lastSync ?? '') },
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
  workbook: ExcelJS.Workbook;
  sheets: SheetPlan[];
}

/** ExcelJS cell values can be rich text / formulas / dates — flatten to text. */
function cellText(v: ExcelJS.CellValue): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  if (typeof v === 'object') {
    const o = v as unknown as Record<string, unknown>;
    if (Array.isArray(o.richText)) return (o.richText as { text: string }[]).map((t) => t.text).join('');
    if ('text' in o) return String(o.text ?? '');
    if ('result' in o) return String(o.result ?? '');
    if ('hyperlink' in o) return String(o.text ?? o.hyperlink ?? '');
  }
  return '';
}

/** Read an uploaded template file and detect each sheet's headers + best dataset. */
export async function analyzeTemplate(file: File): Promise<TemplateAnalysis> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());

  const sheets: SheetPlan[] = workbook.worksheets.map((ws) => {
    const maxCols = Math.max(ws.columnCount, 1);
    // Find the header row = the first row (within the first 10) with >=2 non-empty text cells.
    let headerRowIndex = 0;
    let headers: string[] = [];
    const scan = Math.min(ws.rowCount || 1, 10);
    for (let i = 1; i <= scan; i++) {
      const row = ws.getRow(i);
      const vals: string[] = [];
      for (let c = 1; c <= maxCols; c++) vals.push(cellText(row.getCell(c).value).trim());
      if (vals.filter((v) => v).length >= 2) { headerRowIndex = i - 1; headers = vals; break; }
    }
    // Trim trailing empty header cells.
    while (headers.length && !headers[headers.length - 1]) headers.pop();

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
      sheetName: ws.name, headers, headerRowIndex,
      datasetId: best && best.matched >= 1 ? best.id : null,
      matched: best ? Math.round(best.matched) : 0,
    };
  });
  return { workbook, sheets };
}

/** Fill the analyzed template with discovery data and trigger a style-preserving .xlsx download. */
export async function fillAndDownload(analysis: TemplateAnalysis, ctx: FillContext, filename: string): Promise<{ sheet: string; rows: number }[]> {
  const summary: { sheet: string; rows: number }[] = [];

  for (const plan of analysis.sheets) {
    if (!plan.datasetId) { summary.push({ sheet: plan.sheetName, rows: 0 }); continue; }
    const ds = DATASETS.find((x) => x.id === plan.datasetId);
    if (!ds) { summary.push({ sheet: plan.sheetName, rows: 0 }); continue; }
    const ws = analysis.workbook.getWorksheet(plan.sheetName);
    if (!ws) { summary.push({ sheet: plan.sheetName, rows: 0 }); continue; }
    const records = ds.rows(ctx);

    // Resolve a getter per template column header (unmatched columns stay blank).
    const getters: (FieldGetter | null)[] = plan.headers.map((h) => {
      const nh = norm(h);
      const f = ds.fields.find((fld) => fld.synonyms.includes(nh) || fld.synonyms.some((s) => nh.includes(s) || s.includes(nh)));
      return f ? f.get : null;
    });

    // Capture the styling of the template's first data row (the row under the
    // header) so EVERY written row looks like the template's table — colours,
    // borders, number formats, fonts. Falls back to header style if absent.
    const headerRowNum = plan.headerRowIndex + 1;     // 1-based
    const firstDataRowNum = headerRowNum + 1;
    const templateRow = ws.getRow(firstDataRowNum);
    const headerRow = ws.getRow(headerRowNum);
    const colStyles = plan.headers.map((_h, ci) => {
      const tc = templateRow.getCell(ci + 1);
      const src = tc && tc.style && Object.keys(tc.style).length ? tc.style : headerRow.getCell(ci + 1).style;
      return src ? (JSON.parse(JSON.stringify(src)) as ExcelJS.Style) : null;
    });

    records.forEach((rec, ri) => {
      const row = ws.getRow(firstDataRowNum + ri);
      plan.headers.forEach((_h, ci) => {
        const cell = row.getCell(ci + 1);
        const getter = getters[ci];
        if (getter) {
          const val = getter(rec);
          cell.value = val as ExcelJS.CellValue;
        }
        if (colStyles[ci]) cell.style = colStyles[ci] as ExcelJS.Style;
      });
      row.commit();
    });

    summary.push({ sheet: plan.sheetName, rows: records.length });
  }

  const buf = await analysis.workbook.xlsx.writeBuffer();
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename.endsWith('.xlsx') ? filename : `${filename}.xlsx`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);

  return summary;
}

export const TEMPLATE_DATASETS = DATASETS.map((d) => ({ id: d.id, label: d.label }));
