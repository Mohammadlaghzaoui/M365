# M365 WorkPilot

AI-powered Microsoft 365 Service Provider portal — a working tool (not a handbook) that guides you step by step through helpdesk tickets, migrations, tenant administration, monitoring and security hardening for **Entra ID, Exchange Online, SharePoint Online, Teams, BitTitan MigrationWiz and Syskit**.

## Quick start

```bash
npm install
npm run dev
```

Open http://localhost:5173 — that's it. No backend, no database, no API key needed for the core portal.

### Run it instantly in the cloud (no install)

Open the repo branch in StackBlitz:

```
https://stackblitz.com/github/mohammadlaghzaoui/m365/tree/claude/gallant-shannon-yded1v
```

## What's inside (20 modules)

| Module | What it does |
|---|---|
| Dashboard | Quick actions, open/urgent tickets, migration projects, security progress, daily checklist, portal shortcuts, favorite PowerShell |
| Ticket Assistant | Fill a simple form → generates short/full description, troubleshooting steps, causes, portal paths, PowerShell, evidence list, customer update (EN/NL), escalation note, closure note — all with copy buttons |
| Entra ID Assistant | 15 guided workflows (login, password, MFA, CA, risky users, licenses, groups, guests, hybrid sync, roles, apps) |
| Exchange Online Assistant | 20 workflows incl. message trace, quarantine, permissions, room booking and the special "meeting room disappears after booking" deep-dive |
| SharePoint Online Assistant | 14 workflows (access, sharing, restores, storage, labels, private channel sites, OneDrive sync) |
| Teams Assistant | 15 workflows (access, client, calendar, meetings, AV, channels, guests, federation, recordings, rooms, policies, cache) |
| Migration Assistant | Projects with 15 stages, checklist tasks, owners, risk, due dates, completion %, generated plan + pre/cutover/post/rollback checklists + manual runbooks per migration type (M365→M365 incl. domain move, Google→M365, Exchange on-prem→M365) |
| Cross-Tenant Migration | Full guided setup (app, consent URL, endpoint, org relationships, MailUser prep), script generator from your form values, **Migration Runner with Test Mode → AI error analysis → gated Production runbook**, and a 10-error helper (license, scope, ExchangeGUID, aux archive, …) |
| BitTitan Assistant | 12-step MigrationWiz guide, permission/cutover/post checklists, 8-error helper |
| Syskit Assistant | 12 governance tasks: why, where in Syskit, what to look for, action, risk, ticket note, escalation trigger |
| Tenant Administration | 17 standard procedures with portal path, role, steps, PowerShell, approval-needed flag, risk |
| Monitoring Assistant | 14 checks with frequency, portal path, normal vs risky, escalation triggers |
| Security Hardening | 27 controls with status/evidence tracking, progress bar, quick wins, 30-day and 90-day plans |
| PowerShell Generator | 30+ parameterized commands across Graph/EXO/SPO/Teams/cross-tenant with module, role and warnings |
| Customer Mail Generator | 10 mail types × English/Dutch × 4 tones, optional AI enhancement |
| AI Chat Assistant | Provider-agnostic chat with quick prompts (paste error/ticket/BitTitan log…) |
| Escalation Matrix | 8 worked examples with evidence lists and copy-ready escalation notes |
| Knowledge Base | 16 seeded articles, search/filter/tags/favorites, add & edit your own |
| Personal Notes | Pinned local scratchpad |
| Settings | AI provider (OpenRouter / OpenAI / Claude / Disabled), API key, model, temperature, system prompt, **Test connection** button + local data reset |

## AI setup

The portal is fully usable **without AI** (templates, decision trees, generators). To enable AI:

1. **Settings → Provider → OpenRouter**, paste your `sk-or-v1-…` key, set a model (e.g. `anthropic/claude-sonnet-4.5`), click **Test AI connection**.
2. **Switch to OpenAI later:** select *OpenAI API*, paste `sk-…`, model `gpt-4o` — done.
3. **Switch to Claude API later:** select *Claude API*, paste `sk-ant-…`, model `claude-sonnet-4-5` — done.

All three share one abstraction (`src/services/ai.ts`): only endpoint + auth headers + payload mapping differ. Keys live in your browser localStorage only. For team/production use, put a tiny proxy backend in front instead of browser-side keys.

## Architecture

```
src/
├── main.tsx / App.tsx          # entry + HashRouter routes (20 pages)
├── types.ts                    # all shared TypeScript models
├── index.css                   # Tailwind layers
├── components/
│   ├── Layout.tsx              # sidebar, top search, dark/light toggle
│   ├── ui.tsx                  # Card, Badge, CopyButton, CodeBlock, Stepper bits, fields…
│   ├── WorkflowViewer.tsx      # renders any Workflow (questions→steps→PS→dangers→escalation→ticket)
│   ├── AssistantPage.tsx       # generic list+viewer layout used by all 4 service assistants
│   └── AIHelper.tsx            # reusable "Ask AI about this" panel
├── data/                       # ALL content is data-driven (add workflows without touching UI)
│   ├── entraWorkflows.ts … teamsWorkflows.ts
│   ├── crossTenant.ts          # steps, warnings, script builder, 10 error definitions
│   ├── bittitan.ts / syskit.ts / tenantAdmin.ts / monitoring.ts / security.ts
│   ├── migrationStages.ts      # 15-stage project template
│   ├── migrationGuides.ts      # manual runbooks (T2T domain move, Gmail→M365, on-prem→M365)
│   ├── psTasks.ts / mailTemplates.ts / kbArticles.ts / escalation.ts
│   ├── ticketGenerator.ts      # offline ticket package generation
│   └── dashboard.ts
├── services/
│   ├── ai.ts                   # OpenRouter / OpenAI / Claude abstraction + test
│   └── migrationRunner.ts      # Test-mode validation engine, production runbook builder,
│                               # ExecutionAdapter seam for a future execution backend
└── store/
    ├── useLocalStorage.ts      # typed localStorage hook (workpilot:* keys)
    └── settings.ts             # AI settings persistence
```

### LocalStorage structure (DB-ready)

Every store is one JSON document under a `workpilot:` key — a 1:1 mapping to future tables:

| Key | Content |
|---|---|
| `workpilot:tickets` | `Ticket[]` |
| `workpilot:migration-projects` | `MigrationProject[]` (incl. nested stages/tasks) |
| `workpilot:ct-form` | Cross-tenant configuration |
| `workpilot:security-state` | per-control `{status, evidence, notes}` |
| `workpilot:kb-custom` / `kb-overrides` | custom KB articles / edits+favorites of seeds |
| `workpilot:notes` | `Note[]` |
| `workpilot:ai-settings` | provider/key/model/temperature/system prompt |
| `workpilot:daily-checklist:<date>` | checked daily items per day |
| `workpilot:favorite-ps`, `recent-workflows`, `chat-history`, `dark-mode` | UX state |

To move to SQLite/PostgreSQL later: replace `load`/`save` in `src/store/useLocalStorage.ts` with API calls — the page components don't change.

## Migration Runner (test → production)

In **Cross-Tenant Migration → Migration Runner**:

1. Fill the configuration (tenant IDs, app ID, domains, scope group, user CSV).
2. **Run test migration** — validates everything locally against the rules that cause real failures and shows the exact exception each problem would raise (e.g. `MissingExchangeGuidException`, `UserDuplicateInOtherBatchException`).
3. If it fails, click **AI: analyze & fix** — the AI explains root causes and gives corrected values; adjust and re-run until green.
4. Production unlocks only after a passing test + explicit approval checkbox, then emits the exact validated PowerShell runbook in execution order.

> By design the browser never executes tenant-admin commands directly (credentials in a browser would be a security hole). The `ExecutionAdapter` interface in `src/services/migrationRunner.ts` is the prepared seam for a server-side executor (Azure Function/Automation with certificate auth) that runs the same validated plan.

## Example output

**Generated PowerShell (Room mailbox check):**
```powershell
Get-CalendarProcessing room.amsterdam@contoso.com | Format-List AutomateProcessing,BookingWindowInDays,MaximumDurationInMinutes,AllowConflicts,AllowRecurringMeetings,ConflictPercentageAllowed,MaximumConflictInstances,AllBookInPolicy,BookInPolicy,ResourceDelegates
```

**Generated ticket (short description):**
```
[Exchange Online] Mail not received — Jane Doe (Contoso BV)
```
…plus full description, troubleshooting plan, causes, portal paths, PowerShell, evidence list, customer update (EN/NL), escalation note and closure note — each with a copy button.

## Tech

React 18 · TypeScript (strict) · Vite 5 · Tailwind CSS 3 · react-router (hash routing, works on any static host) · lucide-react icons · localStorage persistence · dark/light mode · fully responsive.

## Build for production

```bash
npm run build    # outputs static files to dist/ — host anywhere (relative base path configured)
npm run preview
```
