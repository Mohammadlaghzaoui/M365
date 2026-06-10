# WorkPilot Migration Agent

The agent is the component that **actually executes** migrations, user provisioning
and endpoint validation. The WorkPilot portal (browser) never holds privileged
credentials or runs PowerShell — it sends signed requests to this agent, which
runs on a machine **with access to your environment** (domain-joined server,
jump host, or Azure VM) and does the real work via local PowerShell and Microsoft
Graph.

```
 Portal (browser, sorrento.cloud)
        │  HTTPS + X-API-Key
        ▼
 WorkPilot Migration Agent  (this service, on a trusted host)
        ├── PowerShell  → On-Prem AD (New-ADUser), Exchange (New-MigrationBatch, MoveRequests), Entra Connect sync
        └── Microsoft Graph (app-only) → cloud user creation, B2B invitations, group assignment
```

## Run it

Requirements: Node.js 18+ on the host. For on-prem work, install the relevant
PowerShell modules (`ActiveDirectory`, `ExchangeOnlineManagement`) and run the
agent under an account with the delegated rights (never Domain Admin).

```bash
cd agent
cp .env.example .env        # fill in AGENT_API_KEY + Graph app credentials
npm install
npm start
```

You should see: `WorkPilot Migration Agent listening on http://localhost:8787`.

Expose it to the portal over HTTPS (reverse proxy / tunnel) and paste the URL +
API key into the portal: **Settings → Integrations → Migration Agent → Test connection**.

## Endpoints

| Method | Path | Purpose |
|---|---|---|
| GET | `/health` | Capabilities (PowerShell available?, Graph configured?) — no auth |
| POST | `/endpoints/verify` | Verify `graph` / `onprem-ad` / `exchange-onprem` connectivity |
| POST | `/provision/user` | Create user (Internal→AD PowerShell, ExternalMember→Graph, Guest→invitation) |
| POST | `/migrate/start` | Submit a mailbox migration batch (runs the validated `New-MigrationBatch`) |
| POST | `/powershell/run` | Operator PowerShell (only if `ALLOW_RAW_POWERSHELL=true`) |
| GET | `/jobs/:id` | Job status + log |
| GET | `/jobs/:id/stream` | Live log via Server-Sent Events |

All routes except `/health` require the `X-API-Key` header.

## Security model

- The agent runs where the work happens; the portal only orchestrates.
- API-key auth + origin allow-list; put it behind HTTPS and, ideally, a VPN.
- Graph uses app-only client credentials scoped to the minimum permissions.
- Raw PowerShell is **off by default**; the provisioning/migration jobs use
  parameterized, purpose-built commands.
- Run as a least-privilege service account. Audit every job via `/jobs`.
