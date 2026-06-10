# WorkPilot — server deployment (all-in-one)

One server runs **everything**: the portal UI, the agent API and the real
executor (PowerShell for on-prem AD/Exchange, Microsoft Graph for cloud).
Pick the option that matches your host.

## What you get

```
        Browser ──HTTPS──►  ONE SERVER  (http://server:8787)
                            ├── Portal UI (the React app)
                            ├── /health, /provision, /migrate, /jobs … (API, X-API-Key)
                            └── Executor: PowerShell (AD/Exchange) + Microsoft Graph
```

Because portal and API share one origin, there are no CORS headaches and the
agent URL in Settings is simply `http://localhost:8787` (or your HTTPS name).

---

## Option A — Windows server (recommended for hybrid: AD + Exchange)

Use this on a **domain-joined** Windows Server (or jump host) so the executor
can run `New-ADUser`, trigger Entra Connect sync and submit Exchange migration
batches.

1. Install **Node.js 18+** (https://nodejs.org).
2. Extract the repo, open **PowerShell as Administrator** in the repo root.
3. Run:

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\scripts\install-windows.ps1 -ApiKey "paste-a-long-random-key" -Port 8787
   ```

   Add Graph app-only for cloud accounts/invitations:

   ```powershell
   .\scripts\install-windows.ps1 -ApiKey "..." `
     -TenantId "<tenant-guid>" -ClientId "<app-id>" -ClientSecret "<secret>"
   ```

The script builds the portal, installs it to `C:\WorkPilot`, writes `.env`,
registers a **Windows service** (auto-start) via NSSM and opens the firewall
port. Browse to `http://<server>:8787`.

> Run the service under a **delegated** account (Recipient/Account Management,
> not Domain Admin). Put it behind HTTPS with IIS + ARR or any reverse proxy.

---

## Option B — Docker / Linux (cloud-only paths)

Good for a Linux VM doing Microsoft 365 (Graph) work. On-prem AD PowerShell
needs Windows, so use Option A for hybrid.

```bash
# edit docker-compose.yml: set AGENT_API_KEY (+ TENANT_ID/CLIENT_ID/CLIENT_SECRET)
docker compose up -d --build
# → http://server:8787
```

---

## Option C — Any Node host (manual)

```bash
npm install
npm run build:server      # builds the portal into agent/public
cd agent
npm install
# set env (or create agent/.env from .env.example): AGENT_API_KEY=...
npm start                 # → http://localhost:8787  (portal + API)
```

Keep it running with `pm2`, `systemd`, or a container.

---

## After install — connect the portal to its own agent

1. Open `http://<server>:8787`, sign in.
2. **Settings → Integrations → WorkPilot Migration Agent**:
   - URL: `http://localhost:8787` (same origin) or your HTTPS URL
   - API key: the value you passed as `AGENT_API_KEY`
   - **Test agent connection** → shows host, PowerShell and Graph status.
3. **Settings → Sign-in (SSO)**: add the Entra app (redirect URI = the server
   URL) and/or Google client ID (JS origin = the server URL).
4. Migration Console now shows **LIVE** and executes for real.

## Microsoft Graph app registration (for the executor)

Entra admin center → App registrations → New (single tenant). Add **application**
permissions and grant admin consent:
`User.ReadWrite.All`, `User.Invite.All`, `Group.ReadWrite.All`
(add `Mailbox.Migration` if you drive cross-tenant moves via Graph).
Create a client secret → put TenantId/ClientId/ClientSecret in the install.

## Security checklist

- [ ] Long random `AGENT_API_KEY`, stored only on the server.
- [ ] HTTPS in front (reverse proxy / IIS ARR), ideally inside a VPN.
- [ ] Service runs as least-privilege account (no Domain Admin).
- [ ] `ALLOW_RAW_POWERSHELL=false` unless you explicitly need it.
- [ ] Review jobs/audit via the portal and `/jobs`.
