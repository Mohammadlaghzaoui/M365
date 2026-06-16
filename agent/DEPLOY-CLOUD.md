# Run the agent in the cloud — get the "enter a code" sign-in (like Codex)

You wanted: a popup that shows a **code to type in**, then a **Microsoft Graph
approval** — and **nothing installed on your PC**. This is exactly what the
device-code flow does. The catch is that a browser **cannot** request that code:
Microsoft's `/devicecode` endpoint returns no CORS header, so the request must
run on a server. This tiny agent does that for you, hosted in the cloud.

Once it's running, the portal's **Migration Discovery → "Easy connect — no
setup, just a code"** card gives you:

1. A code + the `microsoft.com/devicelogin` link.
2. You sign in once as the **customer's admin** and approve the Graph consent.
3. The **read-only** analysis runs automatically and fills the portal.

No app registration. No tenant ID typed. Read-only `*.Read.All` scopes only.

---

## Option 1 — Render (free, one-click blueprint)

1. Open <https://dashboard.render.com> → **New** → **Blueprint**.
2. Select this GitHub repo. Render reads `agent/render.yaml` and creates the service.
3. Wait for the deploy to go green. Note the URL, e.g.
   `https://workpilot-agent-xxxx.onrender.com`.
4. Open the service → **Environment** tab → copy the auto-generated
   `AGENT_API_KEY`.
5. In the portal: **Settings → Integrations → Migration Agent** →
   paste the **URL** and the **AGENT_API_KEY**, toggle **Enabled**, **Save**.
6. Add your portal domain to `ALLOWED_ORIGINS` in the Render service if it isn't
   already (e.g. your one.com domain).

> Free Render services sleep after ~15 min idle and take ~30s to wake on the
> first request. That's fine for occasional assessments.

## Option 2 — Railway / Fly.io / Koyeb / Cloud Run (Docker)

The repo includes `agent/Dockerfile`. Point any container host at it and set:

| Variable          | Value                                                            |
|-------------------|-----------------------------------------------------------------|
| `AGENT_API_KEY`   | a long random string (e.g. `openssl rand -hex 24`)              |
| `ALLOWED_ORIGINS` | your portal URL(s), comma-separated                             |
| `HOST`            | `0.0.0.0` (already set in the Dockerfile)                       |

Then paste the public HTTPS URL + the key into the portal as in step 5 above.

---

## Security notes (unchanged from the rest of the agent)

- **Read-only**: discovery requests only `*.Read.All` Graph scopes and issues
  only `GET` calls. Nothing in the customer tenant is modified.
- **Tokens are never persisted** — they live in memory only, and the session is
  dropped after 30 minutes.
- The agent **refuses to start** on the network with a weak/default key, and CORS
  is locked to `ALLOWED_ORIGINS`. Render/Railway terminate TLS for you (HTTPS).
- The shared `AGENT_API_KEY` is the identity; keep it secret. Rotate it by
  changing the env var and updating the portal.
