# Deploy WorkPilot entirely on one.com (no Render, no agent, nothing local)

This is the **"enter a code → approve Microsoft → read-only analysis"** flow
(like Codex), running **100% on your one.com web hosting**.

## Why this works without Render
A browser can't ask Microsoft for the device code (the `/devicecode` endpoint
sends no CORS header). A **server** can — and **one.com web hosting includes
PHP**, which is a server. So a tiny PHP file on your own domain does that one
step. After sign-in, the read-only Microsoft Graph collection runs in the
browser. No app registration, no tenant ID, no secret.

> Requirement: a one.com **Web Hosting** plan (the normal one — it has PHP with
> cURL/outbound HTTPS). A "domain only / parked" product without hosting won't
> run PHP.

## What you upload (the whole `dist/` folder)
After `npm run build`, the `dist/` folder contains the finished portal **plus**
`dist/api/auth.php` (the device-code helper). Upload the **contents of `dist/`**
to your one.com web root so you end up with:

```
<one.com web root>/
├── index.html
├── assets/...
├── manifest.webmanifest
├── sw.js
└── api/
    └── auth.php        ← the device-code helper (must keep the /api/ path)
```

## Steps
1. **Build** (or use the zip I provide):
   ```bash
   npm install
   npm run build
   ```
2. **Upload** everything inside `dist/` to your one.com site via:
   - one.com **File Manager** (Control panel → File Manager), or
   - **FTP/SFTP** (one.com → Control panel → FTP for host/username; drag `dist/*`
     into the web root, usually the folder your domain serves from).
   Keep the `api/auth.php` file at `/api/auth.php`.
3. **Open your site** (e.g. `https://sorrento.cloud`) → **Migration Discovery**
   → **"Connect a tenant — just enter a code"** → click the button.
4. A **code** appears. Open `microsoft.com/devicelogin`, enter the code, sign in
   as the **customer's admin**, approve the Microsoft Graph consent.
5. The **read-only** analysis runs automatically and fills the portal. Export to
   Excel or fill your own template as usual.

## Verify the helper is live
Open `https://YOURDOMAIN/api/auth.php` in a browser. You should see:
```json
{"error":"unknown_action","hint":"use ?action=start or ?action=poll"}
```
That means PHP works. If you instead see the PHP **source code**, PHP isn't
enabled on that folder — make sure it's a Web Hosting plan and the file ends in
`.php`. If you get **404**, the file isn't at `/api/auth.php`.

## Security (unchanged)
- **Read-only**: only `*.Read.All` Graph scopes, only GET calls. Nothing in the
  customer tenant is changed.
- The helper holds **no secret** — it uses Microsoft's public client. Tokens are
  never stored; they live only in the browser tab for the session.
- CORS on the helper is locked to your own domains (edit the `$allowed` list at
  the top of `api/auth.php` if your domain differs).
