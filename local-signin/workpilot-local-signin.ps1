# WorkPilot — local sign-in helper
# Requests the Microsoft device-code from THIS PC, so the sign-in shows YOUR
# location (Belgium) instead of the web server. No install, no app registration,
# read-only. Just keep this window open while you connect a tenant.

$ErrorActionPreference = 'Stop'
$port   = 8799
$client = '14d82eec-204b-4c2f-b7e8-296a70dab67e'   # Microsoft's public "Graph CLI" client — no app registration
$scope  = 'User.Read.All Group.Read.All Directory.Read.All Organization.Read.All Domain.Read.All Reports.Read.All Sites.Read.All Application.Read.All Policy.Read.All AuditLog.Read.All DeviceManagementConfiguration.Read.All DeviceManagementManagedDevices.Read.All DeviceManagementApps.Read.All SecurityEvents.Read.All CloudPC.Read.All InformationProtectionPolicy.Read.All offline_access openid profile'

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Prefixes.Add("http://127.0.0.1:$port/")
try { $listener.Start() } catch {
  Write-Host ""
  Write-Host "  Could not start on port $port (is the helper already running?)." -ForegroundColor Yellow
  Write-Host "  Close other copies and try again."
  Read-Host "  Press Enter to exit"
  exit
}

Write-Host ""
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host "   WorkPilot local sign-in helper is RUNNING." -ForegroundColor Green
Write-Host "   Keep this window open, then go to the portal and click" -ForegroundColor Green
Write-Host "   'Or sign in with a code'. The sign-in now comes from THIS" -ForegroundColor Green
Write-Host "   PC (your location), not Denmark." -ForegroundColor Green
Write-Host "  ============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "   (Read-only. It only asks Microsoft for a sign-in code.)"
Write-Host "   Close this window when you're done."
Write-Host ""

while ($listener.IsListening) {
  $ctx = $listener.GetContext()
  $req = $ctx.Request
  $res = $ctx.Response
  $res.Headers['Access-Control-Allow-Origin']          = '*'
  $res.Headers['Access-Control-Allow-Headers']         = 'content-type'
  $res.Headers['Access-Control-Allow-Methods']         = 'GET, POST, OPTIONS'
  $res.Headers['Access-Control-Allow-Private-Network']  = 'true'
  if ($req.HttpMethod -eq 'OPTIONS') { $res.StatusCode = 204; $res.Close(); continue }

  $out = '{}'
  try {
    $path = $req.Url.AbsolutePath
    if ($path -like '*devicecode*') {
      $body = "client_id=$client&scope=" + [uri]::EscapeDataString($scope)
      $r = Invoke-WebRequest -Uri 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode' -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -UseBasicParsing
      $out = $r.Content
      Write-Host "   -> Issued a sign-in code from this PC." -ForegroundColor Cyan
    } else {
      $out = '{"status":"ok","helper":"workpilot"}'
    }
  } catch {
    $res.StatusCode = 502
    $msg = ($_.Exception.Message -replace '"', "'")
    $out = '{"error":"helper_failed","error_description":"' + $msg + '"}'
  }

  $bytes = [Text.Encoding]::UTF8.GetBytes($out)
  $res.ContentType     = 'application/json'
  $res.ContentLength64 = $bytes.Length
  $res.OutputStream.Write($bytes, 0, $bytes.Length)
  $res.Close()
}
