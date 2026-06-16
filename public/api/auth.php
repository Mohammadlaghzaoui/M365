<?php
/**
 * WorkPilot device-code broker — runs on one.com PHP (server-side).
 *
 * Why this exists: a browser cannot call Microsoft's /devicecode endpoint
 * (it sends no CORS header). A server CAN — that's exactly what Codex does.
 * This tiny script is that server, hosted on your own one.com web hosting.
 * No Render, no agent, nothing local.
 *
 * It ONLY brokers the OAuth 2.0 device-code handshake (request a code, poll for
 * the token) using Microsoft's PUBLIC client "Microsoft Graph Command Line
 * Tools" — so there is no app registration and no secret. The actual read-only
 * Microsoft Graph collection runs in your browser with the returned token.
 *
 * Upload this file to one.com so it is reachable at:  https://YOURDOMAIN/api/auth.php
 *
 * Endpoints:
 *   POST api/auth.php?action=start              -> { device_code, user_code, verification_uri, ... }
 *   POST api/auth.php?action=poll (device_code) -> { access_token, ... } | { error: authorization_pending }
 */

header('Content-Type: application/json; charset=utf-8');

// Same-origin on one.com, but allow the GitHub Pages preview too.
$origin  = $_SERVER['HTTP_ORIGIN'] ?? '';
$allowed = [
  'https://sorrento.cloud',
  'https://www.sorrento.cloud',
  'https://mohammadlaghzaoui.github.io',
];
if (in_array($origin, $allowed, true)) {
  header("Access-Control-Allow-Origin: $origin");
  header('Vary: Origin');
}
header('Access-Control-Allow-Headers: content-type');
header('Access-Control-Allow-Methods: POST, OPTIONS');
if (($_SERVER['REQUEST_METHOD'] ?? '') === 'OPTIONS') { http_response_code(204); exit; }

// Microsoft's public client — no app registration, READ-ONLY scopes only.
$CLIENT = '14d82eec-204b-4c2f-b7e8-296a70dab67e';
$AUTH   = 'https://login.microsoftonline.com/organizations/oauth2/v2.0';
$SCOPE  = implode(' ', [
  'User.Read.All', 'Group.Read.All', 'Directory.Read.All', 'Organization.Read.All',
  'Domain.Read.All', 'Reports.Read.All', 'Sites.Read.All', 'Application.Read.All',
  'Policy.Read.All', 'AuditLog.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementManagedDevices.Read.All',
  'offline_access', 'openid', 'profile',
]);

/** POST application/x-www-form-urlencoded; works with cURL or the stream wrapper. */
function ms_post(string $url, array $fields): array {
  $body = http_build_query($fields);
  if (function_exists('curl_init')) {
    $ch = curl_init($url);
    curl_setopt_array($ch, [
      CURLOPT_POST           => true,
      CURLOPT_POSTFIELDS     => $body,
      CURLOPT_RETURNTRANSFER => true,
      CURLOPT_HTTPHEADER     => ['Content-Type: application/x-www-form-urlencoded'],
      CURLOPT_TIMEOUT        => 20,
    ]);
    $res  = curl_exec($ch);
    $code = (int) curl_getinfo($ch, CURLINFO_HTTP_CODE);
    if ($res === false) return [502, json_encode(['error' => 'agent_network', 'error_description' => curl_error($ch)])];
    curl_close($ch);
    return [$code ?: 502, $res];
  }
  $ctx = stream_context_create(['http' => [
    'method'        => 'POST',
    'header'        => "Content-Type: application/x-www-form-urlencoded\r\n",
    'content'       => $body,
    'ignore_errors' => true,
    'timeout'       => 20,
  ]]);
  $res  = @file_get_contents($url, false, $ctx);
  $code = 0;
  if (isset($http_response_header[0]) && preg_match('#\s(\d{3})\s#', $http_response_header[0], $m)) $code = (int) $m[1];
  return [$code ?: 502, $res === false ? json_encode(['error' => 'agent_network']) : $res];
}

$action = $_GET['action'] ?? '';

if ($action === 'start') {
  [$code, $res] = ms_post("$AUTH/devicecode", ['client_id' => $CLIENT, 'scope' => $SCOPE]);
  http_response_code($code);
  echo $res;
  exit;
}

if ($action === 'poll') {
  $dc = $_POST['device_code'] ?? '';
  if ($dc === '') { http_response_code(400); echo json_encode(['error' => 'missing_device_code']); exit; }
  [$code, $res] = ms_post("$AUTH/token", [
    'grant_type'  => 'urn:ietf:params:oauth:grant-type:device_code',
    'client_id'   => $CLIENT,
    'device_code' => $dc,
  ]);
  http_response_code($code);
  echo $res;
  exit;
}

http_response_code(400);
echo json_encode(['error' => 'unknown_action', 'hint' => 'use ?action=start or ?action=poll']);
