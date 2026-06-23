<?php
/**
 * WorkPilot server-side store — persists assessments across browsers & devices.
 *
 * Without this, the portal keeps everything in the browser's localStorage, which
 * is per-browser: open another browser/device and it looks empty. This tiny PHP
 * store (on your own one.com hosting) saves the data server-side so any browser
 * you sign in from sees the same source tenants and assessments.
 *
 * Upload to:  https://YOURDOMAIN/api/store.php   (that's it — no key, no setup)
 *
 * Data is linked to the signed-in user (e-mail), kept in api/_store/ which is
 * blocked from direct web access, and only the portal's own domains may call it
 * (CORS allow-list above). Read-only assessment data only — never any token.
 */

header('Content-Type: application/json; charset=utf-8');

$origin  = $_SERVER['HTTP_ORIGIN'] ?? '';
$allowed = ['https://sorrento.cloud', 'https://www.sorrento.cloud', 'https://mohammadlaghzaoui.github.io'];
if (in_array($origin, $allowed, true)) { header("Access-Control-Allow-Origin: $origin"); header('Vary: Origin'); }
header('Access-Control-Allow-Headers: content-type, x-store-key, x-user');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
if (($_SERVER['REQUEST_METHOD'] ?? '') === 'OPTIONS') { http_response_code(204); exit; }

// No key needed — data is linked to the signed-in user (e-mail) and the browser
// origin is restricted above. (Optional: set WORKPILOT_STORE_KEY to also require
// a shared key, but that is NOT required and the portal does not send one.)
$SHARED_KEY = getenv('WORKPILOT_STORE_KEY') ?: '';
if ($SHARED_KEY !== '') {
  $key = $_SERVER['HTTP_X_STORE_KEY'] ?? '';
  if (!is_string($key) || !hash_equals($SHARED_KEY, $key)) {
    http_response_code(401);
    echo json_encode(['error' => 'unauthorized']);
    exit;
  }
}

$user = strtolower(preg_replace('/[^a-z0-9@._\-]/i', '', $_SERVER['HTTP_X_USER'] ?? 'shared'));
if ($user === '') $user = 'shared';

$dir = __DIR__ . '/_store';
if (!is_dir($dir)) {
  @mkdir($dir, 0700, true);
  // Block direct web access to the data folder (Apache).
  @file_put_contents($dir . '/.htaccess', "Require all denied\nDeny from all\n");
}
$file = $dir . '/' . hash('sha256', $user) . '.json';

if ($_SERVER['REQUEST_METHOD'] === 'GET') {
  echo is_file($file) ? file_get_contents($file) : '{}';
  exit;
}

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
  $body = file_get_contents('php://input');
  if ($body === false) { http_response_code(400); echo json_encode(['error' => 'no_body']); exit; }
  if (strlen($body) > 25 * 1024 * 1024) { http_response_code(413); echo json_encode(['error' => 'too_large']); exit; }
  json_decode($body);
  if (json_last_error() !== JSON_ERROR_NONE) { http_response_code(400); echo json_encode(['error' => 'invalid_json']); exit; }
  if (file_put_contents($file, $body, LOCK_EX) === false) { http_response_code(500); echo json_encode(['error' => 'write_failed']); exit; }
  echo json_encode(['ok' => true, 'bytes' => strlen($body)]);
  exit;
}

http_response_code(405);
echo json_encode(['error' => 'method_not_allowed']);
