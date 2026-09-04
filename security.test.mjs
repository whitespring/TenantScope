import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { access, readFile } from 'node:fs/promises';

const [server, nginx, service, auth, app, indexHtml, authHtml, releaseFiles] = await Promise.all([
  readFile(new URL('./server.py', import.meta.url), 'utf8'),
  readFile(new URL('./deploy/nginx-site.conf', import.meta.url), 'utf8'),
  readFile(new URL('./deploy/tenant-report-proxy.service', import.meta.url), 'utf8'),
  readFile(new URL('./auth.js', import.meta.url), 'utf8'),
  readFile(new URL('./app.js', import.meta.url), 'utf8'),
  readFile(new URL('./index.html', import.meta.url), 'utf8'),
  readFile(new URL('./auth.html', import.meta.url), 'utf8'),
  readFile(new URL('./deploy/release-files.txt', import.meta.url), 'utf8'),
]);
const firewall = await readFile(new URL('./deploy/nftables.conf', import.meta.url), 'utf8');

assert.match(server, /BaseHTTPRequestHandler/, 'the token proxy must not inherit static file serving');
assert.match(server, /class Handler\(BaseHTTPRequestHandler\)/, 'the production token proxy must not inherit static file serving');
assert.match(server, /handler = DevHandler if "--serve-static"/, 'static files must be an explicit development-only mode');
assert.match(server, /MAX_REPORT_BYTES = 20 \* 1024 \* 1024/, 'report memory use must be bounded');
assert.match(server, /build_opener\(NoRedirect\)[\s\S]*location/, 'report downloads must reject follow-up redirects');
assert.match(nginx, /frame-ancestors 'none'/, 'CSP must block framing');
assert.match(nginx, /X-Content-Type-Options "nosniff"/, 'responses must disable MIME sniffing');
assert.match(nginx, /location ~ \^\/api\/reports\/\(m365-apps\|copilot\)\$/, 'only known report routes may reach the proxy');
assert.match(nginx, /limit_req zone=tenantscope_reports/, 'report collection must be rate limited');
assert.match(nginx, /limit_conn tenantscope_global 2/, 'report collection must have a global concurrency limit');
assert.match(nginx, /proxy_buffering off/, 'report payloads must not be buffered to disk');
assert.match(nginx, /proxy_max_temp_file_size 0/, 'report payloads must not create temporary files');
assert.match(nginx, /listen 443 ssl/, 'the origin must support encrypted proxy traffic');
assert.match(nginx, /location = \/ \{[\s\S]*try_files \/index\.html =404;/, 'the public webroot must use an allowlist');
assert.doesNotMatch(nginx, /Access-Control-Allow-Origin/, 'the private same-origin API must not enable CORS');
assert.match(auth, /dist\/redirect_bridge\/index\.mjs/, 'the auth bridge must use the installed MSAL entry point');
assert.match(app, /renderRunQueue\(\[scope\]\)[\s\S]*runScopeInventory\(scope, accessToken\)/, 'scope rechecks must reuse verbose queue progress');
assert.doesNotMatch(app, /inventoryRunners\[scope\.id\]\(accessToken, \(\) => \{\}\)/, 'scope rechecks must not discard progress updates');
assert.match(indexHtml, /id="home-link" href="\/"/, 'the wordmark must use a real root URL');
assert.match(app, /#home-link[\s\S]*preventDefault\(\)[\s\S]*showView\('setup'\)/, 'the wordmark must preserve the current SPA session on normal clicks');
for (const html of [indexHtml, authHtml]) {
  const importMap = html.match(/<script type="importmap">([\s\S]*?)<\/script>/)?.[1];
  assert.ok(importMap, 'each app page must have an import map');
  const hash = createHash('sha256').update(importMap).digest('base64');
  assert.ok(nginx.includes(`'sha256-${hash}'`), 'the CSP must contain the current import-map hash');
}
for (const path of releaseFiles.trim().split('\n')) await access(new URL(`./${path}`, import.meta.url));
assert.doesNotMatch(releaseFiles, /(?:server\.py|package|test|deploy\/)/, 'the public webroot allowlist must not contain source or deployment files');
for (const setting of ['NoNewPrivileges=true', 'RestrictNamespaces=true', 'CapabilityBoundingSet=', 'MemoryDenyWriteExecute=true', 'SystemCallFilter=@system-service']) {
  assert.ok(service.includes(setting), `systemd hardening must include ${setting}`);
}
assert.match(firewall, /chain input[\s\S]*policy drop/);
assert.match(nginx, /allow __REVERSE_PROXY_IP__;/);
assert.match(firewall, /ip saddr __REVERSE_PROXY_IP__ tcp dport \{ 80, 443 \}/);
assert.match(firewall, /chain output[\s\S]*policy accept/);

console.log('Security hardening self-check passed');
