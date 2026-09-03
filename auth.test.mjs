import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';

const [app, callback, bridge, index] = await Promise.all([
  readFile(new URL('./app.js', import.meta.url), 'utf8'),
  readFile(new URL('./auth.html', import.meta.url), 'utf8'),
  readFile(new URL('./auth.js', import.meta.url), 'utf8'),
  readFile(new URL('./index.html', import.meta.url), 'utf8'),
]);

assert.match(app, /new URL\('auth\.html'/, 'MSAL must redirect to the dedicated callback');
assert.match(app, /9002326[\s\S]+Single-page application/, 'SPA platform errors must explain the Entra fix');
assert.match(callback, /src="auth\.js"/, 'callback must load its bridge script');
assert.match(bridge, /broadcastResponseToMainFrame/, 'callback must relay the response to MSAL');
assert.match(index, /id="side-nav"/, 'the app must expose its section navigation');
assert.match(index, /class="report-sidebar"[\s\S]+id="report-nav"/, 'report areas must use a left sidebar on desktop');
assert.match(index, /id="report-nav-select"/, 'report areas must use a compact selector on narrow screens');
assert.match(index, /id="continue-to-scope"/, 'tenant setup must expose a visible primary continuation button');
assert.match(index, /id="export-dialog"/, 'all export formats must use the configurable export dialog');
assert.match(index, /id="usage-notice-dialog"/, 'the app must show a usage and liability notice before use');
assert.match(index, /Ich bin zur Prüfung des Tenants berechtigt/, 'the legal notice must require explicit authorization confirmation');
assert.match(index, /https:\/\/www\.whitespring\.de\/datenschutz\//, 'the app must link to whitespring privacy information');
assert.match(index, /https:\/\/www\.whitespring\.de\/impressum\//, 'the app must link to the whitespring imprint');
assert.match(app, /sessionStorage\.setItem\(USAGE_NOTICE_KEY, 'accepted'\)/, 'usage notice acceptance must be scoped to the browser session');
for (const option of ['summary', 'recommendations', 'findings', 'inventory', 'goodPractice', 'details', 'serviceHealth']) {
  assert.match(index, new RegExp(`name="${option}"`), `export option ${option} must be available`);
}
assert.match(app, /filter\(\(\{ id \}\) => id !== 'service'\)/, 'service health must be removable from the exported report');
assert.match(app, /if \(options\.recommendations\)/, 'recommendation summary must be optional');
assert.match(app, /report: 'm365-apps'/, 'M365 app usage must use the local report collector');
assert.match(app, /report: 'copilot'/, 'Copilot usage must use the local report collector');
assert.match(app, /fetch\(`\/api\/reports\/\$\{report\}`/, 'report downloads must use the same-origin server endpoint');
assert.match(app, /setupPagination\(reportContent\)/, 'rendered lists must initialize pagination');
assert.match(app, /const metricFilterRules/, 'area metrics must define their underlying detail filters');
assert.match(app, /function applyMetricFilter/, 'area metrics must filter their underlying detail tables');
assert.match(app, /data-metric-index/, 'area metrics must render as interactive controls');
assert.match(app, /data-table-search/, 'every detail table must expose a text filter');
assert.match(app, /createScopeProgress[\s\S]+queue-elapsed/, 'long inventory steps must show a live elapsed clock');
assert.match(app, /new Set\(\['sharing', 'apps'\]\)/, 'slow sharing and app scans must start early in parallel');
assert.match(app, /Batch \$\{groupIndex \+ 1\}\/\$\{groups\.length\}/, 'batched detail scans must expose live batch progress');
assert.match(app, /\[429, 503, 504\][\s\S]+Retry-After/, 'throttled requests inside Graph batches must be retried');
assert.match(app, /BATCH_INTERVAL_MS = 1000[\s\S]+paceGraphBatch/, 'Graph batches must be paced centrally');
assert.match(app, /waitWithCountdown\(retryAfter[\s\S]+erneuter Versuch in \$\{remaining\} s/, 'throttled batch retries must show a live countdown');
assert.match(app, /runEtaDeadlines[\s\S]+noch etwa \$\{formatRemaining/, 'long-running scans must expose an overall ETA');
assert.match(app, /data-report-panel/, 'report areas must render as separate views');
assert.match(app, /reportView \? 'report'/, 'detail views must keep the report workflow step active');
assert.match(app, /setAttribute\('aria-current', 'page'\)/, 'the active report area must be exposed to assistive technology');
assert.match(app, /data-recheck-scope/, 'each report area must support a live re-check');
assert.match(app, /Ergebnis geändert['"] : ['"]Keine Änderung/, 're-checks must report whether the result changed');
assert.match(app, /result\.checkedAt = new Date\(\)\.toISOString\(\)/, 're-checks must retain a visible timestamp');
assert.match(app, /Lizenzoptimierung & Kosten[\s\S]+Reports\.Read\.All[\s\S]+ReportSettings\.Read\.All/, 'license optimization must request live usage-report access');
assert.match(app, /assignedLicenses,licenseAssignmentStates/, 'license analysis must include assignment source and state');
assert.match(app, /data-age-filter/, 'sharing inventory must expose an age filter');
assert.match(app, /Files\.Read\.All/, 'sharing inventory must request delegated file access');
assert.match(app, /Microsoft-Hilfe[\s\S]+icon-external-link[\s\S]+admin-link/, 'findings must link to Microsoft help and the relevant admin center');
assert.match(app, /Number\(left\.id === 'service'\) - Number\(right\.id === 'service'\)/, 'service health must stay at the end of the report');
assert.match(app, /roleAssignments\?\$expand=principal['"]/, 'role assignments must use only Graph-supported query parameters');
assert.doesNotMatch(app, /roleAssignments\?[^'"\n]*\$top/, 'role assignments must not use unsupported $top');
for (const removedPermission of ['AuditLog.Read.All', 'IdentityRiskyUser.Read.All', 'DeviceManagementManagedDevices.Read.All', 'TeamMember.Read.All', 'SecurityAlert.Read.All', 'AccessReview.Read.All']) {
  assert.doesNotMatch(app, new RegExp(removedPermission.replaceAll('.', '\\.')), `${removedPermission} must not be requested`);
}

console.log('OAuth callback self-check passed');
