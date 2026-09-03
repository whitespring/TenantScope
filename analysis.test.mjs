import assert from 'node:assert/strict';
import {
  analyseAccess,
  analyseApplications,
  analyseCompliance,
  analyseDevices,
  analyseIdentities,
  analyseLicenses,
  analyseRoles,
  analyseSecurity,
  analyseSharing,
  analyseServiceHealth,
  analyseSites,
  analyseStorage,
  analyseTenant,
  analyseTeams,
  analyseUsage,
  compareTableValues,
  escapeMarkdown,
  fitMetricFontSize,
  graphUrl,
  parsePriceCsv,
  screenPublicShares,
  sortFindingsBySeverity,
  trustedMicrosoftUrl,
} from './analysis.mjs';

assert.equal(graphUrl('/users?$top=1'), 'https://graph.microsoft.com/v1.0/users?$top=1');
assert.equal(graphUrl('https://graph.microsoft.com/v1.0/users?$skiptoken=one'), 'https://graph.microsoft.com/v1.0/users?$skiptoken=one');
assert.throws(() => graphUrl('https://example.invalid/steal'), /Unzulässiges/);
assert.equal(trustedMicrosoftUrl('https://tenant.sharepoint.com/sites/one'), 'https://tenant.sharepoint.com/sites/one');
assert.equal(trustedMicrosoftUrl('https://example.invalid/phish'), null);
assert.equal(escapeMarkdown('<img src=x> [Klick](https://example.invalid)'), '&lt;img src=x&gt; \\[Klick\\]\\(https://example.invalid\\)');

const identities = analyseIdentities(
  [
    { id: '1', displayName: 'Ada', accountEnabled: true, userType: 'Member' },
    { id: '2', displayName: 'Grace', accountEnabled: false, userType: 'Member' },
    { id: '3', displayName: 'Linus', accountEnabled: true, userType: 'Guest' },
  ],
);
assert.deepEqual(identities.metrics.map(([value]) => value), ['3', '2', '1', '1']);
assert.equal(identities.findings[0].severity, 'info');
assert.deepEqual(identities.details.rows[0].slice(0, 4), ['Ada', '–', 'Member', 'Ja']);

const prices = parsePriceCsv('sku;preisProMonat\nM365_E3;32,50');
assert.ok(compareTableValues('2 GB', '900 MB') > 0);
assert.ok(compareTableValues('15.7.2026', '1.9.2026') < 0);
assert.ok(compareTableValues('Team 2', 'Team 10') < 0);
for (const empty of ['', '-', '–', '—']) assert.ok(compareTableValues(empty, '0') < 0);
assert.deepEqual(['–', '0', '2', ''].sort((left, right) => -compareTableValues(left, right)), ['2', '0', '–', '']);
assert.equal(fitMetricFontSize(31, 220, 260), 26);
assert.equal(fitMetricFontSize(31, 220, 120), 31);
assert.deepEqual(
  sortFindingsBySeverity(['ok', 'medium', 'info', 'high', 'error', 'low'].map((severity) => ({ severity }))).map(({ severity }) => severity),
  ['high', 'medium', 'low', 'error', 'info', 'ok'],
);
const licenses = analyseLicenses([{ skuId: 'sku-1', skuPartNumber: 'M365_E3', consumedUnits: 80, prepaidUnits: { enabled: 100 }, capabilityStatus: 'Enabled' }], [
  { displayName: 'Disabled Ada', accountEnabled: false, assignedLicenses: [{ skuId: 'sku-1' }] },
], prices);
assert.equal(licenses.metrics[1][0], '1');
assert.equal(licenses.metrics[3][0], '682,50 €');
assert.equal(licenses.findings[0].severity, 'medium');
assert.equal(licenses.details.rows[0][4], '20');
const freeSkuFiltered = analyseLicenses([
  { skuId: 'free', skuPartNumber: 'FLOW_FREE', consumedUnits: 2, prepaidUnits: { enabled: 10000 }, capabilityStatus: 'Enabled' },
  { skuId: 'premium', skuPartNumber: 'O365_BUSINESS_PREMIUM', consumedUnits: 3, prepaidUnits: { enabled: 4 }, capabilityStatus: 'Enabled' },
], [{ displayName: 'Ada', accountEnabled: true, assignedLicenses: [{ skuId: 'premium' }, { skuId: 'free' }] }]);
assert.equal(freeSkuFiltered.extraDetails.find(({ title }) => title === 'Einsparpotenzial nach Hebel').rows[0][1], '1');
assert.equal(freeSkuFiltered.extraDetails.find(({ title }) => title === 'Lizenzzuweisungen je Benutzer').rows[0][3], 'FLOW_FREE, O365_BUSINESS_PREMIUM');

const premiumPlans = [
  ['exchange', 'EXCHANGE_S_STANDARD'], ['sharepoint', 'SHAREPOINTSTANDARD'], ['teams', 'TEAMS1'], ['office', 'OFFICESUBSCRIPTION'],
].map(([servicePlanId, servicePlanName]) => ({ servicePlanId, servicePlanName, provisioningStatus: 'Success', appliesTo: 'User' }));
const basicPlans = premiumPlans.slice(0, 3);
const optimizedLicenses = analyseLicenses([
  { skuId: 'premium', skuPartNumber: 'PREMIUM', consumedUnits: 3, prepaidUnits: { enabled: 3 }, capabilityStatus: 'Enabled', servicePlans: premiumPlans },
  { skuId: 'basic', skuPartNumber: 'BASIC', consumedUnits: 0, prepaidUnits: { enabled: 2 }, capabilityStatus: 'Enabled', servicePlans: basicPlans },
], [
  { id: 'disabled', displayName: 'Disabled', userPrincipalName: 'disabled@example.com', accountEnabled: false, userType: 'Member', assignedLicenses: [{ skuId: 'premium', disabledPlans: [] }], licenseAssignmentStates: [{ skuId: 'premium', state: 'Active', assignedByGroup: null }] },
  { id: 'inactive', displayName: 'Inactive', userPrincipalName: 'inactive@example.com', accountEnabled: true, userType: 'Member', assignedLicenses: [{ skuId: 'premium', disabledPlans: [] }], licenseAssignmentStates: [{ skuId: 'premium', state: 'Active', assignedByGroup: 'group-1' }] },
  { id: 'light', displayName: 'Light User', userPrincipalName: 'light@example.com', accountEnabled: true, userType: 'Member', assignedLicenses: [{ skuId: 'premium', disabledPlans: [] }], licenseAssignmentStates: [{ skuId: 'premium', state: 'Active', assignedByGroup: null }] },
], new Map([['PREMIUM', 30], ['BASIC', 10]]), {
  complete: true,
  concealed: false,
  appReports: [{ userPrincipalName: 'light@example.com', lastActivityDate: '2026-09-01', details: [{ outlook: true, teams: true, web: true }] }],
  activeReports: [{ userPrincipalName: 'light@example.com', exchangeLastActivityDate: '2026-09-01', teamsLastActivityDate: '2026-09-01' }],
});
assert.deepEqual(optimizedLicenses.metrics.map(([value]) => value), ['2', '2', '1', '100,00 €']);
assert.equal(optimizedLicenses.extraDetails[0].rows[1][5], 'Gruppe');
assert.deepEqual(optimizedLicenses.extraDetails[1].rows[0].slice(2, 8), ['PREMIUM', 'Exchange, Teams', 'SharePoint/OneDrive, Microsoft 365 Desktop-Apps', 'BASIC', '20,00 €', 'Mittel']);
assert.equal(optimizedLicenses.extraDetails[2].rows[3][2], '100,00 €');

const storage = analyseStorage(
  [{ displayName: 'Ada', userPrincipalName: 'ada@example.com', userType: 'Member' }],
  [{ userPrincipalName: 'ada@example.com', storageUsedInBytes: 9 * 1024 ** 3, prohibitSendReceiveQuotaInBytes: 10 * 1024 ** 3 }],
  [{ ownerPrincipalName: 'ada@example.com', storageUsedInBytes: 2 * 1024 ** 3, storageAllocatedInBytes: 100 * 1024 ** 3 }],
);
assert.equal(storage.findings[0].severity, 'high');
assert.equal(storage.metrics[2][0], '11 GB / 110 GB');
assert.deepEqual(storage.details.rows[0].slice(2), ['9 GB', '10 GB', '90 %', '2 GB', '100 GB', '2 %', '11 GB']);

const applications = analyseApplications(
  [{ id: 'app-1', passwordCredentials: [{ displayName: 'Rotation 2026', endDateTime: '2026-09-20T00:00:00Z' }] }],
  new Date('2026-09-03T00:00:00Z'),
);
assert.equal(applications.findings[0].severity, 'medium');
assert.deepEqual(applications.details.rows[0].slice(1), ['Rotation 2026', 'Secret', '20.9.2026', 'Läuft ≤ 90 Tage ab']);

assert.equal(analyseTenant([{}], [{ id: 'example.com', isVerified: false }]).findings[0].severity, 'medium');
assert.equal(analyseRoles([{ roleDefinition: { displayName: 'Global Administrator' }, principal: { displayName: 'Ada' } }]).findings[0].severity, 'high');
assert.equal(analyseAccess([]).findings[0].severity, 'high');

const usage = analyseUsage(
  [{ id: '1', displayName: 'Ada Lovelace', userPrincipalName: 'ada@example.com', accountEnabled: true, userType: 'Member', assignedLicenses: [{}] }],
  [{ userPrincipalName: 'ada@example.com', lastActivityDate: '2026-09-01', details: [{ word: true, windows: true }] }],
  [{ userPrincipalName: 'ada@example.com', assignedProducts: ['O365_BUSINESS_PREMIUM', 'FLOW_FREE'] }],
);
assert.equal(usage.metrics[0][0], '1');
assert.equal(usage.details.rows[0][0], 'Ada Lovelace');
assert.equal(usage.details.rows[0][2], 'Word');
assert.equal(usage.details.rows[0][5], 'FLOW_FREE, O365_BUSINESS_PREMIUM');

const teamInventory = new Map([['t1:owners', []]]);
assert.equal(analyseTeams([{ id: 't1', displayName: 'Team 1', visibility: 'Private' }], teamInventory).findings[0].severity, 'high');

const sites = analyseSites(
  [{ id: 's1', displayName: 'Site 1', webUrl: 'https://example.sharepoint.com/sites/one' }],
  [{ siteUrl: 'https://example.sharepoint.com/sites/one', storageUsedInBytes: 10, storageAllocatedInBytes: 100 }],
  {},
  new Map([['s1', [{ link: { scope: 'anonymous' } }]]]),
);
assert.equal(sites.findings[0].severity, 'high');

const publicShareScreening = screenPublicShares(
  [{ id: 'd1' }],
  new Map([['d1', [
    { id: 'public', shared: { scope: 'anonymous' } },
    { id: 'internal', shared: { scope: 'organization' } },
    { id: 'unknown', shared: {} },
    { id: 'deleted', shared: { scope: 'anonymous' }, deleted: {} },
  ]]]),
);
assert.deepEqual([publicShareScreening.screenedItems, publicShareScreening.sharedRoots, publicShareScreening.unclassified, publicShareScreening.candidates.length], [4, 3, 1, 2]);
assert.deepEqual(publicShareScreening.itemsByDrive.get('d1').map(({ id }) => id), ['public', 'unknown']);

const sharing = analyseSharing(
  [{ id: 'd1', name: 'Dokumente', owner: { user: { displayName: 'Ada' } } }],
  new Map([['d1', [{ id: 'i1', name: 'Projekt', folder: {}, webUrl: 'https://example.sharepoint.com/projekt', shared: { sharedDateTime: '2026-01-01T00:00:00Z', sharedBy: { user: { displayName: 'Grace' } } } }]]]),
  new Map([['d1:i1', [{ roles: ['write'], link: { scope: 'anonymous' }, expirationDateTime: '2026-12-31T00:00:00Z' }]]]),
  new Date('2026-09-03T00:00:00Z'),
);
assert.deepEqual(sharing.metrics.map(([value]) => value), ['1', '1', '1', '0']);
assert.deepEqual(sharing.details.rows[0].slice(0, 6), ['Projekt', 'Ordner', 'Dokumente', 'Ada', 'Grace', '1.1.2026']);
assert.equal(sharing.details.ageColumn, 6);
assert.match(sharing.findings[0].title, /Schreibrecht/);

const unprotectedSharing = analyseSharing(
  [{ id: 'd1', name: 'Dokumente' }],
  new Map([['d1', [{ id: 'i1', name: 'Öffentlich', file: {}, shared: { scope: 'anonymous' } }]]]),
  new Map([['d1:i1', [{ roles: ['read'], link: { scope: 'anonymous' } }]]]),
);
assert.equal(unprotectedSharing.metrics[3][0], '1');
assert.match(unprotectedSharing.findings[0].title, /ohne Ablaufdatum/);

const authenticatedSharing = analyseSharing(
  [{ id: 'd1', name: 'Dokumente' }],
  new Map([['d1', [{ id: 'i1', name: 'Intern', file: {}, shared: { scope: 'organization' } }]]]),
  new Map([['d1:i1', [{ roles: ['write'], link: { scope: 'organization' } }]]]),
);
assert.equal(authenticatedSharing.metrics[1][0], '0');
assert.equal(authenticatedSharing.findings[0].severity, 'ok');

const resolvedFallbackSharing = analyseSharing(
  [{ id: 'd1', name: 'Dokumente' }],
  new Map([['d1', [{ id: 'i1', name: 'Ohne Hinweis', file: {}, shared: {} }]]]),
  new Map([['d1:i1', [{ roles: ['read'], link: { scope: 'organization' } }]]]),
  new Date(),
  0,
  new Map(),
  { sharedRoots: 1, unclassified: 1 },
);
assert.equal(resolvedFallbackSharing.unavailable, 0);
assert.equal(resolvedFallbackSharing.findings[0].severity, 'ok');

const throttledSharing = analyseSharing(
  [{ id: 'd1', name: 'Dokumente' }],
  new Map([['d1', [{ id: 'i1', name: 'Projekt', file: {}, shared: {} }]]]),
  new Map(),
  new Date('2026-09-03T00:00:00Z'),
  0,
  new Map([['d1:i1', { status: 429, code: 'TooManyRequests' }]]),
);
assert.match(throttledSharing.findings[0].title, /gedrosselt/);
assert.equal(throttledSharing.details.rows[0][8], 'TooManyRequests · HTTP 429');

const deniedSharing = analyseSharing(
  [{ id: 'd1', name: 'Dokumente' }],
  new Map([['d1', [{ id: 'i1', name: 'Projekt', file: {}, shared: {} }]]]),
  new Map(),
  new Date('2026-09-03T00:00:00Z'),
  0,
  new Map([['d1:i1', { status: 403, code: 'accessDenied' }]]),
);
assert.match(deniedSharing.findings[0].title, /Zugriff/);

const devices = analyseDevices([{ id: 'd1', accountEnabled: true, approximateLastSignInDateTime: '2026-07-01T00:00:00Z' }], new Date('2026-09-03T00:00:00Z'));
assert.equal(devices.findings[0].severity, 'medium');

const security = analyseSecurity([{ currentScore: 50, maxScore: 100 }]);
assert.equal(security.findings[0].severity, 'medium');

assert.equal(analyseCompliance([], []).findings[0].severity, 'medium');
assert.equal(analyseServiceHealth([], [{ title: 'Outage', classification: 'incident', status: 'serviceInterruption' }], []).findings[0].severity, 'high');

console.log('Live analysis self-check passed');
