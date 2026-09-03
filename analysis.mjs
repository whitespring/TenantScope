const day = 24 * 60 * 60 * 1000;

function finding(severity, title, description, action) {
  return { severity, title, description, action };
}

function percent(part, total) {
  return total ? Math.round((part / total) * 100) : 0;
}

function date(value) {
  return value ? new Intl.DateTimeFormat('de-DE').format(new Date(value)) : '–';
}

function bytes(value) {
  const amount = Number(value || 0);
  if (!amount) return '0 B';
  const units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB'];
  const unit = Math.min(Math.floor(Math.log(amount) / Math.log(1024)), units.length - 1);
  return `${new Intl.NumberFormat('de-DE', { maximumFractionDigits: 1 }).format(amount / 1024 ** unit)} ${units[unit]}`;
}

function currency(value) {
  return new Intl.NumberFormat('de-DE', { style: 'currency', currency: 'EUR' }).format(value || 0);
}

function normalized(value) {
  return String(value || '').trim().toLowerCase();
}

function sortedList(values) {
  return [...new Set(values.filter(Boolean))].sort(new Intl.Collator('de-DE', { numeric: true, sensitivity: 'base' }).compare).join(', ');
}

export function graphUrl(path) {
  const value = String(path || '');
  const url = new URL(value.startsWith('/') ? `https://graph.microsoft.com/v1.0${value}` : value);
  if (url.protocol !== 'https:' || url.hostname !== 'graph.microsoft.com' || url.port || url.username || url.password) throw new Error('Unzulässiges Microsoft-Graph-Ziel.');
  return url.href;
}

export function trustedMicrosoftUrl(value) {
  try {
    const url = new URL(String(value));
    const allowed = ['microsoft.com', 'microsoftonline.com', 'office.com', 'sharepoint.com'];
    return url.protocol === 'https:' && !url.port && !url.username && !url.password && allowed.some((domain) => url.hostname === domain || url.hostname.endsWith(`.${domain}`)) ? url.href : null;
  } catch {
    return null;
  }
}

export function escapeMarkdown(value) {
  return String(value ?? '').replaceAll('\\', '\\\\').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replace(/([`*_[\]{}()#+!|])/g, '\\$1').replace(/\r?\n/g, ' ');
}

export function compareTableValues(left, right) {
  const parse = (value) => {
    const text = String(value ?? '').replaceAll('\u00a0', ' ').trim();
    if (!text || /^[-–—]$/.test(text)) return [0, Number.NEGATIVE_INFINITY];
    const dateMatch = text.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:,?\s+(\d{1,2}):(\d{2}))?$/);
    if (dateMatch) return [0, Date.UTC(Number(dateMatch[3]), Number(dateMatch[2]) - 1, Number(dateMatch[1]), Number(dateMatch[4] || 0), Number(dateMatch[5] || 0))];
    const numberMatch = text.match(/^(-?[\d.]+(?:,\d+)?)\s*(B|KB|MB|GB|TB|%|€)?$/i);
    if (numberMatch) {
      const units = { B: 1, KB: 1024, MB: 1024 ** 2, GB: 1024 ** 3, TB: 1024 ** 4, '%': 1, '€': 1 };
      return [0, Number(numberMatch[1].replaceAll('.', '').replace(',', '.')) * (units[numberMatch[2]?.toUpperCase()] || 1)];
    }
    return [1, text];
  };
  const [leftType, leftValue] = parse(left);
  const [rightType, rightValue] = parse(right);
  if (leftType !== rightType) return leftType - rightType;
  return leftType === 0 ? leftValue - rightValue : new Intl.Collator('de-DE', { numeric: true, sensitivity: 'base' }).compare(leftValue, rightValue);
}

export function sortFindingsBySeverity(findings) {
  const rank = { high: 0, medium: 1, low: 2, error: 3, info: 4, ok: 5 };
  return [...findings].sort((left, right) => (rank[left.severity] ?? 99) - (rank[right.severity] ?? 99));
}

export function fitMetricFontSize(base, available, required) {
  return required > available && available > 0 ? Math.max(16, Math.floor(base * available / required)) : base;
}

function latestDate(values) {
  return values.filter(Boolean).sort().at(-1) || null;
}

export function parsePriceCsv(text) {
  const prices = new Map();
  for (const [index, rawLine] of String(text || '').replace(/^\uFEFF/, '').split(/\r?\n/).entries()) {
    const line = rawLine.trim();
    if (!line) continue;
    const separator = line.includes(';') ? ';' : ',';
    const [rawSku, rawPrice] = line.split(separator);
    if (index === 0 && normalized(rawSku).includes('sku')) continue;
    const price = Number(String(rawPrice || '').trim().replace(',', '.'));
    if (rawSku?.trim() && Number.isFinite(price) && price >= 0) prices.set(rawSku.trim().toUpperCase(), price);
  }
  return prices;
}

export function analyseTenant(organizations, domains) {
  const organization = organizations[0] || {};
  const unverified = domains.filter((domain) => !domain.isVerified);
  const federated = domains.filter((domain) => normalized(domain.authenticationType) === 'federated');
  const findings = [];

  if (unverified.length) findings.push(finding('medium', `${unverified.length} nicht verifizierte Domains`, 'Mindestens eine im Mandanten vorhandene Domain ist nicht vollständig verifiziert.', 'DNS-Konfiguration und weitere Verwendung der betroffenen Domains prüfen.'));
  if (federated.length) findings.push(finding('info', `${federated.length} föderierte Domains`, 'Anmeldungen dieser Domains hängen von einer externen Föderationsinfrastruktur ab.', 'Verfügbarkeit, Zertifikate und Notfallzugang der Föderation regelmäßig prüfen.'));
  if (!findings.length) findings.push(finding('ok', 'Tenant-Basis ohne Auffälligkeit', 'Alle gefundenen Domains sind verifiziert und werden cloudverwaltet authentifiziert.', 'Domainbestand und Benachrichtigungskontakte regelmäßig kontrollieren.'));

  return {
    records: domains.length,
    summary: `${domains.length} Domains und Tenant-Basis geprüft`,
    metrics: [[String(domains.length), 'Domains'], [String(domains.length - unverified.length), 'verifiziert'], [organization.onPremisesSyncEnabled ? 'Aktiv' : 'Nein', 'Verzeichnissynchronisierung']],
    findings,
    details: {
      title: 'Domains',
      columns: ['Domain', 'Standard', 'Initial', 'Verifiziert', 'Authentifizierung'],
      rows: domains.map((domain) => [domain.id, domain.isDefault ? 'Ja' : 'Nein', domain.isInitial ? 'Ja' : 'Nein', domain.isVerified ? 'Ja' : 'Nein', domain.authenticationType || '–']),
    },
  };
}

export function analyseRoles(assignments = []) {
  const privilegedNames = new Set(['Global Administrator', 'Privileged Role Administrator', 'Security Administrator', 'Conditional Access Administrator', 'Exchange Administrator', 'SharePoint Administrator', 'Teams Administrator', 'Intune Administrator', 'Application Administrator', 'Cloud Application Administrator', 'User Administrator']);
  const privileged = assignments.filter((assignment) => privilegedNames.has(assignment.roleDefinition?.displayName));
  const globalAdmins = assignments.filter((assignment) => assignment.roleDefinition?.displayName === 'Global Administrator');
  const orphaned = assignments.filter((assignment) => !assignment.principal);
  const findings = [];

  if (globalAdmins.length < 2) findings.push(finding('high', `Nur ${globalAdmins.length} Global Administrator(en) gefunden`, 'Für Wiederherstellung und Notfälle sollte mehr als ein kontrollierter administrativer Zugang verfügbar sein.', 'Mindestens zwei getrennte, besonders geschützte Notfall-Administratorkonten vorsehen.'));
  if (globalAdmins.length > 5) findings.push(finding('medium', `${globalAdmins.length} Global Administratoren`, 'Eine hohe Zahl dauerhaft global privilegierter Identitäten vergrößert die Angriffsfläche.', 'Bedarf prüfen und Aufgaben auf weniger weitreichende Rollen verteilen.'));
  if (orphaned.length) findings.push(finding('medium', `${orphaned.length} Rollenzuweisungen ohne lesbaren Principal`, 'Mindestens eine Zuweisung verweist auf ein nicht mehr auflösbares oder nicht lesbares Objekt.', 'Verwaiste Rollenzuweisungen und gelöschte Identitäten prüfen.'));
  if (!findings.length) findings.push(finding('ok', 'Administrative Rollen ohne Grundauffälligkeit', 'Anzahl und Auflösbarkeit der wichtigsten administrativen Zuweisungen sind unauffällig.', 'Dauerhafte Privilegien zusätzlich mit PIM und Access Reviews kontrollieren.'));

  return {
    records: assignments.length,
    summary: `${assignments.length} aktive Rollenzuweisungen geprüft`,
    metrics: [[String(assignments.length), 'Zuweisungen aktiv'], [String(privileged.length), 'hochkritisch aktiv'], [String(globalAdmins.length), 'Global Admins aktiv']],
    findings,
    details: {
      title: 'Administrative Rollenzuweisungen',
      columns: ['Rolle', 'Principal', 'Typ', 'Scope'],
      rows: assignments.map((assignment) => [assignment.roleDefinition?.displayName || assignment.roleDefinitionId, assignment.principal?.displayName || assignment.principalId, assignment.principal?.['@odata.type']?.split('.').at(-1) || 'Nicht lesbar', assignment.directoryScopeId || '/']),
    },
  };
}

export function analyseAccess(policies = []) {
  const enabledPolicies = policies.filter((policy) => policy.state === 'enabled');
  const disabledPolicies = policies.filter((policy) => policy.state === 'disabled');
  const findings = [];

  if (!enabledPolicies.length) findings.push(finding('high', 'Keine aktive Conditional-Access-Richtlinie gefunden', 'Microsoft Graph liefert keine Richtlinie im Zustand „enabled“. Report-only-Richtlinien setzen keine Zugriffsentscheidung durch.', 'Mindestens MFA, Legacy-Authentifizierung und Gerätezustand über Conditional Access absichern.'));
  if (disabledPolicies.length) findings.push(finding('info', `${disabledPolicies.length} deaktivierte Conditional-Access-Richtlinien`, 'Deaktivierte Richtlinien erhöhen nicht den Schutz und können auf Altbestand hinweisen.', 'Dokumentieren, archivieren oder reaktivieren, sofern weiterhin erforderlich.'));
  if (!findings.length) findings.push(finding('ok', 'Conditional Access ist aktiv', 'Mindestens eine aktive Richtlinie wurde gefunden.', 'Ausschlüsse, Break-glass-Konten und Richtlinienwirkung regelmäßig testen.'));

  return {
    records: policies.length,
    summary: `${policies.length} Conditional-Access-Richtlinien geprüft`,
    metrics: [[String(policies.length), 'Richtlinien'], [String(enabledPolicies.length), 'aktiv'], [String(disabledPolicies.length), 'deaktiviert']],
    findings,
    details: {
      title: 'Conditional-Access-Richtlinien',
      columns: ['Richtlinie', 'Status', 'Erstellt', 'Geändert'],
      rows: policies.map((policy) => [policy.displayName || policy.id, policy.state || '–', date(policy.createdDateTime), date(policy.modifiedDateTime)]),
    },
  };
}

export function analyseIdentities(users = []) {
  const active = users.filter((user) => user.accountEnabled);
  const guests = users.filter((user) => user.userType === 'Guest');
  const disabled = users.filter((user) => !user.accountEnabled);
  const findings = disabled.length
    ? [finding('info', `${disabled.length} deaktivierte Benutzerkonten`, 'Deaktivierte Konten verbleiben im Verzeichnis.', 'Bedarf, Aufbewahrung und Lizenzzuweisungen regelmäßig prüfen.')]
    : [finding('ok', 'Keine deaktivierten Benutzerkonten gefunden', 'Alle gefundenen Konten sind aktiviert.', 'Kontenbestand und Gastzugriffe regelmäßig kontrollieren.')];

  return {
    records: users.length,
    summary: `${users.length} Benutzerkonten geprüft`,
    metrics: [[String(users.length), 'Benutzerkonten'], [String(active.length), 'aktiv'], [String(guests.length), 'Gastkonten'], [String(disabled.length), 'deaktiviert']],
    findings,
    details: {
      title: 'Benutzerbestand',
      columns: ['Benutzer', 'Konto', 'Typ', 'Aktiv', 'Erstellt', 'Passwort geändert', 'On-Prem synchronisiert'],
      rows: users.map((user) => [user.displayName || user.userPrincipalName || user.id, user.userPrincipalName || '–', user.userType || '–', user.accountEnabled ? 'Ja' : 'Nein', date(user.createdDateTime), date(user.lastPasswordChangeDateTime), user.onPremisesSyncEnabled ? 'Ja' : 'Nein']),
    },
  };
}

const licenseWorkloadRules = [
  ['Exchange', /EXCHANGE/i],
  ['SharePoint/OneDrive', /SHAREPOINT|ONEDRIVE/i],
  ['Teams', /TEAMS|MCOSTANDARD/i],
  ['Microsoft 365 Desktop-Apps', /OFFICESUBSCRIPTION|O365PROPLUS|M365APPS|OFFICE_PRO_PLUS/i],
];

function skuWorkloads(sku, disabledPlans = []) {
  const disabled = new Set(disabledPlans.map(normalized));
  const names = (sku?.servicePlans || []).filter((plan) =>
    (!plan.appliesTo || plan.appliesTo === 'User') && plan.provisioningStatus !== 'Disabled' && !disabled.has(normalized(plan.servicePlanId))).map((plan) => plan.servicePlanName || '');
  return new Set(licenseWorkloadRules.filter(([, pattern]) => names.some((name) => pattern.test(name))).map(([label]) => label));
}

function reportedLicenseActivity(user, appByUser, activeByUser) {
  const key = normalized(user.userPrincipalName);
  const apps = appByUser.get(key) || {};
  const appUsage = apps.details?.[0] || apps;
  const active = activeByUser.get(key) || {};
  const desktopApp = ['outlook', 'word', 'excel', 'powerPoint', 'oneNote'].some((name) => appUsage[`${name}Windows`] || appUsage[`${name}Mac`]);
  const workloads = [
    ['Exchange', active.exchangeLastActivityDate || appUsage.outlook],
    ['SharePoint/OneDrive', active.oneDriveLastActivityDate || active.sharePointLastActivityDate],
    ['Teams', active.teamsLastActivityDate || appUsage.teams],
    ['Microsoft 365 Desktop-Apps', desktopApp],
  ].filter(([, used]) => used).map(([name]) => name);
  const lastActivity = latestDate([apps.lastActivityDate, active.exchangeLastActivityDate, active.oneDriveLastActivityDate, active.sharePointLastActivityDate, active.teamsLastActivityDate]);
  return { workloads: new Set(workloads), lastActivity, hasActivity: Boolean(lastActivity || workloads.length) };
}

export function analyseLicenses(skus, users = [], prices = new Map(), usage) {
  // ponytail: Free/viral SKUs and downgrade coverage are inferred from live SKU names/service plans; use a maintained commerce catalog if exact product semantics become available.
  const paid = (sku) => prices.has(String(sku?.skuPartNumber || '').toUpperCase()) || !/(FREE|VIRAL|TRIAL|EXPLORATORY)/i.test(sku?.skuPartNumber || '');
  const capacitySkus = skus.filter(paid);
  const capacityIds = new Set(capacitySkus.map((sku) => normalized(sku.skuId)));
  const skuById = new Map(skus.map((sku) => [normalized(sku.skuId), sku]));
  const enabled = capacitySkus.reduce((sum, sku) => sum + Number(sku.prepaidUnits?.enabled || 0), 0);
  const consumed = capacitySkus.reduce((sum, sku) => sum + Number(sku.consumedUnits || 0), 0);
  const available = Math.max(0, enabled - consumed);
  const freeShare = percent(available, enabled);
  const unhealthy = skus.filter((sku) => !['Enabled', 'Warning'].includes(sku.capabilityStatus));
  const licensedUsers = users.filter((user) => user.assignedLicenses?.length);
  const allAssignments = licensedUsers.flatMap((user) => user.assignedLicenses.map((license) => {
    const sku = skuById.get(normalized(license.skuId));
    const states = (user.licenseAssignmentStates || []).filter((item) => normalized(item.skuId) === normalized(license.skuId));
    const state = states[0];
    const direct = states.some((item) => !item.assignedByGroup);
    const grouped = states.some((item) => item.assignedByGroup);
    return {
      user, license, sku, state,
      price: prices.get(String(sku?.skuPartNumber || '').toUpperCase()),
      source: direct && grouped ? 'Direkt + Gruppe' : grouped ? 'Gruppe' : direct ? 'Direkt' : '–',
      disabledPlans: license.disabledPlans?.length ? license.disabledPlans : state?.disabledPlans || [],
    };
  })).filter(({ sku }) => sku);
  const assignments = allAssignments.filter(({ sku }) => capacityIds.has(normalized(sku.skuId)));
  const appByUser = new Map((usage?.appReports || []).map((entry) => [normalized(entry.userPrincipalName), entry]));
  const activeByUser = new Map((usage?.activeReports || []).map((entry) => [normalized(entry.userPrincipalName), entry]));
  const activity = new Map(users.map((user) => [normalized(user.userPrincipalName), reportedLicenseActivity(user, appByUser, activeByUser)]));
  const removable = assignments.filter(({ user }) => !user.accountEnabled || (usage?.complete && user.userType !== 'Guest' && !activity.get(normalized(user.userPrincipalName))?.hasActivity)).map((item) => ({
    ...item,
    reason: item.user.accountEnabled ? 'Keine Aktivität in 90 Tagen' : 'Konto deaktiviert',
    confidence: item.user.accountEnabled ? 'Mittel' : 'Hoch',
  }));
  const removableKeys = new Set(removable.map(({ user, sku }) => `${user.id}:${sku.skuId}`));
  const downsize = assignments.flatMap((item) => {
    if (removableKeys.has(`${item.user.id}:${item.sku.skuId}`) || !usage?.complete || !item.user.accountEnabled || item.user.userType === 'Guest') return [];
    const observed = activity.get(normalized(item.user.userPrincipalName))?.workloads || new Set();
    const current = skuWorkloads(item.sku, item.disabledPlans);
    const unused = [...current].filter((workload) => !observed.has(workload));
    if (!observed.size || current.size < 3 || unused.length < 2) return [];
    const alternative = item.price == null ? null : capacitySkus.map((sku) => ({ sku, workloads: skuWorkloads(sku), price: prices.get(String(sku.skuPartNumber || '').toUpperCase()) }))
      .filter((candidate) => candidate.sku.skuId !== item.sku.skuId && candidate.price != null && candidate.price < item.price && candidate.workloads.size < current.size
        && [...observed].every((workload) => candidate.workloads.has(workload)) && [...candidate.workloads].every((workload) => current.has(workload)))
      .sort((left, right) => left.price - right.price)[0];
    const nonObservable = (item.sku.servicePlans || []).filter((plan) => !licenseWorkloadRules.some(([, pattern]) => pattern.test(plan.servicePlanName || ''))).length;
    return [{ ...item, observed, current, unused, alternative, saving: alternative ? item.price - alternative.price : null, confidence: alternative && !nonObservable ? 'Mittel' : 'Niedrig', nonObservable }];
  });
  const freeSeatValue = capacitySkus.reduce((sum, sku) => sum + Math.max(0, Number(sku.prepaidUnits?.enabled || 0) - Number(sku.consumedUnits || 0)) * Number(prices.get(String(sku.skuPartNumber || '').toUpperCase()) || 0), 0);
  const removableValue = removable.reduce((sum, item) => sum + Number(item.price || 0), 0);
  const downsizeValue = downsize.reduce((sum, item) => sum + Number(item.saving || 0), 0);
  const potential = freeSeatValue + removableValue + downsizeValue;
  const missingPrices = capacitySkus.filter((sku) => !prices.has(String(sku.skuPartNumber || '').toUpperCase())).length;
  const disabledAssignments = removable.filter(({ user }) => !user.accountEnabled);
  const inactiveAssignments = removable.filter(({ user }) => user.accountEnabled);
  const findings = [];

  if (unhealthy.length) findings.push(finding('high', `${unhealthy.length} Lizenzprodukte nicht im Status „Enabled“`, 'Microsoft Graph meldet mindestens ein abonniertes Produkt in einem abweichenden Bereitstellungsstatus.', 'Betroffene Subscriptions und Zahlungs- beziehungsweise Verlängerungsstatus prüfen.'));
  if (disabledAssignments.length) findings.push(finding('medium', `${disabledAssignments.length} Lizenzzuweisungen an deaktivierten Konten`, prices.size ? `Der erfassbare Monatswert dieser Rückgabe-Kandidaten beträgt ${currency(disabledAssignments.reduce((sum, item) => sum + Number(item.price || 0), 0))}.` : 'Deaktivierte Konten belegen weiterhin kostenpflichtige Lizenzplätze.', 'Zuweisung nach Aufbewahrungs- und Offboarding-Prüfung entfernen; gruppenbasierte Quellen über Gruppe oder Mitgliedschaft korrigieren.'));
  if (inactiveAssignments.length) findings.push(finding('medium', `${inactiveAssignments.length} Lizenzzuweisungen ohne Aktivität in 90 Tagen`, 'Die zuordenbaren M365-Berichte zeigen für diese aktiven Konten keine Kernnutzung. Technische, gemeinsam genutzte, neue oder länger abwesende Konten können legitime Ausnahmen sein.', 'Bedarf mit Owner oder Fachbereich bestätigen und erst danach Lizenz zurückgeben oder Konto passend lizenzieren.'));
  if (downsize.length) findings.push(finding('low', `${downsize.length} potenziell überdimensionierte Lizenzzuweisungen`, 'Die Benutzer verwenden höchstens einen Teil der in ihrer Lizenz enthaltenen, über Reports beobachtbaren Kernworkloads. Security-, Compliance-, Telefonie-, Power-Platform- und Gerätefunktionen sind nicht vollständig als Nutzung messbar.', 'Kandidaten fachlich prüfen und nur dann auf eine kleinere Lizenz wechseln, wenn alle benötigten Zusatzfunktionen, Aufbewahrung und Sicherheitsanforderungen abgedeckt bleiben.'));
  if (available) findings.push(finding('info', `${available} freie Lizenzplätze als Vertragshebel`, prices.size ? `${freeShare} % der aktivierten Kapazität sind frei; der erfassbare Listenwert beträgt ${currency(freeSeatValue)} pro Monat.` : `${freeShare} % der aktivierten Kapazität sind nicht belegt.`, 'Betriebsreserve festlegen und darüber hinaus freie Seats zum nächsten vertraglich möglichen Termin reduzieren.'));
  if (usage?.concealed) findings.push(finding('info', 'Personenbezogene Nutzungszuordnung anonymisiert', 'Microsoft verbirgt die Benutzerkennungen in den Nutzungsreports; dadurch sind Inaktivitäts- und Downgrade-Kandidaten nicht belastbar zuordenbar.', 'Optional Klarnamen für Reports aktivieren oder die Lizenzberatung auf aggregierte Kapazitäten begrenzen.'));
  if (!prices.size) findings.push(finding('info', 'Keine SKU-Preisliste importiert', 'Microsoft Graph liefert Seats und Servicepläne, aber keine kundenspezifischen Einkaufspreise.', 'Optional die tatsächlichen monatlichen Nettopreise je SKU importieren, um Monats- und Jahrespotenziale zu berechnen.'));
  else if (missingPrices) findings.push(finding('info', `${missingPrices} kostenrelevante Lizenzprodukte ohne Preis`, 'Das ausgewiesene Euro-Potenzial umfasst nur Produkte, deren SKU in der importierten Preisliste enthalten ist.', 'Fehlende SKUs ergänzen und die Analyse erneut ausführen.'));
  if (!findings.length) findings.push(finding('ok', 'Keine unmittelbare Lizenzoptimierung erkannt', 'Es wurden weder freie Kapazität noch belastbare Rückgabe- oder Downgrade-Kandidaten identifiziert.', 'Nutzung, Rollenbedarf und Vertragsfristen quartalsweise erneut prüfen.'));

  const name = (user) => user.displayName || user.userPrincipalName || user.id;
  const money = (value) => value == null ? '–' : currency(value);
  return {
    records: skus.length,
    summary: `${skus.length} Lizenzprodukte, ${removable.length} Rückgabe- und ${downsize.length} Downgrade-Prüfkandidaten`,
    metrics: [[String(skus.length), 'Lizenzprodukte'], [String(removable.length), 'Rückgabe prüfen'], [String(downsize.length), 'Downgrade prüfen'], [prices.size ? currency(potential) : '–', 'erfasstes Potenzial/Monat']],
    findings,
    details: {
      title: 'Lizenzprodukte und Kapazitäten',
      columns: ['SKU', 'Status', 'Aktiviert', 'Belegt', 'Frei', 'Kernleistungen', 'Preis/Monat', 'Wert freie Seats'],
      rows: [...skus].sort((left, right) => compareTableValues(left.skuPartNumber, right.skuPartNumber)).map((sku) => {
        const skuEnabled = Number(sku.prepaidUnits?.enabled || 0);
        const skuConsumed = Number(sku.consumedUnits || 0);
        const unitPrice = prices.get(String(sku.skuPartNumber || '').toUpperCase());
        const skuFree = Math.max(0, skuEnabled - skuConsumed);
        return [sku.skuPartNumber || sku.skuId, sku.capabilityStatus || '–', String(skuEnabled), String(skuConsumed), String(skuFree), [...skuWorkloads(sku)].join(', ') || 'Nicht klassifiziert', money(unitPrice), money(unitPrice == null ? null : skuFree * unitPrice)];
      }),
    },
    extraDetails: [
      {
        title: 'Rückgabe- und Stilllegungskandidaten',
        columns: ['Priorität', 'Benutzer', 'Konto', 'Lizenz', 'Grund', 'Quelle', 'Letzte Aktivität', 'Potenzial/Monat', 'Konfidenz'],
        rows: removable.map((item) => [item.user.accountEnabled ? '2' : '1', name(item.user), item.user.userPrincipalName || '–', item.sku.skuPartNumber, item.reason, item.source, date(activity.get(normalized(item.user.userPrincipalName))?.lastActivity), money(item.price), item.confidence]),
      },
      {
        title: 'Downgrade-Prüfkandidaten',
        columns: ['Benutzer', 'Konto', 'Aktuelle Lizenz', 'Beobachtet (90 Tage)', 'Nicht beobachtete Kernleistungen', 'Mögliche Alternative', 'Potenzial/Monat', 'Konfidenz', 'Zusätzlich prüfen'],
        rows: downsize.map((item) => [name(item.user), item.user.userPrincipalName || '–', item.sku.skuPartNumber, [...item.observed].join(', '), item.unused.join(', '), item.alternative?.sku.skuPartNumber || 'Kleinere Lizenzklasse prüfen', money(item.saving), item.confidence, item.nonObservable ? `${item.nonObservable} weitere Servicepläne, insbesondere Security, Compliance und Geräteverwaltung` : 'Fachlichen Funktionsbedarf']),
      },
      {
        title: 'Einsparpotenzial nach Hebel',
        columns: ['Hebel', 'Kandidaten/Seats', 'Monat', 'Jahr', 'Einordnung'],
        rows: [
          ['Freie Kapazität', String(available), prices.size ? currency(freeSeatValue) : '–', prices.size ? currency(freeSeatValue * 12) : '–', 'Nur zum vertraglich möglichen Termin reduzierbar'],
          ['Rückgabe/Stilllegung', String(removable.length), prices.size ? currency(removableValue) : '–', prices.size ? currency(removableValue * 12) : '–', 'Erst nach Owner-, Aufbewahrungs- und Ausnahmeprüfung'],
          ['Downgrade', String(downsize.length), prices.size ? currency(downsizeValue) : '–', prices.size ? currency(downsizeValue * 12) : '–', 'Nur für bepreiste, abdeckende Alternativen berechnet'],
          ['Gesamtpotenzial', String(available + removable.length + downsize.length), prices.size ? currency(potential) : '–', prices.size ? currency(potential * 12) : '–', 'Erfasstes Potenzial, keine garantierte Einsparung'],
        ],
      },
      {
        title: 'Lizenzzuweisungen je Benutzer',
        columns: ['Benutzer', 'Konto', 'Aktiv', 'Lizenzprodukte', 'Quellen', 'Zuweisungsstatus', 'Deaktivierte Servicepläne', 'Monatswert'],
        rows: licensedUsers.map((user) => {
          const items = allAssignments.filter((item) => item.user === user);
          const chargeable = items.filter(({ sku }) => capacityIds.has(normalized(sku.skuId)));
          const monthly = chargeable.every((item) => item.price != null) ? money(chargeable.reduce((sum, item) => sum + Number(item.price || 0), 0)) : '–';
          return [name(user), user.userPrincipalName || '–', user.accountEnabled ? 'Ja' : 'Nein', sortedList(items.map((item) => item.sku.skuPartNumber)), sortedList(items.map((item) => item.source)), sortedList(items.map((item) => item.state?.state || '–')), String(items.reduce((sum, item) => sum + item.disabledPlans.length, 0)), monthly];
        }).sort((left, right) => compareTableValues(left[0], right[0])),
      },
    ],
  };
}

export function analyseStorage(users, mailboxReports, oneDriveReports) {
  const key = (value) => String(value || '').toLowerCase();
  const mailboxes = mailboxReports.filter((item) => !item.isDeleted);
  const drives = oneDriveReports.filter((item) => !item.isDeleted);
  const mailboxByUser = new Map(mailboxes.map((item) => [key(item.userPrincipalName), item]));
  const driveByUser = new Map(drives.map((item) => [key(item.ownerPrincipalName), item]));
  const directoryKeys = new Set(users.map((user) => key(user.userPrincipalName)));
  const reportOnly = [...new Map([...mailboxes, ...drives]
    .filter((item) => !directoryKeys.has(key(item.userPrincipalName || item.ownerPrincipalName)))
    .map((item) => [key(item.userPrincipalName || item.ownerPrincipalName), { displayName: item.displayName || item.ownerDisplayName, userPrincipalName: item.userPrincipalName || item.ownerPrincipalName, userType: 'Bericht' }])).values()];
  const allUsers = [...users, ...reportOnly];
  const rows = allUsers.map((user) => {
    const mailbox = mailboxByUser.get(key(user.userPrincipalName));
    const drive = driveByUser.get(key(user.userPrincipalName));
    const mailUsed = Number(mailbox?.storageUsedInBytes || 0);
    const mailQuota = Number(mailbox?.prohibitSendReceiveQuotaInBytes || mailbox?.prohibitSendQuotaInBytes || 0);
    const driveUsed = Number(drive?.storageUsedInBytes || 0);
    const driveQuota = Number(drive?.storageAllocatedInBytes || 0);
    return { user, mailbox, drive, mailUsed, mailQuota, driveUsed, driveQuota, mailPercent: mailQuota ? percent(mailUsed, mailQuota) : null, drivePercent: driveQuota ? percent(driveUsed, driveQuota) : null };
  }).sort((a, b) => b.mailUsed + b.driveUsed - a.mailUsed - a.driveUsed);
  const mailboxUsed = mailboxes.reduce((sum, item) => sum + Number(item.storageUsedInBytes || 0), 0);
  const mailboxQuota = mailboxes.reduce((sum, item) => sum + Number(item.prohibitSendReceiveQuotaInBytes || item.prohibitSendQuotaInBytes || 0), 0);
  const driveUsed = drives.reduce((sum, item) => sum + Number(item.storageUsedInBytes || 0), 0);
  const driveQuota = drives.reduce((sum, item) => sum + Number(item.storageAllocatedInBytes || 0), 0);
  const critical = rows.filter((item) => item.mailPercent >= 90 || item.drivePercent >= 90);
  const warning = rows.filter((item) => (item.mailPercent >= 80 || item.drivePercent >= 80) && item.mailPercent < 90 && item.drivePercent < 90);
  const findings = [];

  if (critical.length) findings.push(finding('high', `${critical.length} Konten mit mindestens 90 % Speicherauslastung`, 'Mindestens ein Postfach oder OneDrive nähert sich dem gemeldeten Limit.', 'Betroffene Konten prüfen, Daten bereinigen oder Kapazität rechtzeitig erweitern.'));
  if (warning.length) findings.push(finding('medium', `${warning.length} Konten mit mindestens 80 % Speicherauslastung`, 'Mindestens ein Postfach oder OneDrive hat nur noch begrenzte Reserve.', 'Wachstum beobachten und Verantwortliche frühzeitig informieren.'));
  if (reportOnly.length) findings.push(finding('info', `${reportOnly.length} Berichtskonten nicht mit Entra-Benutzern verknüpft`, 'Microsoft 365 liefert für diese Speicherberichte keine übereinstimmenden Benutzerkennungen; häufig ist die Anonymisierung der Nutzungsberichte aktiv.', 'Im Microsoft 365 Admin Center die Einstellung zur Anzeige verborgener Benutzerangaben prüfen.'));
  if (!findings.length) findings.push(finding('ok', 'Keine kritische Speicherauslastung erkannt', 'Kein auswertbares Postfach oder OneDrive liegt bei mindestens 80 % des gemeldeten Limits.', 'Speicherauslastung regelmäßig erneut prüfen.'));

  return {
    records: rows.length,
    summary: `${rows.length} Benutzerkonten und deren Speicher geprüft`,
    metrics: [[`${bytes(mailboxUsed)} / ${bytes(mailboxQuota)}`, 'Postfächer belegt / Limit'], [`${bytes(driveUsed)} / ${bytes(driveQuota)}`, 'OneDrive belegt / Kapazität'], [`${bytes(mailboxUsed + driveUsed)} / ${bytes(mailboxQuota + driveQuota)}`, 'gesamt belegt / Kapazität']],
    findings,
    details: {
      title: 'Speicher je Benutzer',
      columns: ['Benutzer', 'Konto', 'Postfach belegt', 'Postfach-Limit', 'Postfach', 'OneDrive belegt', 'OneDrive-Kapazität', 'OneDrive', 'Gesamt belegt'],
      rows: rows.map(({ user, mailbox, drive, mailUsed, mailQuota, driveUsed, driveQuota, mailPercent, drivePercent }) => [
        user.displayName || user.userPrincipalName || '–', user.userPrincipalName || '–', mailbox ? bytes(mailUsed) : '–', mailbox ? bytes(mailQuota) : '–', mailPercent == null ? '–' : `${mailPercent} %`, drive ? bytes(driveUsed) : '–', drive ? bytes(driveQuota) : '–', drivePercent == null ? '–' : `${drivePercent} %`, bytes(mailUsed + driveUsed),
      ]),
    },
  };
}

export function analyseUsage(users = [], appReports = [], activeReports = [], teamsReports = [], emailReports = [], copilotReports = [], reportSettings = {}) {
  const byKey = (items, property = 'userPrincipalName') => new Map(items.filter((item) => item[property]).map((item) => [normalized(item[property]), item]));
  const appsByUser = byKey(appReports);
  const activeByUser = byKey(activeReports);
  const teamsByUser = byKey(teamsReports);
  const emailByUser = byKey(emailReports);
  const copilotByUser = byKey(copilotReports);
  const concealed = reportSettings.displayConcealedNames === true;
  const reportKeys = new Set([...appsByUser.keys(), ...activeByUser.keys(), ...teamsByUser.keys(), ...emailByUser.keys(), ...copilotByUser.keys()]);
  const directoryByUser = byKey(users);
  const allKeys = concealed ? reportKeys : new Set([...directoryByUser.keys(), ...reportKeys]);
  const rows = [...allKeys].map((key) => {
    const user = directoryByUser.get(key);
    const apps = appsByUser.get(key) || {};
    const appUsage = apps.details?.[0] || apps;
    const active = activeByUser.get(key) || {};
    const teams = teamsByUser.get(key) || {};
    const email = emailByUser.get(key) || {};
    const copilot = copilotByUser.get(key) || {};
    const usedApps = [['Outlook', appUsage.outlook], ['Word', appUsage.word], ['Excel', appUsage.excel], ['PowerPoint', appUsage.powerPoint], ['OneNote', appUsage.oneNote], ['Teams', appUsage.teams]].filter(([, used]) => used).map(([name]) => name);
    const platforms = [['Windows', appUsage.windows], ['Mac', appUsage.mac], ['Mobile', appUsage.mobile], ['Web', appUsage.web]].filter(([, used]) => used).map(([name]) => name);
    const serviceDates = [['Exchange', active.exchangeLastActivityDate], ['OneDrive', active.oneDriveLastActivityDate], ['SharePoint', active.sharePointLastActivityDate], ['Teams', active.teamsLastActivityDate], ['Viva Engage', active.yammerLastActivityDate]].filter(([, value]) => value);
    const copilotApps = [['Chat', copilot.copilotChatLastActivityDate], ['Teams', copilot.microsoftTeamsCopilotLastActivityDate], ['Word', copilot.wordCopilotLastActivityDate], ['Excel', copilot.excelCopilotLastActivityDate], ['PowerPoint', copilot.powerPointCopilotLastActivityDate], ['Outlook', copilot.outlookCopilotLastActivityDate], ['OneNote', copilot.oneNoteCopilotLastActivityDate], ['Loop', copilot.loopCopilotLastActivityDate], ['Edge', copilot.edgeLastActivityDate], ['Agent', copilot.copilotAgentLastActivityDate]].filter(([, value]) => value);
    const lastActivity = latestDate([apps.lastActivityDate, teams.lastActivityDate, email.lastActivityDate, copilot.lastActivityDate, ...serviceDates.map(([, value]) => value)]);
    return {
      user, key, apps, active, teams, email, copilot, usedApps, platforms, serviceDates, copilotApps, lastActivity,
      hasActivity: Boolean(lastActivity || usedApps.length || serviceDates.length),
    };
  }).sort((a, b) => String(b.lastActivity || '').localeCompare(String(a.lastActivity || '')));
  const licensedUsers = users.filter((user) => user.accountEnabled && user.userType !== 'Guest' && user.assignedLicenses?.length);
  const inactiveLicensed = concealed ? [] : licensedUsers.filter((user) => !rows.find((row) => row.key === normalized(user.userPrincipalName))?.hasActivity);
  const copilotInactive = copilotReports.filter((entry) => !entry.lastActivityDate);
  const findings = [];

  if (inactiveLicensed.length) findings.push(finding('medium', `${inactiveLicensed.length} lizenzierte Konten ohne Aktivität im 90-Tage-Fenster`, 'Für diese aktiven Konten wurde in den ausgewerteten Microsoft-365-Reports keine Nutzung gefunden.', 'Fachlichen Bedarf und passende Lizenzstufe prüfen; technische und gemeinsam genutzte Konten separat bewerten.'));
  if (copilotInactive.length) findings.push(finding('medium', `${copilotInactive.length} Copilot-Konten ohne gemeldete Aktivität`, 'Der Copilot-Nutzungsreport enthält lizenzierte oder aktivierte Konten ohne letzte Nutzung.', 'Enablement, Lizenzbedarf und geeignete Anwendungsszenarien prüfen.'));
  if (concealed) findings.push(finding('info', 'Benutzernamen in Reports verborgen', 'Die Tenant-Einstellung anonymisiert Benutzer-, Gruppen- und Site-Angaben in Microsoft-365-Nutzungsreports.', 'Für personenbezogene Lizenzoptimierung muss ein Administrator die Report-Anonymisierung bewusst deaktivieren.'));
  if (!findings.length) findings.push(finding('ok', 'Keine deutliche Nutzungslücke erkannt', 'Alle zuordenbaren aktiven Lizenzkonten zeigen Aktivität in mindestens einem ausgewerteten 90-Tage-Report.', 'Nutzungstrends und Lizenzbedarf quartalsweise neu bewerten.'));

  return {
    records: rows.length,
    summary: `${rows.length} Nutzerprofile aus Microsoft-365-Reports zusammengeführt`,
    metrics: [[String(rows.filter((row) => row.hasActivity).length), 'aktive Nutzerprofile'], [String(inactiveLicensed.length), 'lizenzierte ohne Aktivität'], [String(copilotReports.length), 'Copilot-Profile'], [concealed ? 'Aktiv' : 'Nein', 'Anonymisierung']],
    findings,
    details: {
      title: 'Nutzung je Benutzer (90 Tage)',
      columns: ['Benutzer', 'Letzte Aktivität', 'Apps', 'Plattformen', 'Dienste', 'Lizenzprodukte', 'Teams Chats', 'Calls', 'Meetings', 'E-Mail gesendet/gelesen/empfangen', 'Copilot-Apps'],
      rows: rows.map((row) => [
        row.user?.displayName || row.user?.userPrincipalName || row.apps.userPrincipalName || row.active.userPrincipalName || row.teams.userPrincipalName || row.email.userPrincipalName || row.copilot.userPrincipalName || row.key,
        date(row.lastActivity), row.usedApps.join(', ') || '–', row.platforms.join(', ') || '–', row.serviceDates.map(([name]) => name).join(', ') || '–', sortedList(row.active.assignedProducts || []) || '–',
        String(Number(row.teams.teamChatMessageCount || 0) + Number(row.teams.privateChatMessageCount || 0)), String(row.teams.callCount || 0), String(row.teams.meetingCount || 0),
        `${row.email.sendCount || 0}/${row.email.readCount || 0}/${row.email.receiveCount || 0}`, row.copilotApps.map(([name]) => name).join(', ') || '–',
      ]),
    },
  };
}

export function analyseTeams(teams, inventory = new Map(), groups = teams) {
  const count = (team, kind) => inventory.get(`${team.id}:${kind}`)?.length;
  const ownerless = teams.filter((team) => count(team, 'owners') === 0);
  const unknownOwners = teams.filter((team) => count(team, 'owners') == null);
  const publicTeams = teams.filter((team) => team.visibility === 'Public');
  const findings = [];

  if (ownerless.length) findings.push(finding('high', `${ownerless.length} Teams ohne Owner`, 'Für diese Teams hat Microsoft Graph keine verantwortliche Person geliefert.', 'Mindestens zwei fachlich verantwortliche Owner je aktivem Team festlegen.'));
  if (publicTeams.length) findings.push(finding('low', `${publicTeams.length} öffentliche Teams`, 'Öffentliche Teams sind tenantweit auffindbar und können von Beschäftigten ohne Owner-Freigabe betreten werden.', 'Prüfen, ob die öffentliche Sichtbarkeit fachlich beabsichtigt ist.'));
  if (unknownOwners.length) findings.push(finding('error', `Owner-Status für ${unknownOwners.length} Teams nicht lesbar`, 'Einzelne Owner-Abfragen konnten von Microsoft Graph nicht ausgewertet werden.', 'Berechtigung Group.Read.All und mögliche Graph-Drosselung prüfen.'));
  if (!findings.length) findings.push(finding('ok', 'Team-Ownership ohne Auffälligkeit', 'Alle gefundenen Teams besitzen mindestens eine verantwortliche Person und sind nicht öffentlich.', 'Owner-Lebenszyklus regelmäßig erneut prüfen.'));

  return {
    records: teams.length,
    summary: `${teams.length} Teams geprüft`,
    metrics: [[String(groups.length), 'Entra-Gruppen'], [String(teams.length), 'Teams'], [String(ownerless.length), 'ohne Owner'], [String(publicTeams.length), 'öffentlich']],
    findings,
    details: {
      title: 'Teams und Owner',
      columns: ['Team', 'Sichtbarkeit', 'Owner', 'Erstellt', 'Verlängert'],
      rows: teams.map((team) => [team.displayName || team.id, team.visibility || '–', count(team, 'owners') == null ? 'Nicht lesbar' : String(count(team, 'owners')), date(team.createdDateTime), date(team.renewedDateTime)]),
    },
    extraDetails: [
      {
        title: 'Entra-Gruppen',
        columns: ['Gruppe', 'Typ', 'Sichtbarkeit', 'Mail', 'Security', 'Erstellt', 'Verlängert', 'Team'],
        rows: groups.map((group) => [group.displayName || group.id, group.groupTypes?.includes('Unified') ? 'Microsoft 365' : group.groupTypes?.includes('DynamicMembership') ? 'Dynamisch' : 'Security/Verteiler', group.visibility || '–', group.mailEnabled ? 'Ja' : 'Nein', group.securityEnabled ? 'Ja' : 'Nein', date(group.createdDateTime), date(group.renewedDateTime), group.resourceProvisioningOptions?.includes('Team') ? 'Ja' : 'Nein']),
      },
    ],
  };
}

export function analyseSites(sites, usage = [], settings = {}, permissions = new Map(), now = new Date()) {
  const cutoff = new Date(now.getTime() - 180 * day);
  const usageByUrl = new Map(usage.map((entry) => [normalized(entry.siteUrl), entry]));
  const stale = sites.filter((site) => {
    const last = usageByUrl.get(normalized(site.webUrl))?.lastActivityDate || site.lastModifiedDateTime;
    return last && new Date(last) < cutoff;
  });
  const reportSites = usage.filter((entry) => !entry.isDeleted);
  const used = reportSites.reduce((sum, entry) => sum + Number(entry.storageUsedInBytes || 0), 0);
  const allocated = reportSites.reduce((sum, entry) => sum + Number(entry.storageAllocatedInBytes || 0), 0);
  const anonymousLinks = sites.reduce((sum, site) => sum + (permissions.get(site.id) || []).filter((permission) => normalized(permission.link?.scope) === 'anonymous').length, 0);
  const findings = [];
  if (stale.length) findings.push(finding('medium', `${stale.length} Sites seit mehr als 180 Tagen inaktiv`, 'Nutzungsreport oder Site-Zeitstempel deutet auf mögliche Altbestände hin.', 'Fachliche Verantwortlichkeit, Aufbewahrung und tatsächliche Nutzung dieser Sites prüfen.'));
  if (anonymousLinks) findings.push(finding('high', `${anonymousLinks} anonyme Root-Freigabelinks gefunden`, 'Mindestens ein Root-Drive einer Site enthält einen Link, der ohne Anmeldung verwendet werden kann.', 'Anonyme Links fachlich prüfen und externes Sharing tenantweit begrenzen.'));
  if (settings.sharingCapability && normalized(settings.sharingCapability).includes('externaluserandguestsharing')) findings.push(finding('info', 'Externes und anonymes Sharing tenantweit möglich', 'Die SharePoint-Tenant-Einstellung erlaubt die weiteste Form externer Freigabe.', 'Prüfen, ob restriktivere Standardwerte oder Ablaufregeln erforderlich sind.'));
  if (!findings.length) findings.push(finding('ok', 'SharePoint-Grundprüfung ohne Auffälligkeit', 'Keine lange inaktiven Sites oder anonymen Root-Freigaben erkannt.', 'Berechtigungsvererbung und einzelne Dateien später über einen Deep Collector prüfen.'));

  return {
    records: Math.max(sites.length, reportSites.length),
    summary: `${sites.length} auffindbare SharePoint-Sites und ${reportSites.length} Nutzungsdatensätze geprüft`,
    metrics: [[String(sites.length), 'auffindbare Sites'], [String(stale.length), '> 180 Tage inaktiv'], [`${bytes(used)} / ${bytes(allocated)}`, 'Speicher belegt / zugewiesen'], [String(anonymousLinks), 'anonyme Root-Links']],
    findings,
    details: {
      title: 'Site-Bestand und Nutzung',
      columns: ['Site', 'URL', 'Letzte Aktivität', 'Speicher', 'Kapazität', 'Dateien', 'Root-Berechtigungen', 'Anonyme Links'],
      rows: sites.map((site) => {
        const report = usageByUrl.get(normalized(site.webUrl));
        const sitePermissions = permissions.get(site.id);
        return [site.displayName || site.name || site.id, site.webUrl || '–', date(report?.lastActivityDate || site.lastModifiedDateTime), report ? bytes(report.storageUsedInBytes) : 'Nicht zuordenbar', report ? bytes(report.storageAllocatedInBytes) : 'Nicht zuordenbar', report ? String(report.fileCount || 0) : '–', sitePermissions == null ? 'Nicht lesbar' : String(sitePermissions.length), sitePermissions == null ? '–' : String(sitePermissions.filter((permission) => normalized(permission.link?.scope) === 'anonymous').length)];
      }),
    },
    extraDetails: [
      {
        title: 'SharePoint-Nutzungsreport',
        columns: ['Site/Report-ID', 'Letzte Aktivität', 'Speicher', 'Kapazität', 'Dateien', 'Seitenaufrufe'],
        rows: reportSites.map((item) => [item.siteUrl || item.siteId || 'Verborgen', date(item.lastActivityDate), bytes(item.storageUsedInBytes), bytes(item.storageAllocatedInBytes), String(item.fileCount || 0), String(item.pageViewCount || 0)]),
      },
      {
        title: 'SharePoint- und OneDrive-Tenant-Einstellungen',
        columns: ['Einstellung', 'Wert'],
        rows: Object.entries(settings).filter(([, value]) => ['string', 'number', 'boolean'].includes(typeof value)).map(([key, value]) => [key, String(value)]),
      },
    ],
  };
}

function sharedIdentity(identitySet = {}) {
  const identity = identitySet.user || identitySet.siteUser || identitySet.group || identitySet.application;
  return identity?.displayName || identity?.email || identity?.loginName || identity?.id;
}

export function analyseSharing(drives = [], itemsByDrive = new Map(), permissionsByItem = new Map(), now = new Date(), unreadableDrives = 0, permissionFailures = new Map()) {
  const scopeLabel = { anonymous: 'Jeder mit Link', organization: 'Organisation', users: 'Bestimmte Personen' };
  const entries = drives.flatMap((drive) => (itemsByDrive.get(drive.id) || [])
    .filter((item) => item.shared && !item.deleted)
    .map((item) => {
      const permissions = permissionsByItem.get(`${drive.id}:${item.id}`);
      const permissionFailure = permissionFailures.get(`${drive.id}:${item.id}`);
      const scopes = [...new Set((permissions || []).map((permission) => permission.link?.scope).filter(Boolean))];
      const roles = [...new Set((permissions || []).flatMap((permission) => permission.roles || []))];
      const recipients = [...new Set((permissions || []).flatMap((permission) => [
        sharedIdentity(permission.grantedToV2),
        ...(permission.grantedToIdentitiesV2 || []).map(sharedIdentity),
        permission.invitation?.email,
        permission.link?.scope === 'anonymous' ? 'Jeder mit Link' : null,
        permission.link?.scope === 'organization' ? 'Alle in der Organisation' : null,
      ]).filter(Boolean))];
      const expirations = (permissions || []).map((permission) => permission.expirationDateTime).filter((value) => value && !value.startsWith('0001')).sort();
      const sharedDate = item.shared.sharedDateTime;
      const ageDays = sharedDate ? Math.max(0, Math.floor((now - new Date(sharedDate)) / day)) : null;
      return {
        drive, item, permissions, permissionFailure, ageDays,
        scopes: scopes.length ? scopes : item.shared.scope ? [item.shared.scope] : [],
        roles,
        recipients,
        expiration: expirations[0],
      };
    }))
    .sort((left, right) => (right.ageDays ?? -1) - (left.ageDays ?? -1));
  const anonymous = entries.filter((entry) => entry.scopes.includes('anonymous'));
  const writable = entries.filter((entry) => entry.roles.includes('write'));
  const old = entries.filter((entry) => entry.ageDays >= 180);
  const unreadable = entries.filter((entry) => entry.permissions == null);
  const accessDenied = unreadable.filter((entry) => [401, 403].includes(entry.permissionFailure?.status));
  const throttled = unreadable.filter((entry) => [429, 503, 504].includes(entry.permissionFailure?.status));
  const otherUnreadable = unreadable.filter((entry) => ![401, 403, 429, 503, 504].includes(entry.permissionFailure?.status));
  const findings = [];

  if (anonymous.length) findings.push(finding('high', `${anonymous.length} Dateien oder Ordner mit anonymem Link`, 'Diese Inhalte können ohne Anmeldung über einen weitergegebenen Link geöffnet werden.', 'Geschäftlichen Bedarf prüfen, anonyme Links löschen oder durch Freigaben für bestimmte Personen ersetzen.'));
  if (writable.length) findings.push(finding('medium', `${writable.length} geteilte Inhalte mit Schreibrecht`, 'Mindestens eine aktuelle Freigabe erlaubt Änderungen am Inhalt.', 'Empfänger, Bearbeitungsbedarf und mögliche Nur-Lese-Alternativen prüfen.'));
  if (old.length) findings.push(finding('medium', `${old.length} Freigaben seit mindestens 180 Tagen`, 'Lange bestehende Freigaben können ihren ursprünglichen Zweck überdauern.', 'Owner und fachlichen Bedarf rezertifizieren; nicht mehr benötigte Freigaben entfernen.'));
  if (accessDenied.length) findings.push(finding('error', `${accessDenied.length} Freigaben: Zugriff auf Berechtigungen verweigert`, 'Microsoft Graph hat diese Berechtigungslisten mit HTTP 401 oder 403 abgelehnt.', 'Admin-Consent für Files.Read.All beziehungsweise Sites.Read.All und den Zugriff des angemeldeten Kontos auf die Speicherorte prüfen.'));
  if (throttled.length) findings.push(finding('error', `${throttled.length} Berechtigungsabfragen nach Wiederholung gedrosselt`, 'Microsoft Graph hat die Detailabfragen auch nach automatischer Wartezeit nicht verarbeitet.', 'Diesen Bereich später erneut prüfen; bereits gelesene Ergebnisse bleiben auswertbar.'));
  if (otherUnreadable.length) findings.push(finding('error', `${otherUnreadable.length} Freigaben mit technisch nicht lesbaren Berechtigungen`, `Microsoft Graph konnte die Berechtigungslisten nicht liefern${permissionFailures.size ? ` (${[...new Set(otherUnreadable.map((entry) => entry.permissionFailure?.code).filter(Boolean))].slice(0, 3).join(', ') || 'unbekannter Fehler'})` : ''}.`, 'Prüfstatus in der Detailtabelle kontrollieren und den Bereich erneut prüfen.'));
  if (unreadableDrives) findings.push(finding('error', `${unreadableDrives} Speicherorte nicht vollständig lesbar`, 'Mindestens eine Dokumentbibliothek oder ein OneDrive konnte nicht vollständig inventarisiert werden.', 'Sites.Read.All, Files.Read.All und den SharePoint-Zugriff des angemeldeten Kontos prüfen.'));
  if (!findings.length) findings.push(finding('ok', 'Keine riskante Dateifreigabe erkannt', entries.length ? 'Die gefundenen Freigaben sind jünger als 180 Tage, nicht anonym und ohne erkanntes Schreibrecht.' : 'In den auswertbaren Speicherorten wurden keine geteilten Dateien oder Ordner gefunden.', 'Freigabebestand regelmäßig erneut prüfen.'));

  return {
    records: entries.length,
    unavailable: unreadable.length + unreadableDrives,
    summary: `${entries.length} geteilte Dateien und Ordner in ${drives.length} Speicherorten geprüft`,
    metrics: [[String(entries.length), 'Freigaben'], [String(anonymous.length), 'anonym'], [String(writable.length), 'mit Schreibrecht'], [String(old.length), 'mindestens 180 Tage alt']],
    findings,
    details: {
      title: 'Geteilte Ordner und Dateien',
      ageColumn: 6,
      columns: ['Objekt', 'Typ', 'Speicherort', 'Eigentümer', 'Geteilt von', 'Geteilt seit', 'Alter (Tage)', 'Freigabeart', 'Prüfstatus', 'Rechte', 'Empfänger', 'Ablauf', 'Öffnen'],
      rows: entries.map(({ drive, item, permissions, permissionFailure, ageDays, scopes, roles, recipients, expiration }) => [
        item.name || item.id,
        item.folder ? 'Ordner' : item.file ? 'Datei' : 'Objekt',
        drive.sourceName || drive.name || drive.id,
        sharedIdentity(drive.owner) || '–',
        sharedIdentity(item.shared.sharedBy) || '–',
        date(item.shared.sharedDateTime),
        ageDays == null ? '–' : String(ageDays),
        scopes.map((scope) => scopeLabel[scope] || scope).join(', ') || (permissions == null ? 'Nicht lesbar' : 'Direkte Berechtigung'),
        permissions == null ? `${permissionFailure?.code || 'Nicht lesbar'}${permissionFailure?.status ? ` · HTTP ${permissionFailure.status}` : ''}` : 'Gelesen',
        roles.join(', ') || '–',
        recipients.join(', ') || '–',
        date(expiration),
        item.webUrl || '–',
      ]),
    },
  };
}

export function analyseDevices(devices = [], now = new Date()) {
  const cutoff = new Date(now.getTime() - 30 * day);
  const stale = devices.filter((device) => device.approximateLastSignInDateTime && new Date(device.approximateLastSignInDateTime) < cutoff);
  const disabled = devices.filter((device) => !device.accountEnabled);
  const managed = devices.filter((device) => device.isManaged);
  const findings = [];

  if (stale.length) findings.push(finding('medium', `${stale.length} Entra-Geräte seit 30 Tagen ohne Anmeldung`, 'Der ungefähre letzte Anmeldezeitpunkt deutet auf inaktive oder stillgelegte Geräteobjekte hin.', 'Gerätebestand prüfen und veraltete Objekte kontrolliert bereinigen.'));
  if (disabled.length) findings.push(finding('info', `${disabled.length} deaktivierte Entra-Geräte`, 'Deaktivierte Geräteobjekte verbleiben im Verzeichnis und sollten einem dokumentierten Lifecycle folgen.', 'Bedarf und Aufbewahrung prüfen; nicht mehr benötigte Objekte kontrolliert entfernen.'));
  if (!findings.length) findings.push(finding('ok', 'Keine Auffälligkeit im Entra-Gerätebestand', 'Alle gefundenen Geräte sind aktiv und haben innerhalb der letzten 30 Tage eine Anmeldung gemeldet.', 'Gerätebestand regelmäßig erneut prüfen.'));

  return {
    records: devices.length,
    summary: `${devices.length} Entra-Geräte geprüft`,
    metrics: [[String(devices.length), 'Entra-Geräte'], [String(managed.length), 'als verwaltet markiert'], [String(stale.length), '> 30 Tage inaktiv'], [String(disabled.length), 'deaktiviert']],
    findings,
    details: {
      title: 'Entra-Geräte',
      columns: ['Gerät', 'Betriebssystem', 'Trust-Typ', 'Aktiv', 'Verwaltet', 'Konform', 'Letzte Anmeldung'],
      rows: devices.map((device) => [device.displayName || device.deviceId || device.id, `${device.operatingSystem || '–'} ${device.operatingSystemVersion || ''}`.trim(), device.trustType || '–', device.accountEnabled ? 'Ja' : 'Nein', device.isManaged ? 'Ja' : 'Nein', device.isCompliant ? 'Ja' : 'Nein', date(device.approximateLastSignInDateTime)]),
    },
  };
}

export function analyseApplications(applications, now = new Date(), servicePrincipals = [], grants = [], appRoles = [], owners = new Map()) {
  const warningDate = new Date(now.getTime() + 90 * day);
  const credentials = applications.flatMap((app) => [
    ...(app.passwordCredentials || []).map((credential) => ({ app, credential, type: 'Secret', end: credential.endDateTime && new Date(credential.endDateTime) })),
    ...(app.keyCredentials || []).map((credential) => ({ app, credential, type: 'Zertifikat', end: credential.endDateTime && new Date(credential.endDateTime) })),
  ]);
  const expiredApps = new Set(credentials.filter(({ end }) => end && end < now).map(({ app }) => app.id));
  const expiringApps = new Set(credentials.filter(({ app, end }) => end && end >= now && end <= warningDate && !expiredApps.has(app.id)).map(({ app }) => app.id));
  const servicePrincipalById = new Map(servicePrincipals.map((item) => [item.id, item]));
  const delegated = grants.flatMap((grant) => String(grant.scope || '').split(/\s+/).filter(Boolean).map((permission) => ({
    client: servicePrincipalById.get(grant.clientId)?.displayName || grant.clientId,
    resource: servicePrincipalById.get(grant.resourceId)?.displayName || grant.resourceId,
    permission,
    consentType: grant.consentType,
  })));
  const dangerous = /(^|\.)(readwrite|fullcontrol|manage|send)(\.|$)|directory\.read\.all|rolemanagement/i;
  const broadDelegated = delegated.filter((grant) => dangerous.test(grant.permission));
  const broadAppRoles = appRoles.filter((grant) => dangerous.test(grant.permission));
  const disabledServicePrincipals = servicePrincipals.filter((item) => item.accountEnabled === false);
  const ownerless = applications.filter((app) => owners.has(`application:${app.id}`) && owners.get(`application:${app.id}`).length === 0);
  const findings = [];

  if (expiredApps.size) findings.push(finding('high', `${expiredApps.size} App-Registrierungen mit abgelaufenen Credentials`, 'Mindestens ein hinterlegtes Secret oder Zertifikat ist bereits abgelaufen.', 'Nutzung prüfen, Credentials sicher rotieren oder ungenutzte App-Registrierungen entfernen.'));
  if (broadAppRoles.length) findings.push(finding('high', `${broadAppRoles.length} weitreichende Application-Permission-Zuweisungen`, 'Application Permissions wirken ohne angemeldeten Benutzer und können tenantweit auf Daten zugreifen.', 'Geschäftszweck, Owner, letzte Nutzung und Least-Privilege-Alternativen prüfen.'));
  if (broadDelegated.length) findings.push(finding('medium', `${broadDelegated.length} weitreichende delegierte OAuth-Scopes`, 'Mindestens ein Consent enthält Schreib-, Verwaltungs- oder besonders breite Leserechte.', 'Admin-Consents und weiterhin benötigte Scopes regelmäßig rezertifizieren.'));
  if (expiringApps.size) findings.push(finding('medium', `${expiringApps.size} App-Registrierungen laufen binnen 90 Tagen aus`, 'Mindestens ein Secret oder Zertifikat erreicht zeitnah sein Ablaufdatum.', 'Owner informieren und Rotation vor Ablauf terminieren.'));
  if (ownerless.length) findings.push(finding('medium', `${ownerless.length} App-Registrierungen ohne Owner`, 'Für diese App-Registrierungen wurde keine verantwortliche Identität gefunden.', 'Mindestens zwei fachlich und technisch verantwortliche Owner zuweisen.'));
  if (disabledServicePrincipals.length) findings.push(finding('info', `${disabledServicePrincipals.length} deaktivierte Enterprise Apps`, 'Diese Service Principals sind für die Anmeldung deaktiviert, verbleiben aber im Mandanten.', 'Abhängigkeiten und Aufbewahrungsbedarf prüfen; ungenutzte Objekte kontrolliert entfernen.'));
  if (!findings.length) findings.push(finding('ok', 'Keine zeitkritischen App-Credentials erkannt', 'Für die gefundenen App-Registrierungen laufen innerhalb von 90 Tagen keine Credentials ab.', 'Owner und Credential-Rotation weiterhin regelmäßig kontrollieren.'));

  return {
    records: applications.length + servicePrincipals.length,
    summary: `${applications.length} App-Registrierungen, ${servicePrincipals.length} Enterprise Apps und ihre Grants geprüft`,
    metrics: [[String(applications.length), 'App-Registrierungen'], [String(servicePrincipals.length), 'Enterprise Apps'], [String(expiredApps.size), 'Credentials abgelaufen'], [String(broadAppRoles.length + broadDelegated.length), 'weitreichende Grants']],
    findings,
    details: {
      title: 'Zeitkritische App-Credentials',
      columns: ['App-Registrierung', 'Credential', 'Typ', 'Ablauf', 'Status'],
      rows: credentials
        .filter(({ end }) => end && end <= warningDate)
        .map(({ app, credential, type, end }) => [app.displayName || app.id, credential.displayName || credential.keyId || '–', type, date(end), end < now ? 'Abgelaufen' : 'Läuft ≤ 90 Tage ab']),
    },
    extraDetails: [
      {
        title: 'App-Registrierungen',
        columns: ['App', 'Client-ID', 'Owner', 'Zielgruppe', 'Redirects', 'Credentials'],
        rows: applications.map((item) => {
          const appOwners = owners.get(`application:${item.id}`);
          return [item.displayName || item.id, item.appId || '–', appOwners == null ? 'Nicht lesbar' : appOwners.map((owner) => owner.displayName || owner.userPrincipalName || owner.id).join(', ') || 'Keine', item.signInAudience || '–', String((item.web?.redirectUris?.length || 0) + (item.spa?.redirectUris?.length || 0)), String((item.passwordCredentials?.length || 0) + (item.keyCredentials?.length || 0))];
        }),
      },
      {
        title: 'Delegierte OAuth-Consents',
        columns: ['Client', 'Ressource', 'Permission', 'Consent'],
        rows: delegated.map((grant) => [grant.client, grant.resource, grant.permission, grant.consentType || '–']),
      },
      {
        title: 'Application Permissions',
        columns: ['Client', 'Ressource', 'Permission'],
        rows: appRoles.map((grant) => [grant.client, grant.resource, grant.permission]),
      },
      {
        title: 'Enterprise Apps',
        columns: ['App', 'Typ', 'Aktiv', 'Benutzerzuweisung erforderlich', 'Owner'],
        rows: servicePrincipals.map((item) => {
          const appOwners = owners.get(`servicePrincipal:${item.id}`);
          return [item.displayName || item.appId || item.id, item.servicePrincipalType || '–', item.accountEnabled === false ? 'Nein' : 'Ja', item.appRoleAssignmentRequired ? 'Ja' : 'Nein', appOwners == null ? 'Nicht lesbar' : appOwners.map((owner) => owner.displayName || owner.userPrincipalName || owner.id).join(', ') || 'Keine'];
        }),
      },
    ],
  };
}

export function analyseSecurity(scores = []) {
  const latestScore = [...scores].sort((a, b) => String(b.createdDateTime || '').localeCompare(String(a.createdDateTime || '')))[0];
  const scorePercent = latestScore?.maxScore ? percent(latestScore.currentScore, latestScore.maxScore) : null;
  const findings = [];

  if (scorePercent != null && scorePercent < 60) findings.push(finding('medium', `Secure Score bei ${scorePercent} %`, 'Der aktuelle Secure Score liegt unter 60 Prozent des erreichbaren Maximalwerts.', 'Verbesserungsmaßnahmen nach Wirkung, Aufwand und vorhandenen Lizenzen priorisieren.'));
  if (scorePercent == null) findings.push(finding('info', 'Kein Secure Score gemeldet', 'Microsoft Graph liefert aktuell keinen auswertbaren Sicherheitswert.', 'Secure-Score-Berechnung und Lizenzierung im Microsoft Defender Portal prüfen.'));
  if (!findings.length) findings.push(finding('ok', `Secure Score bei ${scorePercent} %`, 'Der aktuelle Sicherheitswert liegt bei mindestens 60 Prozent.', 'Verbesserungsmaßnahmen weiterhin nach Wirkung und Aufwand priorisieren.'));

  return {
    records: scores.length,
    summary: 'Microsoft Secure Score geprüft',
    metrics: [[scorePercent == null ? '–' : `${scorePercent} %`, 'Secure Score'], [String(latestScore?.currentScore ?? '–'), 'Punkte'], [String(latestScore?.maxScore ?? '–'), 'Maximum']],
    findings,
    details: {
      title: 'Secure-Score-Messung',
      columns: ['Zeitpunkt', 'Punkte', 'Maximum', 'Prozent'],
      rows: latestScore ? [[date(latestScore.createdDateTime), String(latestScore.currentScore ?? '–'), String(latestScore.maxScore ?? '–'), scorePercent == null ? '–' : `${scorePercent} %`]] : [],
    },
  };
}

export function analyseCompliance(labels = [], cases = [], labelsAvailable = true, casesAvailable = true) {
  const activeLabels = labels.filter((item) => !['inactive', 'deleted'].includes(normalized(item.state || item.status)));
  const openCases = cases.filter((item) => !['closed', 'resolved'].includes(normalized(item.status)));
  const findings = [];

  if (labelsAvailable && !labels.length) findings.push(finding('medium', 'Keine Records-Management-Labels gefunden', 'Im auswertbaren Bestand sind keine Aufbewahrungslabels vorhanden.', 'Aufbewahrungsanforderungen klären und passende Labels sowie Policies definieren.'));
  if (casesAvailable && openCases.length) findings.push(finding('info', `${openCases.length} offene eDiscovery-Fälle`, 'Offene Fälle können laufende rechtliche oder interne Untersuchungen abbilden.', 'Owner, Zugriffsberechtigungen, Hold-Umfang und Abschlussstatus kontrollieren.'));
  if (labelsAvailable && casesAvailable && !findings.length) findings.push(finding('ok', 'Compliance-Grundbestand vorhanden', 'Aufbewahrungslabels sind vorhanden und es wurden keine offenen eDiscovery-Fälle erkannt.', 'Policies, Label-Veröffentlichung und tatsächliche Anwendung separat prüfen.'));

  return {
    records: labels.length + cases.length,
    summary: `${labels.length} Aufbewahrungslabels und ${cases.length} eDiscovery-Fälle geprüft`,
    metrics: [[labelsAvailable ? String(labels.length) : '–', 'Labels'], [labelsAvailable ? String(activeLabels.length) : '–', 'Labels aktiv'], [casesAvailable ? String(cases.length) : '–', 'eDiscovery-Fälle'], [casesAvailable ? String(openCases.length) : '–', 'Fälle offen']],
    findings,
    details: {
      title: 'Records-Management-Labels',
      columns: ['Label', 'Verhalten', 'Aktion', 'Dauer', 'In Verwendung'],
      rows: labels.map((item) => [item.displayName || item.name || item.id, item.behaviorDuringRetentionPeriod || '–', item.actionAfterRetentionPeriod || '–', item.retentionDuration?.displayName || item.retentionDuration?.days || '–', item.isInUse ? 'Ja' : 'Nein']),
    },
    extraDetails: [{ title: 'eDiscovery-Fälle', columns: ['Fall', 'Status', 'Erstellt', 'Geschlossen', 'Beschreibung'], rows: cases.map((item) => [item.displayName || item.id, item.status || '–', date(item.createdDateTime), date(item.closedDateTime), item.description || '–']) }],
  };
}

export function analyseServiceHealth(health = [], issues = [], messages = []) {
  const unhealthy = health.filter((item) => !['serviceoperational', 'restoringservice', 'servicerestored'].includes(normalized(item.status)));
  const activeIssues = issues.filter((item) => item.isResolved !== true && !['resolved', 'servicerestored'].includes(normalized(item.status)));
  const actionable = messages.filter((item) => item.actionRequiredByDateTime || normalized(item.category) === 'preventorfixissue');
  const interruptions = activeIssues.filter((item) => normalized(item.classification) === 'incident' || normalized(item.featureGroup) === 'service interruption');
  const findings = [];

  if (interruptions.length) findings.push(finding('high', `${interruptions.length} aktive Service-Unterbrechungen`, 'Microsoft 365 meldet aktive Incidents oder Service-Unterbrechungen.', 'Auswirkung, Workaround und Statuskommunikation unmittelbar prüfen.'));
  if (activeIssues.length - interruptions.length || unhealthy.length) findings.push(finding('medium', `${activeIssues.length} aktive Service-Issues`, 'Mindestens ein Dienst oder Health-Issue ist noch nicht vollständig wiederhergestellt.', 'Betroffene Dienste und Nutzergruppen identifizieren und Updates verfolgen.'));
  if (actionable.length) findings.push(finding('info', `${actionable.length} Service-Nachrichten mit möglichem Handlungsbedarf`, 'Microsoft hat Änderungen oder Maßnahmen mit möglicher administrativer Relevanz veröffentlicht.', 'Fristen, betroffene Dienste und interne Zuständigkeit prüfen.'));
  if (!findings.length) findings.push(finding('ok', 'Microsoft-365-Dienste ohne erkannte Störung', 'Alle auswertbaren Dienste sind operational und es wurden keine aktiven Issues gefunden.', 'Service Health und Message Center weiterhin beobachten.'));

  return {
    records: health.length + issues.length + messages.length,
    summary: `${health.length} Dienste, ${issues.length} Health-Issues und ${messages.length} Nachrichten geprüft`,
    metrics: [[String(health.length), 'Dienste'], [String(unhealthy.length), 'beeinträchtigt'], [String(activeIssues.length), 'Issues aktiv'], [String(actionable.length), 'Nachrichten relevant']],
    findings,
    details: {
      title: 'Service Health',
      columns: ['Dienst', 'Status', 'Details'],
      rows: health.map((item) => [item.service || item.serviceName || item.id, item.status || '–', item.statusDetails?.map?.((detail) => detail.status)?.join(', ') || '–']),
    },
    extraDetails: [
      { title: 'Aktive Health-Issues', columns: ['Titel', 'Dienst', 'Status', 'Klassifikation', 'Beginn', 'Ende'], rows: activeIssues.map((item) => [item.title || item.id, item.service || item.serviceName || '–', item.status || '–', item.classification || '–', date(item.startDateTime), date(item.endDateTime)]) },
      { title: 'Service-Nachrichten', columns: ['Titel', 'Kategorie', 'Dienste', 'Veröffentlicht', 'Aktion bis'], rows: messages.map((item) => [item.title || item.id, item.category || '–', (item.services || []).join(', ') || '–', date(item.startDateTime || item.lastModifiedDateTime), date(item.actionRequiredByDateTime)]) },
    ],
  };
}
