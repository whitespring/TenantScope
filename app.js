import { BrowserCacheLocation, InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser';
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
import { createDocx } from './docx.mjs';

if (window.location.hostname === '127.0.0.1') {
  const localhost = new URL(window.location.href);
  localhost.hostname = 'localhost';
  window.location.replace(localhost);
}

const SESSION_KEY = 'tenant-scope-poc';
const USAGE_NOTICE_KEY = 'tenant-scope-usage-notice-v1';
const AUTH_MODE_KEY = 'tenant-scope-auth-mode';
const ADMIN_CONSENT_STATE_KEY = 'tenant-scope-admin-consent-state';
const scopes = [
  {
    id: 'tenant', name: 'Tenant & Domains', tag: 'Mandant',
    description: 'Organisation, Domains, Synchronisierung und Basisdaten',
    permissions: ['Domain.Read.All'],
  },
  {
    id: 'identity', name: 'Benutzer & Gäste', tag: 'Entra ID',
    description: 'Konten, Status, Typ und Synchronisierung',
    permissions: ['User.Read.All'],
  },
  {
    id: 'roles', name: 'Administrative Rollen', tag: 'Entra RBAC',
    description: 'Aktive Rollenzuweisungen und ihre Inhaber',
    permissions: ['RoleManagement.Read.Directory'],
  },
  {
    id: 'access', name: 'Conditional Access', tag: 'Zero Trust',
    description: 'Status und Bestand der Zugriffsrichtlinien',
    permissions: ['Policy.Read.All'],
  },
  {
    id: 'licenses', name: 'Lizenzoptimierung & Kosten', tag: 'FinOps',
    description: 'Ungenutzte Zuweisungen, Downgrade-Kandidaten und Einsparpotenziale',
    permissions: ['LicenseAssignment.Read.All', 'User.Read.All', 'Reports.Read.All', 'ReportSettings.Read.All'],
  },
  {
    id: 'usage', name: 'Nutzung & Adoption', tag: 'M365 Reports',
    description: 'Apps, Plattformen, Dienste, Teams-Aktivität und Copilot',
    permissions: ['Reports.Read.All', 'ReportSettings.Read.All', 'User.Read.All'],
  },
  {
    id: 'storage', name: 'Postfächer & OneDrive', tag: 'M365 Reports',
    description: 'Speicher je Benutzer und Summen für den Tenant',
    permissions: ['Reports.Read.All', 'User.Read.All'],
  },
  {
    id: 'teams', name: 'Teams & Gruppen', tag: 'Collaboration',
    description: 'Team- und Gruppenbestand, Owner und Sichtbarkeit',
    permissions: ['Group.Read.All'],
  },
  {
    id: 'sites', name: 'SharePoint Sites', tag: 'SharePoint',
    description: 'Sites, Speicher, Aktivität, Root-Freigaben und Tenant-Regeln',
    permissions: ['Sites.Read.All', 'Reports.Read.All', 'SharePointTenantSettings.Read.All'],
  },
  {
    id: 'sharing', name: 'Öffentliche Freigaben', tag: 'SharePoint & OneDrive',
    description: 'Anonyme Links, Schreibrechte, Ablauf und schwache Absicherung',
    permissions: ['Sites.Read.All', 'Files.Read.All', 'User.Read.All'],
  },
  {
    id: 'devices', name: 'Entra-Geräte', tag: 'Entra ID',
    description: 'Gerätebestand, Status und letzte Anmeldung',
    permissions: ['Device.Read.All'],
  },
  {
    id: 'apps', name: 'Apps, OAuth & Credentials', tag: 'App Registry',
    description: 'Registrierungen, Enterprise Apps, Grants und App-Berechtigungen',
    permissions: ['Application.Read.All', 'Directory.Read.All'],
  },
  {
    id: 'security', name: 'Microsoft Secure Score', tag: 'Security',
    description: 'Aktueller Sicherheitswert des Tenants',
    permissions: ['SecurityEvents.Read.All'],
  },
  {
    id: 'compliance', name: 'Purview & Compliance', tag: 'Purview',
    description: 'Aufbewahrungslabels und eDiscovery-Fälle',
    permissions: ['RecordsManagement.Read.All', 'eDiscovery.Read.All'],
  },
  {
    id: 'service', name: 'Service Health', tag: 'Betrieb',
    description: 'Dienststatus, aktive Störungen und Message Center',
    permissions: ['ServiceHealth.Read.All', 'ServiceMessage.Read.All'],
  },
];

const scopeGuidance = {
  tenant: {
    explanation: 'Die Tenant- und Domainbasis bestimmt, welche Anmeldenamen, Standarddomains und Synchronisierungswege im Mandanten gelten. Unklare oder unverifizierte Domains können Anmeldung, Zustellung und Wiederherstellung beeinträchtigen.',
    goodPractice: 'Alle produktiv verwendeten Domains verifizieren, Verantwortliche und technische Kontakte dokumentieren und externe Föderationsabhängigkeiten überwachen. Mindestens zwei cloud-only Notfallkonten mit phishingresistenter Anmeldung regelmäßig testen.',
    helpUrl: 'https://learn.microsoft.com/en-us/microsoft-365/admin/setup/add-domain?view=o365-worldwide',
    adminLinks: [['Domains verwalten', 'https://admin.microsoft.com/Adminportal/Home#/Domains']],
  },
  identity: {
    explanation: 'Der Benutzerbestand macht aktive, deaktivierte, synchronisierte und externe Identitäten sichtbar. Er ist die Grundlage für sauberes Onboarding, Offboarding und die weitere Lizenzprüfung.',
    goodPractice: 'Joiner-, Mover- und Leaver-Prozesse mit klaren Fristen betreiben; deaktivierte Konten, externe Gäste und nicht mehr benötigte Zugriffe regelmäßig rezertifizieren. Administrative Arbeit mit getrennten Konten und Least Privilege ausführen.',
    helpUrl: 'https://learn.microsoft.com/en-us/entra/fundamentals/how-to-manage-user-profile-info',
    adminLinks: [['Benutzer verwalten', 'https://entra.microsoft.com/#view/Microsoft_AAD_UsersAndTenants/UserManagementMenuBlade/~/AllUsers']],
  },
  roles: {
    explanation: 'Dauerhafte administrative Rollenzuweisungen besitzen erhöhte Rechte im Tenant. Entscheidend sind wenige, nachvollziehbare Inhaber, getrennte Notfallkonten und regelmäßige Rezertifizierung.',
    goodPractice: 'Dauerhafte Global-Admin-Zuweisungen minimieren und reguläre Administration über kleinere Rollen sowie PIM/JIT abbilden. Zwei kontrollierte Notfallkonten vorhalten und privilegierte Zuweisungen mindestens quartalsweise prüfen.',
    helpUrl: 'https://learn.microsoft.com/en-us/entra/identity/role-based-access-control/permissions-reference',
    adminLinks: [['Rollen verwalten', 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AllRolesBlade']],
  },
  access: {
    explanation: 'Conditional Access steuert Zugriffe anhand von Identität, Gerät, Standort und Risiko. Diese Inventur bewertet Bestand und Aktivierungsstatus, aber keine vollständige Richtlinienabdeckung oder What-if-Wirkung.',
    goodPractice: 'Basisrichtlinien zunächst im Report-only-Modus testen und gestuft aktivieren: MFA für Administratoren und Benutzer, Blockade von Legacy Authentication und risikobasierte Kontrollen. Nur dokumentierte Notfallkonten ausschließen und deren Nutzung überwachen.',
    helpUrl: 'https://learn.microsoft.com/en-us/entra/identity/conditional-access/overview',
    adminLinks: [['Richtlinien verwalten', 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/Policies']],
  },
  licenses: {
    explanation: 'Die Lizenzberatung verbindet abonnierte SKUs, einzelne Zuweisungen, deren Quelle und 90-Tage-Nutzungsberichte. Deaktivierte Konten sind starke Rückgabe-Kandidaten; Inaktivität und mögliche Downgrades bleiben Prüfhypothesen, weil nicht jede lizenzierte Security-, Compliance-, Telefonie- oder Gerätefunktion als Benutzernutzung messbar ist.',
    goodPractice: 'Deaktivierte Konten und 90-Tage-Inaktivität zuerst prüfen, gruppenbasierte Zuweisungen an der Quelle korrigieren und eine kleine dokumentierte Reserve festlegen. Downgrades nur nach Abgleich von Fachbedarf, Desktop-Apps, Security, Compliance, Intune, Telefonie, Aufbewahrung und Vertragsfristen umsetzen. Euro-Werte auf tatsächlichen Netto-Vertragspreisen statt Listenpreisen basieren.',
    helpUrl: 'https://learn.microsoft.com/en-us/microsoft-365/admin/activity-reports/activity-reports?view=o365-worldwide',
    adminLinks: [['Lizenzen verwalten', 'https://admin.microsoft.com/Adminportal/Home#/licenses'], ['Produkte & Verträge prüfen', 'https://admin.microsoft.com/Adminportal/Home#/subscriptions']],
  },
  usage: {
    explanation: 'Die Nutzungsdaten stammen aus Microsoft-365-Berichten und zeigen Aktivität, nicht Produktivität oder Qualität. Berichte können 24 bis 72 Stunden verzögert und durch die Datenschutzoption anonymisiert sein.',
    goodPractice: '90-Tage-Trends für Enablement und Lizenzentscheidungen verwenden, nicht zur individuellen Leistungskontrolle. Rollen, Saisonverläufe sowie technische Konten berücksichtigen und Zweck, Zugriff und Datenschutz transparent dokumentieren.',
    helpUrl: 'https://learn.microsoft.com/en-us/microsoft-365/admin/activity-reports/activity-reports?view=o365-worldwide',
    adminLinks: [['Nutzungsberichte öffnen', 'https://admin.microsoft.com/Adminportal/Home#/reportsUsage'], ['Reporteinstellungen öffnen', 'https://admin.microsoft.com/Adminportal/Home#/Settings/Services/:/Settings/L1/Reports']],
  },
  storage: {
    explanation: 'Postfach- und OneDrive-Berichte zeigen belegten Speicher und gemeldete Kapazitäten je Benutzer sowie als Tenant-Summe. Die Werte sind Berichtssnapshots und können gegenüber den Admin-Portalen verzögert sein.',
    goodPractice: 'Warnschwellen bei 80 und 90 Prozent etablieren, Wachstum beobachten und Aufbewahrung von bloßer Speicherbereinigung trennen. Verwaiste Daten kontrolliert archivieren oder löschen und Owner frühzeitig einbeziehen.',
    helpUrl: 'https://learn.microsoft.com/en-us/microsoft-365/admin/activity-reports/mailbox-usage?view=o365-worldwide',
    adminLinks: [['Postfächer verwalten', 'https://admin.exchange.microsoft.com/#/mailboxes'], ['OneDrive und Sites verwalten', 'https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true']],
  },
  teams: {
    explanation: 'Die Prüfung betrachtet Teams-fähige Microsoft-365-Gruppen, ihre Sichtbarkeit und lesbare Owner. Inhalte, Kanäle und einzelne Mitgliedschaften gehören nicht zu diesem derzeit freigegebenen Prüfpfad.',
    goodPractice: 'Private Teams als Standard, mindestens zwei Owner, einen dokumentierten Zweck und regelmäßige Verlängerung vorsehen. Gastzugriffe, öffentliche Teams und verwaiste Gruppen rezertifizieren; sensible Zusammenarbeit mit Labels und restriktivem Sharing absichern.',
    helpUrl: 'https://learn.microsoft.com/en-us/microsoftteams/manage-teams-in-modern-portal',
    adminLinks: [['Teams verwalten', 'https://admin.teams.microsoft.com/teams/manage']],
  },
  sites: {
    explanation: 'Die SharePoint-Auswertung verbindet auffindbare Sites mit Nutzungs- und Speicherberichten und prüft Root-Freigaben. Sie ist keine vollständige Berechtigungsanalyse aller Dateien, Ordner und Unterwebsites.',
    goodPractice: 'Für jede Site Zweck, Datenklassifikation und mindestens zwei Owner dokumentieren. Externes Sharing standardmäßig restriktiv setzen, inaktive Sites über einen Lifecycle behandeln und Berechtigungsvererbung regelmäßig prüfen.',
    helpUrl: 'https://learn.microsoft.com/en-us/sharepoint/manage-sites-in-new-admin-center',
    adminLinks: [['Aktive Sites verwalten', 'https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true']],
  },
  sharing: {
    explanation: 'Die Prüfung sucht gezielt nach „Jeder mit Link“-Freigaben in erreichbaren SharePoint- und OneDrive-Speicherorten. Microsoft Graph liefert für Geschäftskonten keinen belastbaren Passwortindikator; ein anonymer Link gilt deshalb als Zugriff ohne Anmeldung. Interne und personengebundene Freigaben werden nicht einzeln abgefragt.',
    goodPractice: 'Standardmäßig „Bestimmte Personen“ verwenden. Öffentliche Links auf begründete Ausnahmen beschränken, nur lesend und mit kurzem Ablaufdatum vergeben sowie regelmäßig durch Owner rezertifizieren. Öffentliche Schreiblinks oder Links ohne Ablauf zeitnah ersetzen.',
    helpUrl: 'https://learn.microsoft.com/en-us/graph/api/driveitem-list-permissions?view=graph-rest-1.0',
    adminLinks: [['SharePoint-Sites verwalten', 'https://admin.microsoft.com/sharepoint?page=siteManagement&modern=true'], ['OneDrive-Freigabeeinstellungen', 'https://admin.microsoft.com/sharepoint?page=sharing&modern=true']],
  },
  devices: {
    explanation: 'Entra-Geräteobjekte zeigen Registrierung, Aktivierungsstatus und den ungefähren letzten Login. Das ist kein Ersatz für eine Intune-Inventur und beweist weder aktuelle Verwaltung noch technische Compliance.',
    goodPractice: 'Zugriff auf sensible Dienste über Conditional Access an registrierte, verwaltete und konforme Geräte binden. Inaktive Geräte erst nach Abgleich mit Intune und Owner kontrolliert deaktivieren oder löschen.',
    helpUrl: 'https://learn.microsoft.com/en-us/entra/identity/devices/manage-device-identities',
    adminLinks: [['Entra-Geräte verwalten', 'https://entra.microsoft.com/#view/Microsoft_AAD_Devices/DevicesMenuBlade/~/Devices']],
  },
  apps: {
    explanation: 'App-Registrierungen, Enterprise Apps, Credentials und Consents bilden nicht-menschliche Zugriffswege ab. Abgelaufene Schlüssel, fehlende Owner und breite Rechte erhöhen Betriebs- und Sicherheitsrisiken.',
    goodPractice: 'Mindestens zwei verantwortliche Owner benennen, kurzlebige Zertifikate oder Managed Identities gegenüber Secrets bevorzugen und Rotation überwachen. Admin-Consents und Application Permissions nach Least Privilege regelmäßig rezertifizieren; ungenutzte Apps entfernen.',
    helpUrl: 'https://learn.microsoft.com/en-us/entra/identity-platform/howto-remove-app',
    adminLinks: [['App-Registrierungen verwalten', 'https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade'], ['Enterprise Apps und Consents', 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/StartboardApplicationsMenuBlade/~/AppAppsPreview']],
  },
  security: {
    explanation: 'Microsoft Secure Score verdichtet umgesetzte Sicherheitsmaßnahmen zu einem Vergleichswert. Er priorisiert Verbesserungen, ersetzt aber weder eine Risikobewertung noch die Prüfung einzelner Empfehlungen.',
    goodPractice: 'Secure-Score-Maßnahmen nach realem Risiko, Wirkung, Aufwand und Lizenzlage priorisieren statt nur den Prozentwert zu maximieren. Für jede Maßnahme Owner, Zieltermin und dokumentierte Ausnahme festlegen.',
    helpUrl: 'https://learn.microsoft.com/en-us/defender-xdr/microsoft-secure-score',
    adminLinks: [['Secure Score öffnen', 'https://security.microsoft.com/securescore']],
  },
  compliance: {
    explanation: 'Aufbewahrungslabels und eDiscovery-Fälle zeigen einen Teil der Purview-Governance. Aussagekraft und Sichtbarkeit hängen zusätzlich von Purview-Lizenzen, Rollen und veröffentlichten Richtlinien ab.',
    goodPractice: 'Aufbewahrung aus rechtlichen und fachlichen Anforderungen ableiten, Labels über Policies veröffentlichen und deren tatsächliche Anwendung testen. eDiscovery-Zugriffe eng begrenzen, Holds dokumentieren und abgeschlossene Fälle schließen.',
    helpUrl: 'https://learn.microsoft.com/en-us/purview/retention',
    adminLinks: [['Data Lifecycle Management öffnen', 'https://purview.microsoft.com/datalifecyclemanagement'], ['eDiscovery öffnen', 'https://purview.microsoft.com/ediscovery']],
  },
  service: {
    explanation: 'Service Health und Message Center zeigen den aktuellen Betriebszustand und angekündigte Änderungen. Diese Informationen ändern sich laufend und stehen deshalb bewusst am Ende des Berichts.',
    goodPractice: 'Für aktive Incidents klare Owner, betroffene Dienste, Nutzerwirkung, Workaround und Kommunikationsrhythmus festlegen. Message-Center-Änderungen mit Frist einer verantwortlichen Person zuweisen und die Umsetzung nachverfolgen.',
    helpUrl: 'https://learn.microsoft.com/en-us/microsoft-365/enterprise/view-service-health?view=o365-worldwide',
    adminLinks: [['Service Health öffnen', 'https://admin.microsoft.com/Adminportal/Home#/servicehealth'], ['Message Center öffnen', 'https://admin.microsoft.com/Adminportal/Home#/MessageCenter']],
  },
};

const form = document.querySelector('#tenant-form');
const scopeGrid = document.querySelector('#scope-grid');
const setupSection = document.querySelector('#setup');
const scopeSection = document.querySelector('#scope');
const continueButton = document.querySelector('#continue-to-scope');
const startButton = document.querySelector('#start-inventory');
const reportSection = document.querySelector('#report');
const runPanel = document.querySelector('#run-panel');
const sideNav = document.querySelector('#side-nav');
const reportNav = document.querySelector('#report-nav');
const reportNavSelect = document.querySelector('#report-nav-select');
const reportOverviewNav = document.querySelector('#report-overview-nav');
const authMessage = document.querySelector('#auth-message');
const toast = document.querySelector('#toast');
const consentButton = document.querySelector('#grant-consent');
const reportNamesButton = document.querySelector('#reveal-report-names');
const exportDialog = document.querySelector('#export-dialog');
const exportForm = document.querySelector('#export-form');
const confirmExportButton = document.querySelector('#confirm-export');
const usageNoticeDialog = document.querySelector('#usage-notice-dialog');
const usageNoticeForm = document.querySelector('#usage-notice-form');
const usageNoticeCheckbox = document.querySelector('#accept-usage-notice');
const confirmUsageNoticeButton = document.querySelector('#confirm-usage-notice');
let currentReport;
let msalClient;
let msalConfigKey;
let authenticationPromise;
let reportNamesBusy = false;
let pendingExportFormat;
const BATCH_INTERVAL_MS = 1000;
const runEtaDeadlines = new Map();
let nextBatchAt = 0;

class GraphError extends Error {
  constructor(status, code, message) {
    super(message);
    this.status = status;
    this.code = code;
  }
}

function escapeHtml(value) {
  const element = document.createElement('span');
  element.textContent = String(value ?? '');
  return element.innerHTML;
}

function tenantData() {
  return Object.fromEntries(new FormData(form));
}

function validateTenant() {
  const input = form.elements.tenant;
  input.setCustomValidity(['common', 'organizations', 'consumers', 'adfs'].includes(input.value.trim().toLowerCase()) ? 'Bitte eine konkrete Tenant-ID oder Tenant-Domain eingeben.' : '');
}

function currentRedirectUri() {
  return new URL('auth.html', window.location.href).href;
}

function showUsageNotice(review = false) {
  usageNoticeDialog.dataset.required = String(!review);
  usageNoticeCheckbox.checked = review;
  confirmUsageNoticeButton.disabled = !review;
  usageNoticeDialog.showModal();
}

function restoreSession() {
  try {
    const saved = JSON.parse(sessionStorage.getItem(SESSION_KEY));
    if (!saved) return;
    for (const name of ['tenant', 'clientId']) {
      if (saved[name]) form.elements[name].value = saved[name];
    }
    setSaveStatus(true);
  } catch {
    sessionStorage.removeItem(SESSION_KEY);
  }
}

function persistSession() {
  const { tenant, clientId } = tenantData();
  sessionStorage.setItem(SESSION_KEY, JSON.stringify({ tenant, clientId }));
  setSaveStatus(true);
  setAuthMessage();
  updateStartState();
}

function setSaveStatus(saved) {
  const status = document.querySelector('#save-status');
  status.textContent = saved ? 'In dieser Sitzung gespeichert' : 'Noch nicht gespeichert';
  status.classList.toggle('saved', saved);
}

function setAuthMessage(message = '', success = false) {
  authMessage.hidden = !message;
  authMessage.textContent = message;
  authMessage.classList.toggle('success', success);
}

function selectedScopes() {
  const ids = [...scopeGrid.querySelectorAll('input:checked')].map((input) => input.value);
  return scopes.filter((scope) => ids.includes(scope.id));
}

function updateStartState() {
  const count = selectedScopes().length;
  document.querySelector('#selection-count').textContent = `${count} ${count === 1 ? 'Bereich' : 'Bereiche'} ausgewählt`;
  continueButton.disabled = !form.checkValidity();
  startButton.disabled = !form.checkValidity() || count === 0;
  consentButton.disabled = startButton.disabled;
  reportNamesButton.disabled = !form.checkValidity() || reportNamesBusy;
}

function renderScopes() {
  scopeGrid.innerHTML = scopes.map((scope) => `
    <label class="scope-card">
      <input type="checkbox" name="scope" value="${scope.id}" />
      <span>
        <strong>${scope.name}</strong>
        <small>${scope.description}</small>
        <span class="scope-tag" title="${escapeHtml(scope.permissions.join(', '))}">${scope.tag} · ${scope.permissions.length} Read-only</span>
      </span>
    </label>`).join('');
  scopeGrid.addEventListener('change', updateStartState);
}

function setStep(step) {
  document.querySelectorAll('[data-step]').forEach((item) => {
    const value = Number(item.dataset.step);
    item.classList.toggle('done', value < step);
    item.querySelector('.nav-index').innerHTML = value < step ? '<span class="icon icon-check" aria-hidden="true"></span>' : value;
  });
}

function showView(view, scroll = true) {
  const reportView = !['setup', 'scope', 'run'].includes(view);
  const selectedReportView = view === 'report' ? 'overview' : view;
  const workflowView = reportView ? 'report' : view === 'run' ? 'scope' : view;
  setupSection.hidden = view !== 'setup';
  scopeSection.hidden = view !== 'scope';
  runPanel.hidden = view !== 'run';
  reportSection.hidden = !reportView;
  sideNav.querySelectorAll('[data-app-view]').forEach((button) => {
    const current = button.dataset.appView === workflowView;
    button.classList.toggle('active', current);
    if (current) button.setAttribute('aria-current', 'step');
    else button.removeAttribute('aria-current');
  });
  if (reportView) selectReportView(selectedReportView);
  if (scroll) (reportView ? reportSection : view === 'run' ? runPanel : view === 'scope' ? scopeSection : setupSection).scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function renderRunQueue(selected) {
  runEtaDeadlines.clear();
  const jobs = [{ id: 'auth', name: 'Microsoft-Anmeldung', detail: 'Wartet auf Start' }, ...selected.map((scope) => ({
    id: scope.id,
    name: scope.name,
    detail: 'In der Warteschlange',
  }))];
  document.querySelector('#run-queue').innerHTML = jobs.map((job) => `
    <li class="waiting" data-job="${job.id}">
      <span class="queue-state" aria-hidden="true">·</span>
      <span><strong>${escapeHtml(job.name)}</strong><small><span class="queue-detail">${escapeHtml(job.detail)}</span><span class="queue-elapsed" aria-hidden="true"></span></small></span>
      <span class="queue-label">wartet</span>
    </li>`).join('');
  updateRunProgress();
}

function updateRunJob(id, status, detail) {
  const item = document.querySelector(`[data-job="${id}"]`);
  if (!item) return;
  item.className = status;
  item.querySelector('.queue-state').innerHTML = ({ waiting: '·', running: '', done: '<span class="icon icon-check" aria-hidden="true"></span>', error: '!' })[status];
  item.querySelector('.queue-detail').textContent = detail;
  if (status !== 'running') item.querySelector('.queue-elapsed').textContent = '';
  item.querySelector('.queue-label').textContent = ({ waiting: 'wartet', running: 'läuft', done: 'fertig', error: 'Fehler' })[status];
  updateRunProgress();
}

function updateRunProgress() {
  const jobs = [...document.querySelectorAll('#run-queue li')];
  const finished = jobs.filter((job) => job.matches('.done, .error')).length;
  const progress = jobs.length ? Math.round((finished / jobs.length) * 100) : 0;
  const deadline = Math.max(0, ...runEtaDeadlines.values());
  const estimate = deadline > Date.now() ? ` · noch etwa ${formatRemaining((deadline - Date.now()) / 1000)}` : '';
  document.querySelector('#progress-bar').style.width = `${progress}%`;
  document.querySelector('#run-progress').textContent = `${finished} von ${jobs.length} Schritten abgeschlossen · ${progress} %${estimate}`;
}

function createScopeProgress(scope) {
  const started = Date.now();
  const elapsed = document.querySelector(`[data-job="${scope.id}"] .queue-elapsed`);
  const tick = () => {
    const seconds = Math.floor((Date.now() - started) / 1000);
    elapsed.textContent = ` · ${String(Math.floor(seconds / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
    if (runEtaDeadlines.has(scope.id)) updateRunProgress();
  };
  const timer = setInterval(tick, 1000);
  tick();
  return {
    update(detail, etaSeconds) {
      if (Number.isFinite(etaSeconds) && etaSeconds > 0) runEtaDeadlines.set(scope.id, Math.max(runEtaDeadlines.get(scope.id) || 0, Date.now() + etaSeconds * 1000));
      else if (etaSeconds === 0) runEtaDeadlines.delete(scope.id);
      updateRunJob(scope.id, 'running', detail);
      document.querySelector('#run-title').textContent = `${scope.name}: ${detail}`;
    },
    stop() {
      clearInterval(timer);
      runEtaDeadlines.delete(scope.id);
      elapsed.textContent = '';
      updateRunProgress();
    },
  };
}

function requestedPermissions(selected) {
  return [...new Set(['openid', 'profile', 'User.Read', 'Organization.Read.All', ...selected.flatMap((scope) => scope.permissions)])];
}

async function getMsalClient(config) {
  const key = `${config.tenant}:${config.clientId}:${config.redirectUri}`;
  if (msalClient && key === msalConfigKey) return msalClient;
  msalConfigKey = key;
  msalClient = new PublicClientApplication({
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenant}`,
      redirectUri: config.redirectUri,
    },
    cache: { cacheLocation: BrowserCacheLocation.SessionStorage },
  });
  await msalClient.initialize();
  if (sessionStorage.getItem(AUTH_MODE_KEY) !== 'popup-bridge-v2') {
    await msalClient.clearCache();
    sessionStorage.setItem(AUTH_MODE_KEY, 'popup-bridge-v2');
  }
  return msalClient;
}

function isInteractionInProgress(error) {
  return error?.errorCode === 'interaction_in_progress' || error?.message?.includes('interaction_in_progress');
}

async function acquireGraphAccessOnce(config, selected) {
  const client = await getMsalClient(config);
  const scopesToRequest = requestedPermissions(selected);
  let account = client.getActiveAccount() || client.getAllAccounts()[0];
  let response;

  if (!account) {
    response = await client.loginPopup({ scopes: scopesToRequest, prompt: 'select_account', redirectUri: config.redirectUri });
    account = response.account;
  } else {
    try {
      response = await client.acquireTokenSilent({ scopes: scopesToRequest, account, redirectUri: config.redirectUri });
    } catch (error) {
      if (!(error instanceof InteractionRequiredAuthError)) throw error;
      response = await client.acquireTokenPopup({ scopes: scopesToRequest, account, redirectUri: config.redirectUri });
    }
  }

  client.setActiveAccount(account);
  return { accessToken: response.accessToken, account };
}

async function acquireGraphAccess(config, selected) {
  if (authenticationPromise) return authenticationPromise;
  authenticationPromise = acquireGraphAccessOnce(config, selected).catch(async (error) => {
    if (!isInteractionInProgress(error)) throw error;
    updateRunJob('auth', 'running', 'Abgebrochene Anmeldung wird bereinigt und neu geöffnet');
    const client = await getMsalClient(config);
    await client.clearCache();
    const response = await client.loginPopup({ scopes: requestedPermissions(selected), prompt: 'select_account', redirectUri: config.redirectUri });
    client.setActiveAccount(response.account);
    return { accessToken: response.accessToken, account: response.account };
  }).finally(() => {
    authenticationPromise = undefined;
  });
  return authenticationPromise;
}

function wait(milliseconds) {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function formatRemaining(seconds) {
  const minutes = Math.ceil(seconds / 60);
  if (seconds < 60) return `${Math.max(1, Math.ceil(seconds))} s`;
  if (minutes < 60) return `${minutes} Min.`;
  return `${Math.floor(minutes / 60)} Std.${minutes % 60 ? ` ${minutes % 60} Min.` : ''}`;
}

async function paceGraphBatch() {
  const start = Math.max(Date.now(), nextBatchAt);
  nextBatchAt = start + BATCH_INTERVAL_MS;
  if (start > Date.now()) await wait(start - Date.now());
}

async function waitWithCountdown(seconds, update) {
  const deadline = Date.now() + seconds * 1000;
  while (Date.now() < deadline) {
    const remaining = Math.ceil((deadline - Date.now()) / 1000);
    update(remaining);
    await wait(Math.min(1000, deadline - Date.now()));
  }
}

async function graphRequest(path, token, options = {}) {
  const url = graphUrl(path);
  const headers = new Headers(options.headers);
  headers.set('Authorization', `Bearer ${token}`);
  headers.set('Accept', 'application/json');
  if (options.body) headers.set('Content-Type', 'application/json');

  for (let attempt = 0; attempt < 3; attempt += 1) {
    const response = await fetch(url, { ...options, headers, signal: options.signal || AbortSignal.timeout(130000) });
    if ([429, 503].includes(response.status) && attempt < 2) {
      await wait(Math.min(Number(response.headers.get('Retry-After') || 1) * 1000, 5000));
      continue;
    }
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new GraphError(response.status, payload.error?.code, payload.error?.message || `Microsoft Graph antwortete mit HTTP ${response.status}.`);
    return payload;
  }
  throw new GraphError(503, 'serviceUnavailable', 'Microsoft Graph ist vorübergehend nicht erreichbar.');
}

async function graphCollection(path, token, limit = Infinity, options = {}) {
  const values = [];
  let next = path;
  while (next && values.length < limit) {
    const page = await graphRequest(next, token, options);
    values.push(...(page.value || []));
    next = page['@odata.nextLink'];
  }
  return values.slice(0, limit);
}

async function serverReport(report, token) {
  let response;
  try {
    response = await fetch(`/api/reports/${report}`, { method: 'POST', headers: { Authorization: `Bearer ${token}`, Accept: 'application/json' }, signal: AbortSignal.timeout(135000) });
  } catch {
    throw new GraphError(503, 'reportProxyUnavailable', 'Der lokale Report-Collector ist nicht erreichbar.');
  }
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new GraphError(response.status, payload.error?.code, payload.error?.message || `Der Report-Collector antwortete mit HTTP ${response.status}.`);
  return payload.value || [];
}

function chunks(values, size) {
  return Array.from({ length: Math.ceil(values.length / size) }, (_, index) => values.slice(index * size, (index + 1) * size));
}

async function graphBatch(requests, token, progress, label = 'Detailabfragen') {
  const results = new Map();
  const failures = new Map();
  const groups = chunks(requests, 20);
  const started = Date.now();
  const estimateRemaining = (completedGroups, extraSeconds = 0) => {
    const secondsPerGroup = completedGroups
      ? Math.max(BATCH_INTERVAL_MS / 1000, (Date.now() - started) / 1000 / completedGroups)
      : BATCH_INTERVAL_MS / 1000;
    return Math.ceil(extraSeconds + secondsPerGroup * (groups.length - completedGroups));
  };
  for (const [groupIndex, group] of groups.entries()) {
    let pending = group;
    for (let attempt = 0; pending.length && attempt < 3; attempt += 1) {
      progress?.(`${label}: Batch ${groupIndex + 1}/${groups.length} · ${results.size}/${requests.length} Objekte verarbeitet${attempt ? ` · Wiederholung ${attempt}/2` : ''}`, estimateRemaining(groupIndex));
      await paceGraphBatch();
      const batch = await graphRequest('/$batch', token, {
        method: 'POST',
        body: JSON.stringify({
          requests: pending.map((request, index) => ({
            id: String(index + 1),
            method: 'GET',
            url: request.url,
          })),
        }),
      });
      const responses = new Map((batch.responses || []).map((response) => [response.id, response]));
      const retry = [];
      let retryAfter = 2 ** attempt;
      for (const [index, request] of pending.entries()) {
        const response = responses.get(String(index + 1));
        if (response?.status === 200) {
          const value = response.body?.value ?? response.body;
          if (Array.isArray(value) && response.body?.['@odata.nextLink']) value.push(...await graphCollection(response.body['@odata.nextLink'], token));
          results.set(request.key, value);
          continue;
        }
        const failure = { status: response?.status || 503, code: response?.body?.error?.code || 'serviceUnavailable' };
        if ([429, 503, 504].includes(failure.status) && attempt < 2) {
          const suggested = Number(response?.headers?.['Retry-After'] || response?.headers?.['retry-after']);
          if (Number.isFinite(suggested) && suggested > 0) retryAfter = Math.max(retryAfter, suggested);
          retry.push(request);
        } else {
          results.set(request.key, null);
          failures.set(request.key, failure);
        }
      }
      pending = retry;
      if (pending.length) {
        const etaAfterRetry = estimateRemaining(groupIndex);
        await waitWithCountdown(retryAfter, (remaining) => progress?.(`${label}: ${pending.length} Teilabfragen gedrosselt · erneuter Versuch in ${remaining} s`, etaAfterRetry + remaining));
      }
    }
  }
  if (requests.length) progress?.(`${label}: ${requests.length}/${requests.length} Objekte verarbeitet${failures.size ? ` · ${failures.size} nicht lesbar` : ''}`);
  results.failures = failures;
  return results;
}

async function readDatasets(token, specifications, progress = () => {}) {
  const entries = Object.entries(specifications);
  const pending = new Set(entries.map(([, specification]) => specification.label));
  progress(`Aktiv: ${[...pending].join(', ')}`);
  const settled = await Promise.allSettled(entries.map(async ([, specification]) => {
    try {
      const value = specification.report
        ? await serverReport(specification.report, token)
        : specification.single
          ? await graphRequest(specification.path, token)
          : await graphCollection(specification.path, token, specification.limit);
      pending.delete(specification.label);
      progress(`${specification.label}: ${Array.isArray(value) ? value.length : 1} Datensätze · ${pending.size ? `aktiv: ${[...pending].join(', ')}` : 'alle Teilabfragen abgeschlossen'}`);
      return value;
    } catch (error) {
      pending.delete(specification.label);
      progress(`${specification.label}: nicht verfügbar · ${pending.size ? `aktiv: ${[...pending].join(', ')}` : 'alle Teilabfragen abgeschlossen'}`);
      throw error;
    }
  }));
  const data = {};
  const failures = [];
  settled.forEach((result, index) => {
    const [key, specification] = entries[index];
    if (result.status === 'fulfilled') data[key] = result.value;
    else {
      data[key] = [];
      failures.push({ label: specification.label, error: result.reason });
    }
  });
  return { data, failures, allFailed: failures.length === entries.length };
}

async function requiredDatasets(token, specifications, progress) {
  const { data, failures } = await readDatasets(token, specifications, progress);
  if (failures.length) throw failures[0].error;
  return data;
}

function includePartialFailures(result, failures, scopeId, allFailed = false) {
  if (!failures.length) return result;
  return {
    ...result,
    unavailable: failures.length,
    summary: allFailed ? `${failures.length} Abfragen nicht verfügbar` : `${result.summary}; ${failures.length} Teilabfrage(n) nicht verfügbar`,
    findings: [
      ...failures.map(({ label, error }) => ({
        severity: 'error',
        title: `${label} nicht auswertbar`,
        description: friendlyError(error, scopeId, label),
        action: 'Lizenz, Benutzerrolle und delegierte Berechtigung für diese Teilabfrage prüfen.',
      })),
      ...(allFailed ? [] : result.findings.filter((item) => item.severity !== 'ok')),
    ],
  };
}

function attachRoleDefinitions(assignments, definitions) {
  const byId = new Map(definitions.map((definition) => [definition.id, definition]));
  return assignments.map((assignment) => ({ ...assignment, roleDefinition: assignment.roleDefinition || byId.get(assignment.roleDefinitionId) }));
}

async function teamInventory(teams, token, progress) {
  return graphBatch(teams.map((team) => ({ key: `${team.id}:owners`, url: `/groups/${team.id}/owners?$select=id&$top=999` })), token, progress, 'Team-Owner');
}

async function sitePermissionInventory(sites, token, progress) {
  return graphBatch(sites.map((site) => ({
    key: site.id,
    url: `/sites/${site.id}/drive/root/permissions?$select=id,roles,grantedToV2,grantedToIdentitiesV2,link,invitation&$top=999`,
  })), token, progress, 'Site-Berechtigungen');
}

async function tenantDriveInventory(sites, users, token, progress) {
  const sources = [
    ...sites.map((site) => ({ key: `site:${site.id}`, name: site.displayName || site.name || site.webUrl, url: `/sites/${site.id}/drives?$select=id,name,driveType,webUrl,owner` })),
    ...users.map((user) => ({ key: `user:${user.id}`, name: user.displayName || user.userPrincipalName, url: `/users/${user.id}/drives?$select=id,name,driveType,webUrl,owner` })),
  ];
  const inventory = await graphBatch(sources, token, progress, 'Speicherorte');
  return [...new Map(sources.flatMap((source) => (inventory.get(source.key) || []).map((drive) => [drive.id, { ...drive, sourceName: source.name }]))).values()];
}

async function sharedFileInventory(drives, token, progress) {
  const itemsByDrive = new Map();
  let unreadableDrives = 0;
  const groups = chunks(drives, 4);
  for (const [groupIndex, group] of groups.entries()) {
    progress?.(`Dateien & Ordner: Speicherorte ${groupIndex * 4 + 1}–${Math.min((groupIndex + 1) * 4, drives.length)} von ${drives.length}`);
    const settled = await Promise.allSettled(group.map((drive) => graphCollection(
      `/drives/${encodeURIComponent(drive.id)}/root/delta?$select=id,name,webUrl,parentReference,file,folder,size,createdDateTime,lastModifiedDateTime,shared,deleted`,
      token,
      Infinity,
      { headers: { Prefer: 'hierarchicalsharing' } },
    )));
    settled.forEach((result, index) => {
      const drive = group[index];
      if (result.status === 'fulfilled') itemsByDrive.set(drive.id, result.value);
      else unreadableDrives += 1;
    });
    progress?.(`Dateien & Ordner: ${Math.min(itemsByDrive.size + unreadableDrives, drives.length)}/${drives.length} Speicherorte · ${[...itemsByDrive.values()].reduce((sum, items) => sum + items.length, 0)} Elemente`);
  }
  const screening = screenPublicShares(drives, itemsByDrive);
  progress?.(`Öffentliche Freigaben: ${screening.screenedItems} Objekte gesichtet · ${screening.sharedRoots} Freigabe-Wurzeln · ${screening.candidates.length} Kandidaten (${screening.unclassified} ohne Linktyp-Hinweis)`);
  const permissions = await graphBatch(screening.candidates.map(({ drive, item }) => ({
    key: `${drive.id}:${item.id}`,
    url: `/drives/${encodeURIComponent(drive.id)}/items/${encodeURIComponent(item.id)}/permissions?$select=id,roles,expirationDateTime,grantedToV2,grantedToIdentitiesV2,invitation,link`,
  })), token, progress, 'Öffentliche Links verifizieren');
  return {
    itemsByDrive: screening.itemsByDrive,
    permissions,
    permissionFailures: permissions.failures,
    unreadableDrives,
    screening,
  };
}

async function appRoleInventory(servicePrincipals, token, progress) {
  const responses = await graphBatch(servicePrincipals.map((servicePrincipal) => ({
    key: servicePrincipal.id,
    url: `/servicePrincipals/${servicePrincipal.id}/appRoleAssignments?$top=999`,
  })), token, progress, 'App-Rollen');
  const byId = new Map(servicePrincipals.map((servicePrincipal) => [servicePrincipal.id, servicePrincipal]));
  return servicePrincipals.flatMap((client) => (responses.get(client.id) || []).map((assignment) => {
    const resource = byId.get(assignment.resourceId);
    const role = resource?.appRoles?.find((item) => item.id === assignment.appRoleId);
    return {
      client: client.displayName || client.appId,
      resource: resource?.displayName || assignment.resourceDisplayName || assignment.resourceId,
      permission: role?.value || assignment.appRoleId,
    };
  }));
}

async function appOwnerInventory(applications, servicePrincipals, token, progress) {
  return graphBatch([
    ...applications.map((app) => ({ key: `application:${app.id}`, url: `/applications/${app.id}/owners?$select=id,displayName,userPrincipalName&$top=999` })),
    ...servicePrincipals.map((app) => ({ key: `servicePrincipal:${app.id}`, url: `/servicePrincipals/${app.id}/owners?$select=id,displayName,userPrincipalName&$top=999` })),
  ], token, progress, 'App-Owner');
}

async function selectedPrices() {
  const file = form.elements.priceFile.files[0];
  return file?.size ? parsePriceCsv(await file.text()) : new Map();
}

function openAdminConsent() {
  if (!form.reportValidity() || !selectedScopes().length) return;
  persistSession();
  const { tenant, clientId } = tenantData();
  const scopesToRequest = requestedPermissions(selectedScopes()).map((permission) =>
    ['openid', 'profile'].includes(permission) ? permission : `https://graph.microsoft.com/${permission}`);
  const redirectUri = new URL('/', window.location.href).href;
  const url = new URL(`https://login.microsoftonline.com/${tenant}/v2.0/adminconsent`);
  const state = crypto.randomUUID();
  sessionStorage.setItem(ADMIN_CONSENT_STATE_KEY, state);
  url.search = new URLSearchParams({ client_id: clientId, scope: scopesToRequest.join(' '), redirect_uri: redirectUri, state });
  window.location.assign(url);
}

async function revealReportNames() {
  if (!form.reportValidity()) return;
  persistSession();
  reportNamesBusy = true;
  updateStartState();
  setAuthMessage('Microsoft-Freigabe für die tenantweite Report-Einstellung wird geöffnet …');
  try {
    const { accessToken } = await acquireGraphAccess(tenantData(), [{ permissions: ['ReportSettings.ReadWrite.All'] }]);
    await graphRequest('/admin/reportSettings', accessToken, {
      method: 'PATCH',
      body: JSON.stringify({ displayConcealedNames: false }),
    });
    setAuthMessage('Klarnamen sind tenantweit aktiviert. Microsoft kann einige Minuten benötigen; danach die Inventur aktualisieren.', true);
    showToast('Klarnamen in M365-Reports aktiviert');
  } catch (error) {
    setAuthMessage(`Klarnamen konnten nicht aktiviert werden: ${friendlyError(error, 'reportSettings')}`);
  } finally {
    reportNamesBusy = false;
    updateStartState();
  }
}

const inventoryRunners = {
  async tenant(token, progress) {
    const data = await requiredDatasets(token, {
      organizations: { label: 'Organisation', path: '/organization?$select=id,displayName,createdDateTime,onPremisesSyncEnabled,technicalNotificationMails,verifiedDomains' },
      domains: { label: 'Domains', path: '/domains?$select=id,isDefault,isInitial,isRoot,isVerified,authenticationType,availabilityStatus' },
    }, progress);
    return analyseTenant(data.organizations, data.domains);
  },
  async identity(token, progress) {
    const data = await requiredDatasets(token, { users: { label: 'Benutzer und Gäste', path: '/users?$select=id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime,lastPasswordChangeDateTime,onPremisesSyncEnabled&$top=999' } }, progress);
    return analyseIdentities(data.users);
  },
  async roles(token, progress) {
    const data = await requiredDatasets(token, {
      definitions: { label: 'Rollendefinitionen', path: '/roleManagement/directory/roleDefinitions' },
      assignments: { label: 'Aktive Rollenzuweisungen', path: '/roleManagement/directory/roleAssignments?$expand=principal' },
    }, progress);
    return analyseRoles(attachRoleDefinitions(data.assignments, data.definitions));
  },
  async access(token, progress) {
    const data = await requiredDatasets(token, { policies: { label: 'Conditional-Access-Richtlinien', path: '/identity/conditionalAccess/policies?$select=id,displayName,state,createdDateTime,modifiedDateTime&$top=999' } }, progress);
    return analyseAccess(data.policies);
  },
  async licenses(token, progress) {
    const [{ data, failures }, prices] = await Promise.all([
      readDatasets(token, {
        skus: { label: 'Lizenzprodukte', path: '/subscribedSkus?$select=skuId,skuPartNumber,consumedUnits,prepaidUnits,capabilityStatus,servicePlans' },
        users: { label: 'Lizenzzuweisungen', path: '/users?$select=id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses,licenseAssignmentStates&$top=999' },
        apps: { label: 'Microsoft-365-App-Nutzung', report: 'm365-apps' },
        active: { label: 'Dienstnutzung', path: "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserDetail(period='D90')?$format=application/json&$top=999" },
      }, progress),
      selectedPrices(),
    ]);
    let reportSettings;
    try {
      progress('Report-Anonymisierung wird geprüft');
      const response = await graphRequest('/admin/reportSettings', token);
      reportSettings = response.value || response;
      progress('Report-Anonymisierung gelesen · Lizenzberatung wird berechnet');
    } catch (error) {
      failures.push({ label: 'Report-Anonymisierung', error });
    }
    const usageLabels = new Set(['Microsoft-365-App-Nutzung', 'Dienstnutzung', 'Report-Anonymisierung']);
    const usage = {
      appReports: data.apps,
      activeReports: data.active,
      concealed: reportSettings?.displayConcealedNames === true,
      complete: reportSettings?.displayConcealedNames === false && !failures.some(({ label }) => usageLabels.has(label)),
    };
    return includePartialFailures(analyseLicenses(data.skus, data.users, prices, usage), failures, 'licenses');
  },
  async usage(token, progress) {
    const { data, failures, allFailed } = await readDatasets(token, {
      users: { label: 'Benutzerverzeichnis', path: '/users?$select=id,displayName,userPrincipalName,accountEnabled,userType,assignedLicenses&$top=999' },
      apps: { label: 'Microsoft-365-App-Nutzung', report: 'm365-apps' },
      active: { label: 'Dienstnutzung', path: "https://graph.microsoft.com/beta/reports/getOffice365ActiveUserDetail(period='D90')?$format=application/json&$top=999" },
      teams: { label: 'Teams-Nutzung', path: "https://graph.microsoft.com/beta/reports/getTeamsUserActivityUserDetail(period='D90')?$format=application/json&$top=999" },
      email: { label: 'E-Mail-Nutzung', path: "https://graph.microsoft.com/beta/reports/getEmailActivityUserDetail(period='D90')?$format=application/json&$top=999" },
      copilot: { label: 'Copilot-Nutzung', report: 'copilot' },
    }, progress);
    let reportSettings = {};
    try {
      progress('Report-Anonymisierung wird geprüft');
      const response = await graphRequest('/admin/reportSettings', token);
      reportSettings = response.value || response;
      progress('Report-Anonymisierung gelesen · Nutzungsprofile werden zusammengeführt');
    } catch (error) {
      failures.push({ label: 'Report-Anonymisierung', error });
    }
    return includePartialFailures(analyseUsage(data.users, data.apps, data.active, data.teams, data.email, data.copilot, reportSettings), failures, 'usage', allFailed);
  },
  async storage(token, progress) {
    const data = await requiredDatasets(token, {
      users: { label: 'Benutzerverzeichnis', path: '/users?$select=id,displayName,userPrincipalName,userType,accountEnabled&$top=999' },
      mailboxes: { label: 'Postfachspeicher', path: "https://graph.microsoft.com/beta/reports/getMailboxUsageDetail(period='D7')?$format=application/json&$top=999" },
      drives: { label: 'OneDrive-Speicher', path: "https://graph.microsoft.com/beta/reports/getOneDriveUsageAccountDetail(period='D7')?$format=application/json&$top=999" },
    }, progress);
    return analyseStorage(data.users, data.mailboxes, data.drives);
  },
  async teams(token, progress) {
    const { groups } = await requiredDatasets(token, { groups: { label: 'Microsoft-365-Gruppen', path: '/groups?$select=id,displayName,description,visibility,groupTypes,mailEnabled,securityEnabled,createdDateTime,renewedDateTime,expirationDateTime,resourceProvisioningOptions&$top=999' } }, progress);
    const teams = groups.filter((group) => group.resourceProvisioningOptions?.includes('Team'));
    return analyseTeams(teams, await teamInventory(teams, token, progress), groups);
  },
  async sites(token, progress) {
    const data = await requiredDatasets(token, {
      sites: { label: 'SharePoint-Sites', path: '/sites?search=*&$select=id,displayName,name,webUrl,createdDateTime,lastModifiedDateTime&$top=999' },
      usage: { label: 'SharePoint-Nutzung', path: "https://graph.microsoft.com/beta/reports/getSharePointSiteUsageDetail(period='D90')?$format=application/json&$top=999" },
      settings: { label: 'SharePoint-Tenanteinstellungen', path: '/admin/sharepoint/settings', single: true },
    }, progress);
    return analyseSites(data.sites, data.usage, data.settings, await sitePermissionInventory(data.sites, token, progress));
  },
  async sharing(token, progress) {
    const data = await requiredDatasets(token, {
      sites: { label: 'SharePoint-Sites', path: '/sites?search=*&$select=id,displayName,name,webUrl&$top=999' },
      users: { label: 'OneDrive-Besitzer', path: '/users?$select=id,displayName,userPrincipalName,accountEnabled,userType&$top=999' },
    }, progress);
    const drives = await tenantDriveInventory(data.sites, data.users.filter((user) => user.accountEnabled && user.userType !== 'Guest'), token, progress);
    const { itemsByDrive, permissions, permissionFailures, unreadableDrives, screening } = await sharedFileInventory(drives, token, progress);
    return analyseSharing(drives, itemsByDrive, permissions, undefined, unreadableDrives, permissionFailures, screening);
  },
  async devices(token, progress) {
    const data = await requiredDatasets(token, { devices: { label: 'Entra-Geräte', path: '/devices?$select=id,deviceId,displayName,operatingSystem,operatingSystemVersion,accountEnabled,approximateLastSignInDateTime,trustType,isCompliant,isManaged&$top=999' } }, progress);
    return analyseDevices(data.devices);
  },
  async apps(token, progress) {
    const data = await requiredDatasets(token, {
      applications: { label: 'App-Registrierungen', path: '/applications?$select=id,appId,displayName,createdDateTime,passwordCredentials,keyCredentials,signInAudience,web,spa&$top=999' },
      servicePrincipals: { label: 'Enterprise Apps', path: '/servicePrincipals?$select=id,appId,displayName,accountEnabled,appRoleAssignmentRequired,servicePrincipalType,passwordCredentials,keyCredentials,appRoles&$top=999' },
      grants: { label: 'Delegierte OAuth-Consents', path: '/oauth2PermissionGrants?$top=999' },
    }, progress);
    const [appRoles, owners] = await Promise.all([appRoleInventory(data.servicePrincipals, token, progress), appOwnerInventory(data.applications, data.servicePrincipals, token, progress)]);
    return analyseApplications(data.applications, undefined, data.servicePrincipals, data.grants, appRoles, owners);
  },
  async security(token, progress) {
    const data = await requiredDatasets(token, { scores: { label: 'Microsoft Secure Score', path: '/security/secureScores?$top=1' } }, progress);
    return analyseSecurity(data.scores);
  },
  async compliance(token, progress) {
    const { data, failures, allFailed } = await readDatasets(token, {
      labels: { label: 'Aufbewahrungslabels', path: '/security/labels/retentionLabels' },
      cases: { label: 'eDiscovery-Fälle', path: '/security/cases/ediscoveryCases?$top=999' },
    }, progress);
    const labelsAvailable = !failures.some((failure) => failure.label === 'Aufbewahrungslabels');
    const casesAvailable = !failures.some((failure) => failure.label === 'eDiscovery-Fälle');
    return includePartialFailures(analyseCompliance(data.labels, data.cases, labelsAvailable, casesAvailable), failures, 'compliance', allFailed);
  },
  async service(token, progress) {
    const { data, failures, allFailed } = await readDatasets(token, {
      health: { label: 'Dienststatus', path: '/admin/serviceAnnouncement/healthOverviews?$top=999' },
      issues: { label: 'Service-Störungen', path: '/admin/serviceAnnouncement/issues?$top=200', limit: 200 },
      messages: { label: 'Message Center', path: '/admin/serviceAnnouncement/messages?$top=200', limit: 200 },
    }, progress);
    return includePartialFailures(analyseServiceHealth(data.health, data.issues, data.messages), failures, 'service', allFailed);
  },
};

async function runScopeInventory(scope, accessToken) {
  const progress = createScopeProgress(scope);
  progress.update('Microsoft-Graph-Abfragen werden vorbereitet');
  try {
    const result = { ...scope, ...(await inventoryRunners[scope.id](accessToken, progress.update)) };
    updateRunJob(scope.id, 'done', result.summary);
    return result;
  } catch (error) {
    const message = friendlyError(error, scope.id);
    updateRunJob(scope.id, 'error', message);
    return { ...scope, error: message };
  } finally {
    progress.stop();
  }
}

async function readTenantInfo(token, config, account) {
  try {
    const organizations = await graphCollection('/organization?$select=id,displayName,verifiedDomains', token);
    const organization = organizations[0];
    const defaultDomain = organization?.verifiedDomains?.find((domain) => domain.isDefault)?.name;
    return {
      name: organization?.displayName || account.name || config.tenant,
      domain: defaultDomain || config.tenant,
      id: organization?.id || account.tenantId || config.tenant,
    };
  } catch {
    return { name: account.name || config.tenant, domain: config.tenant, id: account.tenantId || config.tenant };
  }
}

function friendlyError(error, scopeId, label = '') {
  if (String(error?.message).includes('9002326')) return 'Entra-Konfiguration fehlerhaft: Die Redirect-URI muss unter „Single-page application“ eingetragen sein, nicht unter „Web“.';
  if (isInteractionInProgress(error)) return 'Eine Anmeldung läuft bereits. Bitte das Microsoft-Fenster abschließen oder schließen.';
  if (scopeId === 'storage' && error.status === 403) return 'Speicherberichte nicht lesbar. Prüfe Reports.Read.All mit Admin-Consent und die Entra-Rolle „Reports Reader“ für das angemeldete Konto.';
  if (scopeId === 'reportSettings' && error.status === 403) return 'Erforderlich sind ReportSettings.ReadWrite.All mit Admin-Consent und eine für Reporteinstellungen berechtigte Microsoft-365-Administratorrolle.';
  if (error.code === 'reportProxyUnavailable') return 'Der lokale Report-Collector ist nicht erreichbar. Bitte den Bereich erneut prüfen oder den TenantScope-Dienst kontrollieren.';
  if (scopeId === 'compliance' && error.status === 403) return 'Purview-Daten benötigen die delegierte Read-Berechtigung und zusätzlich eine passende Purview-Rolle, z. B. Records Management oder eDiscovery Manager.';
  if (error.status === 403) return 'Zugriff verweigert. Prüfe delegierte Graph-Berechtigung und Admin-Consent.';
  if (error.status === 404) return 'Endpunkt nicht verfügbar. Prüfe, ob der zugehörige M365-Dienst lizenziert ist.';
  if (error.status === 429) return 'Microsoft Graph hat die Abfragen gedrosselt. Inventur später erneut starten.';
  if (error.code === 'Authorization_RequestDenied') return 'Graph-Berechtigung oder passende Admin-Rolle fehlt.';
  return error.message || 'Dieser Bereich konnte nicht ausgewertet werden.';
}

async function runInventory() {
  if (!form.reportValidity() || !selectedScopes().length) return;
  persistSession();
  setStep(2);
  setAuthMessage();
  showView('run');
  startButton.disabled = true;

  const selected = selectedScopes();
  const config = tenantData();
  renderRunQueue(selected);
  updateRunJob('auth', 'running', 'Bitte Microsoft-Anmeldung im geöffneten Fenster abschließen');
  document.querySelector('#run-title').textContent = 'Warte auf Microsoft-Anmeldung …';

  try {
    const { accessToken, account } = await acquireGraphAccess(config, selected);
    setAuthMessage(`Verbunden als ${account.username || account.name}.`, true);
    updateRunJob('auth', 'running', `Verbunden als ${account.username || account.name} · Tenant-Stammdaten werden gelesen`);
    document.querySelector('#run-title').textContent = 'Tenant-Stammdaten werden gelesen …';
    const tenant = await readTenantInfo(accessToken, config, account);
    updateRunJob('auth', 'done', `Verbunden als ${account.username || account.name}`);

    // ponytail: two early workers keep the slow scans moving without inviting Graph throttling.
    const earlyScopeIds = new Set(['sharing', 'apps']);
    const earlyRuns = new Map(selected.filter((scope) => earlyScopeIds.has(scope.id)).map((scope) => [scope.id, runScopeInventory(scope, accessToken)]));
    const resultsById = new Map();
    for (const scope of selected.filter((item) => !earlyScopeIds.has(item.id))) {
      resultsById.set(scope.id, await runScopeInventory(scope, accessToken));
    }
    for (const scope of selected.filter((item) => earlyScopeIds.has(item.id))) {
      resultsById.set(scope.id, await earlyRuns.get(scope.id));
    }
    const results = selected.map((scope) => resultsById.get(scope.id));

    document.querySelector('#run-title').textContent = 'Bericht wird erstellt …';
    currentReport = buildReport(tenant, account, results);
    renderReport(currentReport);
    await wait(350);
    setStep(3);
    startButton.innerHTML = 'Inventur aktualisieren <span class="icon icon-arrow-right" aria-hidden="true"></span>';
    showView('overview');
  } catch (error) {
    updateRunJob('auth', 'error', friendlyError(error));
    await wait(350);
    setStep(1);
    setAuthMessage(`Anmeldung fehlgeschlagen: ${friendlyError(error)}`);
    showView('setup');
  } finally {
    updateStartState();
  }
}

function buildReport(tenant, account, results) {
  const orderedResults = [...results].sort((left, right) => Number(left.id === 'service') - Number(right.id === 'service'));
  const successful = orderedResults.filter((result) => !result.error);
  const findings = sortFindingsBySeverity(orderedResults.flatMap((result) => result.error
    ? [{ severity: 'error', title: `${result.name} nicht auswertbar`, description: result.error, action: 'Angegebene Berechtigung, Benutzerrolle oder Produktlizenz prüfen.', scope: result.name, scopeId: result.id }]
    : result.findings.map((item) => ({ ...item, scope: result.name, scopeId: result.id }))));
  const totals = Object.fromEntries(['high', 'medium', 'low'].map((severity) => [severity, findings.filter((item) => item.severity === severity).length]));
  const score = successful.length ? Math.max(0, 100 - totals.high * 12 - totals.medium * 6 - totals.low * 2) : null;
  const failed = orderedResults.reduce((sum, result) => sum + (result.error ? 1 : Number(result.unavailable || 0)), 0);
  return {
    tenant,
    account: account.username || account.name,
    results: orderedResults,
    findings,
    totals,
    score,
    failed,
    createdAt: new Intl.DateTimeFormat('de-DE', { dateStyle: 'long', timeStyle: 'short' }).format(new Date()),
  };
}

function severityLabel(severity) {
  return ({ high: 'Hoch', medium: 'Mittel', low: 'Hinweis', info: 'Info', ok: 'In Ordnung', error: 'Ungeprüft' })[severity];
}

function detailGroups(result) {
  return [result.details, ...(result.extraDetails || [])].filter((details) => details?.rows?.length);
}

function guidanceLinks(scopeId) {
  const guidance = scopeGuidance[scopeId];
  if (!guidance) return '';
  return `<a href="${guidance.helpUrl}" target="_blank" rel="noopener noreferrer">Microsoft-Hilfe <span class="icon icon-external-link" aria-hidden="true"></span></a>
    ${guidance.adminLinks.map(([label, url]) => `<a class="admin-link" href="${url}" target="_blank" rel="noopener noreferrer">${escapeHtml(label)} <span class="icon icon-external-link" aria-hidden="true"></span></a>`).join('')}`;
}

function guidanceActions(scopeId, recheckLabel = 'Bereich erneut prüfen', result) {
  if (!scopeGuidance[scopeId]) return '';
  return `<div class="finding-actions">
    ${guidanceLinks(scopeId)}
    <button type="button" data-recheck-scope="${escapeHtml(scopeId)}">${result?.checkedAt ? 'Nochmals prüfen' : escapeHtml(recheckLabel)}</button>
    ${result?.checkedAt ? `<span class="recheck-status"><span class="icon icon-check" aria-hidden="true"></span>${result.recheckChanged ? 'Ergebnis geändert' : 'Keine Änderung'} · ${new Intl.DateTimeFormat('de-DE', { timeStyle: 'short' }).format(new Date(result.checkedAt))}</span>` : ''}
  </div>`;
}

function findingCards(findings, empty = 'Keine Auffälligkeiten in diesem Bereich.', results = []) {
  const resultsById = new Map(results.map((result) => [result.id, result]));
  return findings.map((finding) => {
    const guidance = scopeGuidance[finding.scopeId];
    return `
    <article class="finding" data-report-scope="${escapeHtml(finding.scopeId)}">
      <span class="severity ${finding.severity}">${severityLabel(finding.severity)}</span>
      <div>
        <strong>${escapeHtml(finding.title)}</strong>
        <span class="finding-label">Feststellung</span><p>${escapeHtml(finding.description)}</p>
        ${guidance ? `<span class="finding-label">Einordnung</span><p>${escapeHtml(guidance.explanation)}</p>` : ''}
        ${guidance ? `<div data-export-section="goodPractice"><span class="finding-label">Good Practice</span><p class="finding-practice">${escapeHtml(guidance.goodPractice)}</p></div>` : ''}
        <span class="finding-label">Empfohlene Maßnahme</span><p class="finding-action">${escapeHtml(finding.action)}</p>
        ${guidanceActions(finding.scopeId, 'Bereich erneut prüfen', resultsById.get(finding.scopeId))}
      </div>
    </article>`;
  }).join('') || `<p>${empty}</p>`;
}

function recommendationGroups(findings) {
  const groups = new Map();
  for (const item of findings.filter(({ severity }) => !['ok', 'info'].includes(severity))) {
    if (!groups.has(item.scopeId)) groups.set(item.scopeId, { scopeId: item.scopeId, scope: item.scope, items: [] });
    groups.get(item.scopeId).items.push(item);
  }
  return [...groups.values()];
}

function recommendationSummary(findings) {
  return recommendationGroups(findings).map((group) => `
    <article class="recommendation-group" data-report-scope="${escapeHtml(group.scopeId)}">
      <h4>${escapeHtml(group.scope)}</h4>
      <ul>${group.items.map((item) => `<li><strong>${severityLabel(item.severity)} – ${escapeHtml(item.title)}:</strong> ${escapeHtml(item.action)}</li>`).join('')}</ul>
      <div class="finding-actions">${guidanceLinks(group.scopeId)}</div>
    </article>`).join('') || '<p>Keine priorisierten Empfehlungen.</p>';
}

function tableCell(value) {
  const text = String(value ?? '');
  const link = trustedMicrosoftUrl(text);
  return link
    ? `<a class="table-link" href="${escapeHtml(link)}" target="_blank" rel="noopener noreferrer">Öffnen <span class="icon icon-external-link" aria-hidden="true"></span></a>`
    : escapeHtml(text);
}

const metricFilterRules = {
  tenant: [
    { details: ['Domains'] },
    { details: ['Domains'], conditions: [{ column: 'Verifiziert', operator: 'equals', value: 'Ja' }] },
    { details: ['Domains'] },
  ],
  identity: [
    { details: ['Benutzerbestand'] },
    { details: ['Benutzerbestand'], conditions: [{ column: 'Aktiv', operator: 'equals', value: 'Ja' }] },
    { details: ['Benutzerbestand'], conditions: [{ column: 'Typ', operator: 'equals', value: 'Guest' }] },
    { details: ['Benutzerbestand'], conditions: [{ column: 'Aktiv', operator: 'equals', value: 'Nein' }] },
  ],
  roles: [
    { details: ['Administrative Rollenzuweisungen'] },
    { details: ['Administrative Rollenzuweisungen'], conditions: [{ column: 'Rolle', operator: 'oneOf', value: ['Global Administrator', 'Privileged Role Administrator', 'Security Administrator', 'Conditional Access Administrator', 'Exchange Administrator', 'SharePoint Administrator', 'Teams Administrator', 'Intune Administrator', 'Application Administrator', 'Cloud Application Administrator', 'User Administrator'] }] },
    { details: ['Administrative Rollenzuweisungen'], conditions: [{ column: 'Rolle', operator: 'equals', value: 'Global Administrator' }] },
  ],
  access: [
    { details: ['Conditional-Access-Richtlinien'] },
    { details: ['Conditional-Access-Richtlinien'], conditions: [{ column: 'Status', operator: 'equals', value: 'enabled' }] },
    { details: ['Conditional-Access-Richtlinien'], conditions: [{ column: 'Status', operator: 'equals', value: 'disabled' }] },
  ],
  licenses: [
    { details: ['Lizenzprodukte und Kapazitäten'] },
    { details: ['Rückgabe- und Stilllegungskandidaten'] },
    { details: ['Downgrade-Prüfkandidaten'] },
    { details: ['Einsparpotenzial nach Hebel'] },
  ],
  usage: [
    { details: ['Nutzung je Benutzer (90 Tage)'], conditions: [{ column: 'Letzte Aktivität', operator: 'notMissing' }] },
    { details: ['Nutzung je Benutzer (90 Tage)'], conditions: [{ column: 'Letzte Aktivität', operator: 'missing' }] },
    { details: ['Nutzung je Benutzer (90 Tage)'], conditions: [{ column: 'Copilot-Apps', operator: 'notMissing' }] },
    { details: ['Nutzung je Benutzer (90 Tage)'] },
  ],
  storage: Array.from({ length: 3 }, () => ({ details: ['Speicher je Benutzer'] })),
  teams: [
    { details: ['Entra-Gruppen'] },
    { details: ['Teams und Owner'] },
    { details: ['Teams und Owner'], conditions: [{ column: 'Owner', operator: 'equals', value: '0' }] },
    { details: ['Teams und Owner'], conditions: [{ column: 'Sichtbarkeit', operator: 'equals', value: 'Public' }] },
  ],
  sites: [
    { details: ['Site-Bestand und Nutzung'] },
    { details: ['Site-Bestand und Nutzung'], conditions: [{ column: 'Letzte Aktivität', operator: 'olderThanDays', value: 180 }] },
    { details: ['Site-Bestand und Nutzung'] },
    { details: ['Site-Bestand und Nutzung'], conditions: [{ column: 'Anonyme Links', operator: 'numberAbove', value: 0 }] },
  ],
  sharing: [
    { details: ['Geteilte Ordner und Dateien'] },
    { details: ['Geteilte Ordner und Dateien'], conditions: [{ column: 'Freigabeart', operator: 'contains', value: 'Jeder mit Link' }] },
    { details: ['Geteilte Ordner und Dateien'], conditions: [{ column: 'Rechte', operator: 'contains', value: 'write' }] },
    { details: ['Geteilte Ordner und Dateien'], conditions: [{ column: 'Alter (Tage)', operator: 'numberAtLeast', value: 180 }] },
  ],
  devices: [
    { details: ['Entra-Geräte'] },
    { details: ['Entra-Geräte'], conditions: [{ column: 'Verwaltet', operator: 'equals', value: 'Ja' }] },
    { details: ['Entra-Geräte'], conditions: [{ column: 'Letzte Anmeldung', operator: 'olderThanDays', value: 30 }] },
    { details: ['Entra-Geräte'], conditions: [{ column: 'Aktiv', operator: 'equals', value: 'Nein' }] },
  ],
  apps: [
    { details: ['App-Registrierungen'] },
    { details: ['Enterprise Apps'] },
    { details: ['Zeitkritische App-Credentials'], conditions: [{ column: 'Status', operator: 'equals', value: 'Abgelaufen' }] },
    { details: ['Application Permissions', 'Delegierte OAuth-Consents'], conditions: [{ column: 'Permission', operator: 'broadPermission' }] },
  ],
  security: Array.from({ length: 3 }, () => ({ details: ['Secure-Score-Messung'] })),
  compliance: [
    { details: ['Records-Management-Labels'] },
    { details: ['Records-Management-Labels'] },
    { details: ['eDiscovery-Fälle'] },
    { details: ['eDiscovery-Fälle'], conditions: [{ column: 'Status', operator: 'notOneOf', value: ['closed', 'resolved'] }] },
  ],
  service: [
    { details: ['Service Health'] },
    { details: ['Service Health'], conditions: [{ column: 'Status', operator: 'notOneOf', value: ['serviceOperational', 'restoringService', 'serviceRestored'] }] },
    { details: ['Aktive Health-Issues'] },
    { details: ['Service-Nachrichten'] },
  ],
};

function detailTables(result) {
  return detailGroups(result).map((details) => `
    <details class="detail-group" data-detail-title="${escapeHtml(details.title)}" open>
      <summary><strong>${escapeHtml(details.title)}</strong><span>${details.rows.length} Einträge</span></summary>
      <div class="table-filter">
        <label>Tabelle filtern <input type="search" data-table-search placeholder="Name, Konto, Lizenz …"></label>
        ${Number.isInteger(details.ageColumn) ? `<label>Wie lange schon geteilt?
        <select data-age-filter>
          <option value="0">Alle Freigaben</option>
          <option value="30">Mindestens 30 Tage</option>
          <option value="90">Mindestens 90 Tage</option>
          <option value="180">Mindestens 180 Tage</option>
          <option value="365">Mindestens 1 Jahr</option>
        </select>
        </label>` : ''}
        <span class="metric-filter-label" data-metric-filter-label hidden></span>
        <button type="button" data-clear-table-filter hidden>Filter löschen</button>
      </div>
      <div class="detail-table-wrap">
        <table class="inventory-table detail-table">
          <thead><tr>${details.columns.map(sortableHeader).join('')}</tr></thead>
          <tbody>${details.rows.map((row) => `<tr${Number.isInteger(details.ageColumn) ? ` data-shared-days="${escapeHtml(row[details.ageColumn])}"` : ''}>${row.map((value) => `<td>${tableCell(value)}</td>`).join('')}</tr>`).join('')}</tbody>
        </table>
      </div>
    </details>`).join('') || '<p>Keine Einzelobjekte für diesen Bereich.</p>';
}

function sortableHeader(label) {
  return `<th scope="col" data-column="${escapeHtml(label)}" aria-sort="none"><button class="table-sort" type="button" title="Nach ${escapeHtml(label)} sortieren">${escapeHtml(label)} <span class="icon icon-sort" aria-hidden="true"></span></button></th>`;
}

function sortTable(button) {
  const header = button.closest('th');
  const table = header.closest('table');
  const headers = [...table.tHead.rows[0].cells];
  const column = headers.indexOf(header);
  const direction = header.getAttribute('aria-sort') === 'ascending' ? 'descending' : 'ascending';
  headers.forEach((item) => {
    item.setAttribute('aria-sort', 'none');
    item.querySelector('.table-sort span').className = 'icon icon-sort';
  });
  header.setAttribute('aria-sort', direction);
  button.querySelector('span').className = `icon icon-sort-${direction === 'ascending' ? 'up' : 'down'}`;
  [...table.tBodies[0].rows]
    .map((row, index) => ({ row, index }))
    .sort((a, b) => (compareTableValues(a.row.cells[column]?.textContent, b.row.cells[column]?.textContent) * (direction === 'ascending' ? 1 : -1)) || a.index - b.index)
    .forEach(({ row }) => table.tBodies[0].append(row));
  paginateCollection(table, 1);
}

function collectionItems(collection) {
  const items = collection.matches('table') ? [...collection.tBodies[0].rows] : [...collection.children];
  return items.filter((item) => item.dataset.filteredOut !== 'true');
}

function paginateCollection(collection, page) {
  const allItems = collection.matches('table') ? [...collection.tBodies[0].rows] : [...collection.children];
  const items = collectionItems(collection);
  const pages = Math.ceil(items.length / 10);
  allItems.filter((item) => item.dataset.filteredOut === 'true').forEach((item) => { item.hidden = true; });
  if (!items.length) {
    if (collection.paginationControls) collection.paginationControls.innerHTML = '<span>0 Treffer</span>';
    return;
  }
  const current = Math.max(1, Math.min(page, pages));
  items.forEach((item, index) => { item.hidden = index < (current - 1) * 10 || index >= current * 10; });
  if (!collection.paginationControls) return;
  collection.paginationControls.innerHTML = `
    <button type="button" data-page="${current - 1}" ${current === 1 ? 'disabled' : ''} aria-label="Vorherige Seite"><span class="icon icon-chevron-left" aria-hidden="true"></span></button>
    <span>${(current - 1) * 10 + 1}–${Math.min(current * 10, items.length)} von ${items.length} · Seite ${current}/${pages}</span>
    <button type="button" data-page="${current + 1}" ${current === pages ? 'disabled' : ''} aria-label="Nächste Seite"><span class="icon icon-chevron-right" aria-hidden="true"></span></button>`;
}

function metricConditionMatches(table, row, condition) {
  const column = [...table.tHead.rows[0].cells].findIndex((cell) => cell.dataset.column === condition.column);
  if (column < 0) return true;
  const raw = row.cells[column]?.textContent.trim() || '';
  const value = raw.toLocaleLowerCase('de-DE');
  const expected = String(condition.value ?? '').toLocaleLowerCase('de-DE');
  const missing = !raw || /^[-–—]$/.test(raw);
  if (condition.operator === 'equals') return value === expected;
  if (condition.operator === 'contains') return value.includes(expected);
  if (condition.operator === 'oneOf') return condition.value.some((item) => value === String(item).toLocaleLowerCase('de-DE'));
  if (condition.operator === 'notOneOf') return !condition.value.some((item) => value === String(item).toLocaleLowerCase('de-DE'));
  if (condition.operator === 'missing') return missing;
  if (condition.operator === 'notMissing') return !missing;
  if (condition.operator === 'numberAbove') return Number(raw) > condition.value;
  if (condition.operator === 'numberAtLeast') return Number(raw) >= condition.value;
  if (condition.operator === 'broadPermission') return /(^|\.)(readwrite|fullcontrol|manage|send)(\.|$)|directory\.read\.all|rolemanagement/i.test(raw);
  if (condition.operator === 'olderThanDays') {
    const match = raw.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    return Boolean(match && Date.now() - new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1])).getTime() >= condition.value * 86400000);
  }
  return true;
}

function refreshTableFilter(table) {
  const group = table.closest('.detail-group');
  const query = group.querySelector('[data-table-search]').value.trim().toLocaleLowerCase('de-DE');
  const minimumAge = Number(group.querySelector('[data-age-filter]')?.value || 0);
  const conditions = table.metricFilter?.conditions || [];
  [...table.tBodies[0].rows].forEach((row) => {
    const age = Number(row.dataset.sharedDays);
    const visible = (!query || row.textContent.toLocaleLowerCase('de-DE').includes(query))
      && (!minimumAge || (row.dataset.sharedDays && Number.isFinite(age) && age >= minimumAge))
      && conditions.every((condition) => metricConditionMatches(table, row, condition));
    row.dataset.filteredOut = visible ? 'false' : 'true';
  });
  const visible = collectionItems(table).length;
  const total = table.tBodies[0].rows.length;
  group.querySelector('summary span').textContent = visible === total ? `${total} Einträge` : `${visible} von ${total} Einträgen`;
  const label = group.querySelector('[data-metric-filter-label]');
  label.hidden = !table.metricLabel;
  label.textContent = table.metricLabel ? `Kachel: ${table.metricLabel}` : '';
  group.querySelector('[data-clear-table-filter]').hidden = !query && !minimumAge && !table.metricFilter;
  paginateCollection(table, 1);
}

function applyMetricFilter(button) {
  const panel = button.closest('[data-report-panel]');
  const groups = [...panel.querySelectorAll('.detail-group')];
  const fallback = groups[0]?.dataset.detailTitle;
  const rule = metricFilterRules[panel.dataset.reportScope]?.[Number(button.dataset.metricIndex)] || { details: [fallback] };
  const targets = rule.details || [fallback];
  const targetGroups = groups.filter((group) => targets.includes(group.dataset.detailTitle));
  if (!targetGroups.length) {
    showToast(`${button.querySelector('small').textContent}: keine Einträge`);
    return;
  }
  for (const group of groups) {
    const selected = targets.includes(group.dataset.detailTitle);
    group.open = selected;
    const table = group.querySelector('table');
    table.metricFilter = selected ? rule : null;
    table.metricLabel = selected ? button.querySelector('small').textContent : '';
    group.querySelector('[data-table-search]').value = '';
    if (group.querySelector('[data-age-filter]')) group.querySelector('[data-age-filter]').value = '0';
    refreshTableFilter(table);
  }
  panel.querySelectorAll('[data-metric-index]').forEach((card) => card.setAttribute('aria-pressed', String(card === button)));
  targetGroups[0].scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function clearTableFilter(button) {
  const group = button.closest('.detail-group');
  const table = group.querySelector('table');
  group.querySelector('[data-table-search]').value = '';
  if (group.querySelector('[data-age-filter]')) group.querySelector('[data-age-filter]').value = '0';
  table.metricFilter = null;
  table.metricLabel = '';
  group.closest('[data-report-panel]').querySelectorAll('[data-metric-index]').forEach((card) => card.setAttribute('aria-pressed', 'false'));
  refreshTableFilter(table);
}

function setupPagination(root) {
  root.querySelectorAll('table, .finding-list').forEach((collection) => {
    if (collectionItems(collection).length <= 10) return;
    const controls = document.createElement('nav');
    controls.className = 'pagination';
    controls.setAttribute('aria-label', 'Seitennavigation');
    controls.setAttribute('aria-live', 'polite');
    controls.paginationOwner = collection;
    collection.paginationControls = controls;
    (collection.matches('table') ? collection.closest('.detail-table-wrap') || collection : collection).after(controls);
    paginateCollection(collection, 1);
  });
}

function renderReportNavigation(results) {
  const items = [
    { id: 'overview', name: 'Gesamtübersicht', status: 'Gesamtstatus', statusClass: '' },
    ...results.map((result) => ({
      id: result.id,
      name: result.name,
      status: result.error ? 'Nicht auswertbar' : result.unavailable ? 'Teilweise auswertbar' : 'Live analysiert',
      statusClass: result.error ? 'error' : result.unavailable ? 'partial' : '',
    })),
  ];
  reportOverviewNav.disabled = false;
  reportNav.hidden = false;
  reportNav.innerHTML = items.map((item) => `
    <button class="side-nav-item" type="button" data-report-view="${escapeHtml(item.id)}" aria-label="${escapeHtml(`${item.name}: ${item.status}`)}">
      <span class="nav-status ${item.statusClass}" aria-hidden="true"></span>
      <strong>${escapeHtml(item.name)}</strong>
    </button>`).join('');
  reportNavSelect.innerHTML = items.map((item) => `<option value="${escapeHtml(item.id)}">${escapeHtml(item.name)}</option>`).join('');
}

function fitMetricValues(root) {
  requestAnimationFrame(() => root.querySelectorAll('.score-card strong').forEach((value) => {
    value.style.fontSize = '';
    const base = parseFloat(getComputedStyle(value).fontSize);
    value.style.fontSize = `${fitMetricFontSize(base, value.clientWidth, value.scrollWidth)}px`;
  }));
}

function selectReportView(view) {
  const selected = document.querySelector(`[data-report-panel="${CSS.escape(view)}"]`) || document.querySelector('[data-report-panel="overview"]');
  document.querySelectorAll('[data-report-panel]').forEach((panel) => { panel.hidden = panel !== selected; });
  reportNav.querySelectorAll('[data-report-view]').forEach((button) => {
    const current = button.dataset.reportView === selected.dataset.reportPanel;
    button.classList.toggle('active', current);
    if (current) button.setAttribute('aria-current', 'page');
    else button.removeAttribute('aria-current');
  });
  reportNavSelect.value = selected.dataset.reportPanel;
  document.querySelector('#report-title').textContent = selected.dataset.reportPanel === 'overview' ? 'Governance-Inventur' : selected.dataset.reportTitle;
  fitMetricValues(selected);
}

function renderReport(report) {
  const { tenant, account, results, findings, totals, score, failed, createdAt } = report;
  const scoreLabel = score == null ? 'Nicht berechenbar' : score >= 80 ? 'Gute Basis' : score >= 60 ? 'Handlungsbedarf' : 'Prioritär handeln';
  const reportContent = document.querySelector('#report-content');
  reportContent.innerHTML = `
    <div data-report-panel="overview" data-report-title="Governance-Inventur">
    <div class="report-meta">
      <span>Organisation: <strong>${escapeHtml(tenant.name)}</strong></span>
      <span>Domain: <strong>${escapeHtml(tenant.domain)}</strong></span>
      <span>Angemeldet: <strong>${escapeHtml(account)}</strong></span>
      <span>Erstellt: <strong>${escapeHtml(createdAt)}</strong></span>
    </div>
    <div class="score-row" data-export-section="summary">
      <article class="score-card primary"><small>Governance-Indikator</small><strong>${score == null ? '–' : `${score}/100`}</strong><p>${scoreLabel}</p></article>
      <article class="score-card"><small>Hohes Risiko</small><strong>${totals.high}</strong><p>zeitnah bearbeiten</p></article>
      <article class="score-card"><small>Mittleres Risiko</small><strong>${totals.medium}</strong><p>einplanen</p></article>
      <article class="score-card"><small>Ungeprüfte Abfragen</small><strong>${failed}</strong><p>Rolle, Consent oder Lizenz prüfen</p></article>
    </div>
    <section class="report-section export-recommendations" data-export-section="recommendations">
      <h3>Empfehlungsübersicht</h3>
      <div class="recommendation-summary">${recommendationSummary(findings)}</div>
    </section>
    <section class="report-section" data-export-section="findings">
      <h3>Priorisierte Befunde</h3>
      <div class="finding-list">
        ${findingCards(findings, 'Keine Bereiche konnten ausgewertet werden.', results)}
      </div>
    </section>
    <section class="report-section" data-export-section="inventory">
      <h3>Inventurübersicht</h3>
      <table class="inventory-table">
        <thead><tr>${['Bereich', 'Ergebnis', 'Kennzahlen', 'Status'].map(sortableHeader).join('')}</tr></thead>
        <tbody>${results.map((result) => `<tr data-report-scope="${escapeHtml(result.id)}">
          <td>${escapeHtml(result.name)}</td>
          <td>${escapeHtml(result.error || result.summary)}</td>
          <td>${result.error ? '–' : result.metrics.map(([value, label]) => `${escapeHtml(value)} ${escapeHtml(label)}`).join('<br>')}</td>
          <td class="${result.error || result.unavailable ? 'status-error' : 'status-ok'}">${result.error ? 'Nicht auswertbar' : result.unavailable ? `Teilweise · ${result.unavailable} ungeprüft` : '<span class="icon icon-check" aria-hidden="true"></span> Live analysiert'}</td>
        </tr>`).join('')}</tbody>
      </table>
    </section>
    </div>
    ${results.map((result) => {
      const areaFindings = findings.filter((finding) => finding.scope === result.name);
      const guidance = scopeGuidance[result.id];
      return `<div data-report-panel="${escapeHtml(result.id)}" data-report-title="${escapeHtml(result.name)}" data-report-scope="${escapeHtml(result.id)}" hidden>
        <div class="area-heading">
          <p>${escapeHtml(result.error || result.summary)}</p>
          ${guidance ? `<p class="area-explanation">${escapeHtml(guidance.explanation)}</p><div class="area-practice" data-export-section="goodPractice"><strong>Good Practice</strong><p>${escapeHtml(guidance.goodPractice)}</p>${guidanceActions(result.id, 'Diesen Bereich erneut prüfen', result)}</div>` : ''}
        </div>
        ${result.error ? '' : `<div class="area-metrics" data-export-section="summary">${result.metrics.map(([value, label], index) => `<button class="score-card metric-card" type="button" data-metric-index="${index}" aria-pressed="false" title="${escapeHtml(label)} in den Details anzeigen"><small>${escapeHtml(label)}</small><strong>${escapeHtml(value)}</strong><span class="metric-link">Details filtern <span class="icon icon-arrow-right" aria-hidden="true"></span></span></button>`).join('')}</div>`}
        <section class="report-section" data-export-section="findings"><h3>Befunde</h3><div class="finding-list">${findingCards(areaFindings, undefined, [result])}</div></section>
        <section class="report-section" data-export-section="details"><h3>Details</h3><div class="detail-groups">${detailTables(result)}</div></section>
      </div>`;
    }).join('')}
    <p class="demo-disclaimer"><strong>Datenquelle:</strong> Live-Abfragen der Microsoft Graph API v1.0 sowie ausgewiesener Graph-Beta-Reports. Der Governance-Indikator ist eine eigene Heuristik und nicht der Microsoft Secure Score. Exchange-, Power-Platform- und vollständige SharePoint-Tiefenprüfungen benötigen später einen lokalen Collector. Es werden keine Rohdaten dauerhaft gespeichert.</p>`;
  renderReportNavigation(results);
  setupPagination(reportContent);
  selectReportView('overview');
}

async function recheckScope(scopeId) {
  if (!currentReport || !form.reportValidity()) return;
  const scope = scopes.find((item) => item.id === scopeId);
  if (!scope) return;
  const activeView = document.querySelector('[data-report-panel]:not([hidden])')?.dataset.reportPanel || scopeId;
  const buttons = [...document.querySelectorAll(`[data-recheck-scope="${CSS.escape(scopeId)}"]`)];
  buttons.forEach((button) => { button.disabled = true; button.textContent = 'Prüft live …'; });
  renderRunQueue([scope]);
  updateRunJob('auth', 'running', 'Bestehende Microsoft-Sitzung wird geprüft');
  document.querySelector('#run-title').textContent = `${scope.name}: Anmeldung wird geprüft …`;
  showView('run');
  try {
    const { accessToken, account } = await acquireGraphAccess(tenantData(), [scope]);
    updateRunJob('auth', 'done', `Verbunden als ${account.username || account.name}`);
    const previous = currentReport.results.find((item) => item.id === scopeId);
    const result = await runScopeInventory(scope, accessToken);
    const snapshot = (item) => JSON.stringify({ error: item?.error, summary: item?.summary, metrics: item?.metrics, findings: item?.findings, details: item?.details, extraDetails: item?.extraDetails, unavailable: item?.unavailable });
    result.checkedAt = new Date().toISOString();
    result.recheckChanged = snapshot(previous) !== snapshot(result);
    currentReport = buildReport(currentReport.tenant, { username: currentReport.account }, currentReport.results.map((item) => item.id === scopeId ? result : item));
    renderReport(currentReport);
    await wait(350);
    showView(activeView, false);
    showToast(`${scope.name}: ${result.recheckChanged ? 'Ergebnis geändert' : 'keine Änderung'}`);
  } catch (error) {
    updateRunJob('auth', 'error', friendlyError(error, scope.id));
    buttons.forEach((button) => { button.disabled = false; button.textContent = 'Bereich erneut prüfen'; });
    setAuthMessage(`Re-Check fehlgeschlagen: ${friendlyError(error, scope.id)}`);
    showView(activeView, false);
    showToast('Re-Check fehlgeschlagen');
  }
}

function markdownCell(value) {
  return escapeMarkdown(value);
}

function markdownDetails(results) {
  return results.flatMap((result) => detailGroups(result).map((details) => ({ result, details }))).map(({ result, details }) => {
    const { columns, rows, title } = details;
    return `### ${result.name} – ${title}\n\n| ${columns.map(markdownCell).join(' | ')} |\n|${columns.map(() => '---').join('|')}|\n${rows.map((row) => `| ${row.map(markdownCell).join(' | ')} |`).join('\n')}`;
  }).join('\n\n');
}

function markdownGuidanceLinks(scopeId) {
  const guidance = scopeGuidance[scopeId];
  return [`[Microsoft-Hilfe](${guidance.helpUrl})`, ...guidance.adminLinks.map(([label, url]) => `[${label}](${url})`)].join(' · ');
}

function reportMarkdown(report, options) {
  const sections = [
    `# M365 Governance-Inventur – ${escapeMarkdown(report.tenant.name)}\n\n> Live-Auswertung aus Microsoft Graph v1.0; Speicherberichte über die Graph Reports API beta. Der Governance-Indikator ist nicht der Microsoft Secure Score.\n\n- **Tenant-ID:** ${escapeMarkdown(report.tenant.id)}\n- **Domain:** ${escapeMarkdown(report.tenant.domain)}\n- **Angemeldet:** ${escapeMarkdown(report.account)}\n- **Erstellt:** ${escapeMarkdown(report.createdAt)}`,
  ];
  if (options.summary) sections.push(`## Zusammenfassung\n\n- **Governance-Indikator:** ${report.score == null ? 'nicht berechenbar' : `${report.score}/100`}\n- **Risiken:** ${report.totals.high} hoch, ${report.totals.medium} mittel\n- **Ungeprüfte Abfragen:** ${report.failed}`);
  if (options.recommendations) sections.push(`## Empfehlungsübersicht\n\n${recommendationGroups(report.findings).map((group) => `### ${escapeMarkdown(group.scope)}\n\n${group.items.map((item) => `- **${severityLabel(item.severity)} – ${escapeMarkdown(item.title)}:** ${escapeMarkdown(item.action)}`).join('\n')}\n\n${markdownGuidanceLinks(group.scopeId)}`).join('\n\n') || 'Keine priorisierten Empfehlungen.'}`);
  if (options.findings) sections.push(`## Priorisierte Befunde\n\n${report.findings.map((item) => {
    const guidance = scopeGuidance[item.scopeId];
    return `### ${severityLabel(item.severity)} – ${escapeMarkdown(item.title)}\n\n${escapeMarkdown(item.description)}${guidance ? `\n\n**Einordnung:** ${escapeMarkdown(guidance.explanation)}${options.goodPractice ? `\n\n**Good Practice:** ${escapeMarkdown(guidance.goodPractice)}` : ''}` : ''}\n\n**Empfehlung:** ${escapeMarkdown(item.action)}${guidance ? `\n\n${markdownGuidanceLinks(item.scopeId)}` : ''}`;
  }).join('\n\n') || 'Keine auswertbaren Befunde.'}`);
  if (options.inventory) sections.push(`## Inventurübersicht\n\n| Bereich | Ergebnis | Status |\n|---|---|---|\n${report.results.map((result) => `| ${markdownCell(result.name)} | ${markdownCell(result.error || result.summary)} | ${result.error ? 'Nicht auswertbar' : result.unavailable ? `Teilweise · ${result.unavailable} ungeprüft` : 'Live analysiert'} |`).join('\n')}`);
  if (options.goodPractice) sections.push(`## Good Practice je Bereich\n\n${report.results.map((result) => {
    const guidance = scopeGuidance[result.id];
    return `### ${escapeMarkdown(result.name)}\n\n${escapeMarkdown(guidance.goodPractice)}\n\n${markdownGuidanceLinks(result.id)}`;
  }).join('\n\n')}`);
  if (options.details) sections.push(`## Detailauswertung\n\n${markdownDetails(report.results) || 'Keine betroffenen Einzelobjekte gefunden.'}`);
  return `${sections.join('\n\n')}\n`;
}

function docxInlineLinks(scopeId) {
  const guidance = scopeGuidance[scopeId];
  return [{ text: 'Microsoft-Hilfe', href: guidance.helpUrl }, ...guidance.adminLinks.flatMap(([label, href]) => [{ text: '  ·  ' }, { text: label, href }])];
}

function docxDetailWidths(count) {
  if (count === 1) return [9360];
  const first = [0, 0, 3000, 2600, 2400, 2200, 2000][count];
  const remaining = Math.floor((9360 - first) / (count - 1));
  return [first, ...Array.from({ length: count - 1 }, (_, index) => index === count - 2 ? 9360 - first - remaining * (count - 2) : remaining)];
}

function docxDetailBlocks(result, details) {
  const indexes = details.columns.map((_, index) => index);
  const groups = indexes.length <= 6 ? [indexes] : Array.from({ length: Math.ceil((indexes.length - 1) / 5) }, (_, index) => [0, ...indexes.slice(1 + index * 5, 1 + (index + 1) * 5)]);
  return groups.flatMap((columns, index) => [
    { text: `${result.name} – ${details.title}${groups.length > 1 ? ` · Teil ${index + 1}/${groups.length}` : ''}`, style: 'Heading2' },
    {
      type: 'table',
      columns: columns.map((column) => details.columns[column]),
      rows: details.rows.map((row) => columns.map((column) => row[column])),
      widths: docxDetailWidths(columns.length),
      alignments: columns.map((_, column) => column ? 'center' : 'left'),
    },
  ]);
}

function reportDocxBlocks(report, options) {
  const blocks = [
    { type: 'meta', rows: [['Organisation', report.tenant.name], ['Tenant-ID', report.tenant.id], ['Domain', report.tenant.domain], ['Angemeldet', report.account], ['Erstellt', report.createdAt]] },
    { type: 'callout', label: 'DATENQUELLE', text: 'Live-Auswertung aus Microsoft Graph v1.0; ausgewiesene Nutzungs- und Speicherberichte verwenden zusätzlich Graph beta. Der Governance-Indikator ist nicht der Microsoft Secure Score.' },
  ];
  if (options.summary) blocks.push(
    { text: 'Zusammenfassung', style: 'Heading1' },
    { type: 'metrics', items: [
      { value: report.score == null ? '–' : `${report.score}/100`, label: 'Governance-Indikator' },
      { value: String(report.totals.high), label: 'Hohes Risiko' },
      { value: String(report.totals.medium), label: 'Mittleres Risiko' },
      { value: String(report.failed), label: 'Ungeprüfte Abfragen' },
    ] },
  );
  if (options.recommendations) blocks.push(
    { text: 'Empfehlungsübersicht', style: 'Heading1' },
    {
      type: 'table', columns: ['Priorität', 'Bereich', 'Empfehlung', 'Direkte Links'], widths: [950, 1700, 4550, 2160], alignments: ['center', 'left', 'left', 'left'],
      rows: recommendationGroups(report.findings).flatMap((group) => group.items.map((item) => [severityLabel(item.severity), group.scope, item.action, docxInlineLinks(group.scopeId)])),
    },
  );
  if (options.findings) {
    blocks.push({ text: 'Priorisierte Befunde', style: 'Heading1' });
    for (const item of report.findings) {
      const guidance = scopeGuidance[item.scopeId];
      blocks.push({
        type: 'finding', severity: item.severity, label: severityLabel(item.severity), title: item.title,
        description: item.description, context: guidance?.explanation, goodPractice: options.goodPractice ? guidance?.goodPractice : '',
        action: item.action, links: guidance ? docxInlineLinks(item.scopeId) : [],
      });
    }
  }
  if (options.inventory) blocks.push(
    { text: 'Inventurübersicht', style: 'Heading1' },
    {
      type: 'table', columns: ['Bereich', 'Ergebnis', 'Status'], widths: [1900, 5100, 2360], alignments: ['left', 'left', 'center'],
      rows: report.results.map((result) => [result.name, result.error || result.summary, result.error ? 'Nicht auswertbar' : result.unavailable ? `Teilweise · ${result.unavailable} ungeprüft` : 'Live analysiert']),
    },
  );
  if (options.goodPractice) {
    blocks.push({ text: 'Good Practice je Bereich', style: 'Heading1' });
    for (const result of report.results) {
      const guidance = scopeGuidance[result.id];
      blocks.push({ type: 'callout', title: result.name, text: guidance.goodPractice, links: docxInlineLinks(result.id) });
    }
  }
  if (options.details) {
    blocks.push({ type: 'pageBreak' }, { text: 'Detailauswertung', style: 'Heading1' });
    for (const result of report.results) for (const details of detailGroups(result)) blocks.push(...docxDetailBlocks(result, details));
  }
  return blocks;
}

function safeFilename(organization) {
  return `m365-governance-${organization.toLowerCase().replace(/[^a-z0-9äöüß]+/gi, '-').replace(/^-|-$/g, '') || 'bericht'}`;
}

function download(content, filename, type) {
  const link = document.createElement('a');
  link.href = URL.createObjectURL(content instanceof Blob ? content : new Blob([content], { type }));
  link.download = filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(link.href), 1000);
}

function reportForExport(report, options) {
  if (options.serviceHealth) return report;
  return { ...buildReport(report.tenant, { username: report.account }, report.results.filter(({ id }) => id !== 'service')), createdAt: report.createdAt };
}

function printReport(report, options) {
  const activeView = document.querySelector('[data-report-panel]:not([hidden])')?.dataset.reportPanel || 'overview';
  renderReport(report);
  document.querySelectorAll('[data-export-section]').forEach((element) => {
    if (!options[element.dataset.exportSection]) element.dataset.exportExcluded = 'true';
  });
  document.body.classList.add('print-exporting');
  window.addEventListener('afterprint', () => {
    document.body.classList.remove('print-exporting');
    renderReport(currentReport);
    showView(activeView, false);
  }, { once: true });
  requestAnimationFrame(() => window.print());
}

function exportReport(format, options) {
  if (!currentReport) return;
  const report = reportForExport(currentReport, options);
  const filename = safeFilename(report.tenant.name);
  if (format === 'pdf') {
    printReport(report, options);
  } else if (format === 'md') {
    download(reportMarkdown(report, options), `${filename}.md`, 'text/markdown;charset=utf-8');
    showToast('Markdown-Bericht gespeichert');
  } else {
    const bytes = createDocx('M365 Governance-Inventur', reportDocxBlocks(report, options), { subtitle: `${report.tenant.name} · ${report.createdAt}` });
    download(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), `${filename}.docx`);
    showToast('DOCX-Bericht gespeichert');
  }
}

function openExportDialog(format) {
  if (!currentReport) return;
  pendingExportFormat = format;
  const label = ({ pdf: 'PDF', md: 'Markdown', docx: 'DOCX' })[format];
  document.querySelector('#export-dialog-title').textContent = `${label}-Export konfigurieren`;
  confirmExportButton.textContent = `${label} exportieren`;
  exportDialog.showModal();
}

function showToast(message) {
  toast.textContent = message;
  toast.classList.add('visible');
  setTimeout(() => toast.classList.remove('visible'), 2200);
}

renderScopes();
restoreSession();
validateTenant();
form.elements.redirectUri.value = currentRedirectUri();
if (!sessionStorage.getItem(USAGE_NOTICE_KEY)) showUsageNotice();
const consentResult = new URLSearchParams(window.location.search);
const consentState = sessionStorage.getItem(ADMIN_CONSENT_STATE_KEY);
const hasConsentResult = consentResult.has('admin_consent') || consentResult.has('error');
if (hasConsentResult && (!consentState || consentResult.get('state') !== consentState)) {
  setAuthMessage('Administratorfreigabe verworfen: Die zurückgegebene Sitzung konnte nicht bestätigt werden.');
  sessionStorage.removeItem(ADMIN_CONSENT_STATE_KEY);
  history.replaceState({}, '', window.location.pathname + window.location.hash);
} else if (consentResult.get('admin_consent') === 'True') {
  setAuthMessage('Administratorfreigabe erteilt. Die vollständige Inventur kann gestartet werden.', true);
  sessionStorage.removeItem(ADMIN_CONSENT_STATE_KEY);
  history.replaceState({}, '', window.location.pathname + window.location.hash);
} else if (consentResult.get('error')) {
  setAuthMessage(`Administratorfreigabe fehlgeschlagen: ${consentResult.get('error_description') || consentResult.get('error')}`);
  sessionStorage.removeItem(ADMIN_CONSENT_STATE_KEY);
  history.replaceState({}, '', window.location.pathname + window.location.hash);
}
updateStartState();
showView('setup', false);
form.addEventListener('input', () => { validateTenant(); persistSession(); });
form.addEventListener('change', persistSession);
document.querySelector('#select-all').addEventListener('click', () => {
  scopeGrid.querySelectorAll('input').forEach((input) => { input.checked = true; });
  updateStartState();
});
document.querySelector('#select-none').addEventListener('click', () => {
  scopeGrid.querySelectorAll('input').forEach((input) => { input.checked = false; });
  updateStartState();
});
document.querySelector('#reset-session').addEventListener('click', async () => {
  if (msalClient) await msalClient.clearCache().catch(() => {});
  sessionStorage.removeItem(SESSION_KEY);
  sessionStorage.removeItem(AUTH_MODE_KEY);
  sessionStorage.removeItem(ADMIN_CONSENT_STATE_KEY);
  sessionStorage.removeItem(USAGE_NOTICE_KEY);
  form.reset();
  form.elements.redirectUri.value = currentRedirectUri();
  currentReport = undefined;
  reportOverviewNav.disabled = true;
  reportNav.hidden = true;
  reportNav.innerHTML = '';
  reportNavSelect.innerHTML = '';
  startButton.innerHTML = 'Anmelden & Inventur starten <span class="icon icon-arrow-right" aria-hidden="true"></span>';
  setSaveStatus(false);
  setAuthMessage();
  setStep(1);
  showView('setup');
  updateStartState();
  showToast('Sitzungs- und Anmeldedaten gelöscht');
  showUsageNotice();
});
usageNoticeCheckbox.addEventListener('change', () => {
  confirmUsageNoticeButton.disabled = !usageNoticeCheckbox.checked;
});
usageNoticeDialog.addEventListener('cancel', (event) => {
  if (usageNoticeDialog.dataset.required === 'true') event.preventDefault();
});
usageNoticeForm.addEventListener('submit', (event) => {
  event.preventDefault();
  sessionStorage.setItem(USAGE_NOTICE_KEY, 'accepted');
  usageNoticeDialog.close();
});
document.querySelector('#open-usage-notice').addEventListener('click', () => showUsageNotice(true));
startButton.addEventListener('click', runInventory);
continueButton.addEventListener('click', () => showView('scope'));
consentButton.addEventListener('click', openAdminConsent);
reportNamesButton.addEventListener('click', revealReportNames);
sideNav.addEventListener('click', (event) => {
  const button = event.target.closest('button:not(:disabled)');
  if (button?.dataset.appView) showView(button.dataset.appView);
});
reportNav.addEventListener('click', (event) => {
  const button = event.target.closest('button[data-report-view]');
  if (button) showView(button.dataset.reportView);
});
reportNavSelect.addEventListener('change', () => showView(reportNavSelect.value));
document.querySelector('#report-content').addEventListener('click', (event) => {
  const metricCard = event.target.closest('[data-metric-index]');
  if (metricCard) {
    applyMetricFilter(metricCard);
    return;
  }
  const clearFilter = event.target.closest('[data-clear-table-filter]');
  if (clearFilter) {
    clearTableFilter(clearFilter);
    return;
  }
  const recheckButton = event.target.closest('[data-recheck-scope]');
  if (recheckButton) {
    recheckScope(recheckButton.dataset.recheckScope);
    return;
  }
  const button = event.target.closest('.table-sort');
  if (button) sortTable(button);
  const pageButton = event.target.closest('.pagination button');
  if (pageButton) paginateCollection(pageButton.closest('.pagination').paginationOwner, Number(pageButton.dataset.page));
});
document.querySelector('#report-content').addEventListener('change', (event) => {
  const ageFilter = event.target.closest('[data-age-filter]');
  if (ageFilter) refreshTableFilter(ageFilter.closest('.detail-group').querySelector('table'));
});
document.querySelector('#report-content').addEventListener('input', (event) => {
  const tableFilter = event.target.closest('[data-table-search]');
  if (tableFilter) refreshTableFilter(tableFilter.closest('.detail-group').querySelector('table'));
});
document.querySelectorAll('[data-export]').forEach((button) => button.addEventListener('click', () => openExportDialog(button.dataset.export)));
document.querySelector('#cancel-export').addEventListener('click', () => exportDialog.close());
exportForm.addEventListener('submit', (event) => {
  event.preventDefault();
  const selected = new Set(new FormData(exportForm).keys());
  const options = Object.fromEntries(['summary', 'recommendations', 'findings', 'inventory', 'goodPractice', 'details', 'serviceHealth'].map((name) => [name, selected.has(name)]));
  exportDialog.close();
  exportReport(pendingExportFormat, options);
});
window.addEventListener('resize', () => fitMetricValues(document));
