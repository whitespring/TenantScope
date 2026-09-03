# m365-governance-platform

Lokaler POC für eine echte Microsoft-365-Governance-Inventur über Microsoft Graph.

## Start

```bash
npm install
npm start
```

Danach `http://localhost:4173` öffnen. Die bereitgestellte Instanz läuft unter `https://tenantscope.wspg.org` (`__ORIGIN_IP__:443` hinter dem Reverse Proxy; Port 80 bleibt nur bis zur NPM-Umstellung als Übergang aktiv).

## Enthalten

- Admin-Anmeldung über Microsoft OAuth (Authorization Code + PKCE)
- echte Read-only-Abfragen der Microsoft Graph API v1.0 sowie einzelner ausgewiesener Beta-Reports
- sichtbare Inventur-Queue mit Live-Datensatz, Objekt-/Batch-Zählern und Laufzeit; lange Datei- und App-Prüfungen starten früh parallel
- linke Navigation für Setup, Prüfumfang, Übersicht und einzelne Ergebnisbereiche
- klickbare Kennzahlen als Detailfilter sowie durchsuchbare, sortierbare Tabellen mit 10er-Pagination
- nach Schwere priorisierte Live-Befunde, Good-Practice-Hinweise und Detailtabellen der betroffenen Objekte
- 15 getrennt auswählbare Bereiche: Tenant, Benutzer, Rollen, Conditional Access, Lizenzen/Kosten, Nutzung, Speicher, Teams/Gruppen, SharePoint, öffentliche Freigaben, Entra-Geräte, Apps, Secure Score, Purview und Service Health
- skalierbares Freigabe-Screening: Delta liefert Freigabe-Wurzeln, Detailberechtigungen werden nur für anonyme Link-Kandidaten gelesen; interne und personengebundene Freigaben erzeugen keine Einzelabfragen
- eigener Bereich zur Lizenzoptimierung: freie Seats, deaktivierte oder 90 Tage inaktive Zuweisungen, gruppenbasierte Quellen, vorsichtige Downgrade-Kandidaten und Monats-/Jahrespotenziale
- optionale lokale SKU-Preisliste (`sku;preisProMonat`) für kundenspezifische Monats- und Jahreswerte
- konfigurierbarer Export als PDF, Markdown und echtes DOCX – einschließlich optionaler Empfehlungsübersicht und optionaler Service-Health-Daten
- vorgeschalteter Nutzungshinweis mit Haftungsausschluss sowie dauerhafte Links zu whitespring-Datenschutz und -Impressum

## Einmalige Entra-Konfiguration

1. Im [Microsoft Entra Admin Center](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade) eine Single-Tenant-App registrieren.
2. Unter **Authentication → Add a platform → Single-page application** die Callback-URI eintragen: lokal `http://localhost:4173/auth.html`, produktiv `https://tenantscope.wspg.org/auth.html`. Dieselbe URI darf nicht zusätzlich unter **Web** stehen.
3. Für die Adminfreigabe außerdem die jeweilige Root-URI unter **Web** registrieren: lokal `http://localhost:4173/`, produktiv `https://tenantscope.wspg.org/`.
4. Tenant-ID und Application-ID in TenantScope eingeben, Bereiche wählen und **Adminfreigabe** öffnen. TenantScope fordert die ausgewählten delegierten Read-only-Rechte dynamisch beim Administrator an; sie müssen nicht einzeln im Portal hinzugefügt werden.
5. Danach **Anmelden & Inventur starten**. Neben dem Consent gelten Microsofts dienstspezifische Rollen und Lizenzen; nicht verfügbare Teilabfragen werden einzeln im Bericht ausgewiesen.

## Delegierte Graph-Berechtigungen

| Bereich | Angeforderte Rechte |
|---|---|
| Basis | `User.Read`, `Organization.Read.All` |
| Tenant & Domains | `Domain.Read.All` |
| Benutzer & Gäste | `User.Read.All` |
| Rollen | `RoleManagement.Read.Directory` |
| Conditional Access | `Policy.Read.All` |
| Lizenzoptimierung | `LicenseAssignment.Read.All`, `User.Read.All`, `Reports.Read.All`, `ReportSettings.Read.All` |
| Nutzung/Speicher | `Reports.Read.All`, `ReportSettings.Read.All`, `User.Read.All` |
| Teams & Gruppen | `Group.Read.All` |
| SharePoint | `Sites.Read.All`, `Reports.Read.All`, `SharePointTenantSettings.Read.All` |
| Öffentliche Freigaben | `Sites.Read.All`, `Files.Read.All`, `User.Read.All` |
| Entra-Geräte | `Device.Read.All` |
| Apps & OAuth | `Application.Read.All`, `Directory.Read.All` |
| Microsoft Secure Score | `SecurityEvents.Read.All` |
| Purview | `RecordsManagement.Read.All`, `eDiscovery.Read.All` |
| Service Health | `ServiceHealth.Read.All`, `ServiceMessage.Read.All` |

Für die M365-Reports braucht das anmeldende Konto beispielsweise **Reports Reader**; für SharePoint-Tenant-Einstellungen **Global Reader** oder **SharePoint Administrator**; Purview/eDiscovery benötigt zusätzlich eine passende Purview-Rolle. Copilot-Abfragen liefern nur Daten, wenn der Dienst im Tenant lizenziert ist.

Microsofts anonymisierte Report-Kennungen lassen sich nicht auf Benutzer zurückrechnen. Die Schaltfläche **Tenantweit Klarnamen aktivieren** fordert deshalb bei Bedarf dynamisch `ReportSettings.ReadWrite.All` mit Admin-Consent an und setzt `displayConcealedNames` auf `false`. Diese Änderung betrifft die Microsoft-365-Nutzungsreports tenantweit, wird protokolliert und kann einige Minuten bis zur Wirkung benötigen.

Kein Client-Secret anlegen oder eingeben. MSAL speichert die temporären OAuth-Token im Sitzungsspeicher des Browsers. Microsoft-Reports mit Download-Redirect werden flüchtig über den Serverstandort Deutschland geladen; Token und Reportdaten werden dort nicht gespeichert oder protokolliert. Bereiche ohne passende Lizenz, Benutzerrolle oder Consent werden als nicht auswertbar ausgewiesen; es gibt keine Mock-Daten.

## Sicherheits- und Vertrauensmodell

TenantScope ist eine statische SPA ohne Datenbank. Die meisten Abfragen laufen direkt vom Browser zu Microsoft Graph. Nur zwei Microsoft-Reports, deren Download-Redirect im Browser nicht zuverlässig funktioniert, werden über den lokalen Proxy übertragen. Der Proxy akzeptiert ausschließlich diese festen Report-Routen, folgt nur validierten Microsoft-Download-Zielen, begrenzt Laufzeit und Größe und speichert keine Tokens oder Reportinhalte.

Self-Service bedeutet hier **Bring your own App Registration**: Der Administrator trägt Tenant-ID und Client-ID ein und erteilt den angezeigten delegierten Rechten Consent. Vor dem Consent müssen Tenant, Client-ID, Herausgeber und angeforderte Rechte im Microsoft-Dialog geprüft werden. TenantScope kann technisch nicht beweisen, wem eine frei eingegebene App-ID gehört. Nach einem Test lässt sich der Zugriff im Entra Admin Center über die zugehörige Enterprise Application beziehungsweise deren Consent vollständig widerrufen.

Die produktive Instanz ist auf mehreren Ebenen begrenzt:

- Origin-Zugriff nur vom NPM-Host `__REVERSE_PROXY_IP__`; zusätzliche persistente nftables-Regel im LXC
- kein SSH und kein Maildienst im LXC; Administration nur über Proxmox `pct`
- Nginx-Allowlist für veröffentlichte Assets und zwei API-Routen; Source-, Test-, Deployment- und Dotfiles bleiben außerhalb des Webroots oder werden blockiert
- CSP, Frame-Schutz, MIME-Schutz, restriktive Referrer-/Permissions-Policy und TLS 1.2/1.3
- Rate- und Parallelitätslimits vor dem Report-Proxy
- Proxy als `www-data` ohne Capabilities, mit `NoNewPrivileges`, Syscall-, Adressfamilien-, Speicher- und Task-Limits
- automatische Debian-Sicherheitsupdates; Proxmox-Snapshot `rollback-snapshot` als Rollback-Punkt

Der vollständige Stand, die Prüfnachweise und die offenen Betriebsmaßnahmen stehen in [SECURITY_AUDIT.md](SECURITY_AUDIT.md).

## Vor öffentlicher Freigabe in Nginx Proxy Manager

1. Forward Scheme auf `https`, Forward Host auf `__ORIGIN_IP__` und Forward Port auf `443` setzen.
2. Folgenden Block unter **Advanced → Custom Nginx Configuration** einsetzen. Keine eigene `location /` definieren; sonst werden NPM-Access-List und „Block Common Exploits“ leicht umgangen.

   ```nginx
   proxy_ssl_protocols TLSv1.2 TLSv1.3;
   proxy_ssl_server_name on;
   proxy_ssl_name tenantscope;

   proxy_connect_timeout 5s;
   proxy_send_timeout 10s;
   proxy_read_timeout 130s;

   client_max_body_size 1k;
   proxy_request_buffering off;
   proxy_buffering off;
   proxy_max_temp_file_size 0;
   ```

   NPM setzt `Host`, `X-Forwarded-For`, `X-Real-IP` und `X-Forwarded-Proto` in seiner Standard-Location bereits selbst; diese Header hier nicht doppelt konfigurieren.
3. HSTS entweder auf `max-age=63072000; includeSubDomains; preload` korrigieren, wenn **alle** Subdomains dauerhaft HTTPS nutzen, oder `preload` entfernen.
4. Erst nach erfolgreichem OAuth-, Inventur- und Exporttest die NPM-Access-List von „homelab only“ auf den gewünschten öffentlichen Zugriff umstellen.
5. Danach Port 80 am Origin aus Nginx und nftables entfernen. Bis dahin bleibt er ausschließlich für `__REVERSE_PROXY_IP__` erreichbar.

NPM darf weder den `Authorization`-Header noch Querystrings mit OAuth-Codes protokollieren. Die Datenschutzerklärung muss die technisch notwendigen NPM-Zugriffslogs samt Aufbewahrungsfrist sowie die flüchtige Verarbeitung der zwei Download-Reports beschreiben.

Vorerst ausgeblendet sind MFA-/Sign-in-/Risikoberichte, Teams-Mitglieder und -Channels, Intune, Defender-Alerts/-Incidents sowie Identity Governance. Dafür fehlen im getesteten Tenant derzeit Lizenz, Rolle oder zuverlässig lesbare Daten. Microsoft Graph liefert außerdem keine kundenspezifischen Einkaufspreise; diese kommen optional aus der lokalen CSV. Downgrade-Angaben sind bewusst Prüfkandidaten: Nutzungsreports belegen Kernaktivität, aber nicht den Bedarf an Security, Compliance, Telefonie, Power Platform oder Geräteverwaltung.

## Test

```bash
npm test
```
