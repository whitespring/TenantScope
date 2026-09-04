# TenantScope Security Audit

Stand: 4. September 2026
Zielsystem: `https://tenantscope.wspg.org`
Ergebnis: **bedingt für eine öffentliche Preview freigabefähig**

Die Anwendung selbst und der LXC sind gehärtet. Vor einer uneingeschränkt öffentlichen Freigabe müssen noch die TLS-Verbindung zwischen Nginx Proxy Manager und Origin aktiviert und ein vollständiger OAuth-/Inventur-/Export-Smoke-Test mit einem echten Tenant durchgeführt werden.

## Scope und Methode

Geprüft wurden Quellcode, OAuth- und Graph-Datenfluss, Report-Proxy, Exportpfade, veröffentlichter Webroot, npm-Lieferkette, Nginx, systemd-Sandbox, LXC-Dienste, Host-/Container-Firewall, Betriebssystemupdates, TLS und der über NPM erreichbare Live-Endpunkt. Die Prüfung kombinierte manuelle Codeanalyse, projektspezifische Regressionstests, `npm audit`, Negativtests gegen den Live-Endpunkt und einen Review aus Architektur-, Security-, Test- und Wartbarkeitsperspektive.

Nicht enthalten sind ein externer Penetrationstest, Microsoft-Tenant-Konfigurationen außerhalb der App, der Betrieb des separaten NPM-Hosts sowie die Sicherheitskonfiguration des gesamten Proxmox-Clusters.

## Daten- und Vertrauensgrenzen

1. Der Browser speichert Tenant-ID, Client-ID und MSAL-Token nur im Sitzungsspeicher.
2. OAuth verwendet Authorization Code mit PKCE gegen den konkret eingegebenen Tenant.
3. Normale Inventurabfragen gehen direkt vom Browser an `graph.microsoft.com`.
4. Nur die fünf fest definierten Report-Endpunkte `m365-apps`, `copilot`, `sharepoint-activity`, `onedrive-activity` und `viva-engage` gehen über NPM → Origin-Nginx → loopbackgebundenen Python-Proxy → Microsoft Graph beziehungsweise validierte `reports*.office.com/data/`-Ziele.
5. Berichtsexporte entstehen im Browser. Der Server besitzt keine Tenant-Datenbank und legt keine Reports oder Tokens ab.

## Befunde

| Schwere | Status | Befund und Maßnahme |
|---|---|---|
| Hoch | **Offen vor Public Launch** | NPM leitet aktuell per HTTP an Port 80 des Origins weiter. Damit können Bearer-Token der ausgewählten Download-Reports im internen Segment unverschlüsselt übertragen werden. Origin-TLS ist auf Port 443 vorbereitet; NPM muss auf HTTPS und Port 443 des internen Origin-Hosts umgestellt werden. Anschließend Port 80 entfernen. |
| Hoch | Behoben | Der ursprüngliche Python-Server konnte als Static-File-Server laufen. Produktion nutzt jetzt ausschließlich `BaseHTTPRequestHandler`; statische Auslieferung ist nur im expliziten lokalen Entwicklungsmodus aktiv. |
| Hoch | Behoben | Graph-Folge-URLs und Report-Redirects hätten als Token-/SSRF-Grenze missbraucht werden können. Graph-Ziele sind exakt auf HTTPS `graph.microsoft.com` begrenzt; Download-Redirects ausschließlich auf HTTPS `reports*.office.com`, Port 443 und `/data/`. Authorization wird nicht an das Download-Ziel weitergegeben. |
| Hoch | Behoben | Origin und SSH waren im LAN erreichbar. Eine bootfeste nftables-Policy akzeptiert neue HTTP-/HTTPS-Verbindungen nur vom konfigurierten Reverse-Proxy-Host; SSH und Postfix sind deaktiviert und maskiert. Direkte Verbindungsversuche laufen in einen Timeout. |
| Mittel | **Offen vor Public Launch** | Der echte Ende-zu-Ende-Test mit Admin-Consent, Popup-Rückkehr, allen gewählten Abfragen sowie MD-/DOCX-/PDF-Export benötigt ein reales Microsoft-Konto. Dieser Test kann nicht mit einem synthetischen Token ersetzt werden. |
| Mittel | Offen | Self-Service akzeptiert eine frei eingegebene Client-ID. Das ermöglicht mehrere Tenants, beweist aber nicht die Eigentümerschaft der App-Registrierung. Der Consent-Dialog muss Tenant, Herausgeber, Client-ID und Scopes sichtbar bestätigen. Für ein stärker zentral kontrolliertes Produkt wäre eine feste, verifizierte Multi-Tenant-App die Alternative. |
| Mittel | Offen | Der Proxmox-Host hat ausstehende Sicherheits- und Plattformupdates. Diese betreffen den gemeinsam genutzten Host und müssen in einem eigenen Wartungsfenster mit Cluster-/VM-Folgenabschätzung installiert werden. Der TenantScope-LXC selbst ist aktuell und nutzt unattended upgrades. |
| Mittel | Offen | Die Cluster-Firewall des Virtualisierungshosts ist global deaktiviert. Die vorbereitete VM-Regel wird dadurch nicht angewendet. Die aktive LXC-eigene nftables-Regel kompensiert dies. Eine globale Aktivierung darf erst nach Prüfung der bestehenden Regeln anderer VMs erfolgen. |
| Niedrig | Offen | Der öffentliche HSTS-Header enthält `preload`, aber kein `includeSubDomains`. Entweder alle Subdomains prüfen und `includeSubDomains` ergänzen oder `preload` entfernen. |
| Niedrig | Offen | Schutz gegen volumetrische Angriffe liegt beim Router/NPM/Edge. Der Origin begrenzt Report-Abfragen auf 30 Anfragen pro Minute mit Burst 10 und maximal zwei gleichzeitige Proxy-Anfragen; das ist für eine Preview angemessen, ersetzt aber keinen Edge-DDoS-Schutz. |
| Niedrig | Akzeptiert | Session Storage begrenzt die Token-Lebensdauer auf den Tab, schützt aber nicht vor JavaScript-Ausführung im selben Origin. Eine strikte CSP, keine Drittanbieter-Skripte, eine enge Asset-Allowlist und sichere Linkbehandlung reduzieren dieses Restrisiko. |

## Umgesetzte Kontrollen

### Anwendung

- Tenant nur als konkrete UUID oder Domain; `common`, `organizations`, `consumers` und `adfs` werden abgewiesen.
- Client-ID nur als UUID.
- Berechtigungen werden je gewähltem Analysebereich angefordert; kein Client-Secret.
- Microsoft-Graph-URLs, Folge-Links und externe UI-/DOCX-Links verwenden Host-Allowlisting.
- Dynamische HTML- und Markdown-Inhalte werden escaped; DOCX-Hyperlinks akzeptieren nur vertrauenswürdige Microsoft-HTTPS-Ziele.
- Report-Downloads haben 120–135 Sekunden Timeout, 20 MB Nutzdatenlimit, 64 KB Fehlerlimit und 16 KB Headerlimit.
- Fehlerprotokolle enthalten ausschließlich Reportkennung und bereinigten Fehlercode, keine Tokens, Ziel-URLs oder Dateninhalte.
- Proxy-Request-/Response-Buffering und temporäre Proxydateien sind für Reportdaten deaktiviert.
- MSAL Browser ist exakt auf `5.21.0` festgelegt; `npm audit --omit=dev` meldete 0 bekannte Schwachstellen.

### Origin und LXC

- Debian 12 vollständig aktualisiert; automatische Sicherheitsupdates aktiv.
- Nginx 1.22.1, Python 3.11.2; keine zusätzliche Webframework-Laufzeit.
- Webroot enthält nur explizit freigegebene Anwendungsassets und benötigte MSAL-Module; Source-, Test-, Deployment-, Map- und Dotfiles werden nicht ausgeliefert.
- CSP mit gehashten Importmaps, `frame-ancestors 'none'`, `object-src 'none'`, `base-uri 'self'`, `form-action 'self'`, eingeschränkten Connect-/Frame-Zielen und `upgrade-insecure-requests`.
- `X-Frame-Options: DENY`, `X-Content-Type-Options: nosniff`, `Referrer-Policy: no-referrer`, restriktive Permissions Policy und `Cross-Origin-Resource-Policy: same-origin`.
- TLS 1.2/1.3 am Origin, Session Tickets deaktiviert, privater Schlüssel Modus 0600.
- Proxy nur auf `127.0.0.1:8080`; systemd-Exposure-Score 3,3 („OK“), eigener Benutzer `www-data`, keine Capabilities, `NoNewPrivileges`, `MemoryDenyWriteExecute`, Namespace-/Realtime-/SUID-/Syscall-/Adressfamilien- und Ressourcenbegrenzungen.
- LXC-Firewall: Input standardmäßig DROP, Loopback und bestehende Verbindungen erlaubt, HTTP/HTTPS nur vom NPM-Host; Forward DROP, Output ACCEPT.
- SSH und Postfix deaktiviert/maskiert. Offene Sockets: 80/443, intern 127.0.0.1:8080.
- Rollback-Snapshot vorhanden.

## Verifikation

Am 3. September 2026 erfolgreich ausgeführt:

- `npm test`: OAuth-, Analyse-, DOCX-, Security- und Proxy-Self-Checks bestanden.
- `npm audit --omit=dev`: 0 bekannte Schwachstellen.
- Neustart des LXC: Nginx, Proxy, nftables und unattended upgrades danach aktiv; SSH/Postfix weiter maskiert.
- Öffentliche Assets `/`, `/auth.html` und MSAL-Einstieg jeweils HTTP 200.
- `/.git/config`, `._app.js`, `package.json`, `server.py`, `security.test.mjs`, unbekannte Dateien und unbekannte API-Routen HTTP 403/404.
- GET auf Report-Route abgewiesen; POST ohne Token HTTP 401; manipuliertes `X-Real-IP` bereits am NPM mit HTTP 403 abgewiesen.
- Rate-Limit-Probe: 7 Antworten HTTP 401, danach 17 Antworten HTTP 429.
- Ungültiger syntaktisch langer Bearer-Token erreicht Microsoft Graph, ergibt HTTP 401 und nur `report=m365-apps error=graph-http-401` im Proxyjournal.
- TLS 1.0/1.1 extern abgewiesen; TLS 1.2/1.3 akzeptiert.
- Direkter Origin-Zugriff und TCP/22 aus dem Arbeitsplatznetz nach Reboot nicht erreichbar; öffentlicher NPM-Pfad weiter HTTP 200.

## Public-Launch-Gate

- [ ] NPM-Upstream auf HTTPS/443 umstellen und Live-HTTP-200 bestätigen.
- [ ] NPM muss `X-Real-IP` überschreiben; Authorization und OAuth-Querystrings dürfen nicht in Logs erscheinen.
- [ ] HSTS-Konfiguration korrigieren.
- [ ] OAuth-Popup, Admin-Consent und mindestens ein vollständiger Inventurlauf mit echtem Tenant erfolgreich.
- [ ] MD-, DOCX- und PDF-Export öffnen und stichprobenartig prüfen.
- [ ] Datenschutzerklärung um NPM-Logdaten/Aufbewahrungsfrist und die flüchtigen Report-Downloads ergänzen beziehungsweise bestätigen.
- [ ] Erst dann NPM-Access-List „homelab only“ entfernen; danach Origin-Port 80 schließen.

Nach Erfüllung dieser Punkte ist die Preview aus Sicht dieses Audits öffentlich freigabefähig. Für einen produktiven Dienst mit SLA oder Kundendaten mehrerer Organisationen sollten zusätzlich externes Penetration Testing, zentrales Monitoring/Alerting und ein dokumentierter Incident-/Patch-Prozess vorgesehen werden.
