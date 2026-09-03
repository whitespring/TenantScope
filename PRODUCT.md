# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack

Statische HTML-, CSS- und JavaScript-Anwendung mit einem kleinen Python-Server für gleich-originige Microsoft-Report-Downloads. Microsoft Graph wird über delegiertes OAuth mit PKCE angebunden.

## Users

Primäre Nutzer sind Microsoft-365-Administratoren, die ihren eigenen oder einen von ihnen ausdrücklich betreuten Tenant im Self-Service prüfen. Sie wählen den Prüfumfang, führen die Live-Inventur aus, bewerten Befunde und exportieren einen Bericht.

## Product Purpose

TenantScope macht den aktuellen Governance-, Sicherheits-, Nutzungs-, Speicher- und Lizenzstand eines Microsoft-365-Tenants aus echten Microsoft-Graph-Daten sichtbar. Erfolg bedeutet, dass ein berechtigter Administrator nachvollziehbare Befunde, konkrete nächste Schritte und einen exportierbaren Bericht erhält.

## Positioning

TenantScope führt die Prüfung transparent im Browser aus: Der Nutzer bestimmt Tenant, App-Registrierung und Prüfumfang. Die App verwendet keine Mock-Daten und kennzeichnet nicht auswertbare Teilabfragen, statt Ergebnisse zu erfinden.

## Operating Context

Die Anwendung läuft unter `https://tenantscope.wspg.org`. Der Nutzer benötigt eine passende Entra-App-Registrierung, delegierte Berechtigungen, Admin-Consent sowie die für einzelne Microsoft-Dienste erforderlichen Rollen und Lizenzen. Ergebnisse werden in der App geprüft und als PDF, Markdown oder DOCX exportiert.

## Capabilities and Constraints

- Die Inventur ist grundsätzlich lesend. Die gesondert gekennzeichnete Option zum Aktivieren von Klarnamen ändert tenantweit eine Microsoft-365-Reporteinstellung.
- OAuth-Token und Konfiguration liegen im Sitzungsspeicher des Browsers.
- Zwei Microsoft-Download-Reports werden flüchtig über den Serverstandort Deutschland abgerufen und dort weder gespeichert noch protokolliert.
- Microsoft-Reports können verzögert, anonymisiert oder aufgrund von Lizenz, Rolle oder Consent nicht verfügbar sein.
- Lizenz-Downgrades und Einsparungen sind Prüfhypothesen. Vertragspreise werden nur aus einer optionalen lokalen CSV ergänzt.
- Automatisierte Befunde ersetzen keine fachliche, rechtliche oder sicherheitstechnische Prüfung.

## Brand Commitments

Produktname ist TenantScope. Die Oberfläche orientiert sich an whitespring und spricht deutsch, direkt und sachlich. Verbindliche Markenquelle ist `/Users/thomas/Desktop/CODE/whitespring-ops/brand/`.

## Evidence on Hand

Die vorhandene Anwendung und ihre Selbsttests belegen die implementierten Prüfpfade. Es liegen keine freigegebenen Kundenstimmen, Benchmarks oder Wirksamkeitsnachweise vor; künftige Oberflächen dürfen sie nicht erfinden.

## Product Principles

- Echte Daten, klar gekennzeichnete Grenzen.
- Least Privilege und verständliche Berechtigungen.
- Befunde müssen erklärbar, überprüfbar und handlungsnah sein.
- Tenant-Daten bleiben so weit wie technisch möglich in der Browser-Sitzung.
- Änderungen im Tenant sind Ausnahmen und müssen vor Ausführung deutlich erkennbar sein.
