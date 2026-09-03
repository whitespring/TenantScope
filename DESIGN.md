---
name: TenantScope
description: Ruhiger whitespring-Arbeitsraum für Microsoft-365-Governance-Prüfungen
colors:
  ink: "#111827"
  muted: "#66707c"
  line: "#dfe3e6"
  soft: "#f4f5f3"
  paper: "#ffffff"
  navy: "#102a3a"
  cyan: "#049bc4"
  cyan-dark: "#087896"
  action-yellow: "#f7c80d"
  error-red: "#c94444"
typography:
  display:
    fontFamily: "Exo 2, ui-sans-serif, sans-serif"
    fontSize: "clamp(42px, 4vw, 60px)"
    fontWeight: 700
    lineHeight: 0.98
    letterSpacing: "-0.005em"
  body:
    fontFamily: "ui-sans-serif, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif"
    fontSize: "16px"
    fontWeight: 400
    lineHeight: 1.55
  label:
    fontFamily: "ui-sans-serif, -apple-system, BlinkMacSystemFont, Segoe UI, sans-serif"
    fontSize: "14px"
    fontWeight: 700
    lineHeight: 1.2
rounded:
  field: "3px"
  control: "4px"
  action: "8px"
spacing:
  xs: "8px"
  sm: "16px"
  md: "24px"
  lg: "48px"
components:
  button-primary:
    backgroundColor: "{colors.action-yellow}"
    textColor: "{colors.ink}"
    rounded: "{rounded.control}"
    padding: "0 18px"
    height: "44px"
  input:
    backgroundColor: "{colors.paper}"
    textColor: "{colors.ink}"
    rounded: "{rounded.field}"
    padding: "0 14px"
    height: "50px"
---

# Design System: TenantScope

## Overview

**Creative North Star: "Der Vimeo-Arbeitsraum"**

TenantScope verbindet die ruhige, inhaltszentrierte Arbeitsfläche von Vimeo mit der whitespring-Identität. Das offizielle W, kräftige Exo-2-Titel und wenige präzise Cyan- und Gelbakzente geben Orientierung, während die komplexe Inventur auf einer offenen weißen Bühne arbeitet.

Die Oberfläche ist bewusst nicht als Sammlung umrandeter Karten gedacht. Tonale Flächen, feine Trennlinien und großzügiger Weißraum strukturieren Inhalte; Produktwahrheit und echte Microsoft-Graph-Zustände bleiben vollständig sichtbar.

**Key Characteristics:**

- Offene, ruhige Arbeitsfläche statt Dashboard-Kachelwand.
- whitespring-W und Exo 2 als klare Markenanker.
- Gelb nur für die primäre Aktion, Cyan für Navigation und Status.
- Tabellen, Befunde und Formulare bleiben scanbar und responsiv.

## Colors

Die Palette ist überwiegend neutral; die beiden Markenakzente werden sparsam und funktional eingesetzt.

**The Two-Accent Rule.** Gelb markiert die nächste Hauptaktion; Cyan markiert Orientierung, Links, Fortschritt und Auswahlzustände. Beide konkurrieren nie auf derselben Hierarchieebene.

## Typography

**Display Font:** Exo 2 mit systemischer Sans-Serif-Fallback-Kette
**Body Font:** System UI Sans Serif

**Character:** Exo 2 gibt den Seitentiteln eine eigenständige, technische whitespring-Stimme. Die Systemschrift hält dichte Inventur- und Tabellendaten unaufgeregt lesbar.

### Hierarchy

- **Display** (700, `clamp(42px, 4vw, 60px)`, 0.98): Seitentitel und zentrale Arbeitszustände.
- **Title** (700, 28–42px): Bereichs- und Berichtstitel.
- **Body** (400, 16px, 1.55): Erläuterungen, Hilfen und längere Befunde.
- **Label** (700, 14–18px): Felder, Navigation und kompakte Aktionen.

**The No-Eyebrow Rule.** Überschriften stehen ohne dekorative, wiederholende Kategorienzeile; Kontext kommt aus Navigation, Titel und Einleitung.

## Layout

Desktop nutzt einen Full-Width-Canvas mit einer leichten Kopfzeile. Die Verbindungssicht teilt die Fläche ungefähr 70/30 in Arbeitsbereich und tonale Hilfe. Ergebnisnavigation bleibt horizontal erreichbar. Unter 780px werden Prozessschritte scrollbar und die Hilfespalte fließt unter die Aufgabe; bei 390px darf kein horizontaler Seitenüberlauf entstehen.

## Elevation & Depth

Die Hauptoberfläche verwendet keine Schatten. Tiefe entsteht durch Papierweiß, helle tonale Flächen und einzelne Trennlinien. Nur modale Dialoge dürfen sich mit einem weichen Schatten klar von der Arbeitsfläche lösen.

**The Flat-By-Default Rule.** Laufende Inhalte liegen flach; Schatten sind ausschließlich für echte Überlagerungen reserviert.

## Shapes

Felder und kleine Kontrollen haben fast gerade Kanten (3–4px). Große Hauptaktionen dürfen leicht weicher sein (8px). Kreisformen bleiben Statuspunkten, Spinnern und Warnmarken vorbehalten.

## Components

### Buttons

- **Primary:** warmes Gelb, dunkle Schrift, kompakter Radius; immer genau eine visuell führende Aktion pro Arbeitsschritt.
- **Secondary:** weiß mit feiner neutraler Kontur.
- **Focus:** sichtbarer Cyan-Fokusring bleibt auf allen Varianten erhalten.

### Inputs / Fields

- Weißer Hintergrund, dünne neutrale Kontur und fast gerade Ecken.
- Placeholder erfüllen mindestens WCAG-AA-Kontrast; Fokus wechselt zu Cyan mit zurückhaltendem Ring.

### Navigation

Die aktive Route wird durch Textfarbe und eine schmale Cyan-Unterstreichung markiert. Deaktivierte Schritte bleiben sichtbar, aber deutlich zurückgenommen. Auf Mobilgeräten bleibt die Reihenfolge horizontal erhalten.

### Lists, Findings and Tables

Einträge werden durch horizontale Linien statt einzelne Kartenrahmen getrennt. Tabellen unterstützen Sortierung in beide Richtungen, Filter und Pagination; leere Werte bleiben bei absteigender Sortierung am Ende.

## Do's and Don'ts

### Do:

- **Do** echte Zustände, Fortschritt und den aktuellen Datensatz klar anzeigen.
- **Do** Detailtiefe über Typografie, Weißraum und Linien staffeln.
- **Do** ausschließlich das offizielle whitespring-W und die gemeinsame SVG-Icon-Sprache verwenden.

### Don't:

- **Don't** Inhalte in viele gleich gewichtete, umrandete Boxen zerlegen.
- **Don't** dekorative Eyebrows über selbsterklärende Überschriften setzen.
- **Don't** Gelb als allgemeine Dekorfarbe oder Cyan für unbedeutende Elemente verwenden.
