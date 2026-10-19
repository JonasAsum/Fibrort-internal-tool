# Margince: Deals, Angebote, Verträge, Rechnungen (einfach erklärt)

Quelle: Margince-Handbuch (`docs/handbook/`). Kürzel in Klammern = Dateiname ohne ".md". "nicht belegt" = Handbuch schweigt.

## 1. Was Margince ist
Margince ist ein CRM, also ein gemeinsamer Ort für Kontakte, Firmen, Leads, Deals und Projekte (what-margince-is). Es hat drei Besonderheiten: KI-Agenten arbeiten im System, aber nur mit den Rechten des Kollegen dahinter; es füllt sich aus angebundener Mailbox und Kalender; und es sagt offen "weiß ich nicht" statt zu raten (what-margince-is). Es kann gehostet, auf eigenen Servern oder komplett lokal laufen (what-margince-is). Beispiel: Eine Anfrage nach Spannzangen kommt per Mail, Margince ordnet sie der Firma zu, und Sie sehen den ganzen Verlauf an einer Stelle.

## 2. Deals
**Was ist ein Deal?** Eine Verkaufschance mit genau einer Firma, die auf "gewonnen" oder "verloren" zusteuert (leads-deals-and-projects, what-margince-is). "Opportunity" heißt in Margince "Deal" (what-margince-is). Beispiel: "Müller GmbH, 40 Spannfutter, 12.000 EUR".

**Abgrenzung**
- **Lead**: jemand, an den Sie vielleicht verkaufen; bewusst getrennt von Kontakten. Erst "Qualifizieren" macht daraus einen Kontakt, optional gleichzeitig einen Deal. Kalte Mail ohne Antwort reicht nicht (leads-deals-and-projects).
- **Projekt**: die Arbeit hinter dem Deal; entsteht schon während des Verkaufs und überlebt den Gewinn. Phasen: Initiative, Pursuing, Delivering, Closed (leads-deals-and-projects).
- **Angebot**: das bepreiste Dokument zu einem Deal (offers-and-products).
- **Vertrag**: eine von Hand erfasste Vereinbarung an einer Firma (contracts-and-invoices).

**Pipeline und Stufen.** Eine Pipeline ist eine Stufenleiter. Startpipeline "Sales": Qualified (10 %), Discovery (25 %), Proposal (50 %), Negotiation (75 %), Won (100 %), Lost (0 %). Namen und Reihenfolge sind änderbar, Won=100 und Lost=0 nicht (the-pipeline). Ein Deal bleibt in seiner Pipeline; wechseln geht nicht, nur neu anlegen und alt archivieren (leads-deals-and-projects).

**Open / Won / Lost.** Jede Stufe hat einen Typ: Open, Won oder Lost (the-pipeline). Es gibt keinen "Gewonnen"-Knopf; Sie schieben den Deal in die Won-Stufe (the-pipeline).
- *Verloren*: Ein **Lost reason** ist Pflicht, sonst bleibt "Confirm" gesperrt (the-pipeline).
- *Gewonnen*: Mit signiertem Vertrag schließt der Deal ohne Rückfrage. Ohne fragt das System **"How was it won?"**: Bestellung, mündlich/telefonisch, per Mail verlängert, aus anderem System importiert, oder "Other" mit Pflicht-Details (the-pipeline). Zweck: Gewinne ohne Papier werden auswertbar.
- Zusätzlich gibt es eine freiwillige **Outcome review** (Warum gewonnen/verloren) (the-pipeline). Geschlossene Deals lassen sich wieder öffnen (the-pipeline). Gewinnen/Verlieren geht nur einzeln, nicht in Masse (the-pipeline).

**Board vs. Tabelle.** Unter "Deals" sehen Sie die Liste als Tabelle oder als Board mit einer Spalte je Stufe; Karten ziehen Sie per Drag-and-drop, am Tablet per Stufenleiste (the-pipeline). Die Spaltenköpfe zeigen Anzahl, Summe und gewichtete Summe (the-pipeline).

**Wahrscheinlichkeit, gewichteter Wert, Abschlussdatum.** Die Wahrscheinlichkeit gehört zur *Stufe*, nicht zum einzelnen Deal (leads-deals-and-projects). Gewichteter Wert = Wert × Stufenwahrscheinlichkeit: 10.000 EUR in einer 50-%-Stufe ergibt 5.000 (the-pipeline). Für das Einzel-Urteil gibt es die **Forecast category** (Commit, Best case, Pipeline, Omitted) (the-pipeline). **Expected close** ist das erwartete Abschlussdatum; bei offenen Deals nicht in der Vergangenheit (leads-deals-and-projects). Die Prognose zählt nur "bestätigte" Daten als belastbar, "vorläufige" sind Schätzung (analytics).

**Auf der Deal-Seite**
- *Buying committee*: Ansprechpartner mit Rolle, z. B. Entscheider (leads-deals-and-projects).
- *Mails und Aktivitäten*: Mails senden, Anrufe/Meetings/Notizen mit "Log activity" erfassen; Aktivität setzt das "stalled"-Kennzeichen (über 60 Tage ohne Aktivität) zurück (the-pipeline).
- *Dateien*: Der Deal hat einen Files-Bereich inkl. Mail-Anhängen (deal-rooms).
- *Angebote*: Panel "Offers" (offers-and-products).
- *Deal Room*: Seite für den Kunden mit Dokumenten und Gespräch, ohne Login des Kunden (deal-rooms).
- *Partner*: Feld "via partner" plus "sourced" oder "influenced" (partners).
- *Projekt*: Deal einem Projekt derselben Firma zuordnen (leads-deals-and-projects).
- *Aufgaben direkt am Deal*: nicht belegt. Belegt ist nur: Aufgaben liegen am Projekt ("Open commitments") und in der Worklist auf Home (leads-deals-and-projects, what-margince-is).

## 3. Angebote
Ein Angebot gehört zu einem Deal und wird nach dem Versand nicht mehr geändert, sondern als neue Revision weitergeführt (offers-and-products).
- **Positionen**: Beschreibung, Einheit, Menge, Einzelpreis, Rabatt %, Steuer %, Summe; Netto, Steuer, Brutto werden exakt serverseitig gerechnet. Wiederkehrende Positionen (monatlich bis jährlich) mit Laufzeit; daraus "Annual recurring" und "Committed net" (offers-and-products).
- **Produkte/Preisliste ("Rate card")**: Settings → Products and offers. Beim Einfügen wird Preis kopiert; spätere Preisänderungen ändern alte Angebote nicht (offers-and-products).
- **Status/Aktionen**: Draft, dann **Send** (friert ein), dann **Accept**, **Reject** oder **Regenerate revision** (neue Entwurfs-Revision, alte "superseded") (offers-and-products). *Accept* übernimmt Betrag und Währung ins Deal (offers-and-products).
- **PDF**: "Render PDF" jederzeit; Vorlagen (de-DE/en-US) mit Kopf- und Fußtext (offers-and-products).
- Beispiel: Deal "Müller GmbH": Angebot mit 40 Spannfuttern à 300 EUR, 5 % Rabatt, 19 % Steuer, als PDF erzeugt und von Ihnen per Mail verschickt.

**Was es nicht gibt:** **Send verschickt nichts**; Sie schicken das PDF selbst (offers-and-products). Der Käufer kann nicht im System annehmen; Sie tragen die Antwort ein (offers-and-products). Leeres Angebot nicht sendbar; Versand ohne Wechselkurs verweigert (offers-and-products). Logo in Vorlagen wird nicht eingebettet; ein gelöschtes Template fällt nicht auf den Standard zurück (offers-and-products). Agenten dürfen Angebote entwerfen, aber nicht senden, annehmen, ablehnen oder als PDF erzeugen (offers-and-products). Eine Bestellbestätigung oder Lieferschein: nicht belegt.

## 4. Verträge
Ein Vertrag ist ein **von Hand eingetragener** Datensatz an einer Firma (Tab Documents → "Add contract"), mit Titel, Vertragsnummer und optional Deal/Projekt (contracts-and-invoices). Er wird nie automatisch aus einem Angebot erzeugt (nicht belegt als Funktion; das Handbuch beschreibt nur Handeingabe).
- **Felder**: Title, Value samt Basis ("Gesamtwert der Laufzeit" oder "12 Monate eines unbefristeten Vertrags"), Starts, Ends (leer = unbefristet), Renews, Notice period (Tage), Payment terms (Tage; 0 = sofort fällig), Signed, plus Anhang (contracts-and-invoices). Werte mit unterschiedlicher Basis werden nie addiert (contracts-and-invoices).
- **Lebenszyklus**: Draft, Active, Expired, Canceled, Superseded. Der Status wird **nie aus dem Datum errechnet**, ein Mensch setzt ihn; abgelaufene Laufzeit zeigt "Term ended, status pending". Expired/Canceled/Superseded sind endgültig (contracts-and-invoices).
- **Verlängern**: Neuer Vertrag (Entwurf), alter wird Superseded, seine Konditionen bleiben unverändert (contracts-and-invoices).
- **Kündigen**: erfasst nur Kündigungsdatum und Wirksamkeitsdatum; der Status ändert sich erst, wenn Sie ihn setzen (contracts-and-invoices). Die Warnung vor Verlängerung greift vor Ablauf der Kündigungsfrist (contracts-and-invoices).
- **Verknüpfung**: Firma immer; Deal und Projekt optional (contracts-and-invoices). Projektseite listet "Contracts" (leads-deals-and-projects). Einschränkung: Das Formular hat laut Pipeline-Seite kein Deal-Feld, die Zuordnung zum zu gewinnenden Deal ist umständlich (the-pipeline). Das Verlängern-Formular dagegen bietet eine Deal-Auswahl (contracts-and-invoices).
- **"Signierter Vertrag" für einen Gewinn** heißt: nicht archiviert, nicht Draft, Signed-Datum, plus Anhang unter Contract/Legal im Zustand Current/Final (contracts-and-invoices).
- Beispiel: Rahmenvertrag mit Müller GmbH über 24 Monate, Gesamtwert 60.000 EUR, Zahlungsziel 30 Tage, Kündigungsfrist 90 Tage.

**Kann nicht:** keine E-Signatur, kein Nachweis der Unterschrift (contracts-and-invoices); keine Statusänderung nach Datum (contracts-and-invoices); Agenten dürfen Verträge nicht einmal lesen (contracts-and-invoices). Vertragsvorlagen/Textbausteine: nicht belegt.

## 5. Rechnungen und Geld
Margince **erstellt keine Rechnungen**: kein Formular, kein "New invoice", keine Schnittstelle zum Erstellen (contracts-and-invoices). Rechnungen stammen nur aus einer Spiegelung des Buchhaltungssystems, Tab **Finance** (nur lesbar) (contracts-and-invoices). Gezeigt werden: netto fakturiert (12 Monate), offener Saldo, überfällig, typische Zahlungsverzögerung, letzte Rechnungen, Rechnungsstatus Draft bis Void (contracts-and-invoices). Ein Buchhaltungssystem lässt sich derzeit **nicht anbinden**; vorhanden ist nur ein Demo-Anbieter mit erzeugten Fantasie-Rechnungen (contracts-and-invoices). Fehlt ein Wert, wird er weggelassen statt als 0 gezeigt (contracts-and-invoices). Provisionen für Partner werden berechnet, aber Margince zahlt nichts aus (partners). Währung: EUR, USD, GBP, CHF; beim Abschluss wird der Kurs fixiert (leads-deals-and-projects, the-pipeline).

## 6. Vergleich
| | Zweck | Wer legt an | Beispiel (Zubehör für Werkzeugmaschinen) |
|---|---|---|---|
| **Deal** | Eine Verkaufschance bis Gewonnen/Verloren verfolgen | Vertriebsmitarbeiter, oder beim Qualifizieren eines Leads (leads-deals-and-projects) | "Müller GmbH: 40 Spannfutter, 12.000 EUR, Stufe Proposal" |
| **Angebot** | Das bepreiste Dokument für den Kunden | Mitarbeiter am Deal; Agent darf Entwurf (offers-and-products) | PDF mit 40 × 300 EUR, 19 % Steuer, gültig bis 30.11. |
| **Vertrag** | Unterzeichnete Vereinbarung festhalten (Laufzeit, Fristen) | Nur Mensch, von Hand (contracts-and-invoices) | Rahmenvertrag 24 Monate, Kündigungsfrist 90 Tage |
| **Projekt** | Arbeit hinter dem Deal bündeln, auch nach dem Gewinn | Benutzer; Pflichtfeld Firma (leads-deals-and-projects) | "Müller Werkstatt-Ausstattung", Phase Delivering; Folgedeal für Zusatzteile |

## 7. Was das System NICHT kann
- Rechnungen erstellen oder verschicken; Buchhaltung anbinden (contracts-and-invoices).
- Angebote per Mail versenden; Kunden Angebote online annehmen lassen (offers-and-products).
- E-Signatur (contracts-and-invoices).
- Vertragsstatus automatisch nach Datum setzen (contracts-and-invoices).
- Deals zusammenführen oder in andere Pipeline verschieben; Wahrscheinlichkeit je Einzeldeal setzen (the-pipeline, leads-deals-and-projects).
- Gewinnen/Verlieren in Masse (the-pipeline).
- Archivierte Deals aus der App zurückholen (the-pipeline).
- Deal Room per App archivieren; ein Deal hat praktisch nur einen Raum (deal-rooms).
- Geld bewegen, Provisionen auszahlen (partners).
- Agenten dürfen keine Verträge anfassen, Angebote nicht versenden, Deal Rooms nicht verwalten (contracts-and-invoices, offers-and-products, deal-rooms).
- Keine Zielvorgaben-Logik als "Deal value needed"; das ist ausdrücklich kein Ziel (analytics). Zielwerte für Verkäufer gibt es dagegen unter "Targets" (sales-reporting).
- *Inferenz (nicht ausdrücklich genannt):* kein Lager-/Bestell-/Lieferscheinwesen und keine Maschinenkompatibilität; das Handbuch kennt nur Produkte als Preislistenzeilen (offers-and-products). Rechnungen aus Angeboten zu erzeugen ist folglich ebenfalls nicht möglich.
