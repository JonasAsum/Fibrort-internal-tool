# Optional: small changes to the "Customer Visits in Detail" template (draft for Klaus, not sent)

*Draft by Claude for Jonas, 2026-10-09. Optional: the importer already reads the template as it is today.*

**Context (Klaus, 8 Oct):** the template should live in the CRM and be filled in completely there
("Direkte Eingabe immer"; "das Template soll hochgeladen werden und nicht die jeweiligen Reports").
So the long-term answer is a visit form in the CRM with the same fields. Until then, these Excel reports
are imported automatically.

---

**Vorschlag (DE, kurz):**

Hallo Klaus,

das Visits-in-Detail-Template lesen wir jetzt so ein, wie es ist. Damit nichts mehr nachgeprüft werden
muss (und als Vorlage für das CRM-Formular), würden fünf Kleinigkeiten helfen:

1. **Eine Zeile pro Besuch, mit eigener Spalte „Firma“.** Heute steht der Kunde im Freitext
   („visit to …“), und „Customer name“ ist manchmal eine Messe.
2. **Bei den Teilnehmern eine Spalte „Firma“** neben Name und Funktion, damit jeder Kontakt sicher der
   richtigen Firma zugeordnet wird.
3. **Datum als echtes Datum** (Excel-Datumsfeld statt „2026.9.15“ als Text), WV ebenso.
4. **Status nur als Status** (offen / erledigt / Angebot …). Notizen und Preise gehören in „Details“,
   nicht in Status oder in Spalten rechts neben der Tabelle.
5. **Ort/Stadt und Anlass** (z. B. Messe) als eigene Felder oben.

Das ist zugleich die Feldliste für das Pflichtformular im CRM: erst wenn alle Pflichtfelder gefüllt
sind, geht der Bericht ins System.

Viele Grüße
Jonas

---

**Same in English (for Jason / Betty):** one row per visit with its own Company column; a Company column
in the participant list; real Excel dates for Date and WV; Status used only for status (notes and
prices in Details); city and occasion (e.g. trade fair) as their own fields. These five fields become
the mandatory fields of the CRM visit form.
