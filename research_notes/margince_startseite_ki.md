# Margince: Die Startseite mit KI, einfach erklärt

Quelle: Handbuch unter `margince/docs/handbook/`. Hinter jeder wichtigen Aussage steht in Klammern die Datei. Was das Handbuch nicht sagt, ist als "nicht belegt" markiert.

Vorab: Ein CRM ist eine Kundendatenbank mit Verkaufsverlauf. Ein "Deal" ist ein einzelnes Verkaufsvorhaben (z. B. "Firma X kauft 20 Lizenzen"). Ein "Kontakt" ist eine Person, eine "Company" ihre Firma.

## 1. Die Startseite (Home), Element für Element

**Home** ist der Tagesstart. Sie hat zwei Ansichten (**Morning**, **Weekly**) und zwei Bereiche (**Mine**, **Team**). *(your-day-and-your-week.md)*

- **Morning / Weekly**: Morning ist die Tagesansicht. Weekly zeigt den Rückblick auf die letzte abgeschlossene Woche und Ihren Wochenplan ("This week's commitments"). Der Rückblick wird montags geschrieben und dann eingefroren, er wird nicht neu berechnet. *(your-day-and-your-week.md)*
- **Mine / Team**: Mine zeigt nur Ihre Arbeit. Team erscheint nur, wenn Ihre Rolle über die eigenen Datensätze hinausreicht, etwa bei einer Teamleitung. *(your-day-and-your-week.md)*
- **Show Worklist / Hide Worklist**: Die Worklist ist Ihre nach Dringlichkeit sortierte To-do-Liste. Sie öffnet sich als Fenster über der Startseite. Eine eigene Seite dafür gibt es nicht. *(your-day-and-your-week.md)* Die Zahl auf dem Knopf ist die Zahl des Tages. *(your-day-and-your-week.md)*
- **"Open full Worklist"**: nicht belegt. Das Handbuch kennt nur "Show Worklist". Vermutlich ist es derselbe Weg, das lässt sich aus dem Handbuch nicht beweisen.
- **Die Liste "In the Worklist"**: nicht belegt. Belegt sind die drei Bänder **Now** (dringend), **Keep momentum** (zugesagte Arbeit, die schleift) und **Prospecting** (Neukunden-Arbeit). Prüfungsaufgaben ("To review") stehen in einem eigenen Bereich darunter. *(your-day-and-your-week.md)*
- **Reihenfolge**: Der Server sortiert, der Bildschirm sortiert nie um. Jede Zeile erklärt in Sätzen, warum sie dort steht und was passiert, wenn man sie liegen lässt ("If ignored, the buyer keeps waiting."). *(your-day-and-your-week.md)*
- **Overdue / Due**: Aufgaben sind nach Fälligkeit gruppiert. Überfällige und heutige stehen zuerst, danach "Due tomorrow", "Due this week", "Later". *(your-day-and-your-week.md)* Ein eigenes rotes "Overdue"-Etikett beschreibt das Handbuch nicht, die Gruppierung aber schon.
- **Done**: Erledigt die Aufgabe ("Task completed"). Alternativ "Snooze 1 day", das die Fälligkeit um einen Tag verschiebt. *(your-day-and-your-week.md)* Beispiel: Aufgabe "Angebot an Firma X nachfassen", fällig gestern. Sie klicken Done, und sie verschwindet.
- **Snooze / Pin / Remove from list**: Snooze holt eine Zeile morgen zurück. Pin hält sie oben (höchstens 50). "Not mine" entfernt sie nur bei Ihnen. "Not a customer" entfernt den ganzen Mailverlauf für alle. Jedes hat Undo. *(your-day-and-your-week.md)*
- **Focus-Karte, Label "AI-assisted", "1 of 6", "View details"**: nicht belegt. Das Wort "Focus" kommt im Handbuch nicht als Home-Element vor. Das Label "AI-assisted draft" gibt es nur bei E-Mail-Entwürfen ("This draft was written by AI. Review and edit it before sending."). *(sending-mail.md)* Wahrscheinlich ist die Focus-Karte eine Ansicht auf die Worklist. Das ist eine Vermutung, kein Beleg.
- **"Nothing needs attention"**: nicht belegt. Ähnlich ist der Leerzustand "Nothing is waiting on you." Er kann täuschen: Konnte eine Quelle nicht gelesen werden, steht dort stattdessen "No items in the sources that loaded." *(your-day-and-your-week.md)*

## 2. Was die KI in Margince tut

**Automatisch, ohne Rückfrage** (Stufe "auto-execute"): nachschlagen, suchen, zusammenfassen, Entwürfe schreiben, normale Feldänderungen, Aktivitäten protokollieren, Leads befördern, Datensätze archivieren. *(agents-and-passports.md)*

Konkret:
- **E-Mail-Entwürfe** ("Draft with AI"). Die KI schreibt, nur Sie drücken Send. Sie fügt nie eine Grußformel an. *(what-the-ai-does.md)*
- **Dokumente lesen** für Deal-Felder (Name, Betrag, Währung, Abschlussdatum), jeweils mit der Textstelle als Beleg. Nichts wird gespeichert, bis Sie "Accept" drücken. *(what-the-ai-does.md)*
- **Morning brief über Nacht**: Rangliste der Konten, die Ihre erste Stunde verdienen. Läuft nur, wenn Sie es unter Settings → Connections eingeschaltet haben. Sie darf lesen und schreiben, aber nie senden. *(what-the-ai-does.md, agents-and-passports.md)*
- **Kontakte erfassen** aus Ihrer Mail: Das ist "Capture". Es legt Mails an der Kontakt-Timeline ab. Ein Modell wird nur gefragt, was für ein Absender das ist (Kontakt, Newsletter, Spam ...). *(capture.md)*
- **Wochenrückblick**: Beobachtungen und Zusammenfassung stammen von der KI. *(your-day-and-your-week.md)*

**Automatische Änderungen** an nur drei Dingen: Abschlussdatum, Firmenname, Lebenszyklus-Stufe. Standardmäßig an, pro Person abschaltbar unter Settings → Agents → Automatic changes. Sie erscheinen unter "Handled for you" mit **Undo** (Undo-Knopf nur bei Deals; sonst über die Historie des Datensatzes). *(approvals.md, your-day-and-your-week.md)*

**Erst nach menschlicher Bestätigung** (Stufe "confirm-first", als "Approval" in der Worklist, Knopf **Decide**, dann Accept, Approve edited oder Reject): Anreicherung aus dem Web, Firmenwebsite lesen, eigene Felder anlegen oder ändern, Webhooks, Tags zusammenführen, Deal gewinnen/verlieren/wieder öffnen, E-Mail unter einem Projekt ablegen, und jedes Feld, das zuletzt ein Mensch getippt hat. *(agents-and-passports.md, approvals.md)* Ein Approval zeigt: den genauen Vorschlag, den Beleg, die Sicherheit (hoch/mittel/niedrig) und wer ihn vorschlug. Er verfällt nach 72 Stunden, dann muss neu vorgeschlagen werden. *(approvals.md)*

Beispiel: Die KI will in einem Deal das Abschlussdatum ändern, das Sie selbst eingetragen haben. Sie ändert es nicht, sie legt eine Freigabe-Karte an. *(agents-and-passports.md)*

Wer hat's getan? Jeder Wert trägt eine Herkunftsmarke, z. B. "Typed by you", "Automated by {agent}", "Via {connector}". *(what-the-ai-does.md)* So erklärt sich "Automated by ..." auf den Karten.

## 3. Woher die Daten kommen, Anbieter, ohne KI-Modell, "Summary failed"

**Datenquellen**: Margince liest nur Postfächer und Kalender, die Sie selbst verbunden haben (Gmail, Outlook, IMAP, Google/Outlook Calendar; Settings → Connections). Kalender werden getrennt von Mail verbunden. Ein Agent kann kein Postfach verbinden. *(connecting-mail-and-calendars.md)* Nie gespeichert werden: Spam, Papierkorb, Entwürfe, Mails nur zwischen Kollegen. *(capture.md)* Zusätzlich: Dokumente, die Sie ablegen, und Firmenwebsites. *(what-the-ai-does.md, relationships-and-research.md)*

**KI-Anbieter** (Settings → AI models → Providers): Anthropic, OpenAI, Google Gemini, Gemini auf Vertex AI, OpenAI-kompatible Dienste (OpenRouter, Mistral, Together, Groq, DeepSeek oder eigener Server) und ein "decision model" (TypeSafe Jev). Lokale Modelle (Ollama, vLLM) brauchen keinen Schlüssel, ihre Adresse legt der Betreiber fest. Nur Admin oder Ops dürfen ändern. Für EU-Verarbeitung gibt es "OpenRouter (EU)" und Vertex-Standort "eu". *(ai-providers.md)*

**Ohne KI-Modell**: Capture läuft normal, nur die Absender-Einstufung entfällt. *(capture.md)* "Draft with AI" ist nicht verfügbar, man schreibt selbst. *(sending-mail.md)* Der Wochenrückblick kommt vollständig, aber ohne Beobachtungen und Textzusammenfassung. *(your-day-and-your-week.md)* Ohne Crawler: "Website research is not configured on this server." *(relationships-and-research.md)*
**Fake-KI in der Demo**: nicht belegt. Das Handbuch erwähnt keine Schein-KI. Es nennt nur einen Offline-Demo-Anbieter für Finanzdaten (contracts-and-invoices.md, nur per Suche gesehen, nicht vollständig gelesen).

**Warum "Summary of <Firma> failed" erscheinen kann**: Der genaue Text ist nicht belegt. Belegt ist, dass "summarising" eine von zwölf sichtbaren KI-Arbeiten ist und fünf Zustände haben kann: queued, running, done, degraded, failed. Dazu kommt "stalled" (unüblich lange, vielleicht gestoppt) und "deferred" (das Monats-KI-Kontingent ist aufgebraucht; keine Störung, die Arbeit wird später nachgeholt). *(what-the-ai-does.md, settings.md)* Plausible Ursachen laut Handbuch: kein Modell/Schlüssel, Anbieter lehnt den Schlüssel ab (Test-Knopf prüft das), Kontingent erschöpft. Ob eine davon in Ihrem System zutrifft, zeigt der **AI call log** (Settings → AI): "each model request and its response". Fehlgeschlagene Jobs stehen zusätzlich in "Background jobs" auf System health. *(settings.md)*

## 4. Deal-Vorschlag ("Automated by ...") von der ersten Spur zum echten Deal

Die Karten "Open this deal" / "Not a deal" sind im Handbuch **nicht belegt**. Nur diese Bausteine sind belegt:

1. Eine Mail kommt über Capture herein und wird an Kontakt und Firma abgelegt. **Capture verknüpft eine Mail nie von selbst mit einem Deal**, das geht nur per Relink von Hand. *(capture.md)* Capture legt auch nie eine Firma an. *(capture.md)*
2. Unklare Absender werden zu einer Karte "Add contact from mail". Accept legt den Kontakt an, Reject tut nichts. *(capture.md)*
3. Die Worklist kennt "deal suggestions". Ein per **Dismiss** weggewischter Vorschlag bleibt weg, bis eine neue Aktivität am Deal verknüpft wird. Dann kommt er zurück und sagt warum. *(the-pipeline.md, what-the-ai-does.md)*
4. Anlegen eines Deals ist ein normaler Schritt mit Wert, Währung, Firma. *(leads-deals-and-projects.md, nur im README aufgeführt)*

Wie genau die KI aus einer Mail einen Vorschlag baut und was "Open this deal" auslöst, steht nicht im Handbuch. Wahrscheinlich ist "Open this deal" ein Anlegen-Vorschlag und "Not a deal" eine Ablehnung, aber das ist Vermutung. Bleibt das System bei den belegten Regeln, wird nichts ohne Ihren Klick angelegt.

Beispiel (nur das Belegte): Frau Meier von Firma X schreibt. Capture legt die Mail unter ihrem Kontakt ab. Daraus wird nicht automatisch ein Deal.

## 5. Was die KI NICHT darf / kann

- Nicht mehr als die Person hinter ihr. Sie wird bei jedem Aufruf gegen deren Rechte geprüft. Was Sie nicht sehen, sieht sie nicht. *(agents-and-passports.md)*
- Ihren eigenen Vorschlag nicht freigeben und keinen für jemand anderen. Nur ablehnen darf sie ihn. *(agents-and-passports.md)*
- Ihr Kontingent nicht selbst erhöhen ("step-up"). *(agents-and-passports.md)*
- Einwilligungen (Consent) oder Datenschutzanfragen nicht setzen oder entscheiden. *(agents-and-passports.md)*
- Keine Dokumentensammlung befragen oder definieren. *(agents-and-passports.md, what-the-ai-does.md)*
- Pipeline-Stufen nicht ändern. *(agents-and-passports.md)*
- Keine Einführungsanfrage stellen, beantworten oder zurückziehen. Sie darf nur lesen, wer wen kennt. *(relationships-and-research.md)*
- Kein Postfach verbinden. *(connecting-mail-and-calendars.md)*
- Nicht senden: die App-KI schreibt nur Entwürfe, Senden macht der Mensch. Die Nacht-KI kann nie senden. Ein von Ihnen verbundener externer Agent mit Recht "Send messages" kann dagegen sofort senden, wenn die Installation keine Bestätigungsstufe vorschaltet. Jeder Versand braucht Einwilligung für den Zweck. *(what-the-ai-does.md, agents-and-passports.md)*
- Keine eigenen Rechte vergeben ("An agent that could grant itself standing authority would be deciding its own rights"). *(agents-and-passports.md)*
- Nicht selbst rückgängig machen: Nur ein Mensch kann Undo drücken. *(approvals.md)*
- Nichts erfinden: Jede Angabe trägt ihren Beleg; was die Quelle nicht klar sagt, wird weggelassen. *(what-the-ai-does.md, relationships-and-research.md)*
- Beim Website-Lesen: nur die eigene Seite der Firma, keine Fremddomains, keine Personen ohne Datenschutzhinweis. *(relationships-and-research.md)*
- Eine neue, nicht eingestufte Schreibfunktion ist für Agenten gesperrt. *(agents-and-passports.md)*
- Namen, Mailadressen und Telefonnummern werden vor dem Modell nicht entfernt, nur Geheimnisse wie Schlüssel. Anhänge werden unverändert mitgesendet. *(agents-and-passports.md)*

## Offene Punkte (nicht belegt)
Focus-Karte, "AI-assisted"-Etikett an Aufgaben, "1 of 6", "Open full Worklist", "In the Worklist", "Open this deal"/"Not a deal", genauer Text "Summary ... failed", "Nothing needs attention", Fake-KI in der Demo. Diese Punkte müssten im Code oder im laufenden System geprüft werden.
