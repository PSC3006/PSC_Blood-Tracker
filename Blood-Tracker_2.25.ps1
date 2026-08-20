# PowerShell Script für Blutwert-Tracking mit GUI
# Version 2.26.0 (20. August 2026)
# Zweck: Grafische Analyse, Erfassung und proaktive Überwachung von Blutwerten.
#
#
# CHANGELOG v2.26.0 (Marker-Verwaltung, 3 neue Marker, Report-Seitenumbruch):
# - FEATURE (Aufgabe #1): Tab "Einzelmarker-Analyse" > GroupBox "Blutmarker
#   verwalten" > Button "Zum Bearbeiten laden" aktiviert jetzt automatisch die
#   Checkbox "Erweiterte Einstellungen". Bisher wurden Gruppe, Ref-Min/Max und
#   Optimal-Min/Max zwar in die Felder geladen, waren aber unsichtbar - der
#   Nutzer musste die Checkbox erst manuell setzen, um sie zu sehen.
#   Ergaenzt: Hinweis-Dialog, wenn kein Marker gewaehlt ist, Fallback auf den
#   eingetippten Text und Fokus-Sprung in das Feld "Name".
# - FEATURE (Aufgabe #2): Neuer Marker "Apolipoprotein E-Genotyp (APOE)"
#   (Gruppe "Genetik", neue Gruppe). Da der Datenspeicher ausschliesslich Zahlen
#   haelt, wird der Genotyp als Code 1-6 gespeichert, aufsteigend nach dem
#   Alzheimer-Risiko: 1 = E2/E2, 2 = E2/E3, 3 = E3/E3, 4 = E2/E4, 5 = E3/E4,
#   6 = E4/E4. Erfassung ueber eine Auswahlliste (analog HIV), Anzeige ueberall
#   im Klartext (Cockpit, Chart-Tooltip, Tab "Daten nach Bluttests", Ausdruck,
#   Custom Report). Referenz 1-3 = kein E4-Allel; Optimal 2-3 (E2/E3, E3/E3).
#   E2/E2 liegt bewusst NICHT im Optimalbereich - Homozygotie fuer E2 ist der
#   typische Genotyp der Hyperlipoproteinaemie Typ III.
#   Quellen: MedlinePlus Genetics "APOE gene" (NIH/NLM); Farrer et al., JAMA
#   1997 (APOE and Alzheimer Disease Meta Analysis Consortium); Alzheimer's
#   Drug Discovery Foundation, Cognitive Vitality.
# - FEATURE (Aufgabe #2): Neuer Marker "Zink" (ug/dl, Gruppe Mineralstoffe).
#   Referenz 80-120 ug/dl (12-18 umol/l) nach NIH Office of Dietary Supplements
#   (Zinc - Health Professional Fact Sheet); unzureichender Status < 74 ug/dl
#   (Mann) bzw. < 70 ug/dl (Frau). Bewusst KEIN Optimalbereich - dafuer gibt es
#   keine offizielle Vorgabe.
# - FEATURE (Aufgabe #2): Neuer Marker "Magnesium" (mmol/l, Gruppe
#   Mineralstoffe). Referenz 0,75-0,95 mmol/l, Hypomagnesiaemie < 0,75 mmol/l
#   (NIH ODS, Magnesium - Health Professional Fact Sheet). Optimal 0,85-0,95
#   mmol/l nach Costello et al., Advances in Nutrition 2016 (PMC5105038).
#   HINWEIS in der Description: nur ca. 0,3 % des Koerpermagnesiums liegen im
#   Serum - ein normaler Wert schliesst einen Mangel NICHT aus.
# - HINWEIS: "Zn" und "Mg" stehen bewusst NICHT im Marker-Namen (keine
#   Klammer-Kuerzel), sondern nur in $script:MarkerAliases. Grund:
#   Build-MarkerAliasMap zieht Klammer-Kuerzel automatisch in den PDF-Import;
#   2-stellige Kuerzel erzeugen dort massenhaft Fehltreffer.
# - BUGFIX: $script:PdfExcludedMarkers - der PDF-Import nimmt qualitative Marker
#   jetzt komplett aus. Betroffen war vor allem HIV: Build-MarkerAliasMap
#   zerlegt "(Anti-HIV-1/2)" an "/" und erzeugte daraus den Alias "2" - jede
#   Zeile der Form "2 <Zahl>" konnte einen HIV-Wert (= "reaktiv") erzeugen.
#   APOE ist aus demselben Grund ausgenommen (Genotyp, keine Messgroesse).
# - FEATURE (Aufgabe #3): Tab "Custom Report" > PDF-Export - jeder Blutmarker
#   beginnt jetzt zwingend auf einer NEUEN Seite (erzwungener Seitenumbruch am
#   Section-Ende). Zusaetzlich wird das Papierformat im PDF-Export fest auf
#   DIN A4 gesetzt (Querformat, da die Verlaufsgrafik die Seitenbreite nutzt),
#   sofern der PDF-Drucker A4 anbietet.
# - BUGFIX: Die Fusszeile ("Seite n - Blood-Tracker Custom Report") wurde nur
#   auf der LETZTEN Seite gedruckt, weil alle Seitenumbruch-Pfade vorher aus dem
#   Handler zurueckkehrten. Sie erscheint jetzt auf jeder Seite.
# - HINWEIS: Referenzbereiche sind methoden- und laborabhaengig. Massgeblich ist
#   immer der Referenzbereich des eigenen Befundes; keine aerztliche Beratung.
# CHANGELOG v2.25.0 (Neue Blutmarker: hs-CRP + SHBG):
# - FEATURE: Neuer Marker "Hochsensitives CRP (hs-CRP)" (mg/l, Gruppe Entzündung).
#   Referenz 0-3 mg/l, Optimalbereich 0-1 mg/l. Grundlage: AHA/CDC-Statement
#   (Pearson et al., Circulation 2003) - kardiovaskuläre Risikostratifizierung
#   < 1 mg/l = niedriges, 1-3 mg/l = mittleres, > 3 mg/l = hohes Risiko;
#   bestätigt in StatPearls "C-Reactive Protein" (NCBI Bookshelf NBK441843,
#   Stand 03.05.2025). Werte > 10 mg/l sprechen für eine akute Entzündung
#   (Kontrolle nach ca. 2 Wochen), nicht für ein kardiovaskuläres Risiko.
# - FEATURE: Neuer Marker "Sexualhormon-bindendes Globulin (SHBG)" (nmol/l,
#   Gruppe Hormone). Referenz 18,3-54,1 nmol/l (Mann, 20-49 Jahre).
#   Quelle: Roche Elecsys SHBG (ECLIA), Herstellerbeipackzettel, publiziert im
#   Leistungsverzeichnis Uniklinik Bonn. Weitere Bereiche in der Description
#   hinterlegt (Mann >= 50 J: 20,6-76,7; Frau 20-49 J: 32,4-128;
#   Frau >= 50 J: 27,1-128 nmol/l). Bewusst KEIN Optimalbereich gesetzt -
#   dafür existiert keine offizielle Vorgabe (niedrige UND hohe Werte sind
#   klinisch relevant).
# - BUGFIX (Voraussetzung): Load-Config überschrieb Config.Markers vollständig
#   mit der gespeicherten Config.json. Neue Default-Marker wären in einer
#   bestehenden Installation NIE aufgetaucht. Neu: versionsgesteuerte
#   Marker-Migration ($script:MarkerSetVersion + $script:MarkerSetAdditions).
#   Ergänzt nur die Marker der jeweils neuen Version, überschreibt keine
#   angepassten Marker und stellt bewusst gelöschte Marker nicht wieder her.
#   Die erreichte Marker-Set-Version wird in der Config.json persistiert.
# - FEATURE: Alias-Auflösung getrennt - "hs-CRP", "hsCRP", "CRP hs",
#   "hochsensitives CRP" zeigen jetzt auf den NEUEN Marker; "CRP" und
#   "C-reaktives Protein" weiterhin auf den klassischen Marker.
#   (Bis v2.24.1 lagen die hs-Aliase auf dem klassischen CRP - beide Marker
#   parallel hätten sonst eine mehrdeutige Auflösung erzeugt.)
# - BUGFIX: PDF-Import ordnete die Zeile "hs-CRP 0,8 mg/l" dem klassischen CRP
#   zu (Regex "(?<!\w)CRP(?!\w)" greift auch nach einem Bindestrich). Neu:
#   $script:PdfSkipLinePatterns - Zeilen mit hs-/hochsensitiv-Bezug werden beim
#   klassischen CRP übersprungen.
# - FEATURE: $script:MarkerFallbacks - fehlt bei einer Score-Berechnung der
#   klassische CRP-Wert, wird automatisch hs-CRP verwendet (gleiche Einheit
#   mg/l). Betrifft InflammAging-Score, PhenoAge, PhenoAge-Accel und
#   Longevity-Score, die fachlich ohnehin auf hsCRP basieren.
# - FEATURE: Tab "Longevity": hs-CRP als eigene Zeile inkl. Score-Bewertung.
# - FEATURE: hs-CRP zusätzlich in den genetischen Vorbelastungen "Chronische
#   Veneninsuffizienz" und "Familiäre Hypertonie / KHK-Prädisposition"
#   hinterlegt (wirkt auch in bestehenden Installationen, da BuiltIn-Marker
#   aus den Defaults kommen und nur das Active-Flag gespeichert wird).
# - HINWEIS: Referenzbereiche sind methoden- und laborabhängig. Maßgeblich ist
#   immer der Referenzbereich des eigenen Befundes; keine ärztliche Beratung.
#
# CHANGELOG v2.24.1 (BUGFIX: Dokument-Zuordnung pro Bluttest):
# - BUGFIX (kritisch): Tab "Daten nach Bluttests" > Button "Dokument anzeigen"
#   zeigte die Anzahl ALLER Dokumente des jeweiligen MONATS an und oeffnete das
#   zuletzt geaenderte Dokument des Monats - nicht das des gewaehlten Bluttests.
#   Ursache: Die Tagesdaten liegen physisch in einem Monatsordner
#   (data\YYYY\MM\); gefiltert wurde mit "*_BLUTTEST*", also ohne Datumsbezug,
#   die Auswahl erfolgte anschliessend per "Sort-Object LastWriteTime -Descending".
#   Beispiel: Auswahl 06.08.2026 -> geoeffnet wurde das Dokument vom 17.08.2026.
#   FIX: Neue Funktion Get-BloodTestDocuments filtert strikt per RegEx auf den
#   Datums-Prefix des gewaehlten Tests: "YYYY-MM-DD_BLUTTEST[_n][.ext]" und
#   sortiert nach laufender Nummer (= Upload-Reihenfolge).
# - FEATURE: Sind mehrere Dokumente zu EINEM Bluttest hinterlegt, oeffnet der
#   Button ein Auswahlmenue (Dateiname, Groesse, Zeitstempel) statt blind das
#   neueste zu starten. Zusaetzlicher Eintrag "Ordner oeffnen".
#   Bei genau einem Dokument wird dieses weiterhin direkt geoeffnet.
# - FEATURE: "Dokument hochladen" unterstuetzt jetzt Mehrfachauswahl
#   ($ofd.Multiselect = $true); alle gewaehlten Dateien werden dem AKTUELL
#   markierten Bluttest-Datum zugeordnet und kollisionsfrei durchnummeriert.
# - FEATURE: Tooltip am Button listet alle Dokumente des gewaehlten Bluttests.
# - REFACTORING: Die Button-Logik liegt zentral im Scriptblock
#   $UpdateShowDocButton (bisher dreifach dupliziert in AfterSelect, Upload und
#   Click). Wird nach Auswahl, Upload und PDF-Import aufgerufen - der Zaehler
#   bleibt dadurch nach jedem Upload korrekt (bisher wurde er auf
#   "Dokument anzeigen" ohne Anzahl zurueckgesetzt).
# - HINWEIS: Keine Migration noetig - der Dateiname enthielt das Datum bereits,
#   fehlerhaft war ausschliesslich die Filterung/Anzeige.
# - HINWEIS: Die Kopfzeile wies bisher "Version 2.19.0" aus, $Version stand auf
#   "2.23.0" bei Dateiname _2.24. Beides ist jetzt auf 2.24.1 vereinheitlicht.
#
# CHANGELOG v2.23.0 (Alias-Suche im Blutmarker-Dropdown):
# - FEATURE: Tab "Einzelmarker-Analyse" > GroupBox "Wert eintragen" > Dropdown
#   "Blutmarker" akzeptiert jetzt zusaetzlich zur Listenauswahl eine freie
#   Direkteingabe. Labor-Synonyme werden automatisch auf den kanonischen
#   Marker aufgeloest - z. B. "WBC" ODER "Leukozyten" -> "Leukozyten (WBC)",
#   "GOT"/"AST"/"ASAT" -> "Aspartat-Aminotransferase (AST/GOT)",
#   "hs-CRP" -> "C-reaktives Protein (CRP)".
# - FEATURE: Live-Filter waehrend des Tippens; die Dropdown-Liste zeigt nur noch
#   passende Marker (Sortierung: exakter Treffer > Wortanfang > enthalten).
#   Enter uebernimmt, Esc setzt zurueck, unbekannte Eingabe faerbt das Feld rot.
# - FEATURE: Neue Info-Zeile rechts neben dem Wert-Feld zeigt Einheit und
#   Referenzbereich des erkannten Markers (Schutz vor Einheiten-Verwechslung).
# - NEU: $script:MarkerAliases (65 Marker) plus Get-MarkerToken,
#   Build-MarkerSearchIndex, Get-MarkerMatches, Resolve-MarkerName.
#   Die Tabelle liegt bewusst AUSSERHALB von Config.Markers, da Load-Config
#   eine gespeicherte Config.json vollstaendig ueber die Defaults legt.
# - BUGFIX: Der "Hinzufuegen"-Button las bisher $markerComboBox.SelectedItem.
#   Da die ComboBox editierbar ist, konnte SelectedItem auf einem frueheren
#   Marker stehen bleiben, waehrend im Textfeld bereits ein anderer stand -
#   der Wert waere unter dem falschen Marker gespeichert worden. Es wird nun
#   immer der aufgeloeste Name verwendet; mehrdeutige oder unbekannte Eingaben
#   erzeugen eine klare Fehlermeldung statt eines Fehleintrags.
# - HINWEIS: BUN, PCT und TG sind bewusst NICHT gemappt (BUN x 2,14 = Harnstoff;
#   PCT = Plateletcrit vs. Procalcitonin; TG = Triglyceride vs. Thyreoglobulin).
#
# CHANGELOG v2.22.1 (Feinschliff Dokument-Button + Annotation-Fix):
# - LAYOUT: Button jetzt quadratisch (60x60), Hoehe verdoppelt, Text 2-zeilig und
#   zentriert, weiter nach rechts platziert.
# - BUGFIX: Annotation "letzter Wert" warf "[...Charting.ContentAlignment] wurde
#   nicht gefunden". AnchorAlignment wird nun per String ('BottomCenter') gesetzt,
#   ohne den (je nach .NET-Version nicht aufloesbaren) Enum-Typ zu referenzieren.
#
# CHANGELOG v2.22.0 (Dokument-Anzeige + Chart-Letztwert):
# - FEATURE: Tab "Daten nach Bluttests": Ist fuer den gewaehlten Test ein Dokument
#   (*_BLUTTEST*) archiviert, erscheint rechts von "PDF importieren (Beta)" ein
#   grosser Button "Dokument anzeigen" (giftgruener, fetter Text). Oeffnet das
#   Dokument mit der Standard-App. Button nur sichtbar, wenn Dokument vorhanden;
#   erscheint auch direkt nach einem Upload.
# - FEATURE: Tab "Einzelmarker-Analyse": Der Graph zeigt zusaetzlich den letzten
#   Wert samt Datum - als Untertitel und als Annotation am letzten Datenpunkt.
#
# CHANGELOG v2.21.2 (Hotfix Join-Path):
# - BUGFIX: 3-Argument-Join-Path (PS 7+) brach unter Windows PowerShell 5.1 mit
#   "kein Positionsparameter ... akzeptiert" ab. Betraf den unverschluesselten
#   Export UND den AES-.btbackup-Export. Nun via verschachteltem Join-Path
#   (5.1-kompatibel). 4 Fundstellen korrigiert.
#
# CHANGELOG v2.21.1 (Unverschlüsselter Migrations-Export):
# - FEATURE: Neuer Button "Unverschlüsselt exportieren (Migration)" in Tab "Daten
#   nach Bluttests". Erstellt ein KLARTEXT-ZIP (Config.json + data\*.json + Dokumente)
#   zur einfachen Migration auf einen anderen PC/Windows-Benutzer. Umgeht das
#   DPAPI-Problem (CurrentUser-Scope), das den direkten Kopiervorgang verhindert.
#   Auf dem Ziel-PC werden die Klartext-Dateien automatisch erkannt und beim
#   nächsten Speichern lokal neu verschlüsselt. ZIP enthält LIESMICH_Migration.txt.
# - HINWEIS: Klartext-Export enthält sensible Gesundheitsdaten; nach Migration löschen.
#
# CHANGELOG v2.19.0 (DPAPI-Verschlüsselung + portabler AES-Export):
# - FEATURE: Alle JSON-Dateien (Config + Tagesdaten) werden ab sofort automatisch
#   mit DPAPI (Windows Data Protection API, CurrentUser-Scope) verschlüsselt.
#   Kein Passwort nötig – Windows übernimmt das Schlüsselmanagement.
# - FEATURE: Automatische Migration bestehender Klartext-JSON-Dateien beim Laden.
#   Erkennung via Try-JSON-Parse; bei Erfolg wird der Klartext transparent geladen
#   und beim nächsten Speichern verschlüsselt.
# - FEATURE: Neuer Button "Verschlüsseltes Backup exportieren" in Tab "Daten nach
#   Bluttests". Erstellt ein AES-256-verschlüsseltes, portables .btbackup-Archiv
#   mit frei wählbarem Passwort. Nutzbar für PC-/Profilwechsel.
# - FEATURE: Neuer Button "Backup importieren" in Tab "Daten nach Bluttests".
#   Importiert ein .btbackup-Archiv, entschlüsselt mit Passwort und re-encrypted
#   alle Daten mit dem lokalen DPAPI-Profil. Danach passwortfreie Nutzung.
# - TECHNIK: PBKDF2 (100.000 Iterationen) für AES-Schlüsselableitung,
#   AES-256-CBC mit zufälligem Salt + IV.
# - FEATURE: Integration in AutoBackup: Neue GroupBox "Portables verschlüsseltes
#   Backup (.btbackup)" im AutoBackup-Dialog. Checkbox aktiviert automatische
#   .btbackup-Erstellung bei jedem AutoBackup. Passwort wird DPAPI-geschützt
#   in Config gespeichert. Cleanup erfasst .btbackup-Dateien ebenfalls.
# - NACHTRAG FÜR TEMPLATE-SCRIPT: v2.19.0 (Verschlüsselung) + v2.19.0 (PDF-Import Beta).
# - FEATURE: PDF-Blutbild-Import (Beta Version). Neuer Button "PDF importieren (Beta)"
#   in Tab "Daten nach Bluttests". Nutzt iTextSharp (auto-download von NuGet, einzelne DLL).
#   Automatisches Marker-Matching via Alias-Map (Kürzel aus Klammern, Langnamen).
#   Review-Dialog mit Konfidenz-Anzeige, editierbaren Werten und Datumswahl.
#   PDF wird zusätzlich im Tagesordner archiviert. Alle erkannten Werte erhalten
#   die Note "PDF-Import (Beta)" zur Rückverfolgbarkeit.
#
# CHANGELOG v2.17.0 (Dokument-Upload + Genetische Vorbelastungen):
# - FEATURE: Neuer Button "Dokument hochladen" in Tab "Daten nach Bluttests".
#   Öffnet Windows OpenFileDialog (alle Dateiformate). Die gewählte Datei wird
#   kopiert, nach Schema "YYYY-MM-dd_BLUTTEST.ext" umbenannt und im Tagesordner
#   des ausgewählten Datums archiviert. Button nur aktiv bei Datumsauswahl.
# - FEATURE: Neue GroupBox "Genetische Vorbelastungen" in Tab "Persönliche Metriken"
#   rechts neben der biometrischen GroupBox. Enthält:
#     - Sub-GroupBox "Vorbelastungen berücksichtigen" mit Master-Checkbox
#       (Default: deaktiviert) und 4 vordefinierten Vorbelastungen:
#       * Chronische Veneninsuffizienz (D-Dimer, CRP)
#       * Familiäre Hypertonie / KHK-Prädisposition (Homocystein, Lp(a),
#         NT-proBNP, TnT, Syst.Blutdruck, ApoB, PREVENT-ASCVD-10Y)
#       * Familiäre Hypercholesterinämie (Chol, LDL, HDL, ApoB, Trig,
#         Non-HDL, AIP, Trig/HDL-Ratio, Omega-3-Index)
#       * Typ-2-Diabetes-Mellitus-Prädisposition (Glukose, HbA1c,
#         Nüchterninsulin, HOMA-IR, TyG-Index)
#     - Sub-GroupBox "Neue genetische Vorbelastung/en definieren" mit
#       Eingabemaske (Name + Marker-Checkbox-Auswahl).
# - FEATURE: Cockpit-Hinweise bei aktivierten Vorbelastungen: betroffene Marker
#   erhalten einen ToolTip-Hinweis im Risiko-Cockpit.
# - FEATURE: Custom Report kann Vorbelastungs-Reports generieren (nur die
#   relevanten Marker der jeweiligen Vorbelastung).
# - FEATURE: Vorbelastungs-Konfiguration wird in Config.json persistiert.
# - BUGFIX: ConvertTo-Hashtable crashte bei $null-Property-Werten mit
#   "$null.GetType()" (pre-existing Bug, nun durch Null-Guard behoben).
# - ROBUSTHEIT: GP-Merge-Code in eigenem try-catch, damit ein Fehler bei
#   Vorbelastungs-Daten nicht die gesamte Config-Ladung abbricht.
#
# CHANGELOG v2.15.3 (Wochentage-Layout-Feinschliff):
# - LAYOUT: "Donnerstag"-Checkbox war 72px breit und brach 2-zeilig um. Breite auf
#   90px erhöht, Text ist jetzt einzeilig lesbar.
# - LAYOUT: Wochentage-Checkboxen neu auf 2 Zeilen verteilt:
#     Zeile 1: Montag, Dienstag, Mittwoch, Donnerstag, Freitag
#     Zeile 2: Samstag, Sonntag
# - LAYOUT: Hinweistext in Wochentage-GroupBox eine Zeile + Leerzeile tiefer
#   platziert, damit optisch klarer getrennt von den Buttons.
# - LAYOUT-KASKADE: GroupBox Wochentage 105 -> 170 hoch, Monatstag analog (Y-Slot-
#   Sharing bleibt erhalten). Zeit-Einstellungen Höhe 280 -> 345, lblScheduleInfo,
#   Action-Buttons und Popup-Höhe entsprechend nach unten verschoben.
#
# CHANGELOG v2.15.2 (AutoBackup-Popup Layout-Feinschliff):
# - LAYOUT: "Jetzt sofort Backup ausführen"-Button auf gleiche Y-Höhe wie
#   Abbrechen/Speichern platziert (Y=660). Alle drei Aktions-Buttons nun bündig
#   in einer Linie.
# - LAYOUT: Sub-GroupBox "Backup-Bereinigung" Höhe von 85 auf 105 erhöht. Der
#   Hinweistext hat jetzt Platz für 2 Zeilen (bisher war die zweite Zeile
#   abgeschnitten).
# - LAYOUT: Sub-GroupBox "Wochentage" Höhe von 85 auf 105 erhöht, aus identischem
#   Grund.
# - LAYOUT: Alle nachgelagerten Groupboxes und Controls um 20px nach unten
#   verschoben (Basis-Konfiguration, Zeit-Einstellungen, Button-Zeile, Popup-Höhe),
#   um Überlappungen zu verhindern.
#
# CHANGELOG v2.15.1 (Button-Feinschliff):
# - LAYOUT: Beenden-Button auf Y=8 hochgezogen (vorher Y=12), damit er die
#   TabControl-Obergrenze nicht mehr minimal überlappt.
# - LAYOUT: Ehemaliger "Einstellungen"-Button aus Tab "Einzelmarker-Analyse" wird
#   umbenannt zu "globale Einstellungen" und auf Form-Ebene platziert, direkt
#   links neben dem Beenden-Button. Damit ist der Einstellungs-Dialog jetzt aus
#   JEDEM Tab erreichbar (vorher: nur sichtbar in Tab 1).
# - Folge-Layout: Im Chart-Bereich rücken Drucken/Bearbeiten an die Position vor,
#   die zuvor der Einstellungs-Button belegte.
#
# CHANGELOG v2.15.0 (Beenden-Button global sichtbar + Backup-Cleanup):
# - LAYOUT-FIX: Beenden-Button lag bisher in $tabPage1.Controls und war damit nur
#   in Tab 1 sichtbar - und auch dort unscheinbar wegen unklarer Y-Koordinate.
#   Fix: Button liegt jetzt auf $form direkt, rechts neben dem letzten Tab-Header
#   ("Persönliche Metriken"), auf Tab-Reiter-Höhe. Somit aus JEDEM Tab sichtbar.
# - FEATURE: Neue Sub-GroupBox "Backup-Bereinigung" in Basis-Konfiguration des
#   AutoBackup-Popups:
#     - Checkbox "Ältere Backups löschen" (Default: deaktiviert)
#     - Dropdown "Retention-Regel" (nur sichtbar bei aktivierter Checkbox):
#         "letztes Backup behalten" (N=1)
#         "letzte zwei Backups behalten" (N=2)
#         "letzte drei Backups behalten" (N=3)
#         "alle löschen" (N=0, nur das neu erstellte bleibt)
#   Cleanup erfolgt NACH erfolgreicher Backup-Erstellung, damit bei einem Fehler
#   keine Backups verloren gehen.
# - FEATURE: Config-Felder "CleanupEnabled" (bool) und "CleanupKeep" (Dropdown-String)
#   werden zuverlässig in Config.json persistiert (Save-Data + Load-Merge).
# - SCHUTZMASSNAHME: Cleanup löscht ausschließlich Dateien/Ordner mit dem Prefix
#   "BloodTracker_AutoBackup_" im Zielverzeichnis. Fremde Dateien bleiben unberührt.
# - SCHUTZMASSNAHME: ZIP- und Ordner-Backups werden getrennt behandelt (jeder Typ
#   hat eigenen Retention-Zähler), damit bei Format="Beides" keine unerwünschten
#   Lücken entstehen.
#
# CHANGELOG v2.14.3 (Hotfix: Beenden-Button-Sichtbarkeit):
# - HOTFIX: Der in v2.14.2 rechts bündig am TabControl platzierte Beenden-Button
#   (X=1050, Anchor Top+Right) war praktisch unsichtbar, weil er die Drucken/
#   Bearbeiten/Einstellungen-Kette überdeckte bzw. hinter ihnen lag.
#   Fix: Button steht jetzt in der linken Spalte des Tabs (Y=17), rechts bündig
#   zur Referenz-GroupBox "Persönliche & biometrische Daten" (X=430). Damit liegt
#   er im sichtbaren Bereich oberhalb der Marker-Konfig-GroupBox.
# - REVERT: Settings-Button wieder auf ursprüngliche Location (Chart.Right - 90).
#
# CHANGELOG v2.14.2 (Bugfix Config-Persistierung + Monatstag-Auswahl + Layout):
# - BUGFIX: AutoBackup-Einstellungen wurden nicht in Config.json gespeichert, weil
#   Save-Data den AutoBackup-Block beim Schreiben ignoriert hat. Folge: Nach App-Neustart
#   gingen sämtliche Backup-Einstellungen (inkl. Zielverzeichnis, Enabled-Flag,
#   LastBackup-Timestamp) verloren.
#   Fix: Save-Data schreibt jetzt alle bekannten Config-Bereiche inkl. AutoBackup.
# - BUGFIX: Load-Merge für AutoBackup war nicht robust gegen Hashtable- vs.
#   PSCustomObject-Varianten (ConvertTo-Hashtable-Output). Jetzt unterstützt beide.
# - FEATURE: Bei Interval="monatlich" erscheint analog zu "wöchentlich" eine zusätzliche
#   Sub-GroupBox "Monatstag" mit Dropdown-Menü: "am Ersten eines Monats", "am zweiten
#   Tag", "am dritten Tag", "am vierten Tag", "am fünften Tag", "zur Monatsmitte (15.)",
#   "am Letzten Tag eines Monats". Bei "am Letzten" wird dynamisch der letzte Tag
#   des aktuellen Monats berechnet (28-31, Schaltjahr-aware).
# - LAYOUT: "Beenden"-Button im Tab "Einzelmarker-Analyse" wurde vom unteren Rand der
#   Einstellungs-Groupbox auf die obere Tab-Reiter-Höhe verschoben, rechts bündig
#   (Anchor Top+Right, damit bei Fenster-Resize weiterhin bündig).
#
# CHANGELOG v2.14.1 (AutoBackup: Wochentage-Auswahl + Scheduling-Transparenz):
# - FEATURE: Bei Auswahl "wöchentlich" erscheint eine zusätzliche GroupBox "Wochentage"
#   mit Checkboxen für Mo–So + "Alle auswählen"/"Alle abwählen"-Buttons. Standardmäßig
#   ist Montag aktiv. Die Groupbox wird dynamisch ein-/ausgeblendet abhängig vom
#   gewählten Intervall.
# - FEATURE: Fälligkeits-Logik für "wöchentlich" berücksichtigt die Wochentags-Auswahl.
#   Backup erfolgt an einem gewählten Tag, wenn seit letztem Backup mindestens 1 Tag
#   vergangen ist (verhindert Doppel-Backups am selben Tag).
# - KLARSTELLUNG: Im Popup wird jetzt explizit erklärt, dass KEIN Windows-Scheduled-
#   Task erstellt wird. Backups erfolgen ausschließlich während die App läuft
#   (Start-Check + 15-Min-Scheduler). Bei "täglich 03:00 Uhr" + App geschlossen
#   → Backup wird beim nächsten App-Start nachgeholt.
# - VALIDIERUNG: Bei "wöchentlich" muss mindestens ein Wochentag angehakt sein,
#   sonst verweigert "Speichern" mit entsprechender Meldung.
#
# CHANGELOG v2.14.0 (Automatisches Backup):
# - FEATURE: Neuer Button "Automatisches Backup" im Einstellungen-Popup des
#   Einzelmarker-Analyse-Tabs. Öffnet Show-AutoBackupPopup mit zwei Groupboxen:
#     - Basis-Konfiguration: Zielverzeichnis, Format (ZIP / Ordner-Dateien / Beides)
#     - Zeit-Einstellungen: Intervall-Dropdown + Uhrzeit-Picker
#   Checkbox-Logik: "Beides" aktiviert sich automatisch, wenn ZIP + Ordner-Dateien
#   beide angehakt sind, und graut die Einzel-Checkboxen aus.
# - FEATURE: In-Process-Scheduler (Invoke-AutoBackupCheck) prüft bei Programm-Start
#   und periodisch (alle 15 Min.) ob ein Backup fällig ist.
# - FEATURE: Registry-Export wird mit gesichert (Konsistenz mit Uninstall-Backup).
#   Enthält DataPath-Konfiguration, essentiell für Restore auf anderem Gerät.
# - FEATURE: Validierung: Zielverzeichnis darf nicht unterhalb des DataDir liegen
#   (verhindert rekursive Backup-von-Backup-Spiralen).
# - FEATURE: Backup-Konfiguration + LastBackup-Timestamp im bestehenden Config.json
#   persistiert. Kein zusätzliches Config-File notwendig.
#
# CHANGELOG v2.13.1 (Patch: Post-Export-Komfort):
# - PATCH: Nach jedem Export wird optional (per Schalter)
#     (a) der Export-Ordner im Explorer geöffnet
#     (b) ein Dialog eingeblendet, der fragt, ob die exportierte Datei geöffnet werden soll
#   Beides ist unabhängig steuerbar via globaler Script-Variablen im neuen Config-Block
#   ganz oben im Script ("USER-KONFIGURATION"). Defaults: beide aktiv ($true).
# - REFACTORING: Gemeinsame Helper-Funktion Invoke-PostExportAction konsolidiert die
#   Nach-Export-Routine für alle 5 Export-Endpunkte (JSON/CSV/PDF generisch + Custom-Report
#   JSON/CSV/PDF). DRY-Prinzip, einheitliches Verhalten garantiert.
#
# CHANGELOG v2.13.0 (Bugfix PDF + Grafik-Reports + Datumsformat):
# - BUGFIX: PDF-Export warf bei kleinen Datensätzen "Arrayindex wurde als NULL
#   ausgewertet". Ursache: Automatisches Array-Unwrapping bei 1 Element.
#   Fix: Daten explizit als @() gecastet und Zugriff via $script:-Variablen statt
#   Closure, damit PrintPage-Handler zuverlässig iteriert.
# - FEATURE: Custom-Report PDF/Druck zeigt jetzt pro ausgewähltem Marker:
#     (1) Überschrift mit Marker-Name
#     (2) Liniengrafik der Werte-Entwicklung im gewählten Zeitraum
#     (3) Tabelle der historischen Werte
#   Grafik entfällt automatisch bei <2 Datenpunkten (Hinweistext).
# - FEATURE: Neue Funktion Render-MarkerChartToImage rendert GDI+-Liniencharts
#   direkt als Bitmap für PDF-Einbettung (keine externen Dependencies).
# - UX: Datumsformat in Custom-Report-Vorschau von "YYYY-MM-dd" auf "dd.MM.YYYY"
#   umgestellt. Interner Speicher bleibt weiterhin ISO-Format (yyyy-MM-dd).
#
# CHANGELOG v2.12.0 (Custom Report + PDF-Export):
# - FEATURE: Neuer Tab "Custom Report" (zwischen "Longevity-Indizes" und "Persönliche Metriken").
#   Frei konfigurierbarer Report: Marker-Auswahl per CheckBox + Zeitraum-Filter +
#   Drucken-Button + Exportieren-Button. Marker ohne Datensätze sind deaktiviert.
# - FEATURE: PDF-Export als 3. Option im Exportieren-Popup aller Tabs.
#   Nutzt den systemweiten "Microsoft Print to PDF"-Drucker (Windows 10+).
#   Fallback: Sucht nach erstem verfügbaren PDF-Drucker (Foxit, Adobe etc.).
# - PROAKTIV: Falls kein PDF-Drucker gefunden wird, erscheint eine klare
#   Fehlermeldung mit Handlungsanweisung.
#
# CHANGELOG v2.11.0 (Export-Feature + neuer Marker):
# - FEATURE: Exportieren-Button in allen 4 Haupt-Tabs neben "Drucken":
#     * Risiko-Cockpit         : immer aktiv
#     * Einzelmarker-Analyse   : linksbündig unter Grafik platziert, immer aktiv
#     * Korrelationen          : deaktiviert bis Analyse durchgeführt wurde
#     * Daten nach Bluttests   : deaktiviert bis Tag/Einträge ausgewählt sind
# - FEATURE: Zentrales Popup "Exportieren" (Show-ExportPopup) bietet JSON- und
#   CSV-Export. Orientiert sich am Design des "Einstellungen"-Popups.
# - FEATURE: Neuer Marker "HIV (Anti-HIV-1/2)" in Gruppe "Infektionskrankheiten".
# - HINWEIS: "ApoB" und "Testosteron, gesamt" waren bereits im Marker-Katalog enthalten
#   und wurden nicht dupliziert. Nur der neu geforderte Marker "HIV" wurde ergänzt.
#
# CHANGELOG v2.10.0 (Longevity Science Update):
# - FEATURE: PhenoAge-Proxy (Levine et al. 2018, Aging) - Schätzung des biologischen Alters
#   aus 9 Laborwerten (Albumin, Kreatinin, Glukose, CRP, Lymphozyten, MCV, RDW-CV,
#   Alkalische Phosphatase, Leukozyten) + chronologisches Alter. Spezialformel [PHENOAGE].
# - FEATURE: PhenoAge-Accel (biologisches minus chronologisches Alter) als direkter
#   Longevity-Indikator. Negative Werte = jünger als chronologisches Alter (gut).
# - FEATURE: InflammAging-Score (Franceschi et al.) - komposit aus hsCRP, NLR (Neutrophile/
#   Lymphozyten-Ratio) und PLR (Thrombozyten/Lymphozyten-Ratio). Spezialformel [INFLAMMAGING].
# - FEATURE: Neutrophilen-zu-Lymphozyten-Ratio (NLR) - validierter Marker für systemische
#   Inflammation und Mortalitätsrisiko (Zahorec 2021).
# - FEATURE: TyG-Index (Triglycerides-Glucose Index, Simental-Mendia 2008) - einfacher
#   Surrogatmarker für Insulinresistenz, hoch korreliert mit HOMA-IR.
# - FEATURE: Triglyceride/HDL-Ratio - metabolischer Gesundheitsmarker, starker Prädiktor
#   für Insulinresistenz und kardiovaskuläre Ereignisse.
# - FEATURE: HbA1c-Trajektorie (Slope pro Jahr) - Trend-Analyse zur Früherkennung von
#   Glukose-Dysregulation (%-Punkte/Jahr).
# - FEATURE: Neuer Tab "Longevity-Science" mit detaillierter Aufschlüsselung aller
#   Score-Komponenten, Interpretationshinweisen und Literaturreferenzen.
# - FEATURE: Korrelations-Matrix erweitert um wissenschaftlich etablierte Paarungen
#   (Inflammation x Lipide, HOMA-IR x hsCRP, Vitamin D x CRP, etc.).
# - VERBESSERT: Calculate-LongevityScore erweitert um PhenoAge-Accel als Zusatzkomponente.
# - HINWEIS: Alle Scores sind Indikatoren, keine Diagnose. Wissenschaftliche Limitierungen:
#   PhenoAge wurde an NHANES-Kohorte (USA) validiert, Einheiten-Konvertierung kritisch.
#
# CHANGELOG v2.9.2:
# - FEATURE: Neuer Button "Sicherung wiederherstellen (ZIP)..." im Einstellungen-Popup.
#   Ermöglicht das vollständige Wiederherstellen eines Deinstallations-Backups (UserData + Registry).
#   Unterstützt Clean-Restore (Daten ersetzen) und Merge-Modus (Daten ergänzen).
#   Validiert ZIP-Struktur, behandelt verschachtelte UserData-Ordner und startet nach Restore neu.
#
# CHANGELOG v2.7.4:
# - FEATURE: "Einstellungen"-Button in den Tabs "Risiko-Cockpit" und "Daten nach Bluttests"
#   hinzugefügt (zwischen Tab-Reitern und Tabelleninhalt). Öffnet zunächst ein
#   vereinfachtes Popup nur mit "Schließen"-Button (Platzhalter für spätere Erweiterung).
#
# CHANGELOG v2.7.3:
# - BUGFIX: Tab-Wechsel aktualisiert nun alle Tabs mit aktuellen Daten.
#   Zuvor wurden Änderungen aus "Daten nach Bluttests" (Löschen/Bearbeiten) erst nach
#   Neustart in anderen Tabs (Einzelmarker-Analyse, Risiko-Cockpit, Korrelationen,
#   Longevity-Indizes) sichtbar. SelectedIndexChanged-Handler erweitert.
# - BUGFIX: Gelöschte Tage/Einträge bleiben nicht mehr als "Geister-Dateien" auf der
#   Festplatte bestehen. Save-AllHistoricalData löscht nun zuerst alle bestehenden
#   Tages-JSON-Dateien und leere Verzeichnisse, bevor die verbleibenden Daten neu
#   geschrieben werden. Zuvor wurden beim nächsten Start gelöschte Daten wieder eingelesen.
#
# CHANGELOG v2.7.2:
# - BUGFIX: "Die Liste hatte eine feste Größe" Fehler beim Löschen von Einträgen behoben.
#   Load-AllHistoricalItems Rückgabe wird nun explizit als ArrayList gecastet,
#   da PowerShell ArrayLists beim Zuweisen zu fixed-size Object[] entpackt.
#   Betroffen: Tab "Daten nach Bluttests" (Tag löschen, Einträge löschen), Import, Hinzufügen.
#
# CHANGELOG v2.8:
# - BUGFIX: Array-Initialisierung im Longevity-Tab korrigiert (IndexOutOfRangeException)
#   Zeile 1032: @() * Count funktioniert nicht → @($null) * Count verwenden
#
# CHANGELOG v2.7:
# - BUGFIX: Array-Konvertierungsfehler beim Laden vorhandener Daily-Daten behoben (foreach statt AddRange)
# - BUGFIX: "Parametername -eq" Fehler in Where-Object Duplikatsprüfungen korrigiert
# - BUGFIX: Pipeline-Syntax in Update-FilterDropdown (Zeile 904) korrigiert
# Datum: 04. Februar 2026

# ╔══════════════════════════════════════════════════════════════════════════════════╗
# ║                            USER-KONFIGURATION                                    ║
# ║                                                                                  ║
# ║  Diese Variablen steuern das Verhalten des Scripts und können bei Bedarf         ║
# ║  angepasst werden. Sie gelten global für alle Export-Vorgänge (JSON, CSV, PDF)   ║
# ║  in allen Tabs (Cockpit, Einzelmarker, Korrelationen, Bluttests, Custom Report). ║
# ╚══════════════════════════════════════════════════════════════════════════════════╝

# Nach jedem Export: Ordner im Windows-Explorer automatisch öffnen?
#   $true  = Ordner wird nach erfolgreichem Export geöffnet (Default)
#   $false = Ordner wird NICHT geöffnet
$script:ExportOpenFolderAfter = $true

# Nach jedem Export: Dialog anzeigen "Möchten Sie die Datei jetzt öffnen?"
#   $true  = Dialog erscheint, bei "Ja" wird die Datei mit Standardprogramm geöffnet (Default)
#   $false = Kein Dialog, keine automatische Datei-Öffnung
$script:ExportAskOpenFile = $true

# ---------- Globale Konfiguration ----------
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System # Für ArrayList, falls benötigt
Add-Type -AssemblyName System.Security # v2.19.0: DPAPI (ProtectedData)

$Version = "2.26.0" # Aktuelle Versionierung

$regPath = "HKCU:\Software\PSC\PSC.Blood-Tracker"
if (-not (Test-Path $regPath)) { New-Item -Path $regPath -Force | Out-Null }
$storedPath = (Get-ItemProperty -Path $regPath -Name DataPath -ErrorAction SilentlyContinue).DataPath

if ($storedPath -and (Test-Path $storedPath)) {
    $dataDir = $storedPath
} else {
    $dataDir = Join-Path -Path ([Environment]::GetFolderPath("UserProfile")) -ChildPath "PSC.Blood-Tracker"
    Set-ItemProperty -Path $regPath -Name DataPath -Value $dataDir
}

$dataFile = Join-Path -Path $dataDir -ChildPath "Config.json"
$script:dailyDataDirBase = Join-Path -Path $dataDir -ChildPath "data"
$script:cockpitTimeFilter = "Alle Daten"
$script:cockpitGroupFilter = "Alle"
$script:cockpitGroupFilterCustom = [System.Collections.ArrayList]@()
$script:cockpitRatingFilter = "Alle"
$script:dataMgmtTimeFilter = "Alle Daten"
$script:longevityTimeFilter = "Alle Daten"
$script:customReportTimeFilter = "Alle Daten"

# ---------- Hilfsfunktionen ----------
function Invoke-PostExportAction {
    <#
    .SYNOPSIS
        Zentrale Nach-Export-Routine (v2.13.1).
    .DESCRIPTION
        Einheitliche Behandlung aller Export-Endpunkte (JSON/CSV/PDF, alle Tabs).
        Steuerung über die globalen Schalter:
            $script:ExportOpenFolderAfter  (Ordner öffnen?)
            $script:ExportAskOpenFile      (Datei-öffnen-Dialog?)
        Reihenfolge: (1) Erfolgsmeldung, (2) Ordner öffnen, (3) Öffnen-Dialog.
    .PARAMETER FilePath
        Absoluter Pfad zur soeben exportierten Datei.
    .PARAMETER SuccessTitle
        Optional: Titel der Erfolgs-MessageBox (Default: "Export abgeschlossen").
    .PARAMETER SuccessMessage
        Optional: Text der Erfolgs-MessageBox. Wenn leer, wird ein Default erzeugt.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [string]$SuccessTitle   = "Export abgeschlossen",
        [string]$SuccessMessage = $null
    )
    try {
        if (-not $SuccessMessage) {
            $SuccessMessage = "Export erfolgreich gespeichert:`n$FilePath"
        }
            [System.Windows.Forms.MessageBox]::Show($SuccessMessage, $SuccessTitle, "OK", "Information") | Out-Null

        # (1) Ordner öffnen (optional)
        if ($script:ExportOpenFolderAfter) {
            $folder = Split-Path -Path $FilePath -Parent
            if ($folder -and (Test-Path $folder)) {
                try { Invoke-Item $folder } catch { Write-Warning "Ordner konnte nicht geöffnet werden: $($_.Exception.Message)" }
            }
        }

        # (2) Datei öffnen nachfragen (optional)
        if ($script:ExportAskOpenFile) {
            $askResult = [System.Windows.Forms.MessageBox]::Show(
                "Möchtest du die exportierte Datei jetzt öffnen?`n`n$FilePath",
                "Datei öffnen?",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            if ($askResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                if (Test-Path $FilePath) {
                    try { Invoke-Item $FilePath } catch {
                        [System.Windows.Forms.MessageBox]::Show("Datei konnte nicht geöffnet werden: $($_.Exception.Message)", "Fehler", "OK", "Warning") | Out-Null
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Datei nicht gefunden: $FilePath", "Fehler", "OK", "Warning") | Out-Null
                }
            }
        }
    } catch {
        Write-Warning "Invoke-PostExportAction Fehler: $($_.Exception.Message)"
    }
}

function Parse-Number {
    param([string]$text)
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    $t = $text.Trim() -replace ",", "."
    $out = 0.0
    if ([double]::TryParse($t, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$out)) { return [double]$out }
    throw "Ungültige Zahl: '$text'"
}

# =====================================================================
# v2.23.0: MARKER-ALIAS-SUCHE (Labor-Synonyme)
# ---------------------------------------------------------------------
# Verschiedene Labore benennen denselben Parameter unterschiedlich
# (z. B. "WBC" vs. "Leukozyten"). Diese Tabelle bildet gaengige Synonyme
# auf den kanonischen Marker-Namen der Config ab.
#
# WICHTIG: Die Tabelle liegt bewusst NICHT in $script:data.Config.Markers,
# da Load-Config eine gespeicherte Config.json vollstaendig ueber die
# Defaults legt - Aliases wuerden dadurch verloren gehen.
#
# BEWUSST NICHT GEMAPPT (Verwechslungs-/Umrechnungsgefahr):
#   BUN  -> ist NICHT Harnstoff (BUN x 2,14 = Harnstoff)
#   PCT  -> doppeldeutig: Plateletcrit vs. Procalcitonin
#   TG   -> doppeldeutig: Triglyceride vs. Thyreoglobulin
#
# v2.25.0: 'CRP' bleibt beim klassischen Marker, alle hs-/hochsensitiv-Varianten
# zeigen auf 'Hochsensitives CRP (hs-CRP)'. Neu ausserdem SHBG.
# =====================================================================

$script:MarkerAliases = @{
    'Albumin'                                   = @('ALB', 'Serumalbumin')
    'Alanin-Aminotransferase (ALT/GPT)'         = @('ALT', 'GPT', 'ALAT', 'SGPT', 'GPT (ALAT)')
    'Alkalische Phosphatase (AP)'               = @('AP', 'ALP', 'AKP', 'ALKP', 'Alk. Phosphatase', 'Alkalische Phosphatase')
    'Anteil großer Thrombozyten (P-LCR)'        = @('P-LCR', 'PLCR', 'P LCR')
    'ApoB'                                      = @('Apo B', 'Apolipoprotein B', 'Apolipoprotein-B', 'ApoB-100', 'Apo B100')
    'Aspartat-Aminotransferase (AST/GOT)'       = @('AST', 'GOT', 'ASAT', 'SGOT', 'GOT (ASAT)')
    'Basophile Granulozyten'                    = @('BASO', 'Basophile', 'Basophile Gran.', 'Basos')
    'CEA (Carcinoembryonic Antigen)'            = @('CEA', 'Karzinoembryonales Antigen', 'Carcinoembryonales Antigen')
    'Cholesterin, gesamt (Chol)'                = @('Chol', 'CHOL', 'Cholesterin', 'Gesamtcholesterin', 'Cholesterin gesamt', 'Total Cholesterol', 'Cholesterin ges.')
    'Cortisol'                                  = @('Kortisol', 'Hydrocortison', 'Cortisol Serum')
    # v2.25.0: hs-Aliase entfernt -> liegen jetzt auf 'Hochsensitives CRP (hs-CRP)'.
    # Beide Marker parallel haetten sonst eine mehrdeutige Aufloesung erzeugt
    # (Resolve-MarkerName liefert bei doppelt vergebenem Token bewusst $null).
    'C-reaktives Protein (CRP)'                 = @('CRP', 'C-reaktives Protein', 'CRP quantitativ')
    'Hochsensitives CRP (hs-CRP)'               = @('hs-CRP', 'hsCRP', 'hs CRP', 'CRP hs', 'CRP (hs)', 'hochsensitives CRP', 'CRP hochsensitiv', 'ultrasensitives CRP', 'high sensitivity CRP', 'hsCRP quantitativ')
    'Sexualhormon-bindendes Globulin (SHBG)'    = @('SHBG', 'Sexualhormon-bindendes Globulin', 'Sexualhormonbindendes Globulin', 'Sexualhormon bindendes Globulin', 'Sex Hormone Binding Globulin', 'Sexualhormon-bindendes-Globulin', 'TeBG', 'Testosteron-bindendes Globulin')
    'Cystatin-C'                                = @('Cystatin C', 'CysC', 'Cystatin')
    'D-Dimer'                                   = @('D-Dimere', 'DDimer', 'D Dimer', 'D-Dimere quantitativ')
    'DHEA-S'                                    = @('DHEAS', 'DHEA S', 'DHEA-Sulfat', 'Dehydroepiandrosteron-Sulfat', 'Dehydroepiandrosteronsulfat')
    'Eisen (Fe)'                                = @('Fe', 'Eisen', 'Serumeisen', 'Eisen im Serum')
    'Eosinophile Granulozyten'                  = @('EOS', 'Eosinophile', 'Eosinophile Gran.', 'Eos')
    'Leukozyten (WBC)'                          = @('WBC', 'Leuko', 'Leukos', 'LEU', 'Leukozyten', 'weiße Blutkörperchen', 'weisse Blutkoerperchen', 'White Blood Cells')
    'Erythrozyten (RBC)'                        = @('RBC', 'ERY', 'Erys', 'Erythrozyten', 'rote Blutkörperchen', 'rote Blutkoerperchen', 'Red Blood Cells')
    'Hämoglobin (HGB/Hb)'                       = @('HGB', 'HB', 'Haemoglobin', 'Hämoglobin', 'Hgb')
    'Hämatokrit (HCT/Hk)'                       = @('HCT', 'HKT', 'HK', 'Haematokrit', 'Hämatokrit', 'Hkt')
    'Mittleres Erythrozytenvolumen (MCV)'       = @('MCV', 'Mittleres Erythrozytenvolumen')
    'Mittlerer Hämoglobingehalt (MCH)'          = @('MCH', 'HbE', 'Hb E', 'Mittlerer Hämoglobingehalt')
    'Mittlere Hämoglobinkonzentration (MCHC)'   = @('MCHC', 'Mittlere Hämoglobinkonzentration')
    'Thrombozyten (PLT)'                        = @('PLT', 'THR', 'THRO', 'Thrombo', 'Thrombos', 'Thrombozyten', 'Blutplättchen', 'Blutplaettchen', 'Platelets')
    'Erythrozyten-Verteilungsbreite (RDW-SD)'   = @('RDW-SD', 'RDWSD', 'RDW SD')
    'Erythrozyten-Verteilungsbreite (RDW-CV)'   = @('RDW-CV', 'RDWCV', 'RDW CV', 'RDW')
    'Thrombozyten-Verteilungsbreite (PDW)'      = @('PDW', 'Platelet Distribution Width')
    'Mittleres Thrombozytenvolumen (MPV)'       = @('MPV', 'Mean Platelet Volume', 'Mittleres Thrombozytenvolumen')
    'Thrombozytenvolumen (PCT)'                 = @('Plateletcrit', 'Thrombokrit')
    'Ferritin'                                  = @('FERR', 'Serumferritin', 'Ferritin im Serum')
    'Freies T3 (fT3)'                           = @('fT3', 'FT3', 'freies T3', 'T3 frei', 'Trijodthyronin frei', 'freies Trijodthyronin')
    'Freies T4 (fT4)'                           = @('fT4', 'FT4', 'freies T4', 'T4 frei', 'Thyroxin frei', 'freies Thyroxin')
    'Freies Testosteron'                        = @('freies Testo', 'fTesto', 'Testosteron frei', 'free testosterone', 'Testosteron, frei')
    'Gamma-Glutamyltransferase (GGT)'           = @('GGT', 'gGT', 'Gamma-GT', 'GammaGT', 'Gamma GT', 'Y-GT')
    'Gewicht'                                   = @('Körpergewicht', 'Koerpergewicht', 'KG', 'Weight', 'Masse')
    'Glomeruläre Filtrationsrate (GFR)'         = @('GFR', 'eGFR', 'GFR CKD-EPI', 'CKD-EPI', 'glomeruläre Filtrationsrate', 'Kreatinin-Clearance')
    'Glukose (Glu)'                             = @('GLU', 'Glucose', 'Glukose', 'Blutzucker', 'BZ', 'Nüchternglukose', 'Nuechternglukose', 'Nüchternblutzucker', 'Glucose nüchtern', 'FPG')
    'Harnsäure (UA)'                            = @('UA', 'Urat', 'Harnsäure', 'Harnsaeure', 'Uric Acid')
    'Harnstoff'                                 = @('Urea', 'Harnstoff im Serum')
    'HDL-Cholesterin (HDL)'                     = @('HDL', 'HDL-C', 'HDL Cholesterin', 'HDL-Cholesterin')
    'Homocystein'                               = @('HCY', 'Hcy', 'Homozystein', 'Homocystein gesamt')
    'HIV (Anti-HIV-1/2)'                        = @('HIV', 'Anti-HIV', 'HIV-AK', 'HIV-Test', 'HIV 1/2', 'HIV-Suchtest', 'HIV-Screening')
    'Kalium (K)'                                = @('K', 'K+', 'Kalium', 'Potassium')
    'Körperfettanteil (KFA)'                    = @('KFA', 'Körperfett', 'Koerperfett', 'Body Fat', 'Fettanteil', 'Körperfettanteil')
    'Kreatinin (Krea)'                          = @('KREA', 'CREA', 'Crea', 'Creatinin', 'Kreatinin', 'Kreatinin im Serum')
    'Langzeitzucker (HbA1c)'                    = @('HbA1c', 'HBA1C', 'A1c', 'HbA1c IFCC', 'Langzeitzucker', 'Glykohämoglobin', 'Glykohaemoglobin', 'HbA1C DCCT')
    'LDL-Cholesterin (LDL)'                     = @('LDL', 'LDL-C', 'LDL Cholesterin', 'LDL-Cholesterin')
    'Lipase'                                    = @('LIP', 'Pankreaslipase', 'Lipase im Serum')
    'Lp(a)'                                     = @('Lpa', 'LPA', 'Lp (a)', 'Lipoprotein(a)', 'Lipoprotein a', 'Lipoprotein (a)')
    'Lymphozyten (LYMPH)'                       = @('LYMPH', 'LYM', 'Lympho', 'Lymphos', 'Lymphozyten')
    'Monozyten'                                 = @('MONO', 'Monos', 'Monozyten absolut')
    'NAD+'                                      = @('NAD', 'NADplus', 'NAD plus', 'Nicotinamid-Adenin-Dinukleotid')
    'Neutrophile Granulozyten (NEUT)'           = @('NEUT', 'NEU', 'Neutro', 'Neutros', 'Neutrophile', 'Segmentkernige', 'Neutrophile Gran.')
    'NT-proBNP'                                 = @('NTproBNP', 'NT pro BNP', 'NT-pro-BNP', 'proBNP', 'BNP')
    'Nüchterninsulin'                           = @('Insulin', 'INS', 'Nüchtern-Insulin', 'Nuechterninsulin', 'Insulin nüchtern', 'Fasting Insulin', 'Insulin basal')
    'Omega-3-Index'                             = @('Omega3', 'Omega-3', 'Omega 3', 'Omega 3 Index', 'O3I', 'HS-Omega-3-Index')
    'PSA (Prostataspezifisches Antigen)'        = @('PSA', 'tPSA', 'PSA gesamt', 'Gesamt-PSA', 'Prostataspezifisches Antigen', 'PSA total')
    # v2.26.0: Genotyp- und Mineralstoff-Marker.
    # WICHTIG: 'Zn' und 'Mg' stehen NUR hier (GUI-Alias-Suche) und bewusst NICHT
    # als Klammer-Kuerzel im Marker-Namen - Build-MarkerAliasMap wuerde sie sonst
    # automatisch in den PDF-Import uebernehmen (2-stellige Kuerzel = Fehltreffer).
    'Apolipoprotein E-Genotyp (APOE)'           = @('APOE', 'ApoE', 'APO E', 'Apolipoprotein E', 'APOE-Genotyp', 'ApoE Genotyp', 'APOE Genotypisierung', 'Apolipoprotein-E-Genotyp', 'APO-E')
    'Zink'                                      = @('Zn', 'Zink im Serum', 'Serumzink', 'Zinc', 'Zink Serum', 'Zink (Serum)')
    'Magnesium'                                 = @('Mg', 'Magnesium im Serum', 'Serummagnesium', 'Mg2+', 'Magnesium Serum', 'Magnesium (Serum)')
    'Systolischer Blutdruck'                    = @('SBP', 'RR systolisch', 'RRsys', 'systolisch', 'Blutdruck systolisch', 'SYS', 'RR sys')
    'Testosteron, gesamt'                       = @('Testosteron', 'Testo', 'Gesamttestosteron', 'Testosteron gesamt', 'Total Testosterone', 'Testosteron ges.')
    'TnT (Troponin T)'                          = @('TnT', 'Troponin', 'Troponin T', 'hsTnT', 'hs-TnT', 'Trop T', 'Troponin T hs')
    'Triglyceride (Trig)'                       = @('TRIG', 'Trig', 'Triglyceride', 'Triglyzeride', 'Neutralfette')
    'TSH-basal (TSH)'                           = @('TSH', 'TSH basal', 'TSH-basal', 'Thyreotropin', 'TSH (basal)')
    'Vitamin B12'                               = @('B12', 'Vit B12', 'Vit. B12', 'Cobalamin', 'Vitamin-B12', 'Cyanocobalamin')
    'Vitamin D (25-OH)'                         = @('Vitamin D', 'Vit D', 'Vit. D', '25-OH-D', '25-OH Vitamin D', '25(OH)D', '25-Hydroxy-Vitamin D', 'Calcidiol', 'Vitamin D3', 'Vitamin-D')
    'Workout-Frequenz'                          = @('Workout', 'Training', 'Trainingsfrequenz', 'Sport', 'Trainingseinheiten')
}

# =====================================================================
# v2.25.0: MARKER-SET-VERSIONIERUNG (Migration neuer Default-Marker)
# ---------------------------------------------------------------------
# Problem bis v2.24.1: Load-Config ersetzt Config.Markers vollstaendig durch
# die gespeicherte Config.json. Neu ausgelieferte Default-Marker waren in einer
# bestehenden Installation damit unsichtbar.
#
# Loesung: Jede Version, die Marker ergaenzt, traegt sich hier ein. Beim Laden
# werden ausschliesslich die Marker der NEUEREN Versionen nachgezogen.
#   - Bereits vorhandene Marker (Name) werden NIE ueberschrieben.
#   - Vom Nutzer geloeschte Marker werden NICHT wieder hergestellt, sobald die
#     erreichte MarkerSetVersion in der Config.json steht.
# =====================================================================

$script:MarkerSetVersion = '2.26.0'

$script:MarkerSetAdditions = @(
    @{ Version = '2.25.0'; Markers = @('Hochsensitives CRP (hs-CRP)', 'Sexualhormon-bindendes Globulin (SHBG)'); CalculatedMarkers = @() },
    @{ Version = '2.26.0'; Markers = @('Apolipoprotein E-Genotyp (APOE)', 'Zink', 'Magnesium'); CalculatedMarkers = @() }
)

# ---------------------------------------------------------------------
# v2.25.0: Ersatz-Marker fuer Score-Berechnungen.
# Fehlt der Schluessel-Marker an einem Testdatum, wird der erste verfuegbare
# Fallback verwendet. Voraussetzung: IDENTISCHE Einheit.
# hs-CRP und CRP werden beide in mg/l gemessen; InflammAging, PhenoAge und
# Longevity-Score sind fachlich ohnehin auf hsCRP definiert.
# ---------------------------------------------------------------------
$script:MarkerFallbacks = @{
    'C-reaktives Protein (CRP)' = @('Hochsensitives CRP (hs-CRP)')
}

# ---------------------------------------------------------------------
# v2.25.0: PDF-Import - Zeilen, die fuer einen Marker NICHT gelten duerfen.
# "hs-CRP 0,8 mg/l" wurde bisher vom klassischen CRP eingesammelt, weil das
# Suchmuster "(?<!\w)CRP(?!\w)" auch nach einem Bindestrich greift.
# ---------------------------------------------------------------------
$script:PdfSkipLinePatterns = @{
    'C-reaktives Protein (CRP)' = '(?i)(hs[\s\-]?crp|crp[\s\-]?hs|hochsensitiv|ultrasensitiv|high\s*sensitivity)'
}

# ---------------------------------------------------------------------
# v2.26.0: Marker, die der PDF-Import NIE automatisch befuellen darf.
# Der Parser sucht das Muster "Alias + Zahl". Bei qualitativen Markern erzeugt
# das systematisch falsche Werte:
#   - 'HIV (Anti-HIV-1/2)': Build-MarkerAliasMap zerlegt den Klammerzusatz an
#     '/' und erzeugt daraus u. a. den Alias "2". Jede Zeile der Form
#     "2 <Zahl>" haette einen HIV-Wert erzeugt - bei Wert >= 1 sogar
#     "reaktiv". (BUGFIX v2.26.0)
#   - APOE: Das Ergebnis ist ein Genotyp (E3/E4), keine Messgroesse.
# Beide Marker werden weiterhin manuell ueber ihre Auswahlliste erfasst.
# ---------------------------------------------------------------------
$script:PdfExcludedMarkers = @(
    'HIV (Anti-HIV-1/2)',
    'Apolipoprotein E-Genotyp (APOE)'
)

# ---------------------------------------------------------------------
# v2.26.0: APOE-Genotyp - Kodierung, Anzeige und Rueckwandlung.
# Der Datenspeicher haelt ausschliesslich Zahlen (Value). Der Genotyp wird
# deshalb als Code 1-6 abgelegt, aufsteigend nach dem Alzheimer-Risiko
# (Bezugsgenotyp E3/E3):
#   1 = E2/E2   2 = E2/E3   3 = E3/E3   4 = E2/E4   5 = E3/E4   6 = E4/E4
# Referenzbereich 1-3 = kein E4-Allel.
#
# Bewusst ASCII-Schreibweise "E2/E3" statt "eps2/eps3": das Script wird haeufig
# kopiert und weitergegeben - griechische Zeichen ueberleben ein versehentliches
# Speichern in ANSI/Windows-1252 nicht, deutsche Umlaute schon.
# ---------------------------------------------------------------------
$script:ApoeMarkerName = 'Apolipoprotein E-Genotyp (APOE)'

# Reihenfolge = Code 1..6 (aufsteigendes Risiko). Wird auch als Quelle fuer die
# Auswahlliste im Tab "Einzelmarker-Analyse" verwendet.
$script:ApoeGenotypeOptions = @('E2/E2', 'E2/E3', 'E3/E3', 'E2/E4', 'E3/E4', 'E4/E4')

function Get-ApoeGenotypeText {
    <#
    .SYNOPSIS
        v2.26.0: Wandelt den gespeicherten Zahlencode in den Genotyp-Klartext.
    .OUTPUTS
        z. B. 3 -> "E3/E3". Unbekannte Codes -> "unbekannt (<Wert>)".
    #>
    param($Value)
    $code = -1
    try { $code = [int][Math]::Round([double]$Value) } catch { $code = -1 }
    if ($code -ge 1 -and $code -le $script:ApoeGenotypeOptions.Count) {
        return [string]$script:ApoeGenotypeOptions[$code - 1]
    }
    return "unbekannt ($Value)"
}

function Get-ApoeGenotypeCode {
    <#
    .SYNOPSIS
        v2.26.0: Wandelt einen Genotyp-Text in den Code 1-6.
    .DESCRIPTION
        Akzeptiert "E3/E4", "e3/e4", "3/4", "E4/E3" (Reihenfolge egal).
    .OUTPUTS
        [int] Code 1-6, oder 0 wenn nicht zuordenbar.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return 0 }
    $digits = ($Text -replace '[^2-4]', '')
    if ($digits.Length -ne 2) { return 0 }
    $sorted = -join (($digits.ToCharArray()) | Sort-Object)
    switch ($sorted) {
        '22' { return 1 }
        '23' { return 2 }
        '33' { return 3 }
        '24' { return 4 }
        '34' { return 5 }
        '44' { return 6 }
    }
    return 0
}

function Merge-NewDefaultMarkers {
    <#
    .SYNOPSIS
        v2.25.0: Zieht Default-Marker neuer Versionen in eine bestehende Config nach.
    .PARAMETER ConfigData
        Die (bereits mit der gespeicherten Config befuellte) Config-Hashtable.
    .PARAMETER DefaultMarkers
        Snapshot der Default-Marker VOR dem Ueberschreiben durch die Config.json.
    .PARAMETER DefaultCalculatedMarkers
        Snapshot der berechneten Default-Marker.
    .PARAMETER SavedVersion
        In der Config.json gespeicherte MarkerSetVersion ($null = Altbestand).
    .OUTPUTS
        [bool] $true, wenn Marker ergaenzt wurden (Config muss gespeichert werden).
    #>
    param(
        [Parameter(Mandatory)]$ConfigData,
        $DefaultMarkers,
        $DefaultCalculatedMarkers,
        [string]$SavedVersion
    )
    $changed = $false
    try {
        $savedVer = $null
        if (-not [string]::IsNullOrWhiteSpace($SavedVersion)) {
            [void][version]::TryParse($SavedVersion, [ref]$savedVer)
        }

        $existingNames = @($ConfigData['Markers'] | ForEach-Object { [string]$_.Name })
        $existingCalcNames = @($ConfigData['CalculatedMarkers'] | ForEach-Object { [string]$_.Name })

        foreach ($addition in $script:MarkerSetAdditions) {
            $addVer = $null
            if (-not [version]::TryParse($addition.Version, [ref]$addVer)) { continue }
            # Nur Versionen nachziehen, die neuer sind als der gespeicherte Stand
            if ($savedVer -and $addVer -le $savedVer) { continue }

            foreach ($markerName in @($addition.Markers)) {
                if ($existingNames -contains $markerName) { continue }
                $default = $DefaultMarkers | Where-Object { $_.Name -eq $markerName } | Select-Object -First 1
                if (-not $default) { Write-Warning "Migration: Default-Marker '$markerName' nicht gefunden."; continue }
                [void]$ConfigData['Markers'].Add($default)
                $existingNames += $markerName
                $changed = $true
            }
            foreach ($calcName in @($addition.CalculatedMarkers)) {
                if ($existingCalcNames -contains $calcName) { continue }
                $default = $DefaultCalculatedMarkers | Where-Object { $_.Name -eq $calcName } | Select-Object -First 1
                if (-not $default) { Write-Warning "Migration: Berechneter Default-Marker '$calcName' nicht gefunden."; continue }
                [void]$ConfigData['CalculatedMarkers'].Add($default)
                $existingCalcNames += $calcName
                $changed = $true
            }
        }
    } catch {
        Write-Warning "Marker-Migration fehlgeschlagen: $($_.Exception.Message)"
    }
    return $changed
}

function Get-MarkerToken {
    <#
    .SYNOPSIS
        Normalisiert einen Suchbegriff fuer den Alias-Vergleich.
    .DESCRIPTION
        Kleinschreibung, Umlaut-Aufloesung, Entfernen aller Sonderzeichen.
        "Haemoglobin (HGB/Hb)" -> "haemoglobinhgbhb"
        "hs-CRP"               -> "hscrp"
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return "" }
    $t = $Text.ToLowerInvariant()
    $t = $t -replace 'ä', 'ae' -replace 'ö', 'oe' -replace 'ü', 'ue' -replace 'ß', 'ss'
    $t = $t -replace '[^a-z0-9]', ''
    return $t
}

function Build-MarkerSearchIndex {
    <#
    .SYNOPSIS
        Baut den Suchindex (kanonischer Name -> alle Tokens) auf.
    .DESCRIPTION
        Quellen je Marker: voller Name, Name ohne Klammerzusatz,
        Kuerzel aus Klammern (getrennt per / und ,) sowie $script:MarkerAliases.
        Wird von Update-AllMarkerDropdowns aufgerufen, damit auch selbst
        angelegte Marker sofort auffindbar sind.
    #>
    $index = @()
    $names = @()
    if ($script:data -and $script:data.Config -and $script:data.Config.Markers) {
        $names = @($script:data.Config.Markers | ForEach-Object { [string]$_.Name } | Where-Object { $_ } | Sort-Object)
    }
    foreach ($name in $names) {
        $cands = @($name)
        $base = ($name -replace '\s*\([^)]*\)\s*', '').Trim()
        if ($base) { $cands += $base }
        if ($name -match '\(([^)]+)\)') {
            foreach ($part in ($Matches[1] -split '[/,]')) {
                $p = $part.Trim()
                if ($p) { $cands += $p }
            }
        }
        if ($script:MarkerAliases.ContainsKey($name)) {
            foreach ($a in $script:MarkerAliases[$name]) {
                if ($a) { $cands += [string]$a }
            }
        }
        $tokens = @($cands | ForEach-Object { Get-MarkerToken $_ } | Where-Object { $_ } | Select-Object -Unique)
        $index += [PSCustomObject]@{ Name = $name; Tokens = $tokens }
    }
    $script:MarkerSearchIndex = $index
}

function Get-MarkerMatches {
    <#
    .SYNOPSIS
        Liefert alle passenden Marker-Namen, nach Relevanz sortiert.
    .DESCRIPTION
        Rang 0 = exakter Token-Treffer (z. B. "WBC")
        Rang 1 = Token beginnt mit der Eingabe ("Leuk" -> "Leukozyten (WBC)")
        Rang 2 = Token enthaelt die Eingabe
        Kurz-Kuerzel mit <= 3 Zeichen (K, AP, GOT, GLU) werden NUR exakt
        gematcht, sonst entstehen massenhaft Fehltreffer.
    #>
    param([Parameter(Mandatory)][string]$Query)

    $q = Get-MarkerToken $Query
    if ([string]::IsNullOrEmpty($q)) { return @() }
    if (-not $script:MarkerSearchIndex) { Build-MarkerSearchIndex }

    $rank0 = @(); $rank1 = @(); $rank2 = @()

    foreach ($entry in $script:MarkerSearchIndex) {
        $best = 99
        foreach ($tok in $entry.Tokens) {
            if ($tok -eq $q) { $best = 0; break }
            if ($tok.Length -le 3) { continue }
            if ($tok.StartsWith($q)) { if ($best -gt 1) { $best = 1 }; continue }
            if ($tok.Contains($q))   { if ($best -gt 2) { $best = 2 } }
        }
        switch ($best) {
            0 { $rank0 += $entry.Name }
            1 { $rank1 += $entry.Name }
            2 { $rank2 += $entry.Name }
        }
    }
    return @($rank0) + @($rank1) + @($rank2)
}

function Resolve-MarkerName {
    <#
    .SYNOPSIS
        Loest eine freie Eingabe auf genau einen kanonischen Marker-Namen auf.
    .OUTPUTS
        Kanonischer Marker-Name oder $null bei Mehrdeutigkeit / kein Treffer.
    #>
    param([string]$Query)

    if ([string]::IsNullOrWhiteSpace($Query)) { return $null }
    $q = Get-MarkerToken $Query
    if (-not $q) { return $null }
    if (-not $script:MarkerSearchIndex) { Build-MarkerSearchIndex }

    # 1) Exakter Token-Treffer. Ist ein Token mehrfach vergeben (z. B.
    #    "Erythrozyten-Verteilungsbreite" -> RDW-CV UND RDW-SD), gilt die
    #    Eingabe bewusst als mehrdeutig statt still den ersten zu nehmen.
    $exact = @()
    foreach ($entry in $script:MarkerSearchIndex) {
        if ($entry.Tokens -contains $q) { $exact += [string]$entry.Name }
    }
    if ($exact.Count -eq 1) { return [string]$exact[0] }
    if ($exact.Count -gt 1) { return $null }
    # 2) Eindeutiger Teiltreffer
    $hits = @(Get-MarkerMatches -Query $Query)
    if ($hits.Count -eq 1) { return [string]$hits[0] }
    return $null
}

# =====================================================================
# v2.19.0: Verschlüsselungs-Hilfsfunktionen (DPAPI + AES-256)
# =====================================================================

function Write-ProtectedJsonFile {
    <#
    .SYNOPSIS
        Verschlüsselt einen JSON-String mit DPAPI und schreibt ihn als Datei.
    .DESCRIPTION
        Format der Datei: Base64(DPAPI(UTF8(JSON)))
        Nur der aktuelle Windows-User auf diesem PC kann entschlüsseln.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$JsonString
    )
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($JsonString)
    $encrypted = [System.Security.Cryptography.ProtectedData]::Protect(
        $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser
    )
    $base64 = [Convert]::ToBase64String($encrypted)
    # Schreibe mit Marker-Prefix, damit Read-ProtectedJsonFile erkennen kann
    Set-Content -Path $Path -Value "BTENC:$base64" -Encoding UTF8 -Force
}

function Read-ProtectedJsonFile {
    <#
    .SYNOPSIS
        Liest eine JSON-Datei – automatische Erkennung ob verschlüsselt oder Klartext.
    .DESCRIPTION
        1. Prüft auf BTENC:-Prefix → DPAPI-Entschlüsselung
        2. Sonst: Versucht JSON-Parse → Klartext (Migration von Altdaten)
        Gibt den JSON-String zurück oder $null bei Fehler.
    #>
    param(
        [Parameter(Mandatory)][string]$Path
    )
    if (-not (Test-Path $Path)) { return $null }
    $raw = Get-Content -Path $Path -Raw -ErrorAction Stop
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }

    # Fall 1: Verschlüsselt (BTENC:-Prefix)
    if ($raw.TrimStart().StartsWith("BTENC:")) {
        $base64 = $raw.TrimStart().Substring(6).Trim()
        $encrypted = [Convert]::FromBase64String($base64)
        $decrypted = [System.Security.Cryptography.ProtectedData]::Unprotect(
            $encrypted, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        return [System.Text.Encoding]::UTF8.GetString($decrypted)
    }

    # Fall 2: Klartext-JSON (Altbestand / Migration)
    try {
        $null = $raw | ConvertFrom-Json -ErrorAction Stop
        return $raw
    } catch {
        Write-Warning "Datei '$Path' ist weder verschlüsselt noch gültiges JSON."
        return $null
    }
}

function Export-AesBackup {
    <#
    .SYNOPSIS
        Exportiert das gesamte UserData-Verzeichnis als AES-256-verschlüsseltes .btbackup-Archiv.
    .DESCRIPTION
        Ablauf:
        1. Alle JSON-Dateien entschlüsseln (DPAPI) → temporärer Klartext
        2. ZIP erstellen aus temporärem Verzeichnis
        3. ZIP-Bytes mit AES-256-CBC verschlüsseln (Passphrase → PBKDF2)
        4. [32-byte Salt][16-byte IV][AES-Ciphertext] → .btbackup-Datei
    #>
    param(
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][System.Security.SecureString]$Passphrase
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $tempDir = Join-Path -Path $env:TEMP -ChildPath "BT_export_$timestamp"
    $tempZip = Join-Path -Path $env:TEMP -ChildPath "BT_export_$timestamp.zip"

    try {
        # 1. Alle Dateien entschlüsseln in Temp-Verzeichnis
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
        
        # Config.json
        if (Test-Path $script:dataFile) {
            $json = Read-ProtectedJsonFile -Path $script:dataFile
            if ($json) {
                $configTarget = Join-Path $tempDir "Config.json"
                Set-Content -Path $configTarget -Value $json -Encoding UTF8
            }
        }
        
        # Tagesdaten
        if (Test-Path $script:dailyDataDirBase) {
            $dailyFiles = Get-ChildItem -Path $script:dailyDataDirBase -Filter "*.json" -Recurse -ErrorAction SilentlyContinue
            foreach ($file in $dailyFiles) {
                $json = Read-ProtectedJsonFile -Path $file.FullName
                if ($json) {
                    $relativePath = $file.FullName.Substring($script:dailyDataDirBase.Length).TrimStart('\','/')
                    $targetPath = Join-Path (Join-Path $tempDir "data") $relativePath
                    $targetDir = Split-Path $targetPath -Parent
                    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                    Set-Content -Path $targetPath -Value $json -Encoding UTF8
                }
            }
        }

        # Nicht-JSON-Dateien (hochgeladene Dokumente etc.) kopieren
        if (Test-Path $script:dailyDataDirBase) {
            $otherFiles = Get-ChildItem -Path $script:dailyDataDirBase -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne ".json" }
            foreach ($file in $otherFiles) {
                $relativePath = $file.FullName.Substring($script:dailyDataDirBase.Length).TrimStart('\','/')
                $targetPath = Join-Path (Join-Path $tempDir "data") $relativePath
                $targetDir = Split-Path $targetPath -Parent
                if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                Copy-Item -Path $file.FullName -Destination $targetPath -Force
            }
        }

        # 2. ZIP erstellen
        if (Test-Path $tempZip) { Remove-Item $tempZip -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $tempZip)
        $zipBytes = [System.IO.File]::ReadAllBytes($tempZip)

        # 3. AES-256 verschlüsseln
        $salt = New-Object byte[] 32
        $iv = New-Object byte[] 16
        $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
        $rng.GetBytes($salt)
        $rng.GetBytes($iv)
        $rng.Dispose()

        # Passphrase → Schlüssel via PBKDF2
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Passphrase)
        $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

        $deriveBytes = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($plainPass, $salt, 100000, "SHA256")
        $key = $deriveBytes.GetBytes(32) # 256 Bit
        $deriveBytes.Dispose()
        # Passphrase aus Speicher löschen
        $plainPass = $null

        $aes = [System.Security.Cryptography.Aes]::Create()
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.Key = $key
        $aes.IV = $iv

        $encryptor = $aes.CreateEncryptor()
        $cipherBytes = $encryptor.TransformFinalBlock($zipBytes, 0, $zipBytes.Length)
        $encryptor.Dispose()
        $aes.Dispose()

        # 4. Datei schreiben: [Salt 32][IV 16][Cipher n]
        $outputStream = [System.IO.File]::Create($OutputPath)
        $outputStream.Write($salt, 0, 32)
        $outputStream.Write($iv, 0, 16)
        $outputStream.Write($cipherBytes, 0, $cipherBytes.Length)
        $outputStream.Close()

        return $true
    } catch {
        throw $_
    } finally {
        if (Test-Path $tempDir)  { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tempZip)  { Remove-Item $tempZip -Force -ErrorAction SilentlyContinue }
    }
}

function Import-AesBackup {
    <#
    .SYNOPSIS
        Importiert ein .btbackup-Archiv, entschlüsselt mit Passphrase, re-encrypted mit DPAPI.
    .DESCRIPTION
        1. AES-256 entschlüsseln → ZIP-Bytes
        2. ZIP entpacken → temporäres Verzeichnis
        3. Alle JSON-Dateien mit DPAPI re-encrypten und in UserData-Verzeichnis schreiben
    #>
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [Parameter(Mandatory)][System.Security.SecureString]$Passphrase
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $tempDir = Join-Path -Path $env:TEMP -ChildPath "BT_import_$timestamp"
    $tempZip = Join-Path -Path $env:TEMP -ChildPath "BT_import_$timestamp.zip"

    try {
        # 1. Datei lesen: [Salt 32][IV 16][Cipher n]
        $fileBytes = [System.IO.File]::ReadAllBytes($InputPath)
        if ($fileBytes.Length -lt 49) { throw "Ungültige Backup-Datei (zu klein)." }

        $salt = $fileBytes[0..31]
        $iv = $fileBytes[32..47]
        $cipherBytes = $fileBytes[48..($fileBytes.Length - 1)]

        # Passphrase → Schlüssel
        $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Passphrase)
        $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

        $deriveBytes = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($plainPass, [byte[]]$salt, 100000, "SHA256")
        $key = $deriveBytes.GetBytes(32)
        $deriveBytes.Dispose()
        $plainPass = $null

        # AES entschlüsseln
        $aes = [System.Security.Cryptography.Aes]::Create()
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.Key = $key
        $aes.IV = [byte[]]$iv

        $decryptor = $aes.CreateDecryptor()
        $zipBytes = $decryptor.TransformFinalBlock([byte[]]$cipherBytes, 0, $cipherBytes.Length)
        $decryptor.Dispose()
        $aes.Dispose()

        # 2. ZIP entpacken
        [System.IO.File]::WriteAllBytes($tempZip, $zipBytes)
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force }
        [System.IO.Compression.ZipFile]::ExtractToDirectory($tempZip, $tempDir)

        # 3. Config.json → DPAPI re-encrypt
        $importedConfig = Join-Path $tempDir "Config.json"
        if (Test-Path $importedConfig) {
            $json = Get-Content -Path $importedConfig -Raw -Encoding UTF8
            Write-ProtectedJsonFile -Path $script:dataFile -JsonString $json
        }

        # Tagesdaten → DPAPI re-encrypt
        $importedDataDir = Join-Path $tempDir "data"
        if (Test-Path $importedDataDir) {
            # Bestehende Tagesdaten löschen (vollständiger Import)
            if (Test-Path $script:dailyDataDirBase) {
                Remove-Item $script:dailyDataDirBase -Recurse -Force -ErrorAction SilentlyContinue
            }
            $importedFiles = Get-ChildItem -Path $importedDataDir -Recurse -File -ErrorAction SilentlyContinue
            foreach ($file in $importedFiles) {
                $relativePath = $file.FullName.Substring($importedDataDir.Length).TrimStart('\','/')
                $targetPath = Join-Path $script:dailyDataDirBase $relativePath
                $targetDir = Split-Path $targetPath -Parent
                if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }

                if ($file.Extension -eq ".json") {
                    $json = Get-Content -Path $file.FullName -Raw -Encoding UTF8
                    Write-ProtectedJsonFile -Path $targetPath -JsonString $json
                } else {
                    Copy-Item -Path $file.FullName -Destination $targetPath -Force
                }
            }
        }

        return $true
    } catch {
        throw $_
    } finally {
        if (Test-Path $tempDir)  { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tempZip)  { Remove-Item $tempZip -Force -ErrorAction SilentlyContinue }
    }
}

function Export-PlaintextBackup {
    <#
    .SYNOPSIS
        Exportiert alle Daten UNVERSCHLÜSSELT als portables ZIP-Archiv (für PC-Migration).
    .DESCRIPTION
        Entschlüsselt Config + alle Tagesdaten (DPAPI) und schreibt sie als Klartext-JSON
        in ein ZIP. Auf dem Ziel-PC koennen die Dateien direkt in den Datenordner entpackt
        werden; der Blood-Tracker erkennt Klartext automatisch (Auto-Migration) und
        verschluesselt sie beim naechsten Speichern mit dem dortigen DPAPI-Profil neu.
        ACHTUNG: Das ZIP enthaelt sensible Gesundheitsdaten im Klartext.
    #>
    param([Parameter(Mandatory)][string]$OutputPath)
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $tempDir = Join-Path -Path $env:TEMP -ChildPath "BT_plain_$timestamp"
    try {
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        # Config.json (entschluesselt)
        if (Test-Path $script:dataFile) {
            $json = Read-ProtectedJsonFile -Path $script:dataFile
            if ($json) { Set-Content -Path (Join-Path $tempDir "Config.json") -Value $json -Encoding UTF8 }
        }

        # Tagesdaten (entschluesselt) + Dokumente
        if (Test-Path $script:dailyDataDirBase) {
            $dailyFiles = Get-ChildItem -Path $script:dailyDataDirBase -Filter "*.json" -Recurse -ErrorAction SilentlyContinue
            foreach ($file in $dailyFiles) {
                $json = Read-ProtectedJsonFile -Path $file.FullName
                if ($json) {
                    $relativePath = $file.FullName.Substring($script:dailyDataDirBase.Length).TrimStart('\','/')
                    $targetPath = Join-Path (Join-Path $tempDir "data") $relativePath
                    $targetDir = Split-Path $targetPath -Parent
                    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                    Set-Content -Path $targetPath -Value $json -Encoding UTF8
                }
            }
            $otherFiles = Get-ChildItem -Path $script:dailyDataDirBase -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -ne ".json" }
            foreach ($file in $otherFiles) {
                $relativePath = $file.FullName.Substring($script:dailyDataDirBase.Length).TrimStart('\','/')
                $targetPath = Join-Path (Join-Path $tempDir "data") $relativePath
                $targetDir = Split-Path $targetPath -Parent
                if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
                Copy-Item -Path $file.FullName -Destination $targetPath -Force
            }
        }

        $readme = @"
Blood-Tracker - Unverschluesselter Migrations-Export
=====================================================
Erstellt: $(Get-Date -Format 'yyyy-MM-dd HH:mm')

INHALT:
  Config.json                -> Konfiguration, persoenliche Metriken, Marker
  data\JAHR\MONAT\*.json      -> Tagesdaten (Blutwerte) im Klartext

MIGRATION AUF NEUEN PC:
  1. Blood-Tracker auf dem neuen PC einmal starten und wieder schliessen
     (damit der Datenordner angelegt wird).
  2. Datenordner oeffnen: %UserProfile%\PSC.Blood-Tracker
     (bzw. den in HKCU:\Software\PSC\PSC.Blood-Tracker\DataPath hinterlegten Pfad).
  3. Config.json und den kompletten Ordner "data" aus diesem ZIP dorthin entpacken
     (vorhandene Dateien ueberschreiben).
  4. Blood-Tracker starten. Die Klartext-Dateien werden automatisch erkannt und
     beim naechsten Speichern mit dem lokalen DPAPI-Profil verschluesselt.

WARNUNG: Diese Dateien sind UNVERSCHLUESSELT. Nach erfolgreicher Migration
das ZIP sicher loeschen.
"@
        Set-Content -Path (Join-Path $tempDir "LIESMICH_Migration.txt") -Value $readme -Encoding UTF8

        if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tempDir, $OutputPath)
        return $true
    } catch {
        throw $_
    } finally {
        if (Test-Path $tempDir) { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Show-PassphraseDialog {
    <#
    .SYNOPSIS
        Zeigt einen Passwort-Dialog an (mit optionaler Bestätigung).
    .PARAMETER Title
        Fenstertitel.
    .PARAMETER Confirm
        Wenn $true, wird ein zweites Feld zur Bestätigung angezeigt.
    .OUTPUTS
        SecureString oder $null bei Abbruch.
    #>
    param(
        [string]$Title = "Passwort eingeben",
        [bool]$Confirm = $false
    )
    $ppForm = New-Object System.Windows.Forms.Form
    $ppForm.Text = $Title
    $ppForm.Size = New-Object System.Drawing.Size(420, $(if ($Confirm) { 210 } else { 170 }))
    $ppForm.StartPosition = "CenterParent"
    $ppForm.FormBorderStyle = "FixedDialog"
    $ppForm.MaximizeBox = $false; $ppForm.MinimizeBox = $false

    $lbl1 = New-Object System.Windows.Forms.Label
    $lbl1.Text = "Passwort:"
    $lbl1.Location = New-Object System.Drawing.Point(15, 18)
    $lbl1.Size = New-Object System.Drawing.Size(80, 20)
    $ppForm.Controls.Add($lbl1)

    $txt1 = New-Object System.Windows.Forms.TextBox
    $txt1.UseSystemPasswordChar = $true
    $txt1.Location = New-Object System.Drawing.Point(100, 15)
    $txt1.Size = New-Object System.Drawing.Size(280, 22)
    $ppForm.Controls.Add($txt1)

    $txt2 = $null
    if ($Confirm) {
        $lbl2 = New-Object System.Windows.Forms.Label
        $lbl2.Text = "Bestätigen:"
        $lbl2.Location = New-Object System.Drawing.Point(15, 50)
        $lbl2.Size = New-Object System.Drawing.Size(80, 20)
        $ppForm.Controls.Add($lbl2)

        $txt2 = New-Object System.Windows.Forms.TextBox
        $txt2.UseSystemPasswordChar = $true
        $txt2.Location = New-Object System.Drawing.Point(100, 47)
        $txt2.Size = New-Object System.Drawing.Size(280, 22)
        $ppForm.Controls.Add($txt2)
    }

    $btnY = if ($Confirm) { 85 } else { 55 }
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Location = New-Object System.Drawing.Point(210, $btnY)
    $btnOk.Size = New-Object System.Drawing.Size(80, 28)
    $ppForm.Controls.Add($btnOk)
    $ppForm.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(300, $btnY)
    $btnCancel.Size = New-Object System.Drawing.Size(80, 28)
    $ppForm.Controls.Add($btnCancel)
    $ppForm.CancelButton = $btnCancel

    $script:ppResult = $null
    $btnOk.Add_Click({
        $pw = $txt1.Text
        if ($pw.Length -lt 6) {
            [System.Windows.Forms.MessageBox]::Show("Das Passwort muss mindestens 6 Zeichen lang sein.", "Zu kurz", "OK", "Warning")
            return
        }
        if ($Confirm -and $txt2 -and $pw -ne $txt2.Text) {
            [System.Windows.Forms.MessageBox]::Show("Die Passwörter stimmen nicht überein.", "Keine Übereinstimmung", "OK", "Warning")
            return
        }
        $ss = New-Object System.Security.SecureString
        foreach ($c in $pw.ToCharArray()) { $ss.AppendChar($c) }
        $ss.MakeReadOnly()
        $script:ppResult = $ss
        $ppForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $ppForm.Close()
    })
    $btnCancel.Add_Click({
        $ppForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $ppForm.Close()
    })

    $ppForm.ShowDialog() | Out-Null
    return $script:ppResult
}
# =====================================================================

# =====================================================================
# v2.19.0: PDF-Blutbild-Import (Beta Version) – Nachtrag für Template-Script
# =====================================================================

function Initialize-PdfLibrary {
    <#
    .SYNOPSIS
        Lädt itextsharp.dll – bei Bedarf automatisch von NuGet heruntergeladen.
    .OUTPUTS
        $true wenn Bibliothek geladen, $false bei Fehler.
    #>
    $libDir = Join-Path -Path $script:dataDir -ChildPath "lib"
    $dllPath = Join-Path -Path $libDir -ChildPath "itextsharp.dll"

    # Bereits geladen?
    $loaded = [AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq "itextsharp" }
    if ($loaded) { return $true }

    # DLL vorhanden?
    if (Test-Path $dllPath) {
        try {
            Add-Type -Path $dllPath -ErrorAction Stop
            return $true
        } catch {
            Write-Warning "itextsharp.dll konnte nicht geladen werden: $($_.Exception.Message)"
            return $false
        }
    }

    # Download von NuGet
    try {
        if (-not (Test-Path $libDir)) { New-Item -ItemType Directory -Path $libDir -Force | Out-Null }
        $nugetUrl = "https://www.nuget.org/api/v2/package/iTextSharp/5.5.13.3"
        $nupkgPath = Join-Path -Path $env:TEMP -ChildPath "itextsharp.5.5.13.3.nupkg"

        $wc = New-Object System.Net.WebClient
        $wc.DownloadFile($nugetUrl, $nupkgPath)
        $wc.Dispose()

        # NuGet-Paket ist ein ZIP – DLL extrahieren
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::OpenRead($nupkgPath)
        $entry = $zip.Entries | Where-Object { $_.FullName -like "lib/net4*/itextsharp.dll" } | Select-Object -First 1
        if (-not $entry) {
            $zip.Dispose()
            throw "itextsharp.dll nicht im NuGet-Paket gefunden."
        }
        $stream = $entry.Open()
        $fileStream = [System.IO.File]::Create($dllPath)
        $stream.CopyTo($fileStream)
        $fileStream.Close()
        $stream.Close()
        $zip.Dispose()
        Remove-Item $nupkgPath -Force -ErrorAction SilentlyContinue

        Add-Type -Path $dllPath -ErrorAction Stop
        return $true
    } catch {
        Write-Warning "PDF-Bibliothek Download fehlgeschlagen: $($_.Exception.Message)"
        return $false
    }
}

function Extract-PdfText {
    <#
    .SYNOPSIS
        Extrahiert den gesamten Text aus einer PDF-Datei (Beta Version).
    .OUTPUTS
        String mit dem extrahierten Text oder $null bei Fehler.
    #>
    param([Parameter(Mandatory)][string]$PdfPath)
    try {
        $reader = New-Object iTextSharp.text.pdf.PdfReader($PdfPath)
        $sb = New-Object System.Text.StringBuilder
        for ($i = 1; $i -le $reader.NumberOfPages; $i++) {
            $pageText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i)
            [void]$sb.AppendLine($pageText)
        }
        $reader.Close()
        return $sb.ToString()
    } catch {
        Write-Warning "PDF-Textextraktion fehlgeschlagen: $($_.Exception.Message)"
        return $null
    }
}

function Build-MarkerAliasMap {
    <#
    .SYNOPSIS
        Erstellt eine Mapping-Tabelle: Regex-Alias → Marker-Config-Name.
    .DESCRIPTION
        Aus "Alanin-Aminotransferase (ALT/GPT)" werden die Aliases:
        "ALT", "GPT", "Alanin-Aminotransferase" generiert.
    #>
    $map = @()
    foreach ($marker in $script:data.Config.Markers) {
        $name = $marker.Name
        $unit = $marker.Unit
        $aliases = @()

        # Kürzel aus Klammern extrahieren: "Name (ABC/DEF)" → ABC, DEF
        if ($name -match '\(([^)]+)\)') {
            $innerParts = $Matches[1] -split '[/,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
            $aliases += $innerParts
        }

        # Langname ohne Klammern
        $baseName = ($name -replace '\s*\([^)]*\)\s*', '').Trim()
        if ($baseName) { $aliases += $baseName }

        # Vollständiger Name als Fallback
        $aliases += $name

        # Duplikate entfernen, längste zuerst (Priorität)
        $aliases = $aliases | Select-Object -Unique | Sort-Object { $_.Length } -Descending

        $map += @{
            ConfigName = $name
            Unit       = $unit
            Aliases    = $aliases
            RefMin     = $marker.RefMin
            RefMax     = $marker.RefMax
        }
    }
    return $map
}

function Parse-BloodValuesFromText {
    <#
    .SYNOPSIS
        Parst extrahierten PDF-Text und findet Marker-Wert-Zuordnungen (Beta Version).
    .DESCRIPTION
        Sucht für jeden bekannten Marker nach dessen Aliases im Text und
        extrahiert den zugehörigen numerischen Wert. Unterstützt Komma-
        und Punkt-Dezimaltrennzeichen.
    .OUTPUTS
        Array von @{ MarkerName; Value; Unit; Confidence; Line }
    #>
    param([Parameter(Mandatory)][string]$Text)

    $aliasMap = Build-MarkerAliasMap
    $results = @()
    $alreadyMatched = @{} # Verhindert Doppel-Matches

    # Text in Zeilen aufteilen und bereinigen
    $lines = $Text -split "`r?`n" | Where-Object { $_.Trim() }

    foreach ($entry in $aliasMap) {
        if ($alreadyMatched.ContainsKey($entry.ConfigName)) { continue }
        # v2.26.0: qualitative Marker (HIV, APOE) nie automatisch befuellen
        if ($script:PdfExcludedMarkers -contains $entry.ConfigName) { continue }

        foreach ($alias in $entry.Aliases) {
            if ($alreadyMatched.ContainsKey($entry.ConfigName)) { break }

            # Regex: Alias, gefolgt von optionalem Trennzeichen, dann Zahl
            # Unterstützt: "ALT  42", "ALT: 42", "ALT   42,5", "ALT 4.2"
            $escapedAlias = [regex]::Escape($alias)
            $pattern = "(?i)(?<!\w)${escapedAlias}(?!\w)\s*[:\.\s]*\s*(?<value>\d+[.,]?\d*)"

            foreach ($line in $lines) {
                # v2.25.0: Zeilen ueberspringen, die fachlich zu einem ANDEREN Marker gehoeren
                # (z. B. "hs-CRP 0,8" darf nicht vom klassischen CRP eingesammelt werden).
                if ($script:PdfSkipLinePatterns.ContainsKey($entry.ConfigName) -and
                    ($line -match $script:PdfSkipLinePatterns[$entry.ConfigName])) { continue }
                if ($line -match $pattern) {
                    $rawValue = $Matches['value'] -replace ',', '.'
                    $numValue = 0.0
                    if ([double]::TryParse($rawValue, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numValue)) {
                        # Plausibilitätsprüfung: Wert sollte in einem realistischen Bereich liegen
                        $confidence = "Hoch"
                        if ($entry.RefMin -and $entry.RefMax) {
                            $rangeSpan = $entry.RefMax - $entry.RefMin
                            if ($rangeSpan -gt 0) {
                                $lowerBound = $entry.RefMin - ($rangeSpan * 3)
                                $upperBound = $entry.RefMax + ($rangeSpan * 3)
                                if ($numValue -lt $lowerBound -or $numValue -gt $upperBound) {
                                    $confidence = "Niedrig"
                                }
                            }
                        }

                        $results += @{
                            MarkerName = $entry.ConfigName
                            Value      = $numValue
                            Unit       = $entry.Unit
                            Confidence = $confidence
                            MatchedBy  = $alias
                            SourceLine = $line.Trim().Substring(0, [Math]::Min($line.Trim().Length, 80))
                        }
                        $alreadyMatched[$entry.ConfigName] = $true
                        break
                    }
                }
            }
        }
    }

    return $results
}

function Show-PdfImportReviewDialog {
    <#
    .SYNOPSIS
        Review-Dialog für PDF-Import: Zeigt erkannte Werte zur Bestätigung (Beta Version).
    .OUTPUTS
        Array der bestätigten Werte oder $null bei Abbruch.
    #>
    param(
        [Parameter(Mandatory)][array]$ParsedValues,
        [Parameter(Mandatory)][string]$PdfFileName
    )

    $reviewForm = New-Object System.Windows.Forms.Form
    $reviewForm.Text = "PDF-Import Review (Beta Version) – $PdfFileName"
    $reviewForm.Size = New-Object System.Drawing.Size(850, 620)
    $reviewForm.StartPosition = "CenterParent"
    $reviewForm.FormBorderStyle = "FixedDialog"
    $reviewForm.MaximizeBox = $false; $reviewForm.MinimizeBox = $false

    # Warnhinweis
    $lblWarning = New-Object System.Windows.Forms.Label
    $lblWarning.Text = "⚠ Beta Version – Bitte prüfen Sie jeden erkannten Wert sorgfältig gegen das Original-PDF!"
    $lblWarning.Location = New-Object System.Drawing.Point(15, 12)
    $lblWarning.Size = New-Object System.Drawing.Size(800, 20)
    $lblWarning.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $lblWarning.ForeColor = [System.Drawing.Color]::DarkRed
    $reviewForm.Controls.Add($lblWarning)

    # Datum-Auswahl
    $lblDate = New-Object System.Windows.Forms.Label
    $lblDate.Text = "Datum des Bluttests:"
    $lblDate.Location = New-Object System.Drawing.Point(15, 42)
    $lblDate.Size = New-Object System.Drawing.Size(140, 20)
    $reviewForm.Controls.Add($lblDate)

    $dtPicker = New-Object System.Windows.Forms.DateTimePicker
    $dtPicker.Format = [System.Windows.Forms.DateTimePickerFormat]::Short
    $dtPicker.Location = New-Object System.Drawing.Point(160, 39)
    $dtPicker.Size = New-Object System.Drawing.Size(120, 22)
    $reviewForm.Controls.Add($dtPicker)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Text = "$($ParsedValues.Count) Werte erkannt"
    $lblCount.Location = New-Object System.Drawing.Point(300, 42)
    $lblCount.Size = New-Object System.Drawing.Size(200, 20)
    $lblCount.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $reviewForm.Controls.Add($lblCount)

    # DataGridView für Review
    $grid = New-Object System.Windows.Forms.DataGridView
    $grid.Location = New-Object System.Drawing.Point(15, 70)
    $grid.Size = New-Object System.Drawing.Size(805, 450)
    $grid.AllowUserToAddRows = $false
    $grid.AllowUserToDeleteRows = $false
    $grid.AutoSizeColumnsMode = "Fill"
    $grid.SelectionMode = "FullRowSelect"

    # Spalten
    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.HeaderText = "✓"; $colCheck.Name = "Import"; $colCheck.Width = 35; $colCheck.AutoSizeMode = "None"
    $grid.Columns.Add($colCheck) | Out-Null

    $colMarker = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colMarker.HeaderText = "Marker"; $colMarker.Name = "Marker"; $colMarker.ReadOnly = $true; $colMarker.FillWeight = 35
    $grid.Columns.Add($colMarker) | Out-Null

    $colValue = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colValue.HeaderText = "Wert"; $colValue.Name = "Value"; $colValue.FillWeight = 12
    $grid.Columns.Add($colValue) | Out-Null

    $colUnit = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colUnit.HeaderText = "Einheit"; $colUnit.Name = "Unit"; $colUnit.ReadOnly = $true; $colUnit.FillWeight = 10
    $grid.Columns.Add($colUnit) | Out-Null

    $colConf = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colConf.HeaderText = "Konfidenz"; $colConf.Name = "Confidence"; $colConf.ReadOnly = $true; $colConf.FillWeight = 10
    $grid.Columns.Add($colConf) | Out-Null

    $colSource = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colSource.HeaderText = "Quelle (PDF-Zeile)"; $colSource.Name = "Source"; $colSource.ReadOnly = $true; $colSource.FillWeight = 33
    $grid.Columns.Add($colSource) | Out-Null

    # Daten befüllen
    foreach ($val in ($ParsedValues | Sort-Object { $_.MarkerName })) {
        $rowIdx = $grid.Rows.Add($true, $val.MarkerName, $val.Value, $val.Unit, $val.Confidence, $val.SourceLine)
        if ($val.Confidence -eq "Niedrig") {
            $grid.Rows[$rowIdx].DefaultCellStyle.BackColor = [System.Drawing.Color]::LemonChiffon
            $grid.Rows[$rowIdx].Cells["Import"].Value = $false # Niedrige Konfidenz: default nicht importieren
        }
    }
    $reviewForm.Controls.Add($grid)

    # Buttons
    $btnAll = New-Object System.Windows.Forms.Button
    $btnAll.Text = "Alle auswählen"
    $btnAll.Location = New-Object System.Drawing.Point(15, 530)
    $btnAll.Size = New-Object System.Drawing.Size(120, 28)
    $reviewForm.Controls.Add($btnAll)
    $btnAll.Add_Click({ foreach ($row in $grid.Rows) { $row.Cells["Import"].Value = $true } })

    $btnNone = New-Object System.Windows.Forms.Button
    $btnNone.Text = "Alle abwählen"
    $btnNone.Location = New-Object System.Drawing.Point(145, 530)
    $btnNone.Size = New-Object System.Drawing.Size(120, 28)
    $reviewForm.Controls.Add($btnNone)
    $btnNone.Add_Click({ foreach ($row in $grid.Rows) { $row.Cells["Import"].Value = $false } })

    $btnImport = New-Object System.Windows.Forms.Button
    $btnImport.Text = "Ausgewählte importieren"
    $btnImport.Location = New-Object System.Drawing.Point(570, 530)
    $btnImport.Size = New-Object System.Drawing.Size(120, 28)
    $reviewForm.Controls.Add($btnImport)
    $reviewForm.AcceptButton = $btnImport

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(700, 530)
    $btnCancel.Size = New-Object System.Drawing.Size(120, 28)
    $reviewForm.Controls.Add($btnCancel)
    $reviewForm.CancelButton = $btnCancel

    $script:pdfImportResult = $null
    $btnImport.Add_Click({
        $confirmed = @()
        $dateStr = $dtPicker.Value.ToString("yyyy-MM-dd")
        foreach ($row in $grid.Rows) {
            if ($row.Cells["Import"].Value -eq $true) {
                $valText = [string]$row.Cells["Value"].Value -replace ',', '.'
                $numVal = 0.0
                if ([double]::TryParse($valText, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$numVal)) {
                    $confirmed += @{
                        Date  = $dateStr
                        Name  = $row.Cells["Marker"].Value
                        Value = $numVal
                        Unit  = $row.Cells["Unit"].Value
                        Note  = "PDF-Import (Beta)"
                    }
                }
            }
        }
        if ($confirmed.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Werte zum Import ausgewählt.", "Keine Auswahl", "OK", "Information")
            return
        }
        $script:pdfImportResult = $confirmed
        $reviewForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $reviewForm.Close()
    })
    $btnCancel.Add_Click({
        $reviewForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $reviewForm.Close()
    })

    $reviewForm.ShowDialog() | Out-Null
    return $script:pdfImportResult
}

function Import-PdfBloodValues {
    <#
    .SYNOPSIS
        Hauptfunktion: PDF öffnen → Text extrahieren → parsen → Review → importieren (Beta Version).
    #>
    param([string]$PreselectedDate)

    # 1. Bibliothek laden
    $libOk = Initialize-PdfLibrary
    if (-not $libOk) {
        [System.Windows.Forms.MessageBox]::Show(
            "Die PDF-Bibliothek (iTextSharp) konnte nicht geladen werden.`n`nMögliche Ursachen:`n• Keine Internetverbindung für den automatischen Download`n• Firewall blockiert den Zugriff auf nuget.org`n`nSie können die Datei manuell herunterladen und unter`n$($script:dataDir)\lib\itextsharp.dll ablegen.",
            "PDF-Import – Bibliothek fehlt (Beta Version)", "OK", "Error"
        )
        return
    }

    # 2. PDF-Datei auswählen
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Title = "Blutbild-PDF auswählen (Beta Version)"
    $ofd.Filter = "PDF-Dateien (*.pdf)|*.pdf"
    if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

    # 3. Text extrahieren
    $text = Extract-PdfText -PdfPath $ofd.FileName
    if (-not $text -or $text.Trim().Length -lt 20) {
        [System.Windows.Forms.MessageBox]::Show(
            "Aus dieser PDF konnte kein verwertbarer Text extrahiert werden.`n`nDies kann passieren wenn das PDF gescannt (Bild) statt textbasiert ist.`nBitte geben Sie die Werte in diesem Fall manuell ein.",
            "Kein Text erkannt (Beta Version)", "OK", "Warning"
        )
        return
    }

    # 4. Werte parsen
    $parsed = Parse-BloodValuesFromText -Text $text
    if (-not $parsed -or $parsed.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Es konnten keine bekannten Blutmarker im PDF erkannt werden.`n`nMöglicherweise verwendet dieses Labor ein unbekanntes Format.`nBitte geben Sie die Werte manuell ein.",
            "Keine Marker erkannt (Beta Version)", "OK", "Warning"
        )
        return
    }

    # 5. Review-Dialog
    $pdfName = [System.IO.Path]::GetFileName($ofd.FileName)
    $confirmed = Show-PdfImportReviewDialog -ParsedValues $parsed -PdfFileName $pdfName
    if (-not $confirmed -or $confirmed.Count -eq 0) { return }

    # 6. Werte importieren
    $dateStr = $confirmed[0].Date
    $importCount = 0
    foreach ($item in $confirmed) {
        try {
            $newItem = [PSCustomObject]@{
                Date  = $item.Date
                Name  = $item.Name
                Value = [double]$item.Value
                Unit  = $item.Unit
                Note  = $item.Note
            }
            Save-Data -data @{ NewItem = $newItem } -type "Daily" -dateString $item.Date
            $importCount++
        } catch {
            Write-Warning "Import-Fehler bei $($item.Name): $($_.Exception.Message)"
        }
    }

    # 7. Daten neu laden
    $script:allHistoricalItems = [System.Collections.ArrayList]@(Load-AllHistoricalItems)

    # 8. PDF-Dokument zusätzlich archivieren
    try {
        $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $dateStr
        if (-not (Test-Path $dailyDataDir)) { New-Item -Path $dailyDataDir -ItemType Directory -Force | Out-Null }
        $targetName = "${dateStr}_BLUTTEST.pdf"
        $targetPath = Join-Path -Path $dailyDataDir -ChildPath $targetName
        $counter = 1
        while (Test-Path $targetPath) {
            $targetName = "${dateStr}_BLUTTEST_${counter}.pdf"
            $targetPath = Join-Path -Path $dailyDataDir -ChildPath $targetName
            $counter++
        }
        Copy-Item -Path $ofd.FileName -Destination $targetPath -Force
    } catch {
        Write-Warning "PDF-Archivierung fehlgeschlagen: $($_.Exception.Message)"
    }

    [System.Windows.Forms.MessageBox]::Show(
        "$importCount Blutwerte erfolgreich importiert (Datum: $dateStr).`n`nDas Original-PDF wurde im Tagesordner archiviert.`n`n⚠ Beta Version: Bitte gleichen Sie die importierten Werte mit dem Original-Befund ab.",
        "PDF-Import erfolgreich (Beta Version)", "OK", "Information"
    )
}

# =====================================================================

function Get-DailyDataFilePath {
    param($dateString)
    try {
        $date = [datetime]::ParseExact($dateString, "yyyy-MM-dd", $null)
        $year = $date.ToString("yyyy"); $month = $date.ToString("MM"); $day = $date.ToString("yyyy.MM.dd")
        $dailyDataDir = Join-Path -Path $script:dailyDataDirBase -ChildPath $year | Join-Path -ChildPath $month
        $dailyDataFile = Join-Path -Path $dailyDataDir -ChildPath "$day.json"
        return $dailyDataFile, $dailyDataDir
    } catch { throw "Ungültiges Datumsformat: $dateString" }
}

function Get-BloodTestDocuments {
    <#
    .SYNOPSIS
        v2.24.1: Liefert alle archivierten Dokumente EINES Bluttest-Datums.
    .DESCRIPTION
        Die Tagesdaten liegen physisch in einem MONATS-Ordner (data\YYYY\MM\).
        Bis v2.24.0 wurde dort mit "*_BLUTTEST*" gefiltert - dadurch wurden die
        Dokumente ALLER Bluttests desselben Monats getroffen (falsche Anzahl im
        Button, Öffnen des falschen Dokuments).

        Diese Funktion filtert strikt auf den Datums-Prefix des gewählten Tests:
            YYYY-MM-DD_BLUTTEST[_n][.ext]
        Sortierung nach laufender Nummer = Upload-Reihenfolge (ohne Suffix = 1.).
    .PARAMETER DateString
        Bluttest-Datum im Format yyyy-MM-dd (entspricht $treeNode.Tag).
    .OUTPUTS
        Array von System.IO.FileInfo (leeres Array, wenn keine Dokumente existieren).
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$DateString
    )
    try {
        if ([string]::IsNullOrWhiteSpace($DateString)) { return @() }
        $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $DateString
        if (-not (Test-Path $dailyDataDir)) { return @() }

        # Strikter Datumsbezug: nur Dokumente DIESES Tests, nicht des Monats
        $pattern = '^' + [regex]::Escape($DateString) + '_BLUTTEST(_(?<idx>\d+))?(\..+)?$'

        $files = @(Get-ChildItem -Path $dailyDataDir -File -ErrorAction SilentlyContinue |
                   Where-Object { $_.Name -match $pattern })
        if ($files.Count -eq 0) { return @() }

        return @($files | Sort-Object `
                    @{ Expression = {
                            if ($_.Name -match $pattern -and $Matches['idx']) { [int]$Matches['idx'] } else { 0 }
                        }
                     }, `
                    @{ Expression = { $_.Name } })
    } catch {
        Write-Warning "Dokument-Erkennung für $DateString fehlgeschlagen: $($_.Exception.Message)"
        return @()
    }
}

function ConvertTo-Hashtable {
    param ([Parameter(ValueFromPipeline)] $InputObject)
    $hash = @{}
    $properties = $null
    if ($InputObject -is [System.Management.Automation.PSCustomObject] -or $InputObject.PSObject -ne $null) {
        $properties = $InputObject.PSObject.Properties
    } elseif ($InputObject -is [hashtable]) {
        $properties = $InputObject.GetEnumerator()
    } else {
        return $InputObject
    }
    
    foreach ($prop in $properties) {
        $name = if ($prop -is [System.Collections.DictionaryEntry]) { $prop.Key } else { $prop.Name }
        $value = if ($prop -is [System.Collections.DictionaryEntry]) { $prop.Value } else { $prop.Value }

        if ($null -eq $value) {
            # v2.17.0 Bugfix: $null-Werte direkt übernehmen, sonst crasht $value.GetType()
            $hash[$name] = $null
        } elseif ($value -is [System.Management.Automation.PSCustomObject] -or $value -is [hashtable]) {
            $hash[$name] = ConvertTo-Hashtable -InputObject $value
        } elseif ($value -is [Array] -or $value -is [System.Collections.ArrayList] -or $value.GetType().IsArray) {
            $arrayList = New-Object System.Collections.ArrayList
            foreach ($item in $value) {
                if ($item -is [System.Management.Automation.PSCustomObject] -or $item -is [hashtable]) {
                    $arrayList.Add((ConvertTo-Hashtable -InputObject $item)) | Out-Null
                } else {
                    $arrayList.Add($item) | Out-Null
                }
            }
            $hash[$name] = $arrayList
        } else {
            $hash[$name] = $value
        }
    }
    return $hash
}

function Load-Config {
    $defaultConfig = @{
        Markers = [System.Collections.ArrayList]@(
            [PSCustomObject]@{ Name = 'Albumin'; Unit = 'g/dl'; RefMin = 3.5; RefMax = 5.2; OptimalMin = 4.0; OptimalMax = 5.0; Group = "Proteine"; Description = "Wichtigstes Bluteiweiß; Indikator für Ernährungsstatus und Leber-/Nierenfunktion." },
            [PSCustomObject]@{ Name = 'Alanin-Aminotransferase (ALT/GPT)'; Unit = 'U/l'; RefMin = 0; RefMax = 50; OptimalMin = 10; OptimalMax = 30; Group = "Leber & Galle"; Description = "Ein primärer Lebermarker. Erhöhte Werte deuten auf Leberzellschäden hin." },
            [PSCustomObject]@{ Name = 'Alkalische Phosphatase (AP)'; Unit = 'U/l'; RefMin = 40; RefMax = 130; OptimalMin = 50; OptimalMax = 100; Group = "Leber & Galle"; Description = "Enzym aus Leber, Knochen und Galle. Erhöht bei Knochenumbau, Lebererkrankungen oder Cholestase. Bestandteil der PhenoAge-Berechnung (Levine 2018)." },
            [PSCustomObject]@{ Name = 'Anteil großer Thrombozyten (P-LCR)'; Unit = '%'; RefMin = 13; RefMax = 43; Group = "Blutbild"; Description = "Prozentsatz der großen Blutplättchen. Kann Hinweise auf die Thrombozytenproduktion geben." },
            [PSCustomObject]@{ Name = 'ApoB'; Unit = 'mg/dl'; RefMin = 0; RefMax = 90; OptimalMin = 0; OptimalMax = 80; Group = 'Blutfette'; Description = "Apolipoprotein B. Gilt als genauerer Marker für das Risiko von Herz-Kreislauf-Erkrankungen als LDL allein." },
            # v2.26.0: APOE-Genotyp. Speicherung als Code 1-6 (siehe $script:ApoeGenotypeOptions),
            # Eingabe und Anzeige immer als Genotyp-Klartext (Get-ApoeGenotypeText).
            # Referenz 1-3 = kein E4-Allel. Optimal 2-3 = E2/E3 + E3/E3.
            [PSCustomObject]@{ Name = 'Apolipoprotein E-Genotyp (APOE)'; Unit = 'Genotyp'; RefMin = 1; RefMax = 3; OptimalMin = 2; OptimalMax = 3; Group = "Genetik"; Description = "Genetische Anlage - einmalige Bestimmung, verändert sich lebenslang nicht. Kodierung im Tracker nach aufsteigendem Alzheimer-Risiko: 1 = E2/E2, 2 = E2/E3, 3 = E3/E3 (Bezugsgenotyp, häufigster Genotyp), 4 = E2/E4, 5 = E3/E4, 6 = E4/E4. Referenzbereich 1-3 = kein E4-Allel. Risiko für die spät beginnende Alzheimer-Erkrankung relativ zu E3/E3: ein E4-Allel ca. 3-4-fach, zwei E4-Allele ca. 12-15-fach; E2-Träger ca. 40 % niedriger (Farrer et al., JAMA 1997, APOE and Alzheimer Disease Meta Analysis Consortium; Alzheimer's Drug Discovery Foundation). Allelhäufigkeit in Europa: E2 ca. 8-10 %, E3 ca. 70 %, E4 ca. 15-20 %. E4 erhöht zusätzlich LDL/ApoB sowie das Atherosklerose- und Schlaganfallrisiko. E2/E2 ist der typische Genotyp der Hyperlipoproteinämie Typ III (MedlinePlus Genetics, APOE gene, NIH/NLM) - deshalb liegt Code 1 bewusst NICHT im Optimalbereich. WICHTIG: Der Genotyp ist ein Risiko-, kein Diagnosemarker; ein E4-Allel bedeutet KEINE Erkrankung und E3/E3 keinen Schutz. In Deutschland ist vor einer genetischen Untersuchung eine ärztliche Aufklärung bzw. genetische Beratung nach GenDG vorgesehen - Befund und Konsequenzen immer ärztlich besprechen." },
            [PSCustomObject]@{ Name = 'Aspartat-Aminotransferase (AST/GOT)'; Unit = 'U/l'; RefMin = 0; RefMax = 50; OptimalMin = 10; OptimalMax = 30; Group = "Leber & Galle"; Description = "Ein Enzym, das in Leber, Herz und Muskeln vorkommt. Erhöht bei Zellschäden in diesen Bereichen." },
            [PSCustomObject]@{ Name = 'Basophile Granulozyten'; Unit = '%'; RefMin = 0; RefMax = 2; OptimalMin = 0; OptimalMax = 1; Group = "Immunsystem"; Description = "Eine Art weißer Blutkörperchen, beteiligt an allergischen Reaktionen." },
            [PSCustomObject]@{ Name = 'CEA (Carcinoembryonic Antigen)'; Unit = 'ng/ml'; RefMin = 0; RefMax = 3; OptimalMin = 0; OptimalMax = 2.5; Group = 'Krebsvorsorge'; Description = "Ein Tumormarker, der vor allem bei bestimmten Darmkrebsarten erhöht sein kann." },
            [PSCustomObject]@{ Name = 'Cholesterin, gesamt (Chol)'; Unit = 'mg/dl'; RefMin = 0; RefMax = 200; Group = "Blutfette"; Description = "Summe aller Cholesterinarten im Blut." },
            [PSCustomObject]@{ Name = 'Cortisol'; Unit = 'µg/dl'; RefMin = 5; RefMax = 25; OptimalMin = 10; OptimalMax = 20; Group = 'Hormone'; Description = "Das 'Stresshormon'. Wichtig für die Regulierung von Stoffwechsel und Immunantwort." },
            [PSCustomObject]@{ Name = 'C-reaktives Protein (CRP)'; Unit = 'mg/l'; RefMin = 0; RefMax = 5; OptimalMin = 0; OptimalMax = 1; Group = "Entzündung"; Description = "Ein allgemeiner Marker für Entzündungen im Körper." },
            [PSCustomObject]@{ Name = 'Cystatin-C'; Unit = 'mg/l'; RefMin = 0.6; RefMax = 1.0; Group = "Niere"; Description = "Ein genauer Marker für die Nierenfiltrationsrate, unabhängiger von der Muskelmasse als Kreatinin." },
            [PSCustomObject]@{ Name = 'D-Dimer'; Unit = 'ng/ml'; RefMin = 0; RefMax = 500; OptimalMin = 0; OptimalMax = 250; Group = 'Venen'; Description = "Ein Spaltprodukt von Fibrin, das beim Abbau von Blutgerinnseln entsteht. Wichtig bei Thromboseverdacht." },
            [PSCustomObject]@{ Name = 'DHEA-S'; Unit = 'µg/dl'; RefMin = 80; RefMax = 560; OptimalMin = 200; OptimalMax = 400; Group = 'Hormone'; Description = "Eine Vorstufe von Sexualhormonen. Gilt als 'Anti-Aging'-Hormon." },
            [PSCustomObject]@{ Name = 'Eisen (Fe)'; Unit = 'µg/dl'; RefMin = 65; RefMax = 175; OptimalMin = 80; OptimalMax = 150; Group = "Mineralstoffe"; Description = "Wichtig für den Sauerstofftransport im Blut (Bestandteil von Hämoglobin)." },
            [PSCustomObject]@{ Name = 'Eosinophile Granulozyten'; Unit = '%'; RefMin = 0; RefMax = 6; OptimalMin = 1; OptimalMax = 3; Group = "Immunsystem"; Description = "Art weißer Blutkörperchen, oft erhöht bei Allergien, Asthma oder Parasiteninfektionen." },
            # Hämogramm Marker
            [PSCustomObject]@{ Name = 'Leukozyten (WBC)'; Unit = 'Tsd./µl'; RefMin = 4.0; RefMax = 10.0; Group = "Blutbild"; Description = "Weiße Blutkörperchen (White Blood Cells), die Hauptzellen des Immunsystems." },
            [PSCustomObject]@{ Name = 'Erythrozyten (RBC)'; Unit = 'Mio./µl'; RefMin = 4.5; RefMax = 5.9; Group = "Blutbild"; Description = "Rote Blutkörperchen (Red Blood Cells), verantwortlich für den Sauerstofftransport." },
            [PSCustomObject]@{ Name = 'Hämoglobin (HGB/Hb)'; Unit = 'g/dl'; RefMin = 14; RefMax = 18; Group = "Blutbild"; Description = "Der rote Blutfarbstoff (Hemoglobin), der Sauerstoff bindet und transportiert." },
            [PSCustomObject]@{ Name = 'Hämatokrit (HCT/Hk)'; Unit = '%'; RefMin = 0.38; RefMax = 0.49; Group = "Blutbild"; Description = "Anteil der festen Blutzellen (vor allem Erythrozyten) am Gesamtblutvolumen." },
            [PSCustomObject]@{ Name = 'Mittleres Erythrozytenvolumen (MCV)'; Unit = 'fl'; RefMin = 80; RefMax = 96; Group = "Blutbild"; Description = "Durchschnittliches Volumen der roten Blutkörperchen (Mean Corpuscular Volume). Wichtig für die Anämie-Diagnostik." },
            [PSCustomObject]@{ Name = 'Mittlerer Hämoglobingehalt (MCH)'; Unit = 'pg'; RefMin = 27; RefMax = 33; Group = "Blutbild"; Description = "Durchschnittliche Masse des Hämoglobins in einem Erythrozyten (Mean Corpuscular Hemoglobin)." },
            [PSCustomObject]@{ Name = 'Mittlere Hämoglobinkonzentration (MCHC)'; Unit = 'g/dl'; RefMin = 33; RefMax = 36; Group = "Blutbild"; Description = "Durchschnittliche Konzentration des Hämoglobins im Erythrozyten (Mean Corpuscular Hemoglobin Concentration)." },
            [PSCustomObject]@{ Name = 'Thrombozyten (PLT)'; Unit = 'Tsd./µl'; RefMin = 150; RefMax = 400; Group = "Blutbild"; Description = "Blutplättchen (Platelets), die für die Blutgerinnung und Wundheilung entscheidend sind." },
            [PSCustomObject]@{ Name = 'Erythrozyten-Verteilungsbreite (RDW-SD)'; Unit = 'fl'; RefMin = 37; RefMax = 54; Group = "Blutbild"; Description = "Maß für die Größenvariation der Erythrozyten als Standardabweichung (Standard Deviation)." },
            [PSCustomObject]@{ Name = 'Erythrozyten-Verteilungsbreite (RDW-CV)'; Unit = '%'; RefMin = 11.5; RefMax = 14.5; Group = "Blutbild"; Description = "Maß für die Größenvariation der roten Blutkörperchen als Variationskoeffizient (Coefficient of Variation)." },
            [PSCustomObject]@{ Name = 'Thrombozyten-Verteilungsbreite (PDW)'; Unit = 'fl'; RefMin = 9; RefMax = 14; Group = "Blutbild"; Description = "Maß für die Größenvariation der Thrombozyten (Platelet Distribution Width)." },
            [PSCustomObject]@{ Name = 'Mittleres Thrombozytenvolumen (MPV)'; Unit = 'fl'; RefMin = 9.4; RefMax = 12.3; Group = "Blutbild"; Description = "Durchschnittliches Volumen der Thrombozyten (Mean Platelet Volume)." },
            [PSCustomObject]@{ Name = 'Thrombozytenvolumen (PCT)'; Unit = '%'; RefMin = 0.15; RefMax = 0.4; Group = "Blutbild"; Description = "Anteil des Thrombozytenvolumens am Gesamtblutvolumen (Plateletcrit)." },
            [PSCustomObject]@{ Name = 'Ferritin'; Unit = 'ng/ml'; RefMin = 30; RefMax = 400; OptimalMin = 50; OptimalMax = 200; Group = 'Mineralstoffe'; Description = "Speicherform des Eisens. Wichtigster Marker zur Beurteilung des Eisenstatus." },
            [PSCustomObject]@{ Name = 'Freies T3 (fT3)'; Unit = 'pg/ml'; RefMin = 2.0; RefMax = 4.4; OptimalMin = 3.0; OptimalMax = 4.0; Group = "Schilddrüse"; Description = "Das aktive Schilddrüsenhormon, das den Stoffwechsel direkt beeinflusst." },
            [PSCustomObject]@{ Name = 'Freies T4 (fT4)'; Unit = 'ng/dl'; RefMin = 0.9; RefMax = 1.7; OptimalMin = 1.1; OptimalMax = 1.5; Group = "Schilddrüse"; Description = "Das Speicher-Schilddrüsenhormon, das bei Bedarf in aktives fT3 umgewandelt wird." },
            [PSCustomObject]@{ Name = 'Freies Testosteron'; Unit = 'pg/ml'; RefMin = 50; RefMax = 210; OptimalMin = 100; OptimalMax = 180; Group = 'Hormone'; Description = "Der biologisch aktive Anteil des Testosterons, der nicht an Proteine gebunden ist." },
            [PSCustomObject]@{ Name = 'Gamma-Glutamyltransferase (GGT)'; Unit = 'U/l'; RefMin = 0; RefMax = 60; OptimalMin = 10; OptimalMax = 35; Group = "Leber & Galle"; Description = "Sensitiver Marker für Leber- und Gallenerkrankungen. Oft erhöht bei Alkoholkonsum." },
            [PSCustomObject]@{ Name = 'Gewicht'; Unit = 'kg'; Group = 'Lebensstil'; Description = "Ihr aktuelles Körpergewicht in Kilogramm." },
            [PSCustomObject]@{ Name = 'Glomeruläre Filtrationsrate (GFR)'; Unit = 'ml/min'; RefMin = 60; RefMax = 200; OptimalMin = 90; OptimalMax = 200; Group = "Niere"; Description = "Zentraler Wert zur Beurteilung der Nierenfunktion; gibt an, wie gut die Nieren das Blut filtern." },
            [PSCustomObject]@{ Name = 'Glukose (Glu)'; Unit = 'mg/dl'; RefMin = 70; RefMax = 100; OptimalMin = 75; OptimalMax = 90; Group = "Stoffwechsel"; Description = "Blutzucker. Wichtigster Energielieferant für die Körperzellen." },
            [PSCustomObject]@{ Name = 'Harnsäure (UA)'; Unit = 'mg/dl'; RefMin = 3.4; RefMax = 7.0; Group = "Niere"; Description = "Abbauprodukt von Purinen. Erhöhte Werte können zu Gicht führen." },
            [PSCustomObject]@{ Name = 'Harnstoff'; Unit = 'mg/dl'; RefMin = 20; RefMax = 45; Group = "Niere"; Description = "Abbauprodukt des Proteinstoffwechsels, wird über die Nieren ausgeschieden." },
            [PSCustomObject]@{ Name = 'HDL-Cholesterin (HDL)'; Unit = 'mg/dl'; RefMin = 40; RefMax = 100; OptimalMin = 50; OptimalMax = 80; Group = "Blutfette"; Description = "High-Density Lipoprotein, oft als 'gutes' Cholesterin bezeichnet; transportiert Cholesterin zur Leber." },
            [PSCustomObject]@{ Name = 'Homocystein'; Unit = 'µmol/l'; RefMin = 0; RefMax = 15; OptimalMin = 0; OptimalMax = 7; Group = 'Herz-Kreislauf'; Description = "Aminosäure, deren erhöhte Werte als Risikofaktor für Herz-Kreislauf-Erkrankungen gelten." },
            [PSCustomObject]@{ Name = 'HIV (Anti-HIV-1/2)'; Unit = 'Index'; RefMin = 0; RefMax = 1; OptimalMin = 0; OptimalMax = 1; Group = "Infektionskrankheiten"; Description = "HIV-Suchtest (ELISA, 4. Generation). Erfasst Antikörper gegen HIV-1/2 sowie p24-Antigen. Werte < 1 = negativ (unauffällig). Positive Werte müssen durch Bestätigungstest (Western Blot/Immunoblot) abgeklärt werden." },
            # v2.25.0: hs-CRP - Risikostratifizierung nach AHA/CDC (Pearson et al., Circulation 2003),
            # bestätigt in StatPearls "C-Reactive Protein" (NCBI Bookshelf NBK441843, Stand 03.05.2025).
            [PSCustomObject]@{ Name = 'Hochsensitives CRP (hs-CRP)'; Unit = 'mg/l'; RefMin = 0; RefMax = 3; OptimalMin = 0; OptimalMax = 1; Group = "Entzündung"; Description = "Hochsensitives C-reaktives Protein zur Erfassung niedriggradiger (stiller) Entzündung. Kardiovaskuläre Risikostratifizierung nach AHA/CDC (Pearson 2003): < 1 mg/l = niedriges, 1-3 mg/l = mittleres, > 3 mg/l = hohes Risiko. Werte > 10 mg/l sprechen für eine akute Entzündung/Infektion und sind NICHT als Herz-Kreislauf-Risiko verwertbar - Kontrolle nach ca. 2 Wochen im infektfreien Intervall. Messung nüchtern nicht erforderlich, aber Abstand zu Infekten, Impfungen, Verletzungen und intensiver Belastung (jeweils ca. 1-2 Wochen) einhalten." },
            [PSCustomObject]@{ Name = 'Kalium (K)'; Unit = 'mmol/l'; RefMin = 3.5; RefMax = 5.1; Group = "Mineralstoffe"; Description = "Wichtiger Elektrolyt für Nerven- und Muskelfunktion." },
            [PSCustomObject]@{ Name = 'Körperfettanteil (KFA)'; Unit = '%'; Group = 'Lebensstil'; Description = "Prozentualer Anteil des Fettgewebes am Gesamtkörpergewicht." },
            [PSCustomObject]@{ Name = 'Kreatinin (Krea)'; Unit = 'mg/dl'; RefMin = 0.7; RefMax = 1.2; Group = "Niere"; Description = "Abbauprodukt aus dem Muskelstoffwechsels, zur Beurteilung der Nierenfunktion." },
            [PSCustomObject]@{ Name = 'Langzeitzucker (HbA1c)'; Unit = '%'; RefMin = 4.0; RefMax = 5.7; OptimalMin = 4.5; OptimalMax = 5.3; Group = "Stoffwechsel"; Description = "Gibt Auskunft über den durchschnittlichen Blutzuckerspiegel der letzten 8-12 Wochen." },
            [PSCustomObject]@{ Name = 'LDL-Cholesterin (LDL)'; Unit = 'mg/dl'; RefMin = 0; RefMax = 130; OptimalMin = 50; OptimalMax = 100; Group = "Blutfette"; Description = "Low-Density Lipoprotein, oft als 'schlechtes' Cholesterin bezeichnet." },
            [PSCustomObject]@{ Name = 'Lipase'; Unit = 'U/l'; RefMin = 13; RefMax = 60; Group = "Leber & Galle"; Description = "Enzym der Bauchspeicheldrüse zur Fettverdauung. Wichtiger Marker bei Bauchspeicheldrüsenentzündung." },
            [PSCustomObject]@{ Name = 'Lp(a)'; Unit = 'nmol/l'; RefMin = 0; RefMax = 75; OptimalMin = 0; OptimalMax = 50; Group = 'Herz-Kreislauf'; Description = "Lipoprotein(a), ein genetisch bedingter, unabhängiger Risikofaktor für Herz-Kreislauf-Erkrankungen." },
            [PSCustomObject]@{ Name = 'Lymphozyten (LYMPH)'; Unit = '%'; RefMin = 20; RefMax = 45; Group = "Immunsystem"; Description = "Spezialisierte weiße Blutkörperchen für die gezielte Immunabwehr." },
            # v2.26.0: Magnesium - Referenz NIH ODS, Optimalbereich Costello et al. 2016 (Adv Nutr).
            [PSCustomObject]@{ Name = 'Magnesium'; Unit = 'mmol/l'; RefMin = 0.75; RefMax = 0.95; OptimalMin = 0.85; OptimalMax = 0.95; Group = "Mineralstoffe"; Description = "Magnesium im Serum. Referenzbereich 0,75-0,95 mmol/l (entspricht ca. 1,82-2,31 mg/dl); eine Hypomagnesiämie liegt ab < 0,75 mmol/l vor (NIH Office of Dietary Supplements, Magnesium - Health Professional Fact Sheet). Optimalbereich 0,85-0,95 mmol/l nach Costello et al., Advances in Nutrition 2016: unterhalb 0,85 mmol/l steigt in Kohortenstudien das Risiko für Herz-Kreislauf-Erkrankungen und Typ-2-Diabetes (chronisch latenter Mangel). WICHTIGE EINSCHRÄNKUNG: Nur ca. 0,3 % des Körpermagnesiums befinden sich im Serum - ein normaler Serumwert schließt einen Mangel NICHT aus. Aussagekräftiger sind Vollblut- bzw. Erythrozyten-Magnesium oder ein Magnesium-Belastungstest. Praxisrelevanz bei intensivem Krafttraining (Muskelkrämpfe, Schlafqualität), unter Protonenpumpenhemmern oder Diuretika sowie bei Insulinresistenz. Probe muss hämolysefrei sein - Erythrozyten enthalten ein Vielfaches der Serumkonzentration." },
            [PSCustomObject]@{ Name = 'Monozyten'; Unit = '%'; RefMin = 2; RefMax = 10; OptimalMin = 3; OptimalMax = 7; Group = "Immunsystem"; Description = "Die größten weißen Blutkörperchen; entwickeln sich zu Fresszellen (Makrophagen) im Gewebe." },
            [PSCustomObject]@{ Name = 'NAD+'; Unit = 'µM'; RefMin = 20; RefMax = 100; OptimalMin = 30; OptimalMax = 80; Group = "Longevity"; Description = "Ein entscheidendes Coenzym für den zellulären Energiestoffwechsel und die DNA-Reparatur." },
            [PSCustomObject]@{ Name = 'Neutrophile Granulozyten (NEUT)'; Unit = '%'; RefMin = 40; RefMax = 70; Group = "Immunsystem"; Description = "Die häufigsten weißen Blutkörperchen, primär für die Abwehr von bakteriellen Infektionen." },
            [PSCustomObject]@{ Name = 'NT-proBNP'; Unit = 'pg/ml'; RefMin = 0; RefMax = 125; OptimalMin = 0; OptimalMax = 100; Group = 'Herz-Kreislauf'; Description = "Marker zur Diagnose und Verlaufsbeurteilung von Herzinsuffizienz (Herzschwäche)." },
            [PSCustomObject]@{ Name = 'Nüchterninsulin'; Unit = 'µU/ml'; RefMin = 2.6; RefMax = 24.9; OptimalMin = 3; OptimalMax = 8; Group = "Stoffwechsel"; Description = "Insulinspiegel im nüchternen Zustand; Marker für Insulinresistenz." },
            [PSCustomObject]@{ Name = 'Omega-3-Index'; Unit = '%'; RefMin = 4; RefMax = 12; OptimalMin = 8; OptimalMax = 12; Group = 'Blutfette'; Description = "Anteil der Omega-3-Fettsäuren EPA und DHA in den roten Blutkörperchen. Wichtig für die Herzgesundheit." },
            [PSCustomObject]@{ Name = 'PSA (Prostataspezifisches Antigen)'; Unit = 'ng/ml'; RefMin = 0; RefMax = 4; OptimalMin = 0; OptimalMax = 2.5; Group = 'Krebsvorsorge'; Description = "Ein Protein, das von der Prostata gebildet wird. Dient der Früherkennung von Prostatakrebs." },
            # v2.25.0: SHBG - Referenz Roche Elecsys SHBG (ECLIA), Herstellerbeipackzettel.
            # Voreinstellung = Mann 20-49 Jahre. Bewusst KEIN Optimalbereich: dafür existiert
            # keine offizielle Vorgabe (niedrige UND hohe Werte sind klinisch relevant).
            [PSCustomObject]@{ Name = 'Sexualhormon-bindendes Globulin (SHBG)'; Unit = 'nmol/l'; RefMin = 18.3; RefMax = 54.1; Group = 'Hormone'; Description = "Transportprotein, das Testosteron und Östradiol im Blut bindet und damit steuert, wie viel Hormon biologisch aktiv (frei) verfügbar ist. Referenz (Roche Elecsys, ECLIA): Mann 20-49 J. 18,3-54,1 / Mann ab 50 J. 20,6-76,7 / Frau 20-49 J. 32,4-128 / Frau ab 50 J. 27,1-128 nmol/l. Niedriges SHBG ist mit Insulinresistenz, Adipositas, Fettleber und metabolischem Syndrom assoziiert; hohes SHBG senkt das freie Testosteron und tritt u.a. bei Hyperthyreose, Lebererkrankungen, Östrogeneinfluss und im Alter auf. Nur zusammen mit Gesamt-Testosteron und Albumin sinnvoll beurteilbar (Grundlage der Vermeulen-Berechnung des freien Testosterons)." },
            [PSCustomObject]@{ Name = 'Systolischer Blutdruck'; Unit = 'mmHg'; Group = 'Lebensstil'; Description = "Der obere Blutdruckwert; misst den Druck in den Arterien, wenn das Herz schlägt." },
            [PSCustomObject]@{ Name = 'Testosteron, gesamt'; Unit = 'ng/dl'; RefMin = 250; RefMax = 950; OptimalMin = 500; OptimalMax = 800; Group = "Hormone"; Description = "Gesamtmenge des an Proteine gebundenen und freien Testosterons im Blut." },
            [PSCustomObject]@{ Name = 'TnT (Troponin T)'; Unit = 'ng/L'; RefMin = 0; RefMax = 14; OptimalMin = 0; OptimalMax = 10; Group = 'Herz-Kreislauf'; Description = "Ein sehr spezifischer Marker für eine Schädigung des Herzmuskels, z.B. bei einem Herzinfarkt." },
            [PSCustomObject]@{ Name = 'Triglyceride (Trig)'; Unit = 'mg/dl'; RefMin = 0; RefMax = 150; OptimalMin = 50; OptimalMax = 100; Group = "Blutfette"; Description = "Neutralfette im Blut, die als Energiereserve dienen." },
            [PSCustomObject]@{ Name = 'TSH-basal (TSH)'; Unit = 'mU/l'; RefMin = 0.4; RefMax = 4.0; OptimalMin = 0.5; OptimalMax = 2.0; Group = "Schilddrüse"; Description = "Steuerhormon der Hirnanhangsdrüse. Wichtigster Suchtest für Schilddrüsenerkrankungen." },
            [PSCustomObject]@{ Name = 'Vitamin B12'; Unit = 'pg/ml'; RefMin = 200; RefMax = 900; OptimalMin = 400; OptimalMax = 700; Group = 'Vitamine'; Description = "Wichtig für die Blutbildung, Zellteilung und Nervenfunktion." },
            [PSCustomObject]@{ Name = 'Vitamin D (25-OH)'; Unit = 'ng/ml'; RefMin = 30; RefMax = 80; OptimalMin = 40; OptimalMax = 60; Group = "Vitamine"; Description = "Speicherform von Vitamin D. Wichtig für Knochen, Immunsystem und viele andere Körperfunktionen." },
            [PSCustomObject]@{ Name = 'Workout-Frequenz'; Unit = 'x/Woche'; Group = 'Lebensstil'; Description = "Anzahl Ihrer Trainingseinheiten pro Woche." },
            # v2.26.0: Zink - Referenz NIH ODS. Bewusst ohne Optimalbereich (keine offizielle Vorgabe).
            [PSCustomObject]@{ Name = 'Zink'; Unit = 'µg/dl'; RefMin = 80; RefMax = 120; Group = "Mineralstoffe"; Description = "Zink im Serum. Referenzbereich Gesunder 80-120 µg/dl (12-18 µmol/l); ein unzureichender Zinkstatus liegt laut NIH Office of Dietary Supplements unter 74 µg/dl (Mann) bzw. 70 µg/dl (Frau) vor (Zinc - Health Professional Fact Sheet). Viele deutsche Labore geben 70-120 µg/dl an - maßgeblich ist der Referenzbereich des eigenen Befundes. Bewusst KEIN Optimalbereich hinterlegt: dafür existiert keine offizielle Vorgabe. Zink ist Cofaktor von über 300 Enzymen und wichtig für Immunfunktion, Wundheilung, DNA-Synthese und Testosteronstoffwechsel. PRÄANALYTIK BEACHTEN: Serum-Zink schwankt mit Tageszeit (morgens höher), Nahrungsaufnahme, Infekten und Entzündung (ein CRP-Anstieg senkt Zink), Steroidhormonen sowie Muskelabbau und korreliert nur eingeschränkt mit der Zufuhr. Abnahme möglichst nüchtern am Morgen, Probe hämolysefrei. Tolerierbare Gesamtzufuhr (UL) für Erwachsene: 40 mg/Tag." }
        )
        CalculatedMarkers = [System.Collections.ArrayList]@(
            [PSCustomObject]@{ Name = "Non-HDL-Cholesterin"; Unit = "mg/dl"; RequiredMarkers = @("Cholesterin, gesamt (Chol)", "HDL-Cholesterin (HDL)"); Formula = "%0% - %1%"; Group = "Blutfette"; RefMin = 0; RefMax = 160; OptimalMin = 80; OptimalMax = 130; Description = "Gesamtcholesterin abzüglich des 'guten' HDL-Cholesterins. Umfasst alle atherogenen Partikel." },
            [PSCustomObject]@{ Name = "HOMA-IR Index"; Unit = ""; RequiredMarkers = @("Glukose (Glu)", "Nüchterninsulin"); Formula = "(%0% * %1%) / 405"; Group = "Stoffwechsel"; RefMin = 0; RefMax = 2.5; OptimalMin = 0.5; OptimalMax = 1.5; Description = "Modell zur Abschätzung der Insulinresistenz. Höhere Werte deuten auf eine schlechtere Insulinwirkung hin." },
            [PSCustomObject]@{ Name = "Atherogener Index (AIP)"; Unit = ""; RequiredMarkers = @("Triglyceride (Trig)", "HDL-Cholesterin (HDL)"); Formula = "[Math]::Log10(%0% / %1%)"; Group = "Blutfette"; RefMin = -0.3; RefMax = 0.1; OptimalMin = -0.2; OptimalMax = 0; Description = "Der 'Atherogenic Index of Plasma' korreliert mit dem Risiko für Herz-Kreislauf-Erkrankungen." },
            [PSCustomObject]@{ Name = "BMI"; Unit = "kg/m²"; RequiredMarkers = @("Gewicht"); Formula = "%0% / ([Math]::Pow($personal.Groesse / 100, 2))"; Group = "Stoffwechsel"; RefMin = 18.5; RefMax = 24.9; OptimalMin = 20; OptimalMax = 22; Description = "Body-Mass-Index. Verhältnis von Körpergewicht zu Körpergröße." },
            [PSCustomObject]@{ Name = "TyG-Index"; Unit = ""; RequiredMarkers = @("Triglyceride (Trig)", "Glukose (Glu)"); Formula = "[Math]::Log(((%0%) * (%1%)) / 2)"; Group = "Stoffwechsel"; RefMin = 4.5; RefMax = 8.5; OptimalMin = 4.5; OptimalMax = 4.68; Description = "Triglyceride-Glucose Index (Simental-Mendia 2008). Einfacher Surrogatmarker für Insulinresistenz, stark mit HOMA-IR korreliert. Werte >4.68 sprechen für Insulinresistenz." },
            [PSCustomObject]@{ Name = "Triglyceride/HDL-Ratio"; Unit = ""; RequiredMarkers = @("Triglyceride (Trig)", "HDL-Cholesterin (HDL)"); Formula = "%0% / %1%"; Group = "Blutfette"; RefMin = 0; RefMax = 3.5; OptimalMin = 0; OptimalMax = 2.0; Description = "Metabolischer Gesundheitsmarker. Starker Prädiktor für Insulinresistenz, small-dense LDL und kardiovaskuläre Ereignisse (Salazar 2012). Optimal < 2." },
            [PSCustomObject]@{ Name = "Neutrophile/Lymphozyten-Ratio (NLR)"; Unit = ""; RequiredMarkers = @("Neutrophile Granulozyten (NEUT)", "Lymphozyten (LYMPH)"); Formula = "%0% / %1%"; Group = "Inflammation"; RefMin = 0.78; RefMax = 3.53; OptimalMin = 1.0; OptimalMax = 2.0; Description = "NLR (Zahorec 2021) - validierter Marker für systemische Inflammation. Erhöhte Werte (>3) assoziieren mit InflammAging, kardiovaskulärem Risiko und Gesamtmortalität." },
            [PSCustomObject]@{ Name = "Thrombozyten/Lymphozyten-Ratio (PLR)"; Unit = ""; RequiredMarkers = @("Thrombozyten (PLT)", "Lymphozyten (LYMPH)"); Formula = "%0% / %1%"; Group = "Inflammation"; RefMin = 36; RefMax = 180; OptimalMin = 70; OptimalMax = 130; Description = "PLR - komplementärer Inflammationsmarker. Erhöht bei chronischer Inflammation, onkologischer Belastung und kardiovaskulärem Risiko." },
            [PSCustomObject]@{ Name = "InflammAging-Score"; Unit = "Punkte"; RequiredMarkers = @("C-reaktives Protein (CRP)", "Neutrophile Granulozyten (NEUT)", "Lymphozyten (LYMPH)", "Thrombozyten (PLT)"); Formula = "[INFLAMMAGING]"; Group = "Longevity"; RefMin = 0; RefMax = 100; OptimalMin = 70; OptimalMax = 100; Description = "Komposit-Score (Franceschi et al.) aus hsCRP, NLR und PLR. 100 = optimal (keine systemische Inflammation), 0 = ausgeprägtes InflammAging. Chronische niedriggradige Entzündung ist Haupttreiber des biologischen Alterns." },
            [PSCustomObject]@{ Name = "PhenoAge (biologisch)"; Unit = "Jahre"; RequiredMarkers = @("Albumin", "Kreatinin (Krea)", "Glukose (Glu)", "C-reaktives Protein (CRP)", "Lymphozyten (LYMPH)", "Mittleres Erythrozytenvolumen (MCV)", "Erythrozyten-Verteilungsbreite (RDW-CV)", "Alkalische Phosphatase (AP)", "Leukozyten (WBC)"); Formula = "[PHENOAGE]"; Group = "Longevity"; RefMin = 20; RefMax = 80; OptimalMin = 20; OptimalMax = 35; Description = "Biologisches Alter nach Levine et al. (2018, Aging). Berechnet aus 9 klinischen Laborwerten via Gompertz-Mortalitäts-Score. Validiert an NHANES-Kohorte. Prädiktor für All-Cause-Mortality, CVD und Krebs." },
            [PSCustomObject]@{ Name = "PhenoAge-Accel"; Unit = "Jahre"; RequiredMarkers = @("Albumin", "Kreatinin (Krea)", "Glukose (Glu)", "C-reaktives Protein (CRP)", "Lymphozyten (LYMPH)", "Mittleres Erythrozytenvolumen (MCV)", "Erythrozyten-Verteilungsbreite (RDW-CV)", "Alkalische Phosphatase (AP)", "Leukozyten (WBC)"); Formula = "[PHENOAGE_ACCEL]"; Group = "Longevity"; RefMin = -20; RefMax = 20; OptimalMin = -20; OptimalMax = -2; Description = "Biologisches minus chronologisches Alter. Negative Werte = jünger als das Kalenderalter (günstig). Werte >0 = beschleunigtes Altern. Direkter, sofort interpretierbarer Longevity-Marker." },
            [PSCustomObject]@{ Name = "Longevity-Score"; Unit = "Punkte"; RequiredMarkers = @("Langzeitzucker (HbA1c)", "Nüchterninsulin", "Vitamin D (25-OH)", "Homocystein", "C-reaktives Protein (CRP)", "NAD+", "Albumin"); Formula = "[LONGEVITY_SCORE]"; Group = "Longevity"; RefMin = 0; RefMax = 100; OptimalMin = 80; OptimalMax = 100; Description = "Ein zusammenfassender Score, der wichtige Biomarker für ein gesundes Altern bewertet." },
            [PSCustomObject]@{ Name = "PREVENT-ASCVD-10Y"; Unit = "%"; RequiredMarkers = @("Cholesterin, gesamt (Chol)", "HDL-Cholesterin (HDL)", "Systolischer Blutdruck", "Langzeitzucker (HbA1c)"); Formula = "[PREVENT_SCORE]"; Group = "Risiko"; RefMin = 0; RefMax = 7.5; Description = "Schätzt das 10-Jahres-Risiko für atherosklerotische Herz-Kreislauf-Erkrankungen." }
        )
        WarningThreshold = 10
        Personal = @{
            Name = ""
            Age = 37
            Geschlecht = "männlich"
            Groesse = 0.0
            Gewicht = 64.5
            KFA = 0.0
            Grundumsatz = 1550
            WorkoutFreq = "5-6x/Woche"
            SystBlutdruck = 120
            Raucher = $false
            PersistentTooltips = $false      # v2.20.0: Hilfstexte dauerhaft anzeigen?
        }
        AutoBackup = @{
            Enabled       = $false          # Feature aktiv?
            TargetPath    = ""              # Zielverzeichnis
            FormatZip     = $false          # Format: ZIP-Archiv
            FormatFolder  = $false          # Format: Ordner-und-Dateien (Kopie)
            Interval      = "wöchentlich"   # täglich | wöchentlich | monatlich | bei Programm-Start
            TimeOfDay     = "03:00"         # HH:mm - nicht relevant bei "bei Programm-Start"
            Weekdays      = @("Mo")         # Wochentage bei Interval="wöchentlich" (Mo,Di,Mi,Do,Fr,Sa,So)
            MonthDay      = "am Ersten eines Monats"   # Tag-Auswahl bei Interval="monatlich"
            IncludeRegistry = $true         # Registry-Export mit sichern (empfohlen)
            CleanupEnabled = $false         # Alte Backups automatisch löschen?
            CleanupKeep   = "letzte drei Backups behalten"  # Retention-Regel
            LastBackup    = ""              # ISO-Timestamp des letzten erfolgreichen Backups
            PortableBackup = $false          # v2.19.0: Portables .btbackup bei jedem AutoBackup erstellen?
            PortableBackupPassword = ""      # v2.19.0: DPAPI-verschlüsseltes Passwort (Base64)
        }
        GeneticPredispositions = @{
            Enabled = $false   # Master-Schalter: Vorbelastungen berücksichtigen?
            BuiltIn = @(
                @{
                    Name    = "Chronische Veneninsuffizienz (Venenleiden)"
                    Active  = $false
                    Markers = @("D-Dimer", "C-reaktives Protein (CRP)", "Hochsensitives CRP (hs-CRP)")
                    Hint    = "Venenleiden-Vorbelastung: Achten Sie besonders auf diesen Marker."
                },
                @{
                    Name    = "Familiäre Hypertonie / KHK-Prädisposition (Herz-Kreislauf-Probleme)"
                    Active  = $false
                    Markers = @("Homocystein", "Lp(a)", "NT-proBNP", "TnT (Troponin T)", "Systolischer Blutdruck", "ApoB", "PREVENT-ASCVD-10Y", "Hochsensitives CRP (hs-CRP)")
                    Hint    = "KHK/Hypertonie-Vorbelastung: Achten Sie besonders auf diesen Marker."
                },
                @{
                    Name    = "Familiäre Hypercholesterinämie (Cholesterin)"
                    Active  = $false
                    Markers = @("Cholesterin, gesamt (Chol)", "LDL-Cholesterin (LDL)", "HDL-Cholesterin (HDL)", "ApoB", "Triglyceride (Trig)", "Non-HDL-Cholesterin", "Atherogener Index (AIP)", "Triglyceride/HDL-Ratio", "Omega-3-Index")
                    Hint    = "Hypercholesterinämie-Vorbelastung: Achten Sie besonders auf diesen Marker."
                },
                @{
                    Name    = "Typ-2-Diabetes-Mellitus-Prädisposition (Diabetes)"
                    Active  = $false
                    Markers = @("Glukose (Glu)", "Langzeitzucker (HbA1c)", "Nüchterninsulin", "HOMA-IR Index", "TyG-Index")
                    Hint    = "Diabetes-Vorbelastung: Achten Sie besonders auf diesen Marker."
                }
            )
            Custom = @()   # Benutzerdefinierte Vorbelastungen: @( @{ Name="..."; Active=$false; Markers=@(...); Hint="..." }, ... )
        }
    }
    
    $configData = $defaultConfig

    # v2.25.0: Snapshot der Defaults, BEVOR die gespeicherte Config sie ersetzt.
    # ($configData zeigt auf dieselbe Hashtable wie $defaultConfig.)
    $defaultMarkerSnapshot     = @($defaultConfig.Markers           | ForEach-Object { $_ })
    $defaultCalcMarkerSnapshot = @($defaultConfig.CalculatedMarkers | ForEach-Object { $_ })
    $savedMarkerSetVersion     = $null
    $script:MarkerMigrationPending = $false

    if (Test-Path $dataFile) {
        try {
            # v2.19.0: Automatische Erkennung verschlüsselt vs. Klartext
            $jsonContent = Read-ProtectedJsonFile -Path $dataFile
            if (-not $jsonContent) { throw "Config-Datei konnte nicht gelesen werden." }
            $loadedObject = $jsonContent | ConvertFrom-Json
            $loadedData = ConvertTo-Hashtable -InputObject $loadedObject

            if ($loadedData.Markers) {
                $configData['Markers'] = [System.Collections.ArrayList]@($loadedData.Markers | ForEach-Object { [PSCustomObject]$_ })
            }
            if ($loadedData.CalculatedMarkers) {
                $configData['CalculatedMarkers'] = [System.Collections.ArrayList]@($loadedData.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ })
            }
            if ($loadedData.WarningThreshold) {
                $configData['WarningThreshold'] = $loadedData.WarningThreshold
            }
            # v2.25.0: Erreichter Marker-Set-Stand (fehlt bei Configs <= v2.24.1)
            if ($loadedData.MarkerSetVersion) {
                $savedMarkerSetVersion = [string]$loadedData.MarkerSetVersion
            }
            if ($loadedData.Personal) {
                $configData['Personal'] = $loadedData.Personal
                # v2.20.0: Default für neues Feld, falls nicht in gespeicherter Config vorhanden
                if ($configData['Personal'] -is [hashtable]) {
                    if (-not $configData['Personal'].ContainsKey('PersistentTooltips')) { $configData['Personal']['PersistentTooltips'] = $false }
                } elseif ($configData['Personal'].PSObject -and ($configData['Personal'].PSObject.Properties.Name -notcontains 'PersistentTooltips')) {
                    $configData['Personal'] | Add-Member -NotePropertyName 'PersistentTooltips' -NotePropertyValue $false -Force
                }
            }
            if ($loadedData.AutoBackup) {
                # v2.14.2 Bugfix: Robust gegen Hashtable UND PSCustomObject
                # (ConvertTo-Hashtable liefert je nach Tiefe unterschiedliche Typen).
                $ab = $loadedData.AutoBackup
                foreach ($k in @("Enabled","TargetPath","FormatZip","FormatFolder","Interval","TimeOfDay","Weekdays","MonthDay","IncludeRegistry","CleanupEnabled","CleanupKeep","LastBackup","PortableBackup","PortableBackupPassword")) {
                    $val = $null
                    $hasKey = $false
                    if ($ab -is [hashtable]) {
                        if ($ab.ContainsKey($k)) { $val = $ab[$k]; $hasKey = $true }
                    } elseif ($ab.PSObject -and ($ab.PSObject.Properties.Name -contains $k)) {
                        $val = $ab.$k
                        $hasKey = $true
                    }
                    if ($hasKey) {
                        $configData['AutoBackup'][$k] = $val
                    }
                }
            }
            # v2.17.0: Genetische Vorbelastungen laden (eigener try-catch, damit
            # ein Fehler hier nicht die gesamte Config-Ladung abbricht)
            try {
                if ($loadedData.GeneticPredispositions) {
                    $gp = $loadedData.GeneticPredispositions
                    if ($gp -is [hashtable]) {
                        if ($gp.ContainsKey('Enabled'))  { $configData['GeneticPredispositions']['Enabled'] = [bool]$gp['Enabled'] }
                        if ($gp.ContainsKey('BuiltIn') -and $gp['BuiltIn'])  {
                            $savedBuiltIn = @($gp['BuiltIn'])
                            foreach ($savedEntry in $savedBuiltIn) {
                                if ($null -eq $savedEntry) { continue }
                                $se = if ($savedEntry -is [hashtable]) { $savedEntry } else { ConvertTo-Hashtable -InputObject $savedEntry }
                                if ($se -and $se['Name']) {
                                    $match = $configData['GeneticPredispositions']['BuiltIn'] | Where-Object { $_['Name'] -eq $se['Name'] }
                                    if ($match) { $match['Active'] = [bool]$se['Active'] }
                                }
                            }
                        }
                        if ($gp.ContainsKey('Custom') -and $gp['Custom'] -and @($gp['Custom']).Count -gt 0) {
                            $configData['GeneticPredispositions']['Custom'] = @($gp['Custom'] | ForEach-Object {
                                if ($null -eq $_) { return }
                                $entry = if ($_ -is [hashtable]) { $_ } else { ConvertTo-Hashtable -InputObject $_ }
                                if ($entry -and $entry['Name']) {
                                    @{
                                        Name    = [string]$entry['Name']
                                        Active  = [bool]$entry['Active']
                                        Markers = @(if ($entry['Markers']) { $entry['Markers'] } else { @() })
                                        Hint    = if ($entry['Hint']) { [string]$entry['Hint'] } else { "Vorbelastung: Achten Sie besonders auf diesen Marker." }
                                    }
                                }
                            } | Where-Object { $_ })
                        }
                    } elseif ($null -ne $gp -and $gp.PSObject) {
                        if ($gp.PSObject.Properties.Name -contains 'Enabled') { $configData['GeneticPredispositions']['Enabled'] = [bool]$gp.Enabled }
                        if ($gp.PSObject.Properties.Name -contains 'BuiltIn' -and $gp.BuiltIn) {
                            $savedBuiltIn = @($gp.BuiltIn)
                            foreach ($savedEntry in $savedBuiltIn) {
                                if ($null -eq $savedEntry) { continue }
                                $se = if ($savedEntry -is [hashtable]) { $savedEntry } else { ConvertTo-Hashtable -InputObject $savedEntry }
                                if ($se -and $se['Name']) {
                                    $match = $configData['GeneticPredispositions']['BuiltIn'] | Where-Object { $_['Name'] -eq $se['Name'] }
                                    if ($match) { $match['Active'] = [bool]$se['Active'] }
                                }
                            }
                        }
                        if ($gp.PSObject.Properties.Name -contains 'Custom' -and $gp.Custom -and @($gp.Custom).Count -gt 0) {
                            $configData['GeneticPredispositions']['Custom'] = @($gp.Custom | ForEach-Object {
                                if ($null -eq $_) { return }
                                $entry = if ($_ -is [hashtable]) { $_ } else { ConvertTo-Hashtable -InputObject $_ }
                                if ($entry -and $entry['Name']) {
                                    @{
                                        Name    = [string]$entry['Name']
                                        Active  = [bool]$entry['Active']
                                        Markers = @(if ($entry['Markers']) { $entry['Markers'] } else { @() })
                                        Hint    = if ($entry['Hint']) { [string]$entry['Hint'] } else { "Vorbelastung: Achten Sie besonders auf diesen Marker." }
                                    }
                                }
                            } | Where-Object { $_ })
                        }
                    }
                }
            } catch {
                Write-Warning "GeneticPredispositions Merge-Fehler: $($_.Exception.Message) - Defaults werden verwendet."
            }

            $defaultMarkerProps = $defaultConfig.Markers[0].PSObject.Properties.Name
            foreach($marker in $configData.Markers){
                if ($marker.PSObject -ne $null) {
                     foreach($prop in $defaultMarkerProps){
                        if(-not $marker.PSObject.Properties.Match($prop)){
                            $marker | Add-Member -MemberType NoteProperty -Name $prop -Value $null -ErrorAction SilentlyContinue
                        }
                    }
                }
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Laden der Konfigurationsdatei: $($_.Exception.Message). Standardkonfiguration wird verwendet.")
        }
    }
    
    # v2.25.0: Neue Default-Marker nachziehen (versionsgesteuert, additiv).
    $script:MarkerMigrationPending = Merge-NewDefaultMarkers `
        -ConfigData $configData `
        -DefaultMarkers $defaultMarkerSnapshot `
        -DefaultCalculatedMarkers $defaultCalcMarkerSnapshot `
        -SavedVersion $savedMarkerSetVersion

    $finalConfig = @{
        Markers = [System.Collections.ArrayList]@($configData.Markers | ForEach-Object { [PSCustomObject]$_ })
        CalculatedMarkers = [System.Collections.ArrayList]@($configData.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ })
        WarningThreshold = $configData.WarningThreshold
        Personal = $configData.Personal
        AutoBackup = $configData.AutoBackup
        GeneticPredispositions = $configData.GeneticPredispositions
        # v2.25.0: erreichter Marker-Set-Stand (verhindert erneutes Nachziehen)
        MarkerSetVersion = $script:MarkerSetVersion
    }

    Save-Data -data @{ Config = $finalConfig } -type "Config"
    return $finalConfig
}

function Save-Data {
    param ($data, $type, $dateString)
    try {
        if (-not (Test-Path $dataDir)) { New-Item -Path $dataDir -ItemType Directory -Force | Out-Null }
        if ($type -eq "Config") {
            # v2.14.2 Bugfix: AutoBackup MUSS mit geschrieben werden, sonst gehen alle
            # Backup-Einstellungen nach Neustart verloren.
            $configToSave = @{
                Markers = $data.Config.Markers
                CalculatedMarkers = $data.Config.CalculatedMarkers
                WarningThreshold = $data.Config.WarningThreshold
                Personal = $data.Config.Personal
                # v2.25.0: Marker-Set-Stand mitschreiben, sonst wuerde die Migration
                # bei jedem Start erneut laufen und geloeschte Marker zurueckholen.
                MarkerSetVersion = if ($data.Config.MarkerSetVersion) { $data.Config.MarkerSetVersion } else { $script:MarkerSetVersion }
            }
            if ($data.Config.AutoBackup) {
                $configToSave['AutoBackup'] = $data.Config.AutoBackup
            }
            if ($data.Config.GeneticPredispositions) {
                $configToSave['GeneticPredispositions'] = $data.Config.GeneticPredispositions
            }
            # v2.19.0: DPAPI-verschlüsselt schreiben
            $jsonString = $configToSave | ConvertTo-Json -Depth 5
            Write-ProtectedJsonFile -Path $dataFile -JsonString $jsonString
        } elseif ($type -eq "Daily") {
            $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $dateString
            if (-not (Test-Path $dailyDataDir)) { New-Item -Path $dailyDataDir -ItemType Directory -Force | Out-Null }
            
            $itemsForDay = New-Object System.Collections.ArrayList
            
            if (Test-Path $dailyDataFile) {
                # v2.19.0: Verschlüsselt lesen
                $existingJson = Read-ProtectedJsonFile -Path $dailyDataFile
                if ($existingJson) {
                    $loadedData = $existingJson | ConvertFrom-Json
                    if ($loadedData -and $loadedData.Items) {
                        $itemsArray = @($loadedData.Items)
                        foreach ($item in $itemsArray) {
                            [void]$itemsForDay.Add([PSCustomObject]$item)
                        }
                    }
                }
            }
            
            $itemsForDay.Add([PSCustomObject]$data.NewItem) | Out-Null
            
            $savableItems = $itemsForDay | ForEach-Object {
                 [PSCustomObject]@{
                    Date = $_.Date; Name = $_.Name; Value = [double]$_.Value; Unit = $_.Unit;
                    Note = if ($_.PSObject.Properties['Note']) { $_.Note } else { "" }
                }
            }

            # v2.19.0: DPAPI-verschlüsselt schreiben
            $jsonString = @{ Items = $savableItems } | ConvertTo-Json -Depth 5
            Write-ProtectedJsonFile -Path $dailyDataFile -JsonString $jsonString
        }
    } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern der $type-Daten: $($_.Exception.Message)") }
}

function Load-AllHistoricalItems {
    $allItems = New-Object System.Collections.ArrayList
    try {
        $files = Get-ChildItem -Path $script:dailyDataDirBase -Filter "*.json" -Recurse -ErrorAction SilentlyContinue
        foreach ($file in $files) {
            # v2.19.0: Verschlüsselt lesen (auto-detect)
            $json = Read-ProtectedJsonFile -Path $file.FullName
            if ($json) {
                $data = $json | ConvertFrom-Json
                if ($data.Items) {
                    [void]$allItems.AddRange(@($data.Items | ForEach-Object { [PSCustomObject]$_ }))
                }
            }
        }
    } catch { Write-Warning "Fehler beim Laden der Verlaufsdaten." }
    return $allItems
}

function Save-AllHistoricalData {
    param ($AllItems)
    try {
        # BUGFIX v2.7.3: Zuerst ALLE bestehenden Tages-JSON-Dateien löschen,
        # damit komplett gelöschte Tage nicht als "Geister-Dateien" bestehen bleiben.
        if (Test-Path $script:dailyDataDirBase) {
            $existingFiles = Get-ChildItem -Path $script:dailyDataDirBase -Filter "*.json" -Recurse -ErrorAction SilentlyContinue
            foreach ($file in $existingFiles) {
                Remove-Item -Path $file.FullName -Force -ErrorAction SilentlyContinue
            }
            # Leere Unterverzeichnisse (Monat/Jahr) aufräumen
            Get-ChildItem -Path $script:dailyDataDirBase -Directory -Recurse -ErrorAction SilentlyContinue |
                Sort-Object { $_.FullName.Length } -Descending |
                Where-Object { (Get-ChildItem -Path $_.FullName -ErrorAction SilentlyContinue).Count -eq 0 } |
                ForEach-Object { Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue }
        }
        
        # Verbleibende Daten neu schreiben
        $groupedData = $AllItems | Group-Object -Property Date
        foreach ($group in $groupedData) {
            $dateString = $group.Name
            $itemsForDay = $group.Group
            $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $dateString
            if (-not (Test-Path $dailyDataDir)) { New-Item -Path $dailyDataDir -ItemType Directory -Force | Out-Null }
            $savableItems = $itemsForDay | ForEach-Object {
                $obj = [PSCustomObject]@{
                    Date = $_.Date
                    Name = $_.Name
                    Value = [double]$_.Value
                    Unit = $_.Unit
                }
                if ($_.PSObject.Properties['Note']) {
                    $obj | Add-Member -NotePropertyName 'Note' -NotePropertyValue $_.Note -Force
                }
                $obj
            }
            # v2.19.0: DPAPI-verschlüsselt schreiben
            $jsonString = @{ Items = $savableItems } | ConvertTo-Json -Depth 5
            Write-ProtectedJsonFile -Path $dailyDataFile -JsonString $jsonString
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ein schwerwiegender Fehler ist beim Speichern der Daten aufgetreten: $($_.Exception.Message)", "Speicherfehler", "OK", "Error")
    }
}

function Calculate-PhenoAge {
    <#
    .SYNOPSIS
        Berechnet das biologische Alter nach Levine et al. (2018, Aging).
    .DESCRIPTION
        Morgan E. Levine et al., "An epigenetic biomarker of aging for lifespan
        and healthspan", Aging (Albany NY), 2018; 10(4):573-591.
        Nutzt Gompertz-Mortalitätsmodell mit 9 Laborparametern + chronologisches Alter.
        ACHTUNG: Einheiten MÜSSEN US-klinisch sein:
          Albumin in g/L (nicht g/dL! → *10 wenn g/dL)
          Kreatinin in µmol/L (*88.4 wenn mg/dL)
          Glukose in mmol/L (/18 wenn mg/dL)
          CRP in mg/L (natürlicher log: ln(CRP/10) im NHANES-Original, hier CRP/10+0.1 für ln-Schutz)
          Lymphozyten in % (direkt)
          MCV in fL (direkt)
          RDW in % (direkt)
          Alkalische Phosphatase in U/L (direkt)
          Leukozyten in 1000 cells/µL (direkt)
    .PARAMETER values
        Array [Albumin, Kreatinin, Glukose, CRP, Lympho%, MCV, RDW-CV, AP, WBC]
    .PARAMETER chronoAge
        Chronologisches Alter in Jahren.
    .OUTPUTS
        Biologisches Alter in Jahren.
    #>
    param(
        [double[]]$values,
        [double]$chronoAge
    )
    try {
        # Einheiten-Normalisierung (Input ist in typisch deutschen Laborangaben)
        $albumin_gL    = if ($values[0] -lt 10) { $values[0] * 10 } else { $values[0] }   # g/dL → g/L
        $krea_umolL    = if ($values[1] -lt 10) { $values[1] * 88.4 } else { $values[1] } # mg/dL → µmol/L
        $glu_mmolL     = if ($values[2] -gt 20) { $values[2] / 18.0 } else { $values[2] } # mg/dL → mmol/L
        $crp_mgL       = $values[3]
        # hsCRP oft in mg/dL; wenn < 1, könnte mg/dL sein → normalisieren wir nicht auto, da unsicher
        $lympho_pct    = $values[4]
        $mcv_fL        = $values[5]
        $rdw_pct       = $values[6]
        $ap_UL         = $values[7]
        $wbc_1000uL    = $values[8]

        # Original-Koeffizienten Levine 2018 (Table 1, model with chronological age)
        $xb = -19.907 `
              + (-0.0336 * $albumin_gL) `
              + ( 0.0095 * $krea_umolL) `
              + ( 0.1953 * $glu_mmolL) `
              + ( 0.0954 * [Math]::Log([Math]::Max($crp_mgL, 0.01) / 10.0 + 0.001)) `
              + (-0.0120 * $lympho_pct) `
              + ( 0.0268 * $mcv_fL) `
              + ( 0.3306 * $rdw_pct) `
              + ( 0.00188 * $ap_UL) `
              + ( 0.0554 * $wbc_1000uL) `
              + ( 0.0804 * $chronoAge)

        # Gompertz: 10-Jahres-Mortalitätsrisiko
        $g = 0.0076927  # Gompertz-Parameter (γ) Levine 2018
        $mortScore = 1 - [Math]::Exp(-[Math]::Exp($xb) * ([Math]::Exp(120 * $g) - 1) / $g)
        $mortScore = [Math]::Min([Math]::Max($mortScore, 0.0001), 0.9999)

        # Rücktransformation in Alter
        $phenoAge = 141.50225 + [Math]::Log(-0.00553 * [Math]::Log(1 - $mortScore)) / 0.09165

        return [Math]::Round($phenoAge, 1)
    } catch {
        Write-Warning "PhenoAge-Berechnungsfehler: $($_.Exception.Message)"
        return $null
    }
}

function Calculate-InflammAgingScore {
    <#
    .SYNOPSIS
        InflammAging-Komposit aus hsCRP, NLR, PLR.
    .DESCRIPTION
        Franceschi et al., "Inflammaging" - chronische niedriggradige
        systemische Entzündung als Haupttreiber des biologischen Alterns.
        Skala: 100 = optimal, 0 = stark entzündlich.
    .PARAMETER values
        Array [CRP mg/L, Neutrophile %, Lymphozyten %, Thrombozyten Tsd./µL]
    #>
    param([double[]]$values)
    try {
        $crp    = [double]$values[0]
        $neut   = [double]$values[1]
        $lymph  = [double]$values[2]
        $plt    = [double]$values[3]

        if ($lymph -le 0) { return 0 }

        $nlr = $neut / $lymph
        # PLT ist in Tsd./µl, Lympho in %, also echte Ratios-Skala ist anders;
        # wir normalisieren PLR wie in Papers üblich mit absoluten Zellzahlen,
        # nehmen aber Prozent-Näherung: PLR ≈ PLT / Lympho% * Faktor ~1
        $plr = $plt / $lymph

        # Score-Komponenten (0-100, höher = besser)
        $crpScore  = [Math]::Max(0, [Math]::Min(100, (1 - $crp / 3.0) * 100))         # <1 optimal, >3 schlecht
        $nlrScore  = [Math]::Max(0, [Math]::Min(100, (1 - ($nlr - 1.0) / 2.5) * 100)) # 1-3 Range
        $plrScore  = [Math]::Max(0, [Math]::Min(100, (1 - ($plr - 70) / 130) * 100))  # 70-200 Range

        # Gewichteter Mittelwert: CRP dominant (Evidenzstärke)
        $score = ($crpScore * 0.5) + ($nlrScore * 0.3) + ($plrScore * 0.2)
        return [Math]::Round($score, 1)
    } catch {
        Write-Warning "InflammAging-Berechnungsfehler: $($_.Exception.Message)"
        return $null
    }
}

function Calculate-HbA1cTrajectory {
    <#
    .SYNOPSIS
        Trend-Slope von HbA1c-Werten (%-Punkte pro Jahr).
    .DESCRIPTION
        Positive Werte = Verschlechterung, negative = Verbesserung.
        Trajektorie ist oft aussagekräftiger als Einzelwert.
    .PARAMETER hbA1cItems
        Array von Items mit .Date und .Value
    #>
    param($hbA1cItems)
    if (-not $hbA1cItems -or $hbA1cItems.Count -lt 2) { return $null }
    try {
        $sortedItems = $hbA1cItems | Sort-Object { [datetime]::ParseExact($_.Date, "yyyy-MM-dd", $null) }
        $firstDate   = [datetime]::ParseExact($sortedItems[0].Date, "yyyy-MM-dd", $null)
        $xValues     = $sortedItems | ForEach-Object { ([datetime]::ParseExact($_.Date, "yyyy-MM-dd", $null) - $firstDate).TotalDays / 365.25 }
        $yValues     = $sortedItems.Value
        $n           = $sortedItems.Count
        $sumX        = ($xValues | Measure-Object -Sum).Sum
        $sumY        = ($yValues | Measure-Object -Sum).Sum
        $sumXY       = 0; for ($i = 0; $i -lt $n; $i++) { $sumXY += $xValues[$i] * $yValues[$i] }
        $sumX2       = ($xValues | ForEach-Object { $_ * $_ } | Measure-Object -Sum).Sum
        $denom       = $n * $sumX2 - $sumX * $sumX
        if ($denom -eq 0) { return $null }
        $slope       = ($n * $sumXY - $sumX * $sumY) / $denom
        return [Math]::Round($slope, 3)
    } catch {
        Write-Warning "HbA1c-Trajektorie-Fehler: $($_.Exception.Message)"
        return $null
    }
}

function Render-MarkerChartToImage {
    <#
    .SYNOPSIS
        Rendert eine Liniengrafik eines Blutmarkers als Bitmap für PDF/Print-Embedding.
    .DESCRIPTION
        v2.13.0. Erzeugt ein GDI+-Bitmap mit Achsen, Referenzbereichen und Datenpunkten.
        Kein externer Chart-Control nötig - funktioniert in asynchronen PrintPage-Handlern.
    .PARAMETER items
        Array von Datenpunkten mit .Date (yyyy-MM-dd) und .Value (numerisch).
    .PARAMETER markerName
        Titel der Grafik.
    .PARAMETER unit
        Einheit für Y-Achse.
    .PARAMETER refMin / refMax
        Optional: Referenzbereich (grün schattiert).
    .PARAMETER optMin / optMax
        Optional: Optimalbereich (stärker grün).
    .PARAMETER width / height
        Bildgröße in Pixel.
    .OUTPUTS
        [System.Drawing.Bitmap]
    #>
    param(
        [array]$items,
        [string]$markerName,
        [string]$unit = "",
        $refMin = $null, $refMax = $null,
        $optMin = $null, $optMax = $null,
        [int]$width  = 900,
        [int]$height = 300
    )

    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap $width, $height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::ClearTypeGridFit
    $g.Clear([System.Drawing.Color]::White)

    $titleFont = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $axisFont  = New-Object System.Drawing.Font("Segoe UI", 8)

    # Padding
    $padL = 60; $padR = 20; $padT = 30; $padB = 50
    $plotW = $width  - $padL - $padR
    $plotH = $height - $padT - $padB

    # Titel
    $g.DrawString("$markerName  ($unit)", $titleFont, [System.Drawing.Brushes]::Black, 10, 5)

    # Daten sortieren + numerisch casten
    $arr = @($items | Sort-Object { [datetime]::ParseExact($_.Date, "yyyy-MM-dd", $null) })
    if ($arr.Count -lt 2) {
        $g.DrawString("Zu wenig Datenpunkte für eine Verlaufsgrafik (mindestens 2 erforderlich).", $axisFont, [System.Drawing.Brushes]::DarkGray, ($padL), ($padT + $plotH / 2))
        $g.Dispose()
        return $bmp
    }

    $dates  = @($arr | ForEach-Object { [datetime]::ParseExact($_.Date, "yyyy-MM-dd", $null) })
    $values = @($arr | ForEach-Object { [double]$_.Value })

    $minDate = $dates[0]; $maxDate = $dates[-1]
    $spanDays = ($maxDate - $minDate).TotalDays
    if ($spanDays -le 0) { $spanDays = 1 }

    # Y-Skala: Range inkl. Ref-Bereich, mit 5 % Padding
    $yMin = ($values | Measure-Object -Minimum).Minimum
    $yMax = ($values | Measure-Object -Maximum).Maximum
    if ($refMin -ne $null -and $refMin -is [double]) { if ($refMin -lt $yMin) { $yMin = $refMin } }
    if ($refMax -ne $null -and $refMax -is [double]) { if ($refMax -gt $yMax) { $yMax = $refMax } }
    $yRange = $yMax - $yMin
    if ($yRange -le 0) { $yRange = [Math]::Max([Math]::Abs($yMin) * 0.1, 1) }
    $yPad = $yRange * 0.1
    $yMin = $yMin - $yPad
    $yMax = $yMax + $yPad
    $yRange = $yMax - $yMin

    # Plot-Area mit Rahmen
    $plotRect = New-Object System.Drawing.Rectangle($padL, $padT, $plotW, $plotH)
    $g.FillRectangle([System.Drawing.Brushes]::WhiteSmoke, $plotRect)
    $g.DrawRectangle([System.Drawing.Pens]::DarkGray, $plotRect)

    # Referenzbereiche schattieren
    $toY = { param($v) $padT + $plotH * (1 - ($v - $yMin) / $yRange) }
    if ($refMin -ne $null -and $refMax -ne $null -and $refMin -is [double] -and $refMax -is [double]) {
        $y1 = & $toY $refMax; $y2 = & $toY $refMin
        if ($y2 -gt $y1) {
            $refBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(40, 144, 238, 144))
            $g.FillRectangle($refBrush, $padL, $y1, $plotW, ($y2 - $y1))
        }
    }
    if ($optMin -ne $null -and $optMax -ne $null -and $optMin -is [double] -and $optMax -is [double]) {
        $y1 = & $toY $optMax; $y2 = & $toY $optMin
        if ($y2 -gt $y1) {
            $optBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(90, 50, 205, 50))
            $g.FillRectangle($optBrush, $padL, $y1, $plotW, ($y2 - $y1))
        }
    }

    # Y-Achse: 5 Gridlines + Labels
    $gridPen  = New-Object System.Drawing.Pen ([System.Drawing.Color]::LightGray)
    $blackPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Black, 1)
    for ($i = 0; $i -le 5; $i++) {
        $v = $yMin + ($yRange * $i / 5)
        $y = & $toY $v
        $g.DrawLine($gridPen, $padL, $y, ($padL + $plotW), $y)
        $label = if ([Math]::Abs($v) -gt 1000) { "{0:N0}" -f $v } else { "{0:N2}" -f $v }
        $g.DrawString($label, $axisFont, [System.Drawing.Brushes]::Black, 5, ($y - 7))
    }

    # X-Achse: Datums-Labels (max. 6 Ticks, kompakt)
    $tickCount = [Math]::Min(6, $arr.Count)
    for ($i = 0; $i -lt $tickCount; $i++) {
        $frac = if ($tickCount -eq 1) { 0 } else { $i / ($tickCount - 1) }
        $d = $minDate.AddDays($spanDays * $frac)
        $x = $padL + $plotW * $frac
        $g.DrawLine($gridPen, $x, $padT, $x, ($padT + $plotH))
        $g.DrawString($d.ToString("dd.MM.yy"), $axisFont, [System.Drawing.Brushes]::Black, ($x - 20), ($padT + $plotH + 5))
    }

    # Linie + Punkte zeichnen
    $linePen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(46, 117, 182), 2)
    $ptBrush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromArgb(46, 117, 182))
    $points = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $arr.Count; $i++) {
        $frac = ($dates[$i] - $minDate).TotalDays / $spanDays
        $x = $padL + ($plotW * $frac)
        $y = & $toY $values[$i]
        [void]$points.Add((New-Object System.Drawing.PointF($x, $y)))
    }
    if ($points.Count -ge 2) {
        $g.DrawLines($linePen, ($points.ToArray([System.Drawing.PointF])))
    }
    foreach ($p in $points) {
        $g.FillEllipse($ptBrush, ($p.X - 3), ($p.Y - 3), 6, 6)
    }

    # Achsen-Rand
    $g.DrawRectangle($blackPen, $plotRect)

    $g.Dispose()
    return $bmp
}

function Convert-ToLongevityScore {
    <#
    .SYNOPSIS
        Mappt Werte der neuen Longevity-Marker (v2.10.0) auf einheitliche 0-100-Skala.
    .DESCRIPTION
        Verschiedene Marker haben völlig unterschiedliche Einheiten (Jahre, Punkte, Ratios).
        Diese Funktion normalisiert sie auf eine einheitliche Bewertungsskala für das UI.
        100 = optimal, 0 = stark suboptimal.
    #>
    param(
        [string]$markerName,
        [double]$value
    )
    $score = 50
    switch ($markerName) {
        "PhenoAge (biologisch)" {
            # Score basiert auf Differenz zu chronol. Alter - wird über PhenoAge-Accel abgedeckt
            $chronoAge = [double]$script:data.Config.Personal.Age
            $accel = $value - $chronoAge
            if ($accel -le -5)     { $score = 100 }
            elseif ($accel -le -2) { $score = 90 }
            elseif ($accel -le 0)  { $score = 80 }
            elseif ($accel -le 2)  { $score = 60 }
            elseif ($accel -le 5)  { $score = 40 }
            else                   { $score = 20 }
        }
        "PhenoAge-Accel" {
            # Direkt: negativ = gut, positiv = schlecht
            if ($value -le -5)     { $score = 100 }
            elseif ($value -le -2) { $score = 90 }
            elseif ($value -le 0)  { $score = 80 }
            elseif ($value -le 2)  { $score = 60 }
            elseif ($value -le 5)  { $score = 40 }
            else                   { $score = 20 }
        }
        "InflammAging-Score" {
            # Skala ist bereits 0-100
            $score = [Math]::Round($value, 0)
        }
        "Neutrophile/Lymphozyten-Ratio (NLR)" {
            # Optimal 1.0-2.0, > 3.0 kritisch
            if ($value -le 1.5)      { $score = 100 }
            elseif ($value -le 2.0)  { $score = 85 }
            elseif ($value -le 2.5)  { $score = 70 }
            elseif ($value -le 3.0)  { $score = 50 }
            elseif ($value -le 4.0)  { $score = 30 }
            else                     { $score = 10 }
        }
        "TyG-Index" {
            # Optimal < 4.68 (Simental-Mendia Cutoff)
            if ($value -lt 4.5)      { $score = 100 }
            elseif ($value -lt 4.68) { $score = 85 }
            elseif ($value -lt 4.9)  { $score = 65 }
            elseif ($value -lt 5.1)  { $score = 45 }
            else                     { $score = 20 }
        }
        "Triglyceride/HDL-Ratio" {
            # Optimal < 2.0, > 3.5 kritisch
            if ($value -lt 1.5)      { $score = 100 }
            elseif ($value -lt 2.0)  { $score = 85 }
            elseif ($value -lt 2.5)  { $score = 70 }
            elseif ($value -lt 3.0)  { $score = 50 }
            elseif ($value -lt 3.5)  { $score = 30 }
            else                     { $score = 10 }
        }
        default { $score = 50 }
    }
    return [Math]::Round($score, 0)
}

function Calculate-LongevityScore {
    param($values)
    $scores = New-Object System.Collections.ArrayList
    if ($values[0]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( (5.2 - $values[0]) / 1.2 ) * 100)) ) | Out-Null }
    if ($values[1]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( (5.0 - $values[1]) / 5.0 ) * 100)) ) | Out-Null }
    if ($values[2]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( ($values[2] - 30) / 30 ) * 100)) ) | Out-Null }
    if ($values[3]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( (7.0 - $values[3]) / 7.0 ) * 100)) ) | Out-Null }
    if ($values[4]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( (1.0 - $values[4]) / 1.0 ) * 100)) ) | Out-Null }
    if ($values[5]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( ($values[5] - 20) / 30 ) * 100)) ) | Out-Null }
    if ($values[6]) { $scores.Add( [Math]::Max(0, [Math]::Min(100, ( ($values[6] - 3.5) / 1.5 ) * 100)) ) | Out-Null }

    if ($scores.Count -eq 0) { return 0 }
    return ($scores | Measure-Object -Average).Average
}
function Calculate-PreventScore {
    param($values, $personal)
    $alter = $personal.Age
    $chol = $values[0]
    $hdl = $values[1]
    $bp = $values[2]
    $hba1c = $values[3]
    $raucher = if ($personal.Raucher) { 1 } else { 0 }
    $diabetes = if ($hba1c -ge 6.5) { 1 } else { 0 }
    
    $exponent = 0
    if ($personal.Geschlecht -eq 'männlich') {
        $exponent = -6.4 + (0.058 * $alter) + (0.02 * ($chol - $hdl)) + (0.01 * $bp) + (0.5 * $raucher) + (0.8 * $diabetes)
    } else {
        $exponent = -7.2 + (0.065 * $alter) + (0.018 * ($chol - $hdl)) + (0.012 * $bp) + (0.6 * $raucher) + (0.9 * $diabetes)
    }
    $risk = 1 / (1 + [Math]::Exp(-$exponent)) * 100
    return $risk
}

function Get-CalculatedValuesForMarker {
    param($markerName)
    $calculatedValues = New-Object System.Collections.ArrayList
    $markerConfig = [System.Collections.ArrayList]@($script:data.Config.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ }) | Where-Object { $_.Name -eq $markerName }
    if (-not $markerConfig) { return $calculatedValues }

    $specialFormulas = @("[LONGEVITY_SCORE]", "[PREVENT_SCORE]", "[PHENOAGE]", "[PHENOAGE_ACCEL]", "[INFLAMMAGING]")
    $isSpecialFormula = $specialFormulas -contains $markerConfig.Formula
    $groupedByDate = $script:allHistoricalItems | Group-Object -Property Date
    
    foreach ($dateGroup in $groupedByDate) {
        $date = $dateGroup.Name
        $valuesOnDate = $dateGroup.Group
        
        $formula = $markerConfig.Formula
        $formula = $formula -replace '\$personal\.(\w+)', ('$script:data.Config.Personal.{0}' -f '$1')

        $allMarkersFound = $true
        $markerValues = @()
        for ($i = 0; $i -lt $markerConfig.RequiredMarkers.Count; $i++) {
            $requiredMarkerName = $markerConfig.RequiredMarkers[$i]
            $foundMarker = $valuesOnDate | Where-Object { $_.Name -eq $requiredMarkerName } | Select-Object -First 1
            # v2.25.0: Ersatz-Marker mit identischer Einheit verwenden (z. B. hs-CRP statt CRP)
            if (-not $foundMarker -and $script:MarkerFallbacks.ContainsKey($requiredMarkerName)) {
                foreach ($fallbackName in $script:MarkerFallbacks[$requiredMarkerName]) {
                    $foundMarker = $valuesOnDate | Where-Object { $_.Name -eq $fallbackName } | Select-Object -First 1
                    if ($foundMarker) { break }
                }
            }
            if ($foundMarker) {
                if($isSpecialFormula) {
                    $markerValues += $foundMarker.Value
                } else {
                    $formula = $formula -replace "%$i%", $foundMarker.Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
                }
            } else {
                $allMarkersFound = $false; break
            }
        }

        if ($allMarkersFound) {
            try {
                $resultValue = 0
                if ($markerConfig.Formula -eq "[LONGEVITY_SCORE]") {
                    $resultValue = Calculate-LongevityScore -values $markerValues
                } elseif ($markerConfig.Formula -eq "[PREVENT_SCORE]") {
                    $resultValue = Calculate-PreventScore -values $markerValues -personal $script:data.Config.Personal
                } elseif ($markerConfig.Formula -eq "[PHENOAGE]") {
                    $chronoAge = [double]$script:data.Config.Personal.Age
                    $resultValue = Calculate-PhenoAge -values $markerValues -chronoAge $chronoAge
                    if ($null -eq $resultValue) { continue }
                } elseif ($markerConfig.Formula -eq "[PHENOAGE_ACCEL]") {
                    $chronoAge = [double]$script:data.Config.Personal.Age
                    $bioAge = Calculate-PhenoAge -values $markerValues -chronoAge $chronoAge
                    if ($null -eq $bioAge) { continue }
                    $resultValue = $bioAge - $chronoAge
                } elseif ($markerConfig.Formula -eq "[INFLAMMAGING]") {
                    $resultValue = Calculate-InflammAgingScore -values $markerValues
                    if ($null -eq $resultValue) { continue }
                } else {
                    $resultValue = Invoke-Expression $formula
                }
                
                $calculatedItem = [PSCustomObject]@{
                    Date = $date
                    Name = $markerName
                    Value = [Math]::Round($resultValue, 2)
                    Unit = $markerConfig.Unit
                }
                $calculatedValues.Add($calculatedItem) | Out-Null
            } catch {
                Write-Warning "Fehler bei der Berechnung von '$markerName' am ${date}: $($_.Exception.Message)"
            }
        }
    }
    return $calculatedValues
}
function Calculate-Correlation {
    param ($values1, $values2)
    $n = $values1.Count
    if ($n -eq 0) { return 0 }
    $sumX = $values1 | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumY = $values2 | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumXY = 0
    for ($i = 0; $i -lt $n; $i++) {
        $sumXY += $values1[$i] * $values2[$i]
    }
    $sumX2 = $values1 | ForEach-Object { $_ * $_ } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumY2 = $values2 | ForEach-Object { $_ * $_ } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $numerator = $n * $sumXY - $sumX * $sumY
    $denominator = [Math]::Sqrt(($n * $sumX2 - $sumX * $sumX) * ($n * $sumY2 - $sumY * $sumY))
    if ($denominator -eq 0) { return 0 }
    return $numerator / $denominator
}
function Calculate-LinearRegression {
    param($dataPoints)
    $n = $dataPoints.Count
    if ($n -lt 2) { return $null }
    $xValues = $dataPoints | ForEach-Object { ([datetime]$_.Date).ToOADate() }
    $yValues = $dataPoints.Value
    $sumX = $xValues | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumY = $yValues | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $sumXY = 0
    for ($i = 0; $i -lt $n; $i++) {
        $sumXY += $xValues[$i] * $yValues[$i]
    }
    $sumX2 = $xValues | ForEach-Object { $_ * $_ } | Measure-Object -Sum | Select-Object -ExpandProperty Sum
    $denominator = $n * $sumX2 - $sumX * $sumX
    if ($denominator -eq 0) { return $null }
    $slope = ($n * $sumXY - $sumX * $sumY) / $denominator
    $intercept = ($sumY - $slope * $sumX) / $n
    return [PSCustomObject]@{ Slope = $slope; Intercept = $intercept }
}

function Import-CsvData {
    param($filePath)
    try {
        $csvData = Import-Csv -Path $filePath
        $importedCount = 0
        $skippedCount = 0
        $markerNames = $script:data.Config.Markers.Name
        
        $newItems = New-Object System.Collections.ArrayList

        foreach ($row in $csvData) {
            $dateStr = $row.Date
            $markerName = $row.MarkerName
            $valueStr = $row.Value
            $note = if ($row.PSObject.Properties.Name -contains 'Note') { $row.Note } else { "" }
            try {
                $date = [datetime]::Parse($dateStr).ToString("yyyy-MM-dd")
                $value = Parse-Number $valueStr
                if ($value -eq $null) { throw "Leerer Wert." }
                if (-not ($markerNames -contains $markerName)) { throw "Unbekannter Marker: '$markerName'." }
                
                # Prüfen, ob der Eintrag bereits in den geladenen historischen Daten existiert
                $existingHistorical = @($script:allHistoricalItems | Where-Object { ($_.Date -eq $date) -and ($_.Name -eq $markerName) })
                if ($existingHistorical.Count -gt 0) { throw "Doppelter Eintrag (existiert bereits in historischen Daten)." }
                
                # Prüfen, ob der Eintrag bereits in den neu zu importierenden Daten existiert
                $existingInImport = @($newItems | Where-Object { ($_.Date -eq $date) -and ($_.Name -eq $markerName) })
                if ($existingInImport.Count -gt 0) { throw "Doppelter Eintrag (existiert bereits in CSV-Import)." }

                $markerConfig = $script:data.Config.Markers | Where-Object { $_.Name -eq $markerName } | Select-Object -First 1
                $newItem = [PSCustomObject]@{ Date = $date; Name = $markerName; Value = $value; Unit = $markerConfig.Unit; Note = $note }
                
                $newItems.Add($newItem) | Out-Null
                $importedCount++
            } catch {
                Write-Warning "Überspringe Zeile: $($row | Out-String) - Grund: $($_.Exception.Message)"
                $skippedCount++
            }
        }

        # KORREKTUR: Explizites Hinzufügen der Sammlung zur globalen Liste ELEMENT FÜR ELEMENT
        if ($newItems.Count -gt 0) {
            foreach($item in $newItems) {
                [void]$script:allHistoricalItems.Add($item)
            }
            Save-AllHistoricalData -AllItems $script:allHistoricalItems
        }
        
        [System.Windows.Forms.MessageBox]::Show("$importedCount Einträge erfolgreich importiert. $skippedCount Einträge wurden übersprungen.", "Import abgeschlossen", "OK", "Information")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Importieren der CSV-Datei: $($_.Exception.Message)", "Importfehler", "OK", "Error")
    }
}


# ---------- UI Framework ----------
function Show-EditPopup {
    param($markerName, $historicalData)
    $popupForm = New-Object System.Windows.Forms.Form; $popupForm.Size = New-Object System.Drawing.Size(500, 450); $popupForm.Text = "Einträge für '$markerName' bearbeiten"; $popupForm.StartPosition = "CenterParent"; $popupForm.FormBorderStyle = "FixedDialog"; $popupForm.MinimizeBox = $false
    $changesMade = $false
    $dataGridView = New-Object System.Windows.Forms.DataGridView; $dataGridView.Location = New-Object System.Drawing.Point(10, 10); $dataGridView.Size = New-Object System.Drawing.Size(460, 320); $dataGridView.Anchor = "Top, Left, Right"; $dataGridView.AutoSizeColumnsMode = "Fill"; $dataGridView.AllowUserToAddRows = $false; $dataGridView.AllowUserToDeleteRows = $true; $dataGridView.SelectionMode = "FullRowSelect"; $popupForm.Controls.Add($dataGridView)
    
    $dateCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $dateCol.Name = "Datum"; $dateCol.HeaderText = "Datum"; $dateCol.ReadOnly = $true
    $valueCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $valueCol.Name = "Wert"; $valueCol.HeaderText = "Wert"
    $noteCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $noteCol.Name = "Notiz"; $noteCol.HeaderText = "Notiz"

    $dataGridView.Columns.Add($dateCol); $dataGridView.Columns.Add($valueCol); $dataGridView.Columns.Add($noteCol)

    $markerData = $historicalData | Where-Object { $_.Name -eq $markerName } | Sort-Object Date -Descending
    foreach ($item in $markerData) {
        $rowIndex = $dataGridView.Rows.Add($item.Date, $item.Value, $item.Note)
        $dataGridView.Rows[$rowIndex].Tag = $item
    }

    $deleteButton = New-Object System.Windows.Forms.Button; $deleteButton.Text = "Auswahl löschen"; $deleteButton.Location = New-Object System.Drawing.Point(10, 340); $deleteButton.Size = New-Object System.Drawing.Size(150, 30)
    $saveButton = New-Object System.Windows.Forms.Button; $saveButton.Text = "Änderungen übernehmen && Schließen"; $saveButton.Location = New-Object System.Drawing.Point(220, 370); $saveButton.Size = New-Object System.Drawing.Size(250, 40)
    $closeButton = New-Object System.Windows.Forms.Button; $closeButton.Text = "Abbrechen"; $closeButton.Location = New-Object System.Drawing.Point(10, 370); $closeButton.Size = New-Object System.Drawing.Size(120, 40)
    $popupForm.Controls.AddRange(@($deleteButton, $saveButton, $closeButton))
    $deleteButton.Add_Click({ if ($dataGridView.SelectedRows.Count -eq 0) { return }; $result = [System.Windows.Forms.MessageBox]::Show("Möchten Sie die ausgewählten $($dataGridView.SelectedRows.Count) Einträge wirklich dauerhaft löschen?", "Löschen bestätigen", "YesNo", "Warning"); if ($result -eq "Yes") { $rowsToDelete = New-Object System.Collections.ArrayList; $rowsToDelete.AddRange($dataGridView.SelectedRows); foreach ($row in $rowsToDelete) { $itemToRemove = $row.Tag; $historicalData.Remove($itemToRemove); $dataGridView.Rows.Remove($row) }; $changesMade = $true } })
    $saveButton.Add_Click({ try { foreach ($row in $dataGridView.Rows) { $originalItem = $row.Tag; $newValueStr = $row.Cells["Wert"].Value; $newValue = Parse-Number $newValueStr; $newNote = $row.Cells["Notiz"].Value; if ($originalItem.Value -ne $newValue) { $originalItem.Value = $newValue; $changesMade = $true }; if ($originalItem.Note -ne $newNote) { $originalItem.Note = $newNote; $changesMade = $true } }; if ($changesMade) { $popupForm.DialogResult = [System.Windows.Forms.DialogResult]::OK } else { $popupForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel }; $popupForm.Close() } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Verarbeiten der Werte: $($_.Exception.Message)", "Fehler", "OK", "Error") } })
    $closeButton.Add_Click({ $popupForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $popupForm.Close() }); return $popupForm.ShowDialog()
}
function Show-ExportPopup {
    <#
    .SYNOPSIS
        Kontextabhängiges Export-Popup für alle Tabs (v2.11.0, PDF ergänzt in v2.12.0).
    .DESCRIPTION
        Bietet JSON-, CSV- und PDF-Export der im jeweiligen Tab sichtbaren Daten.
        Orientiert sich am Design des "Einstellungen"-Popups.
    .PARAMETER source
        Kontext-String: "Cockpit" | "Chart" | "Correlation" | "DataMgmt" | "Custom Report"
    .PARAMETER dataProvider
        ScriptBlock, der ein Array von Objekten zurückliefert (oder $null bei leer).
    .PARAMETER defaultFileName
        Default-Dateiname ohne Endung. Wird per Dialog vom Nutzer änderbar.
    .PARAMETER pdfTitle
        Optional: Titel für den PDF-Header. Default = $source.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$source,
        [Parameter(Mandatory=$true)][scriptblock]$dataProvider,
        [string]$defaultFileName = "Export_$(Get-Date -Format 'yyyy-MM-dd')",
        [string]$pdfTitle = $null
    )

    if (-not $pdfTitle) { $pdfTitle = "Report: $source" }

    $popupForm = New-Object System.Windows.Forms.Form
    $popupForm.Size = New-Object System.Drawing.Size(450, 310)
    $popupForm.Text = "Exportieren – $source"
    $popupForm.StartPosition = "CenterParent"
    $popupForm.FormBorderStyle = "FixedDialog"
    $popupForm.MaximizeBox = $false
    $popupForm.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Wähle ein Format zum Exportieren der aktuellen Daten."
    $label.Location = New-Object System.Drawing.Point(20, 20)
    $label.Size = New-Object System.Drawing.Size(400, 20)
    $popupForm.Controls.Add($label)

    $jsonButton = New-Object System.Windows.Forms.Button
    $jsonButton.Text = "Als JSON exportieren (.json)..."
    $jsonButton.Location = New-Object System.Drawing.Point(20, 60)
    $jsonButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($jsonButton)

    $csvButton = New-Object System.Windows.Forms.Button
    $csvButton.Text = "Als CSV exportieren (.csv)..."
    $csvButton.Location = New-Object System.Drawing.Point(20, 105)
    $csvButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($csvButton)

    $pdfButton = New-Object System.Windows.Forms.Button
    $pdfButton.Text = "Als PDF exportieren (.pdf)..."
    $pdfButton.Location = New-Object System.Drawing.Point(20, 150)
    $pdfButton.Size = New-Object System.Drawing.Size(400, 35)
    $popupForm.Controls.Add($pdfButton)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Schließen"
    $closeButton.Location = New-Object System.Drawing.Point(170, 220)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $popupForm.Controls.Add($closeButton)
    $popupForm.CancelButton = $closeButton

    # Gemeinsame Export-Routine für JSON und CSV
    $exportAction = {
        param($format)
        try {
            $data = & $dataProvider
            if (-not $data -or ($data | Measure-Object).Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", "OK", "Information")
                return
            }

            $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
            if ($format -eq "json") {
                $saveDlg.Filter = "JSON-Datei (*.json)|*.json"
                $saveDlg.FileName = "$defaultFileName.json"
            } else {
                $saveDlg.Filter = "CSV-Datei (*.csv)|*.csv"
                $saveDlg.FileName = "$defaultFileName.csv"
            }
            $saveDlg.Title = "Exportziel wählen"

            if ($saveDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

            $exportPath = $saveDlg.FileName
            if ($format -eq "json") {
                $data | ConvertTo-Json -Depth 5 | Out-File -FilePath $exportPath -Encoding UTF8
            } else {
                $data | Export-Csv -Path $exportPath -Delimiter ";" -Encoding UTF8 -NoTypeInformation
            }
            $popupForm.Close()
            Invoke-PostExportAction -FilePath $exportPath -SuccessTitle "Export"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)", "Exportfehler", "OK", "Error")
        }
    }.GetNewClosure()

    # PDF-Export via systemweitem "Microsoft Print to PDF"-Drucker (Windows 10+)
    # v2.13.0: Fix für "Arrayindex wurde als NULL ausgewertet" bei 1-elementigen Daten.
    $pdfAction = {
        try {
            $rawData = & $dataProvider
            if (-not $rawData) {
                [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", "OK", "Information")
                return
            }
            # Explizit als Array casten (verhindert Auto-Unwrapping bei 1 Element)
            $arrData = @($rawData)
            if ($arrData.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", "OK", "Information")
                return
            }

            $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
            $saveDlg.Filter = "PDF-Datei (*.pdf)|*.pdf"
            $saveDlg.FileName = "$defaultFileName.pdf"
            $saveDlg.Title = "PDF-Exportziel wählen"
            if ($saveDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $pdfPath = $saveDlg.FileName

            # PDF-Drucker ermitteln (Windows 10+: "Microsoft Print to PDF")
            Add-Type -AssemblyName System.Drawing
            $installed = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters
            $pdfPrinterName = $null
            foreach ($p in $installed) {
                if ($p -match "Microsoft Print to PDF|Microsoft PDF|PDF|Foxit|Adobe PDF") {
                    $pdfPrinterName = $p
                    if ($p -eq "Microsoft Print to PDF") { break }
                }
            }
            if (-not $pdfPrinterName) {
                [System.Windows.Forms.MessageBox]::Show("Kein PDF-Drucker gefunden.`n`nBitte aktiviere 'Microsoft Print to PDF' unter Windows-Features, oder installiere einen PDF-Drucker (z.B. Foxit).", "PDF-Drucker fehlt", "OK", "Warning")
                return
            }

            # Script-scope statt Closure (zuverlässiger bei asynchronem PrintPage)
            $script:pdfData      = $arrData
            $script:pdfProps     = @($arrData[0].PSObject.Properties.Name)
            $script:pdfRowIndex  = 0
            $script:pdfPageNum   = 0
            $script:pdfTitleStr  = $pdfTitle

            $printDoc = New-Object System.Drawing.Printing.PrintDocument
            $printDoc.PrinterSettings.PrinterName = $pdfPrinterName
            $printDoc.PrinterSettings.PrintToFile = $true
            $printDoc.PrinterSettings.PrintFileName = $pdfPath
            $printDoc.DefaultPageSettings.Landscape = $true
            $printDoc.DocumentName = $pdfTitle

            $printDoc.Add_PrintPage({
                param($sender, $ev)
                $g = $ev.Graphics
                $margins = $ev.MarginBounds
                $titleFont  = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
                $headerFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $cellFont   = New-Object System.Drawing.Font("Segoe UI", 9)
                $smallFont  = New-Object System.Drawing.Font("Segoe UI", 8)
                $pen        = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 0.5)
                $yPos       = $margins.Top
                $rowHeight  = 22

                if ($script:pdfPageNum -eq 0) {
                    $g.DrawString($script:pdfTitleStr, $titleFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += 28
                    $g.DrawString("Exportdatum: $(Get-Date -Format 'dd.MM.yyyy HH:mm')", $smallFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, $yPos)
                    $yPos += 20
                }
                $script:pdfPageNum++

                # Spalten-Breiten proportional verteilen
                $props = $script:pdfProps
                $totalW = $margins.Width
                $colW   = [Math]::Floor($totalW / $props.Count)

                # Kopfzeile
                $xPos = $margins.Left
                foreach ($prop in $props) {
                    $g.FillRectangle([System.Drawing.Brushes]::LightGray, $xPos, $yPos, $colW, $rowHeight)
                    $g.DrawRectangle($pen, $xPos, $yPos, $colW, $rowHeight)
                    $g.DrawString($prop, $headerFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 4))
                    $xPos += $colW
                }
                $yPos += $rowHeight

                # Datenzeilen
                $totalRows = $script:pdfData.Count
                while ($script:pdfRowIndex -lt $totalRows) {
                    if (($yPos + $rowHeight) -gt ($margins.Top + $margins.Height - 20)) {
                        $ev.HasMorePages = $true
                        return
                    }
                    $item = $script:pdfData[$script:pdfRowIndex]
                    $xPos = $margins.Left
                    foreach ($prop in $props) {
                        $val = [string]$item.$prop
                        if ($val.Length -gt 60) { $val = $val.Substring(0, 57) + "..." }
                        $g.DrawRectangle($pen, $xPos, $yPos, $colW, $rowHeight)
                        $g.DrawString($val, $cellFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 4))
                        $xPos += $colW
                    }
                    $yPos += $rowHeight
                    $script:pdfRowIndex++
                }

                # Footer
                $footerText = "Seite $($script:pdfPageNum) – erzeugt durch Blood-Tracker"
                $g.DrawString($footerText, $smallFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, ($margins.Top + $margins.Height - 15))

                $ev.HasMorePages = $false
            })

            $printDoc.Print()
            $popupForm.Close()
            Invoke-PostExportAction -FilePath $pdfPath -SuccessTitle "PDF-Export" -SuccessMessage "PDF erfolgreich erstellt:`n$pdfPath"
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim PDF-Export: $($_.Exception.Message)", "Exportfehler", "OK", "Error")
        }
    }.GetNewClosure()

    $jsonButton.Add_Click({ & $exportAction "json" }.GetNewClosure())
    $csvButton.Add_Click({ & $exportAction "csv" }.GetNewClosure())
    $pdfButton.Add_Click({ & $pdfAction }.GetNewClosure())

    $popupForm.ShowDialog() | Out-Null
}

# ---------- AutoBackup-System (v2.14.0) ----------
#
# Architektur:
#   Invoke-AutoBackupIfDue  - Zentraler Scheduler-Check, bei Programm-Start und periodisch
#   Invoke-AutoBackupNow    - Führt das Backup tatsächlich aus (ZIP und/oder Ordner-Kopie)
#   Show-AutoBackupPopup    - Konfigurations-Dialog
#
# Registry wird mit gesichert (Konsistent mit Uninstall-Backup), da der DataPath-Wert
# essentiell für Restore auf anderem Gerät ist.

function Test-BackupPathSafe {
    param([string]$TargetPath, [string]$DataDir)
    if ([string]::IsNullOrWhiteSpace($TargetPath)) { return "Zielverzeichnis darf nicht leer sein." }
    try {
        $t = [System.IO.Path]::GetFullPath($TargetPath.TrimEnd('\'))
        $d = [System.IO.Path]::GetFullPath($DataDir.TrimEnd('\'))
        if ($t -ieq $d) { return "Zielverzeichnis darf nicht das Datenverzeichnis selbst sein." }
        if ($t.StartsWith($d + [System.IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            return "Zielverzeichnis darf nicht unterhalb des Datenverzeichnisses liegen (rekursive Backups)."
        }
    } catch {
        return "Ungültiger Pfad: $($_.Exception.Message)"
    }
    return $null
}

function Invoke-BackupCleanup {
    <#
    .SYNOPSIS
        Löscht ältere Backups gemäß Retention-Regel (v2.15.0).
    .DESCRIPTION
        Betrachtet ausschließlich Einträge mit Prefix "BloodTracker_AutoBackup_" im
        Zielverzeichnis. ZIP-Archive und Backup-Ordner werden getrennt verwaltet,
        damit bei Format="Beides" keine ungewollten Lücken entstehen.
    .PARAMETER TargetPath
        Zielverzeichnis mit den Backups.
    .PARAMETER Keep
        Anzahl der zu behaltenden Backups (pro Typ). 0 = alle außer dem neuesten löschen.
    #>
    param(
        [Parameter(Mandatory=$true)][string]$TargetPath,
        [Parameter(Mandatory=$true)][int]$Keep
    )
    try {
        if (-not (Test-Path $TargetPath)) { return }

        # ZIP-Backups (sortiert nach Erstellungsdatum, neueste zuerst)
        $zipBackups = @(Get-ChildItem -Path $TargetPath -Filter "BloodTracker_AutoBackup_*.zip" -File -ErrorAction SilentlyContinue |
                        Sort-Object -Property CreationTime -Descending)
        # Ordner-Backups (sortiert nach Erstellungsdatum, neueste zuerst)
        $folderBackups = @(Get-ChildItem -Path $TargetPath -Filter "BloodTracker_AutoBackup_*" -Directory -ErrorAction SilentlyContinue |
                           Sort-Object -Property CreationTime -Descending)

        # "Alle löschen" = Keep=0 bedeutet: nur das neueste bleibt (das gerade erstellte)
        # Daher immer mindestens 1 Element behalten, damit kein Total-Wipe passiert
        $effectiveKeep = [Math]::Max($Keep, 1)

        # ZIP-Cleanup
        if ($zipBackups.Count -gt $effectiveKeep) {
            $toDelete = $zipBackups | Select-Object -Skip $effectiveKeep
            foreach ($f in $toDelete) {
                try {
                    Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                    Write-Verbose "Cleanup: ZIP gelöscht: $($f.Name)"
                } catch {
                    Write-Warning "Cleanup-Fehler bei $($f.Name): $($_.Exception.Message)"
                }
            }
        }

        # Ordner-Cleanup
        if ($folderBackups.Count -gt $effectiveKeep) {
            $toDelete = $folderBackups | Select-Object -Skip $effectiveKeep
            foreach ($d in $toDelete) {
                try {
                    Remove-Item -Path $d.FullName -Recurse -Force -ErrorAction Stop
                    Write-Verbose "Cleanup: Ordner gelöscht: $($d.Name)"
                } catch {
                    Write-Warning "Cleanup-Fehler bei $($d.Name): $($_.Exception.Message)"
                }
            }
        }

        # v2.19.0: .btbackup-Cleanup (Nachtrag für Template-Script)
        $btbackupFiles = @(Get-ChildItem -Path $TargetPath -Filter "BloodTracker_AutoBackup_*.btbackup" -File -ErrorAction SilentlyContinue |
                           Sort-Object -Property CreationTime -Descending)
        if ($btbackupFiles.Count -gt $effectiveKeep) {
            $toDelete = $btbackupFiles | Select-Object -Skip $effectiveKeep
            foreach ($f in $toDelete) {
                try {
                    Remove-Item -Path $f.FullName -Force -ErrorAction Stop
                    Write-Verbose "Cleanup: .btbackup gelöscht: $($f.Name)"
                } catch {
                    Write-Warning "Cleanup-Fehler bei $($f.Name): $($_.Exception.Message)"
                }
            }
        }
    } catch {
        Write-Warning "Invoke-BackupCleanup-Fehler: $($_.Exception.Message)"
    }
}

function Invoke-AutoBackupNow {
    <#
    .SYNOPSIS
        Führt ein automatisches Backup gemäß Config aus.
    .OUTPUTS
        $true = Erfolgreich, $false = Fehler
    #>
    try {
        $bk = $script:data.Config.AutoBackup
        if (-not $bk -or -not $bk.Enabled) { return $false }
        if (-not $bk.FormatZip -and -not $bk.FormatFolder) { return $false }
        $err = Test-BackupPathSafe -TargetPath $bk.TargetPath -DataDir $script:dataDir
        if ($err) { Write-Warning "AutoBackup: $err"; return $false }
        if (-not (Test-Path $bk.TargetPath)) {
            New-Item -ItemType Directory -Path $bk.TargetPath -Force | Out-Null
        }
        $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

        # Registry exportieren (optional, aber empfohlen)
        $tempRegFile = $null
        if ($bk.IncludeRegistry) {
            try {
                $tempRegFile = Join-Path -Path $env:TEMP -ChildPath "BloodTracker_reg_$timestamp.reg"
                # HKCU\Software\PSC\PSC.Blood-Tracker exportieren
                $regKey = "HKCU\Software\PSC\PSC.Blood-Tracker"
                $null = & reg.exe export $regKey $tempRegFile /y 2>&1
                if (-not (Test-Path $tempRegFile)) { $tempRegFile = $null }
            } catch {
                Write-Warning "Registry-Export fehlgeschlagen: $($_.Exception.Message)"
                $tempRegFile = $null
            }
        }

        # ZIP-Format
        if ($bk.FormatZip) {
            $zipPath = Join-Path -Path $bk.TargetPath -ChildPath "BloodTracker_AutoBackup_$timestamp.zip"
            # ZIP erstellen mit .NET-API, da Compress-Archive nur aus einer Quelle lesen kann
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
            $tempStagingDir = Join-Path -Path $env:TEMP -ChildPath "BloodTracker_stage_$timestamp"
            New-Item -ItemType Directory -Path $tempStagingDir -Force | Out-Null
            try {
                # UserData in Staging-Unterordner kopieren
                $stageData = Join-Path -Path $tempStagingDir -ChildPath "UserData"
                Copy-Item -Path $script:dataDir -Destination $stageData -Recurse -Force
                if ($tempRegFile) {
                    Copy-Item -Path $tempRegFile -Destination (Join-Path $tempStagingDir "Registry.reg") -Force
                }
                [System.IO.Compression.ZipFile]::CreateFromDirectory($tempStagingDir, $zipPath)
            } finally {
                if (Test-Path $tempStagingDir) { Remove-Item $tempStagingDir -Recurse -Force -ErrorAction SilentlyContinue }
            }
        }

        # Ordner-Kopie-Format (Robocopy-basiert für Inkrementalität)
        if ($bk.FormatFolder) {
            $folderBackup = Join-Path -Path $bk.TargetPath -ChildPath "BloodTracker_AutoBackup_$timestamp"
            $folderUserData = Join-Path -Path $folderBackup -ChildPath "UserData"
            New-Item -ItemType Directory -Path $folderUserData -Force | Out-Null
            # Robocopy (nur geänderte/neue Dateien kopieren)
            $robocopyArgs = @($script:dataDir, $folderUserData, "/E", "/R:1", "/W:1", "/NP", "/NFL", "/NDL", "/NJH", "/NJS")
            $null = & robocopy.exe @robocopyArgs 2>&1
            # Robocopy exit codes 0-7 sind OK
            if ($LASTEXITCODE -ge 8) {
                Write-Warning "Robocopy Exit-Code $LASTEXITCODE - Kopie evtl. unvollständig"
            }
            if ($tempRegFile) {
                Copy-Item -Path $tempRegFile -Destination (Join-Path $folderBackup "Registry.reg") -Force
            }
        }

        # Temp-Registry-Datei aufräumen
        if ($tempRegFile -and (Test-Path $tempRegFile)) {
            Remove-Item $tempRegFile -Force -ErrorAction SilentlyContinue
        }

        # v2.19.0: Portables .btbackup erstellen (Nachtrag für Template-Script)
        if ($bk.PortableBackup -and $bk.PortableBackupPassword) {
            try {
                $encPass = [Convert]::FromBase64String($bk.PortableBackupPassword)
                $decPass = [System.Security.Cryptography.ProtectedData]::Unprotect(
                    $encPass, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser
                )
                $plainPass = [System.Text.Encoding]::UTF8.GetString($decPass)
                $ss = New-Object System.Security.SecureString
                foreach ($c in $plainPass.ToCharArray()) { $ss.AppendChar($c) }
                $ss.MakeReadOnly()
                $plainPass = $null
                $btbackupPath = Join-Path -Path $bk.TargetPath -ChildPath "BloodTracker_AutoBackup_$timestamp.btbackup"
                Export-AesBackup -OutputPath $btbackupPath -Passphrase $ss
                Write-Verbose "Portables Backup erstellt: $btbackupPath"
            } catch {
                Write-Warning "Portables Backup fehlgeschlagen: $($_.Exception.Message)"
            }
        }

        # Last-Backup-Timestamp aktualisieren
        $script:data.Config.AutoBackup.LastBackup = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
        Save-Data -data $script:data -type "Config"

        # v2.15.0: Backup-Cleanup NACH erfolgreicher Erstellung (damit bei Fehler nichts verloren geht)
        if ($bk.CleanupEnabled) {
            $keep = switch ([string]$bk.CleanupKeep) {
                "letztes Backup behalten"           { 1 }
                "letzte zwei Backups behalten"      { 2 }
                "letzte drei Backups behalten"      { 3 }
                "alle löschen"                      { 0 }   # wird intern zu 1 (neuestes bleibt)
                default                             { 3 }
            }
            Invoke-BackupCleanup -TargetPath $bk.TargetPath -Keep $keep
        }

        return $true
    } catch {
        Write-Warning "AutoBackup-Fehler: $($_.Exception.Message)"
        return $false
    }
}

function Invoke-AutoBackupIfDue {
    <#
    .SYNOPSIS
        Prüft ob ein Backup fällig ist und führt es ggf. aus.
        Wird beim Programm-Start und periodisch (via Timer) aufgerufen.
    #>
    try {
        $bk = $script:data.Config.AutoBackup
        if (-not $bk -or -not $bk.Enabled) { return }
        if (-not $bk.FormatZip -and -not $bk.FormatFolder) { return }
        if ([string]::IsNullOrWhiteSpace($bk.TargetPath)) { return }

        $now = Get-Date
        $lastBackup = $null
        if (-not [string]::IsNullOrWhiteSpace($bk.LastBackup)) {
            try { $lastBackup = [datetime]::Parse($bk.LastBackup) } catch { $lastBackup = $null }
        }

        $due = $false
        switch ($bk.Interval) {
            "bei Programm-Start" {
                # Nur einmal pro Session ausführen - via Script-Flag
                if (-not $script:autoBackupRanThisSession) { $due = $true }
            }
            "täglich"     { if (-not $lastBackup -or $lastBackup.Date -lt $now.Date) { $due = $true } }
            "wöchentlich" {
                # v2.14.1: Prüft Wochentags-Auswahl. Backup erfolgt nur an gewählten
                # Wochentagen UND wenn seit letztem Backup >=1 Tag vergangen ist.
                $selectedDays = @($bk.Weekdays)
                if ($selectedDays.Count -eq 0) {
                    # Fallback: wenn keine Tage gewählt, klassische 7-Tage-Logik
                    if (-not $lastBackup -or ($now - $lastBackup).TotalDays -ge 7) { $due = $true }
                } else {
                    # Wochentag-Mapping (.NET DayOfWeek enum → deutsche Kurzform)
                    $dayMap = @{
                        "Monday"="Mo"; "Tuesday"="Di"; "Wednesday"="Mi"; "Thursday"="Do";
                        "Friday"="Fr"; "Saturday"="Sa"; "Sunday"="So"
                    }
                    $todayShort = $dayMap[[string]$now.DayOfWeek]
                    if ($selectedDays -contains $todayShort) {
                        # Heute ist ein gewählter Tag - Backup wenn heute noch kein Backup lief
                        if (-not $lastBackup -or $lastBackup.Date -lt $now.Date) { $due = $true }
                    }
                }
            }
            "monatlich"   {
                # v2.14.2: Mit MonthDay-Auswahl. Backup erfolgt am gewählten Tag des Monats,
                # wenn im aktuellen Monat noch kein Backup lief.
                $dayOfMonth = $now.Day
                $daysInMonth = [datetime]::DaysInMonth($now.Year, $now.Month)
                $targetDay = 1
                switch ([string]$bk.MonthDay) {
                    "am Ersten eines Monats"         { $targetDay = 1 }
                    "am zweiten Tag"                 { $targetDay = 2 }
                    "am dritten Tag"                 { $targetDay = 3 }
                    "am vierten Tag"                 { $targetDay = 4 }
                    "am fünften Tag"                 { $targetDay = 5 }
                    "zur Monatsmitte (15.)"          { $targetDay = 15 }
                    "am Letzten Tag eines Monats"    { $targetDay = $daysInMonth }
                    default                          { $targetDay = 1 }
                }
                if ($dayOfMonth -ge $targetDay) {
                    # Prüfen ob im aktuellen Monat schon ein Backup lief
                    if (-not $lastBackup -or
                        $lastBackup.Year -lt $now.Year -or
                        ($lastBackup.Year -eq $now.Year -and $lastBackup.Month -lt $now.Month)) {
                        $due = $true
                    }
                }
            }
        }

        if (-not $due) { return }

        # Bei intervallbasierten Backups: Uhrzeit-Check (gilt nicht für "bei Programm-Start")
        if ($bk.Interval -ne "bei Programm-Start") {
            try {
                $targetTime = [datetime]::ParseExact($bk.TimeOfDay, "HH:mm", $null)
                $todayTarget = $now.Date.Add($targetTime.TimeOfDay)
                if ($now -lt $todayTarget) {
                    # Heute noch nicht fällig (Zeitpunkt noch nicht erreicht)
                    return
                }
            } catch {
                Write-Warning "AutoBackup: Ungültige Uhrzeit '$($bk.TimeOfDay)', Backup wird trotzdem ausgeführt."
            }
        }

        $success = Invoke-AutoBackupNow
        if ($success) {
            $script:autoBackupRanThisSession = $true
        }
    } catch {
        Write-Warning "AutoBackupIfDue-Fehler: $($_.Exception.Message)"
    }
}

function Show-AutoBackupPopup {
    <#
    .SYNOPSIS
        Konfigurations-Dialog für automatisches Backup (v2.14.0).
    #>
    $bk = $script:data.Config.AutoBackup

    $popup = New-Object System.Windows.Forms.Form
    $popup.Size = New-Object System.Drawing.Size(600, 950)
    $popup.Text = "Automatisches Backup konfigurieren"
    $popup.StartPosition = "CenterParent"
    $popup.FormBorderStyle = "FixedDialog"
    $popup.MaximizeBox = $false; $popup.MinimizeBox = $false

    # --- GroupBox 1: Basis-Konfiguration ---
    $groupBasis = New-Object System.Windows.Forms.GroupBox
    $groupBasis.Text = "Basis-Konfiguration"
    $groupBasis.Location = New-Object System.Drawing.Point(15, 15)
    $groupBasis.Size = New-Object System.Drawing.Size(555, 340)
    $popup.Controls.Add($groupBasis)

    # Enabled-Checkbox
    $chkEnabled = New-Object System.Windows.Forms.CheckBox
    $chkEnabled.Text = "Automatisches Backup aktivieren"
    $chkEnabled.Location = New-Object System.Drawing.Point(15, 25)
    $chkEnabled.Size = New-Object System.Drawing.Size(300, 22)
    $chkEnabled.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $chkEnabled.Checked = [bool]$bk.Enabled
    $groupBasis.Controls.Add($chkEnabled)

    # Zielverzeichnis
    $lblTarget = New-Object System.Windows.Forms.Label
    $lblTarget.Text = "Zielverzeichnis:"
    $lblTarget.Location = New-Object System.Drawing.Point(15, 60)
    $lblTarget.Size = New-Object System.Drawing.Size(100, 20)
    $groupBasis.Controls.Add($lblTarget)

    $txtTarget = New-Object System.Windows.Forms.TextBox
    $txtTarget.Location = New-Object System.Drawing.Point(120, 57)
    $txtTarget.Size = New-Object System.Drawing.Size(320, 22)
    $txtTarget.Text = [string]$bk.TargetPath
    $groupBasis.Controls.Add($txtTarget)

    $btnBrowse = New-Object System.Windows.Forms.Button
    $btnBrowse.Text = "Durchsuchen..."
    $btnBrowse.Location = New-Object System.Drawing.Point(445, 55)
    $btnBrowse.Size = New-Object System.Drawing.Size(100, 26)
    $groupBasis.Controls.Add($btnBrowse)
    $btnBrowse.Add_Click({
        $folderDlg = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderDlg.Description = "Zielverzeichnis für automatische Backups wählen"
        if ($txtTarget.Text -and (Test-Path $txtTarget.Text)) { $folderDlg.SelectedPath = $txtTarget.Text }
        if ($folderDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $txtTarget.Text = $folderDlg.SelectedPath
        }
    })

    # Format-Label
    $lblFormat = New-Object System.Windows.Forms.Label
    $lblFormat.Text = "Format:"
    $lblFormat.Location = New-Object System.Drawing.Point(15, 100)
    $lblFormat.Size = New-Object System.Drawing.Size(100, 20)
    $groupBasis.Controls.Add($lblFormat)

    # Format-Checkboxen
    $chkZip = New-Object System.Windows.Forms.CheckBox
    $chkZip.Text = "ZIP (Export)"
    $chkZip.Location = New-Object System.Drawing.Point(120, 100)
    $chkZip.Size = New-Object System.Drawing.Size(140, 22)
    $groupBasis.Controls.Add($chkZip)

    $chkFolder = New-Object System.Windows.Forms.CheckBox
    $chkFolder.Text = "Ordner und Dateien"
    $chkFolder.Location = New-Object System.Drawing.Point(270, 100)
    $chkFolder.Size = New-Object System.Drawing.Size(170, 22)
    $groupBasis.Controls.Add($chkFolder)

    $chkBoth = New-Object System.Windows.Forms.CheckBox
    $chkBoth.Text = "Beides"
    $chkBoth.Location = New-Object System.Drawing.Point(450, 100)
    $chkBoth.Size = New-Object System.Drawing.Size(90, 22)
    $groupBasis.Controls.Add($chkBoth)

    # Initial-State aus Config
    $chkZip.Checked    = [bool]$bk.FormatZip
    $chkFolder.Checked = [bool]$bk.FormatFolder
    $chkBoth.Checked   = ([bool]$bk.FormatZip -and [bool]$bk.FormatFolder)

    # Registry-Checkbox
    $chkRegistry = New-Object System.Windows.Forms.CheckBox
    $chkRegistry.Text = "Registry-Einstellungen mitsichern (empfohlen)"
    $chkRegistry.Location = New-Object System.Drawing.Point(15, 140)
    $chkRegistry.Size = New-Object System.Drawing.Size(500, 22)
    $chkRegistry.Checked = [bool]$bk.IncludeRegistry
    $groupBasis.Controls.Add($chkRegistry)

    $lblRegInfo = New-Object System.Windows.Forms.Label
    $lblRegInfo.Text = "Sichert den DataPath-Konfigurationseintrag. Wichtig für Restore auf anderem Gerät."
    $lblRegInfo.Location = New-Object System.Drawing.Point(35, 162)
    $lblRegInfo.Size = New-Object System.Drawing.Size(500, 20)
    $lblRegInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblRegInfo.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupBasis.Controls.Add($lblRegInfo)

    # Status-Label (Last-Backup)
    $lblStatus = New-Object System.Windows.Forms.Label
    $lblStatus.Location = New-Object System.Drawing.Point(15, 195)
    $lblStatus.Size = New-Object System.Drawing.Size(530, 20)
    $lblStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    if ($bk.LastBackup) {
        try {
            $d = [datetime]::Parse($bk.LastBackup)
            $lblStatus.Text = "Letztes erfolgreiches Backup: $($d.ToString('dd.MM.yyyy HH:mm'))"
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        } catch {
            $lblStatus.Text = "Letztes Backup: unbekannt"
        }
    } else {
        $lblStatus.Text = "Noch kein Backup durchgeführt."
        $lblStatus.ForeColor = [System.Drawing.Color]::DarkSlateGray
    }
    $groupBasis.Controls.Add($lblStatus)

    # Sub-GroupBox "Backup-Bereinigung" (v2.15.0, Höhe erhöht in v2.15.2)
    $groupCleanup = New-Object System.Windows.Forms.GroupBox
    $groupCleanup.Text = "Backup-Bereinigung"
    $groupCleanup.Location = New-Object System.Drawing.Point(15, 222)
    $groupCleanup.Size = New-Object System.Drawing.Size(525, 105)
    $groupBasis.Controls.Add($groupCleanup)

    $chkCleanup = New-Object System.Windows.Forms.CheckBox
    $chkCleanup.Text = "Ältere Backups löschen"
    $chkCleanup.Location = New-Object System.Drawing.Point(15, 25)
    $chkCleanup.Size = New-Object System.Drawing.Size(200, 22)
    $chkCleanup.Checked = [bool]$bk.CleanupEnabled
    $groupCleanup.Controls.Add($chkCleanup)

    $lblCleanupRule = New-Object System.Windows.Forms.Label
    $lblCleanupRule.Text = "Regel:"
    $lblCleanupRule.Location = New-Object System.Drawing.Point(230, 28)
    $lblCleanupRule.Size = New-Object System.Drawing.Size(50, 20)
    $groupCleanup.Controls.Add($lblCleanupRule)

    $cmbCleanupKeep = New-Object System.Windows.Forms.ComboBox
    $cmbCleanupKeep.Location = New-Object System.Drawing.Point(285, 25)
    $cmbCleanupKeep.Size = New-Object System.Drawing.Size(225, 22)
    $cmbCleanupKeep.DropDownStyle = "DropDownList"
    $cmbCleanupKeep.Items.AddRange(@(
        "letztes Backup behalten",
        "letzte zwei Backups behalten",
        "letzte drei Backups behalten",
        "alle löschen"
    ))
    $cmbCleanupKeep.SelectedItem = [string]$bk.CleanupKeep
    if (-not $cmbCleanupKeep.SelectedItem) { $cmbCleanupKeep.SelectedItem = "letzte drei Backups behalten" }
    $groupCleanup.Controls.Add($cmbCleanupKeep)

    $lblCleanupHint = New-Object System.Windows.Forms.Label
    $lblCleanupHint.Text = "Bereinigung erfolgt NACH erfolgreichem Backup. Betrifft nur Dateien mit`nPrefix 'BloodTracker_AutoBackup_'. Fremde Dateien bleiben unberührt."
    $lblCleanupHint.Location = New-Object System.Drawing.Point(15, 55)
    $lblCleanupHint.Size = New-Object System.Drawing.Size(500, 40)
    $lblCleanupHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblCleanupHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupCleanup.Controls.Add($lblCleanupHint)

    # Toggle: Dropdown nur bei aktivierter Checkbox sichtbar
    $updateCleanupVisibility = {
        $on = $chkCleanup.Checked
        $lblCleanupRule.Visible = $on
        $cmbCleanupKeep.Visible = $on
    }
    $chkCleanup.Add_CheckedChanged($updateCleanupVisibility)
    & $updateCleanupVisibility

    # Checkbox-Intelligenz (SuspendEvent-Pattern, verhindert Kaskaden)
    $script:abSuspend = $false
    $syncFromIndividual = {
        if ($script:abSuspend) { return }
        $script:abSuspend = $true
        try {
            $bothChecked = $chkZip.Checked -and $chkFolder.Checked
            $chkBoth.Checked = $bothChecked
            # Wenn beide aktiv: Einzeln ausgrauen
            $chkZip.Enabled    = -not $bothChecked
            $chkFolder.Enabled = -not $bothChecked
        } finally { $script:abSuspend = $false }
    }
    $syncFromBoth = {
        if ($script:abSuspend) { return }
        $script:abSuspend = $true
        try {
            if ($chkBoth.Checked) {
                $chkZip.Checked = $true
                $chkFolder.Checked = $true
                $chkZip.Enabled = $false
                $chkFolder.Enabled = $false
            } else {
                # "Beides" deaktivieren → alle Checkboxen zurücksetzen (laut Anforderung)
                $chkZip.Checked = $false
                $chkFolder.Checked = $false
                $chkZip.Enabled = $true
                $chkFolder.Enabled = $true
            }
        } finally { $script:abSuspend = $false }
    }
    $chkZip.Add_CheckedChanged($syncFromIndividual)
    $chkFolder.Add_CheckedChanged($syncFromIndividual)
    $chkBoth.Add_CheckedChanged($syncFromBoth)
    # Initial-Sync für ausgegrauten Zustand bei "Beides"
    & $syncFromIndividual

    # --- v2.19.0: GroupBox Portables Backup (Nachtrag für Template-Script) ---
    $groupPortable = New-Object System.Windows.Forms.GroupBox
    $groupPortable.Text = "Portables verschlüsseltes Backup (.btbackup)"
    $groupPortable.Location = New-Object System.Drawing.Point(15, 365)
    $groupPortable.Size = New-Object System.Drawing.Size(555, 115)
    $popup.Controls.Add($groupPortable)

    $chkPortable = New-Object System.Windows.Forms.CheckBox
    $chkPortable.Text = "Bei jedem AutoBackup zusätzlich portables .btbackup erstellen"
    $chkPortable.Location = New-Object System.Drawing.Point(15, 25)
    $chkPortable.Size = New-Object System.Drawing.Size(520, 22)
    $chkPortable.Checked = [bool]$bk.PortableBackup
    $groupPortable.Controls.Add($chkPortable)

    $btnSetPassword = New-Object System.Windows.Forms.Button
    $btnSetPassword.Text = "Backup-Passwort festlegen / ändern"
    $btnSetPassword.Location = New-Object System.Drawing.Point(15, 55)
    $btnSetPassword.Size = New-Object System.Drawing.Size(260, 28)
    $groupPortable.Controls.Add($btnSetPassword)

    $lblPasswordStatus = New-Object System.Windows.Forms.Label
    $lblPasswordStatus.Location = New-Object System.Drawing.Point(285, 60)
    $lblPasswordStatus.Size = New-Object System.Drawing.Size(260, 20)
    $lblPasswordStatus.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    if ($bk.PortableBackupPassword) {
        $lblPasswordStatus.Text = "✔ Passwort ist gesetzt"
        $lblPasswordStatus.ForeColor = [System.Drawing.Color]::DarkGreen
    } else {
        $lblPasswordStatus.Text = "✘ Kein Passwort gesetzt"
        $lblPasswordStatus.ForeColor = [System.Drawing.Color]::DarkRed
    }
    $groupPortable.Controls.Add($lblPasswordStatus)

    $lblPortableHint = New-Object System.Windows.Forms.Label
    $lblPortableHint.Text = "Das Passwort wird DPAPI-geschützt gespeichert (lokal automatisch).`nFür Import auf neuem PC benötigen Sie das Passwort."
    $lblPortableHint.Location = New-Object System.Drawing.Point(15, 85)
    $lblPortableHint.Size = New-Object System.Drawing.Size(530, 30)
    $lblPortableHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblPortableHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupPortable.Controls.Add($lblPortableHint)

    $btnSetPassword.Add_Click({
        $pp = Show-PassphraseDialog -Title "Backup-Passwort festlegen" -Confirm $true
        if ($pp) {
            # Passwort DPAPI-verschlüsselt in Config speichern
            $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($pp)
            $plainPass = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            $passBytes = [System.Text.Encoding]::UTF8.GetBytes($plainPass)
            $plainPass = $null
            $encrypted = [System.Security.Cryptography.ProtectedData]::Protect(
                $passBytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser
            )
            $script:data.Config.AutoBackup['PortableBackupPassword'] = [Convert]::ToBase64String($encrypted)
            Save-Data -data $script:data -type "Config"
            $lblPasswordStatus.Text = "✔ Passwort ist gesetzt"
            $lblPasswordStatus.ForeColor = [System.Drawing.Color]::DarkGreen
            [System.Windows.Forms.MessageBox]::Show("Backup-Passwort wurde gespeichert.`n`nBitte merken Sie sich dieses Passwort – es wird für den Import auf einem neuen PC benötigt.", "Passwort gesetzt", "OK", "Information")
        }
    })

    # --- GroupBox 2: Zeit-Einstellungen ---
    $groupTime = New-Object System.Windows.Forms.GroupBox
    $groupTime.Text = "Zeit-Einstellungen"
    $groupTime.Location = New-Object System.Drawing.Point(15, 490)
    $groupTime.Size = New-Object System.Drawing.Size(555, 345)
    $popup.Controls.Add($groupTime)

    $lblInterval = New-Object System.Windows.Forms.Label
    $lblInterval.Text = "Zeitraum:"
    $lblInterval.Location = New-Object System.Drawing.Point(15, 35)
    $lblInterval.Size = New-Object System.Drawing.Size(100, 20)
    $groupTime.Controls.Add($lblInterval)

    $cmbInterval = New-Object System.Windows.Forms.ComboBox
    $cmbInterval.Location = New-Object System.Drawing.Point(120, 32)
    $cmbInterval.Size = New-Object System.Drawing.Size(220, 22)
    $cmbInterval.DropDownStyle = "DropDownList"
    $cmbInterval.Items.AddRange(@("täglich","wöchentlich","monatlich","bei Programm-Start"))
    $cmbInterval.SelectedItem = [string]$bk.Interval
    if (-not $cmbInterval.SelectedItem) { $cmbInterval.SelectedItem = "wöchentlich" }
    $groupTime.Controls.Add($cmbInterval)

    $lblTime = New-Object System.Windows.Forms.Label
    $lblTime.Text = "Uhrzeit:"
    $lblTime.Location = New-Object System.Drawing.Point(15, 75)
    $lblTime.Size = New-Object System.Drawing.Size(100, 20)
    $groupTime.Controls.Add($lblTime)

    $dtTime = New-Object System.Windows.Forms.DateTimePicker
    $dtTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Time
    $dtTime.ShowUpDown = $true
    $dtTime.Location = New-Object System.Drawing.Point(120, 72)
    $dtTime.Size = New-Object System.Drawing.Size(120, 22)
    try {
        $parsedTime = [datetime]::ParseExact([string]$bk.TimeOfDay, "HH:mm", $null)
        $dtTime.Value = (Get-Date).Date.Add($parsedTime.TimeOfDay)
    } catch { $dtTime.Value = (Get-Date).Date.AddHours(3) }
    $groupTime.Controls.Add($dtTime)

    $lblTimeHint = New-Object System.Windows.Forms.Label
    $lblTimeHint.Text = ""
    $lblTimeHint.Location = New-Object System.Drawing.Point(250, 76)
    $lblTimeHint.Size = New-Object System.Drawing.Size(290, 20)
    $lblTimeHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblTimeHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupTime.Controls.Add($lblTimeHint)

    # Sub-GroupBox "Wochentage" (v2.14.1, Höhe erhöht v2.15.3 für 2-Zeilen-Layout) - nur sichtbar bei "wöchentlich"
    $groupWeekdays = New-Object System.Windows.Forms.GroupBox
    $groupWeekdays.Text = "Wochentage"
    $groupWeekdays.Location = New-Object System.Drawing.Point(15, 105)
    $groupWeekdays.Size = New-Object System.Drawing.Size(525, 170)
    $groupTime.Controls.Add($groupWeekdays)

    $weekdayKeys = @("Mo","Di","Mi","Do","Fr","Sa","So")
    $weekdayLabels = @{
        "Mo"="Montag"; "Di"="Dienstag"; "Mi"="Mittwoch"; "Do"="Donnerstag";
        "Fr"="Freitag"; "Sa"="Samstag"; "So"="Sonntag"
    }
    $savedDays = @()
    if ($bk.Weekdays) { $savedDays = @($bk.Weekdays) }
    $wdCheckboxes = @{}

    # v2.15.3: Zwei-Zeilen-Layout. "Donnerstag" braucht 90px (sonst 2-zeilig umgebrochen).
    # Zeile 1 (Y=25): Mo, Di, Mi, Do (breiter), Fr
    # Zeile 2 (Y=55): Sa, So
    $wdSpec = @(
        @{ Key="Mo"; X=15;  Y=25; W=72 }
        @{ Key="Di"; X=87;  Y=25; W=72 }
        @{ Key="Mi"; X=159; Y=25; W=72 }
        @{ Key="Do"; X=231; Y=25; W=90 }
        @{ Key="Fr"; X=321; Y=25; W=72 }
        @{ Key="Sa"; X=15;  Y=55; W=72 }
        @{ Key="So"; X=87;  Y=55; W=72 }
    )
    foreach ($spec in $wdSpec) {
        $k = $spec.Key
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $weekdayLabels[$k]
        $cb.Location = New-Object System.Drawing.Point($spec.X, $spec.Y)
        $cb.Size = New-Object System.Drawing.Size($spec.W, 22)
        $cb.Checked = ($savedDays -contains $k)
        $cb.Tag = $k
        $groupWeekdays.Controls.Add($cb)
        $wdCheckboxes[$k] = $cb
    }

    $btnWdAll = New-Object System.Windows.Forms.Button
    $btnWdAll.Text = "Alle auswählen"
    $btnWdAll.Location = New-Object System.Drawing.Point(15, 85)
    $btnWdAll.Size = New-Object System.Drawing.Size(110, 24)
    $groupWeekdays.Controls.Add($btnWdAll)
    $btnWdAll.Add_Click({
        foreach ($k in $weekdayKeys) { $wdCheckboxes[$k].Checked = $true }
    }.GetNewClosure())

    $btnWdNone = New-Object System.Windows.Forms.Button
    $btnWdNone.Text = "Alle abwählen"
    $btnWdNone.Location = New-Object System.Drawing.Point(135, 85)
    $btnWdNone.Size = New-Object System.Drawing.Size(110, 24)
    $groupWeekdays.Controls.Add($btnWdNone)
    $btnWdNone.Add_Click({
        foreach ($k in $weekdayKeys) { $wdCheckboxes[$k].Checked = $false }
    }.GetNewClosure())

    # v2.15.3: Hinweistext eine Zeile + Leerzeile unter den Alle-Buttons (Y=115).
    $lblWdHint = New-Object System.Windows.Forms.Label
    $lblWdHint.Text = "Backup erfolgt an den ausgewählten Tagen, sofern die Uhrzeit`nerreicht ist und heute noch kein Backup durchgeführt wurde."
    $lblWdHint.Location = New-Object System.Drawing.Point(15, 135)
    $lblWdHint.Size = New-Object System.Drawing.Size(500, 26)
    $lblWdHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblWdHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupWeekdays.Controls.Add($lblWdHint)

    # Sub-GroupBox "Monatstag" (v2.14.2, Höhe v2.15.3) - nur sichtbar bei "monatlich"
    # Gleicher Y-Bereich wie Wochentage-Box (sie schließen sich gegenseitig aus)
    $groupMonthDay = New-Object System.Windows.Forms.GroupBox
    $groupMonthDay.Text = "Monatstag"
    $groupMonthDay.Location = New-Object System.Drawing.Point(15, 105)
    $groupMonthDay.Size = New-Object System.Drawing.Size(525, 170)
    $groupTime.Controls.Add($groupMonthDay)

    $lblMonthDay = New-Object System.Windows.Forms.Label
    $lblMonthDay.Text = "Durchführung am:"
    $lblMonthDay.Location = New-Object System.Drawing.Point(15, 28)
    $lblMonthDay.Size = New-Object System.Drawing.Size(130, 20)
    $groupMonthDay.Controls.Add($lblMonthDay)

    $cmbMonthDay = New-Object System.Windows.Forms.ComboBox
    $cmbMonthDay.Location = New-Object System.Drawing.Point(150, 25)
    $cmbMonthDay.Size = New-Object System.Drawing.Size(260, 22)
    $cmbMonthDay.DropDownStyle = "DropDownList"
    $cmbMonthDay.Items.AddRange(@(
        "am Ersten eines Monats",
        "am zweiten Tag",
        "am dritten Tag",
        "am vierten Tag",
        "am fünften Tag",
        "zur Monatsmitte (15.)",
        "am Letzten Tag eines Monats"
    ))
    $cmbMonthDay.SelectedItem = [string]$bk.MonthDay
    if (-not $cmbMonthDay.SelectedItem) { $cmbMonthDay.SelectedItem = "am Ersten eines Monats" }
    $groupMonthDay.Controls.Add($cmbMonthDay)

    $lblMdHint = New-Object System.Windows.Forms.Label
    $lblMdHint.Text = "Backup erfolgt am gewählten Tag des Monats, sofern Uhrzeit erreicht ist."
    $lblMdHint.Location = New-Object System.Drawing.Point(15, 55)
    $lblMdHint.Size = New-Object System.Drawing.Size(500, 20)
    $lblMdHint.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblMdHint.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $groupMonthDay.Controls.Add($lblMdHint)

    # Toggle: Wochentage-/Monatstag-Sichtbarkeit & Uhrzeit-Enable basierend auf Interval
    $updateTimeEnable = {
        $isStart   = ($cmbInterval.SelectedItem -eq "bei Programm-Start")
        $isWeekly  = ($cmbInterval.SelectedItem -eq "wöchentlich")
        $isMonthly = ($cmbInterval.SelectedItem -eq "monatlich")
        $dtTime.Enabled = -not $isStart
        $lblTime.Enabled = -not $isStart
        if ($isStart) {
            $lblTimeHint.Text = "(nicht relevant: Backup beim Programm-Start)"
        } else {
            $lblTimeHint.Text = "Backup erfolgt nicht vor dieser Uhrzeit."
        }
        $groupWeekdays.Visible = $isWeekly
        $groupMonthDay.Visible = $isMonthly
    }
    $cmbInterval.Add_SelectedIndexChanged($updateTimeEnable)
    & $updateTimeEnable

    $lblScheduleInfo = New-Object System.Windows.Forms.Label
    $lblScheduleInfo.Text = "Hinweis: Es wird KEIN Windows-Aufgabenplanungs-Task erstellt. Backups erfolgen nur während das Programm läuft (Start-Check + Scheduler alle 15 Min.). Verpasste Intervalle werden beim nächsten Programmstart nachgeholt."
    $lblScheduleInfo.Location = New-Object System.Drawing.Point(15, 285)
    $lblScheduleInfo.Size = New-Object System.Drawing.Size(530, 55)
    $lblScheduleInfo.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $lblScheduleInfo.ForeColor = [System.Drawing.Color]::DarkRed
    $groupTime.Controls.Add($lblScheduleInfo)

    # --- Jetzt-Backup-Button ---
    $btnBackupNow = New-Object System.Windows.Forms.Button
    $btnBackupNow.Text = "Jetzt sofort Backup ausführen"
    $btnBackupNow.Location = New-Object System.Drawing.Point(15, 870)
    $btnBackupNow.Size = New-Object System.Drawing.Size(220, 32)
    $popup.Controls.Add($btnBackupNow)
    $btnBackupNow.Add_Click({
        # Temporär die aktuellen Dialog-Werte in Config übernehmen, damit Invoke-AutoBackupNow sie nutzt
        $err = Test-BackupPathSafe -TargetPath $txtTarget.Text -DataDir $script:dataDir
        if ($err) {
            [System.Windows.Forms.MessageBox]::Show($err, "Ungültiges Zielverzeichnis", "OK", "Warning")
            return
        }
        if (-not $chkZip.Checked -and -not $chkFolder.Checked) {
            [System.Windows.Forms.MessageBox]::Show("Bitte mindestens ein Format (ZIP oder Ordner) auswählen.", "Format fehlt", "OK", "Warning")
            return
        }
        $prev = @{
            Enabled       = $script:data.Config.AutoBackup.Enabled
            TargetPath    = $script:data.Config.AutoBackup.TargetPath
            FormatZip     = $script:data.Config.AutoBackup.FormatZip
            FormatFolder  = $script:data.Config.AutoBackup.FormatFolder
            IncludeRegistry = $script:data.Config.AutoBackup.IncludeRegistry
        }
        $script:data.Config.AutoBackup.Enabled         = $true
        $script:data.Config.AutoBackup.TargetPath      = $txtTarget.Text
        $script:data.Config.AutoBackup.FormatZip       = $chkZip.Checked
        $script:data.Config.AutoBackup.FormatFolder    = $chkFolder.Checked
        $script:data.Config.AutoBackup.IncludeRegistry = $chkRegistry.Checked
        $ok = Invoke-AutoBackupNow
        # Enabled ggf. zurückstellen falls vorher nicht aktiv
        if (-not $prev.Enabled) { $script:data.Config.AutoBackup.Enabled = $prev.Enabled }
        if ($ok) {
            [System.Windows.Forms.MessageBox]::Show("Backup erfolgreich erstellt in:`n$($txtTarget.Text)", "Backup", "OK", "Information")
            # Status-Label refreshen
            $now = Get-Date
            $lblStatus.Text = "Letztes erfolgreiches Backup: $($now.ToString('dd.MM.yyyy HH:mm'))"
            $lblStatus.ForeColor = [System.Drawing.Color]::DarkGreen
        } else {
            [System.Windows.Forms.MessageBox]::Show("Backup fehlgeschlagen. Details siehe PowerShell-Warnings.", "Backup-Fehler", "OK", "Error")
        }
    })

    # --- Abbrechen / Speichern ---
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Abbrechen"
    $btnCancel.Location = New-Object System.Drawing.Point(355, 870)
    $btnCancel.Size = New-Object System.Drawing.Size(105, 32)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $popup.Controls.Add($btnCancel)
    $popup.CancelButton = $btnCancel

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text = "Speichern"
    $btnSave.Location = New-Object System.Drawing.Point(465, 870)
    $btnSave.Size = New-Object System.Drawing.Size(105, 32)
    $popup.Controls.Add($btnSave)
    $popup.AcceptButton = $btnSave

    $btnSave.Add_Click({
        # Wochentage aus Checkboxen sammeln (v2.14.1)
        $selectedWeekdays = @()
        foreach ($k in $weekdayKeys) {
            if ($wdCheckboxes[$k].Checked) { $selectedWeekdays += $k }
        }

        # Validierung wenn Feature aktiviert
        if ($chkEnabled.Checked) {
            $err = Test-BackupPathSafe -TargetPath $txtTarget.Text -DataDir $script:dataDir
            if ($err) {
                [System.Windows.Forms.MessageBox]::Show($err, "Ungültiges Zielverzeichnis", "OK", "Warning")
                return
            }
            if (-not $chkZip.Checked -and -not $chkFolder.Checked) {
                [System.Windows.Forms.MessageBox]::Show("Bitte mindestens ein Format (ZIP oder Ordner) auswählen.", "Format fehlt", "OK", "Warning")
                return
            }
            if ($cmbInterval.SelectedItem -eq "wöchentlich" -and $selectedWeekdays.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Bitte mindestens einen Wochentag auswählen, an dem das wöchentliche Backup erfolgen soll.", "Wochentag fehlt", "OK", "Warning")
                return
            }
        }
        $script:data.Config.AutoBackup.Enabled         = $chkEnabled.Checked
        $script:data.Config.AutoBackup.TargetPath      = $txtTarget.Text
        $script:data.Config.AutoBackup.FormatZip       = $chkZip.Checked
        $script:data.Config.AutoBackup.FormatFolder    = $chkFolder.Checked
        $script:data.Config.AutoBackup.IncludeRegistry = $chkRegistry.Checked
        $script:data.Config.AutoBackup.Interval        = [string]$cmbInterval.SelectedItem
        $script:data.Config.AutoBackup.TimeOfDay       = $dtTime.Value.ToString("HH:mm")
        $script:data.Config.AutoBackup.Weekdays        = $selectedWeekdays
        $script:data.Config.AutoBackup.MonthDay        = [string]$cmbMonthDay.SelectedItem
        $script:data.Config.AutoBackup.CleanupEnabled  = $chkCleanup.Checked
        $script:data.Config.AutoBackup.CleanupKeep     = [string]$cmbCleanupKeep.SelectedItem
        # v2.19.0: Portables Backup
        $script:data.Config.AutoBackup.PortableBackup  = $chkPortable.Checked
        Save-Data -data $script:data -type "Config"
        [System.Windows.Forms.MessageBox]::Show("Backup-Einstellungen gespeichert.", "Gespeichert", "OK", "Information")
        $popup.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $popup.Close()
    })

    $popup.ShowDialog() | Out-Null
}

function Show-SettingsPopup {
    param($mainForm)
    $popupForm = New-Object System.Windows.Forms.Form; $popupForm.Size = New-Object System.Drawing.Size(450, 500); $popupForm.Text = "Einstellungen & Datenverwaltung"; $popupForm.StartPosition = "CenterParent"; $popupForm.FormBorderStyle = "FixedDialog"; $popupForm.MaximizeBox = $false; $popupForm.MinimizeBox = $false
    $label = New-Object System.Windows.Forms.Label; $label.Text = "Verwalten Sie hier Ihre Konfigurationsdateien und Exporte."; $label.Location = New-Object System.Drawing.Point(20, 20); $label.Size = New-Object System.Drawing.Size(400, 20); $popupForm.Controls.Add($label)
    $openConfigFileButton = New-Object System.Windows.Forms.Button; $openConfigFileButton.Text = "Konfigurationsdatei öffnen (.json)"; $openConfigFileButton.Location = New-Object System.Drawing.Point(20, 60); $openConfigFileButton.Size = New-Object System.Drawing.Size(400, 35); $popupForm.Controls.Add($openConfigFileButton)
    $openDataPathButton = New-Object System.Windows.Forms.Button; $openDataPathButton.Text = "Daten-Verzeichnis im Explorer öffnen"; $openDataPathButton.Location = New-Object System.Drawing.Point(20, 105); $openDataPathButton.Size = New-Object System.Drawing.Size(400, 35); $popupForm.Controls.Add($openDataPathButton)
    $exportDataButton = New-Object System.Windows.Forms.Button; $exportDataButton.Text = "Benutzerdaten als ZIP-Archiv exportieren..."; $exportDataButton.Location = New-Object System.Drawing.Point(20, 150); $exportDataButton.Size = New-Object System.Drawing.Size(400, 35); $popupForm.Controls.Add($exportDataButton)
    $importDataButton = New-Object System.Windows.Forms.Button; $importDataButton.Text = "Labordaten aus CSV importieren..."; $importDataButton.Location = New-Object System.Drawing.Point(20, 195); $importDataButton.Size = New-Object System.Drawing.Size(400, 35); $popupForm.Controls.Add($importDataButton)
    $moveDataButton = New-Object System.Windows.Forms.Button; $moveDataButton.Text = "Datenpfad ändern..."; $moveDataButton.Location = New-Object System.Drawing.Point(20, 240); $moveDataButton.Size = New-Object System.Drawing.Size(400, 35); $popupForm.Controls.Add($moveDataButton)
    $autoBackupButton = New-Object System.Windows.Forms.Button; $autoBackupButton.Text = "Automatisches Backup..."; $autoBackupButton.Location = New-Object System.Drawing.Point(20, 285); $autoBackupButton.Size = New-Object System.Drawing.Size(400, 35); $autoBackupButton.ForeColor = [System.Drawing.Color]::DarkBlue; $popupForm.Controls.Add($autoBackupButton)
    $restoreBackupButton = New-Object System.Windows.Forms.Button; $restoreBackupButton.Text = "Sicherung wiederherstellen (ZIP)..."; $restoreBackupButton.Location = New-Object System.Drawing.Point(20, 330); $restoreBackupButton.Size = New-Object System.Drawing.Size(400, 35); $restoreBackupButton.ForeColor = [System.Drawing.Color]::DarkGreen; $popupForm.Controls.Add($restoreBackupButton)
    $uninstallButton = New-Object System.Windows.Forms.Button; $uninstallButton.Text = "Deinstallieren"; $uninstallButton.Location = New-Object System.Drawing.Point(20, 375); $uninstallButton.Size = New-Object System.Drawing.Size(400, 35); $uninstallButton.ForeColor = [System.Drawing.Color]::DarkRed; $popupForm.Controls.Add($uninstallButton)
    $closeButton = New-Object System.Windows.Forms.Button; $closeButton.Text = "Schließen"; $closeButton.Location = New-Object System.Drawing.Point(310, 425); $closeButton.Size = New-Object System.Drawing.Size(100, 30); $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::OK; $popupForm.Controls.Add($closeButton); $popupForm.AcceptButton = $closeButton
    # v2.20.0: Checkbox "Hilfstexte dauerhaft anzeigen" links neben Schließen
    $chkPersistentTooltips = New-Object System.Windows.Forms.CheckBox
    $chkPersistentTooltips.Text = "Hilfstexte dauerhaft anzeigen"
    $chkPersistentTooltips.Location = New-Object System.Drawing.Point(20, 430)
    $chkPersistentTooltips.Size = New-Object System.Drawing.Size(260, 20)
    $chkPersistentTooltips.Checked = [bool]$script:data.Config.Personal.PersistentTooltips
    $popupForm.Controls.Add($chkPersistentTooltips)
    $chkPersistentTooltips.Add_CheckedChanged({
        $script:data.Config.Personal.PersistentTooltips = $chkPersistentTooltips.Checked
        Save-Data -data $script:data -type "Config"
    })
    $autoBackupButton.Add_Click({ Show-AutoBackupPopup })
    $openConfigFileButton.Add_Click({ if (Test-Path $script:dataFile) { try { Invoke-Item $script:dataFile } catch { [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)") } } else { [System.Windows.Forms.MessageBox]::Show("Konfigurationsdatei nicht gefunden.") } })
    $openDataPathButton.Add_Click({ if (Test-Path $script:dataDir) { try { Invoke-Item $script:dataDir } catch { [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)") } } else { [System.Windows.Forms.MessageBox]::Show("Datenverzeichnis nicht gefunden.") } })
    $exportDataButton.Add_Click({ $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog; $saveFileDialog.Filter = "ZIP-Archiv (*.zip)|*.zip"; $saveFileDialog.Title = "Benutzerdaten exportieren"; $saveFileDialog.FileName = "BlutwertTracker_Export_$(Get-Date -Format 'yyyy-MM-dd').zip"; if ($saveFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $exportPath = $saveFileDialog.FileName; try { Compress-Archive -Path "$($script:dataDir)\*" -DestinationPath $exportPath -Force; [System.Windows.Forms.MessageBox]::Show("Daten erfolgreich nach `"$exportPath`" exportiert.", "Export abgeschlossen", "OK", "Information") } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Exportieren: $($_.Exception.Message)") } } })
    $importDataButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog; $openFileDialog.Filter = "CSV-Datei (*.csv)|*.csv"; $openFileDialog.Title = "CSV-Datei für Import auswählen"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Import-CsvData -filePath $openFileDialog.FileName
            $script:allHistoricalItems = [System.Collections.ArrayList]@(Load-AllHistoricalItems)
            Update-FilterDropdown
        }
    })
    $moveDataButton.Add_Click({ $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog; $folderBrowser.Description = "Wählen Sie ein neues, leeres Verzeichnis für die Anwendungsdaten."; if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $newPath = $folderBrowser.SelectedPath; if ($newPath -eq $script:dataDir) { [System.Windows.Forms.MessageBox]::Show("Der ausgewählte Pfad ist bereits der aktuelle Speicherort.", "Information", "OK", "Information"); return }; if (Get-ChildItem -Path $newPath) { [System.Windows.Forms.MessageBox]::Show("Das ausgewählte Verzeichnis ist nicht leer. Bitte wählen Sie einen leeren Ordner aus, um Datenkonflikte zu vermeiden.", "Fehler", "OK", "Error"); return }; try { Move-Item -Path "$($script:dataDir)\*" -Destination $newPath -Force; Remove-Item -Path $script:dataDir -Recurse -Force; Set-ItemProperty -Path $script:regPath -Name "DataPath" -Value $newPath; [System.Windows.Forms.MessageBox]::Show("Der Datenpfad wurde erfolgreich nach `"$newPath`" geändert. Das Programm wird nun neu gestartet.", "Erfolg", "OK", "Information"); $scriptPath = $MyInvocation.MyCommand.Path; Start-Process powershell -ArgumentList "-File `"$scriptPath`""; $popupForm.Close(); $mainForm.Close() } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Verschieben der Daten: $($_.Exception.Message)", "Fehler", "OK", "Error") } } })
    $restoreBackupButton.Add_Click({
        # Schritt 1: ZIP-Datei auswählen
        $openDlg = New-Object System.Windows.Forms.OpenFileDialog
        $openDlg.Filter = "ZIP-Archiv (*.zip)|*.zip"
        $openDlg.Title = "Backup-ZIP auswählen..."
        if ($openDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $zipPath = $openDlg.FileName

        try {
            # Schritt 2: ZIP in temporäres Verzeichnis entpacken und validieren
            $tempDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "PSC_BloodTracker_Restore_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
            New-Item -Path $tempDir -ItemType Directory -Force | Out-Null
            Expand-Archive -Path $zipPath -DestinationPath $tempDir -Force

            # Prüfe ob UserData-Ordner vorhanden ist (Pflicht-Inhalt)
            $userDataSource = Join-Path -Path $tempDir -ChildPath "UserData"
            # Unterstütze auch verschachtelte Struktur (UserData\PSC.Blood-Tracker\...)
            $hasUserData = Test-Path $userDataSource
            $hasRegBackup = (Test-Path (Join-Path $tempDir "Registry_Backup.reg")) -or (Test-Path (Join-Path $tempDir "Registry_Backup.txt"))

            if (-not $hasUserData) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Die gewählte ZIP-Datei enthält keinen 'UserData'-Ordner und ist kein gültiges PSC.Blood-Tracker Backup.",
                    "Ungültiges Backup", "OK", "Error")
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                return
            }

            # Schritt 3: Bestehende Daten prüfen und Merge/Überschreiben fragen
            $existingDataCount = 0
            if (Test-Path $script:dataDir) {
                $existingDataCount = @(Get-ChildItem -Path $script:dataDir -Recurse -File -ErrorAction SilentlyContinue).Count
            }

            $restoreMode = "overwrite"
            if ($existingDataCount -gt 0) {
                $choice = [System.Windows.Forms.MessageBox]::Show(
                    "Es sind bereits $existingDataCount Dateien im aktuellen Datenverzeichnis vorhanden.`n`nMöchten Sie die bestehenden Daten vorher löschen und komplett durch das Backup ersetzen?`n`n[Ja] = Alles löschen und Backup wiederherstellen`n[Nein] = Backup-Daten ergänzen (bestehende Dateien werden überschrieben)`n[Abbrechen] = Nichts tun",
                    "Bestehende Daten gefunden", "YesNoCancel", "Question")
                if ($choice -eq "Cancel") {
                    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return
                }
                if ($choice -eq "Yes") { $restoreMode = "clean" }
                else { $restoreMode = "merge" }
            }

            # Schritt 4: Daten wiederherstellen
            # 4a: Bei clean-Modus bestehende Daten löschen
            if ($restoreMode -eq "clean" -and (Test-Path $script:dataDir)) {
                Remove-Item -Path "$($script:dataDir)\*" -Recurse -Force
            }

            # 4b: Datenverzeichnis sicherstellen
            if (-not (Test-Path $script:dataDir)) {
                New-Item -Path $script:dataDir -ItemType Directory -Force | Out-Null
            }

            # 4c: UserData-Inhalt kopieren
            # Prüfe ob der UserData-Ordner einen Unterordner mit App-Name enthält (Backup-Struktur: UserData\PSC.Blood-Tracker\...)
            $innerAppDir = Join-Path -Path $userDataSource -ChildPath "PSC.Blood-Tracker"
            if (Test-Path $innerAppDir) {
                # Verschachtelte Struktur → Inhalt des inneren Ordners kopieren
                Copy-Item -Path "$innerAppDir\*" -Destination $script:dataDir -Recurse -Force
            } else {
                # Flache Struktur → Inhalt direkt kopieren
                Copy-Item -Path "$userDataSource\*" -Destination $script:dataDir -Recurse -Force
            }

            # 4d: Registry wiederherstellen (falls vorhanden)
            $regRestored = $false
            $regFile = Join-Path -Path $tempDir -ChildPath "Registry_Backup.reg"
            if (Test-Path $regFile) {
                $regImportResult = & reg import $regFile 2>&1
                $regRestored = $true
            } else {
                # Fallback: .txt-Datei → DataPath aus Backup nicht automatisch übernehmbar,
                # aber Registry-Pfad wird beim nächsten Start ohnehin korrekt initialisiert
                $regRestored = $false
            }

            # Registry: DataPath auf aktuellen Pfad setzen (überschreibt ggf. alten Pfad aus Backup)
            if (-not (Test-Path $script:regPath)) { New-Item -Path $script:regPath -Force | Out-Null }
            Set-ItemProperty -Path $script:regPath -Name "DataPath" -Value $script:dataDir

            # Schritt 5: Aufräumen
            Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue

            # Schritt 6: Zusammenfassung und Neustart
            $summary = "Wiederherstellung erfolgreich abgeschlossen!`n`n"
            $summary += "- Benutzerdaten: Wiederhergestellt ($restoreMode)`n"
            $summary += "- Registry: $(if ($regRestored) { 'Importiert' } else { 'Übersprungen (wird beim Start initialisiert)' })`n"
            $summary += "`nDas Programm wird nun neu gestartet."

            [System.Windows.Forms.MessageBox]::Show($summary, "Wiederherstellung abgeschlossen", "OK", "Information")

            # Neustart
            $scriptPath = $PSCommandPath
            if (-not $scriptPath) { $scriptPath = $MyInvocation.ScriptName }
            if ($scriptPath -and (Test-Path $scriptPath)) {
                Start-Process powershell -ArgumentList "-File `"$scriptPath`""
            }
            $popupForm.Close()
            $mainForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei der Wiederherstellung: $($_.Exception.Message)",
                "Fehler", "OK", "Error")
            # Temp-Verzeichnis aufräumen falls vorhanden
            if ($tempDir -and (Test-Path $tempDir)) {
                Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    })
    $uninstallButton.Add_Click({
        # Schritt 1: Bestätigung
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie den PSC.Blood-Tracker wirklich vollständig deinstallieren?`n`nFolgende Daten werden unwiderruflich entfernt:`n- Alle gespeicherten Blutwert-Daten`n- Registry-Einträge`n- Das Script selbst`n`nDieser Vorgang kann nicht rückgängig gemacht werden!",
            "Deinstallation bestätigen", "YesNo", "Warning")
        if ($confirm -ne "Yes") { return }

        # Schritt 2: Sicherung anbieten
        $backup = [System.Windows.Forms.MessageBox]::Show(
            "Möchten Sie vorher eine vollständige Sicherung aller Daten erstellen?`n`n(Enthält: Benutzerdaten, Registry-Einträge und das Script)",
            "Sicherung erstellen?", "YesNoCancel", "Question")
        if ($backup -eq "Cancel") { return }

        $scriptPath = $PSCommandPath
        if (-not $scriptPath) { $scriptPath = $MyInvocation.ScriptName }
        if (-not $scriptPath) { $scriptPath = & { $MyInvocation.ScriptName } }

        if ($backup -eq "Yes") {
            # Schritt 3: Zielpfad auswählen und ZIP erstellen
            $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
            $saveDlg.Filter = "ZIP-Archiv (*.zip)|*.zip"
            $saveDlg.Title = "Sicherung speichern unter..."
            $saveDlg.FileName = "PSC.Blood-Tracker_Backup_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').zip"
            if ($saveDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
            $backupZip = $saveDlg.FileName

            try {
                # Temporäres Staging-Verzeichnis
                $stagingDir = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "PSC_BloodTracker_Uninstall_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                New-Item -Path $stagingDir -ItemType Directory -Force | Out-Null

                # 3a: Benutzerdaten kopieren
                if (Test-Path $script:dataDir) {
                    $dataBackupDir = Join-Path -Path $stagingDir -ChildPath "UserData"
                    Copy-Item -Path $script:dataDir -Destination $dataBackupDir -Recurse -Force
                }

                # 3b: Registry-Schlüssel exportieren
                $regExportFile = Join-Path -Path $stagingDir -ChildPath "Registry_Backup.reg"
                $regKeyFull = "HKCU\Software\PSC\PSC.Blood-Tracker"
                $regExportResult = & reg export $regKeyFull $regExportFile /y 2>&1
                if (-not (Test-Path $regExportFile)) {
                    # Fallback: Manueller Export als Text
                    $regExportFile = Join-Path -Path $stagingDir -ChildPath "Registry_Backup.txt"
                    $regValues = Get-ItemProperty -Path $script:regPath -ErrorAction SilentlyContinue
                    if ($regValues) { $regValues | Out-String | Set-Content -Path $regExportFile -Encoding UTF8 }
                }

                # 3c: Script-Datei kopieren
                if ($scriptPath -and (Test-Path $scriptPath)) {
                    Copy-Item -Path $scriptPath -Destination $stagingDir -Force
                }

                # 3d: ZIP erstellen
                if (Test-Path $backupZip) { Remove-Item -Path $backupZip -Force }
                Compress-Archive -Path "$stagingDir\*" -DestinationPath $backupZip -Force

                # Staging aufräumen
                Remove-Item -Path $stagingDir -Recurse -Force -ErrorAction SilentlyContinue

                [System.Windows.Forms.MessageBox]::Show(
                    "Sicherung erfolgreich erstellt:`n$backupZip`n`nDie Deinstallation wird nun fortgesetzt.",
                    "Sicherung abgeschlossen", "OK", "Information")
            } catch {
                $errMsg = $_.Exception.Message
                $abortChoice = [System.Windows.Forms.MessageBox]::Show(
                    "Fehler bei der Sicherung: $errMsg`n`nMöchten Sie die Deinstallation trotzdem fortsetzen?",
                    "Sicherungsfehler", "YesNo", "Error")
                if ($abortChoice -ne "Yes") { return }
            }
        }

        # Schritt 4: Daten entfernen
        try {
            # 4a: Datenverzeichnis löschen
            if (Test-Path $script:dataDir) {
                Remove-Item -Path $script:dataDir -Recurse -Force
            }

            # 4b: Registry-Schlüssel entfernen (inkl. übergeordnetem PSC-Schlüssel wenn leer)
            if (Test-Path $script:regPath) {
                Remove-Item -Path $script:regPath -Recurse -Force
            }
            $parentRegPath = "HKCU:\Software\PSC"
            if ((Test-Path $parentRegPath) -and @(Get-ChildItem -Path $parentRegPath -ErrorAction SilentlyContinue).Count -eq 0) {
                Remove-Item -Path $parentRegPath -Force -ErrorAction SilentlyContinue
            }

            # 4c: Script-Datei löschen (verzögert per Hintergrundjob)
            if ($scriptPath -and (Test-Path $scriptPath)) {
                $deleteCmd = "Start-Sleep -Seconds 2; Remove-Item -Path '$($scriptPath -replace "'","''")' -Force -ErrorAction SilentlyContinue"
                Start-Process powershell -ArgumentList "-WindowStyle Hidden -Command $deleteCmd" -WindowStyle Hidden
            }

            [System.Windows.Forms.MessageBox]::Show(
                "PSC.Blood-Tracker wurde erfolgreich deinstalliert.`nAlle Daten und Einstellungen wurden entfernt.`n`nDas Programm wird jetzt beendet.",
                "Deinstallation abgeschlossen", "OK", "Information")

            # Programm beenden
            $popupForm.Close()
            $mainForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fehler bei der Deinstallation: $($_.Exception.Message)`n`nEinige Daten konnten möglicherweise nicht entfernt werden.",
                "Fehler", "OK", "Error")
        }
    })
    $popupForm.ShowDialog()
}

function Show-CockpitSettingsPopup {
    # Verfügbare Gruppen ermitteln (nur die, für die Daten existieren)
    $allMarkerConfigs = @($script:data.Config.Markers) + @($script:data.Config.CalculatedMarkers)
    $groupsWithData = New-Object System.Collections.ArrayList
    foreach ($marker in $allMarkerConfigs) {
        $hasData = $false
        $isCalc = $marker.PSObject.Properties.Name -contains "RequiredMarkers"
        if ($isCalc) {
            $calcData = Get-CalculatedValuesForMarker -markerName $marker.Name
            if ($calcData.Count -gt 0) { $hasData = $true }
        } else {
            $histData = @($script:allHistoricalItems | Where-Object { $_.Name -eq $marker.Name })
            if ($histData.Count -gt 0) { $hasData = $true }
        }
        if ($hasData -and $marker.Group -and -not $groupsWithData.Contains($marker.Group)) {
            $groupsWithData.Add($marker.Group) | Out-Null
        }
    }
    $groupsWithData.Sort()

    $bewertungBaseY = 165
    $checkboxAreaHeight = [Math]::Max(120, ($groupsWithData.Count * 24) + 20)
    $filterGroupBaseH = 290
    $collapsedHeight = 400
    $expandedHeight = $collapsedHeight + $checkboxAreaHeight

    $popupForm = New-Object System.Windows.Forms.Form
    $popupForm.Size = New-Object System.Drawing.Size(450, $collapsedHeight)
    $popupForm.Text = "Einstellungen"
    $popupForm.StartPosition = "CenterParent"
    $popupForm.FormBorderStyle = "FixedDialog"
    $popupForm.MaximizeBox = $false
    $popupForm.MinimizeBox = $false

    $filterGroup = New-Object System.Windows.Forms.GroupBox
    $filterGroup.Text = "Filter"
    $filterGroup.Location = New-Object System.Drawing.Point(20, 20)
    $filterGroup.Size = New-Object System.Drawing.Size(400, $filterGroupBaseH)
    $popupForm.Controls.Add($filterGroup)

    # --- Sub-GroupBox: Zeitraum ---
    $zeitraumGroup = New-Object System.Windows.Forms.GroupBox
    $zeitraumGroup.Text = "Zeitraum"
    $zeitraumGroup.Location = New-Object System.Drawing.Point(15, 25)
    $zeitraumGroup.Size = New-Object System.Drawing.Size(370, 60)
    $filterGroup.Controls.Add($zeitraumGroup)
    $zeitraumLabel = New-Object System.Windows.Forms.Label
    $zeitraumLabel.Text = "Zeitraum:"
    $zeitraumLabel.Location = New-Object System.Drawing.Point(15, 24)
    $zeitraumLabel.Size = New-Object System.Drawing.Size(70, 20)
    $zeitraumGroup.Controls.Add($zeitraumLabel)
    $zeitraumCombo = New-Object System.Windows.Forms.ComboBox
    $zeitraumCombo.Location = New-Object System.Drawing.Point(90, 22)
    $zeitraumCombo.Size = New-Object System.Drawing.Size(260, 20)
    $zeitraumCombo.DropDownStyle = "DropDownList"
    $zeitraumCombo.Items.AddRange(@("Aktuelles Jahr", "Letzte 3 Jahre", "Letzte 5 Jahre", "Letzte 10 Jahre", "Alle Daten"))
    $zeitraumCombo.SelectedItem = $script:cockpitTimeFilter
    $zeitraumGroup.Controls.Add($zeitraumCombo)
    $zeitraumCombo.Add_SelectedIndexChanged({ $script:cockpitTimeFilter = $zeitraumCombo.SelectedItem })

    # --- Sub-GroupBox: Funktionsbereich ---
    $fbGroup = New-Object System.Windows.Forms.GroupBox
    $fbGroup.Text = "Funktionsbereich"
    $fbGroup.Location = New-Object System.Drawing.Point(15, 95)
    $fbGroup.Size = New-Object System.Drawing.Size(370, 60)
    $filterGroup.Controls.Add($fbGroup)
    $fbLabel = New-Object System.Windows.Forms.Label
    $fbLabel.Text = "Funktionsbereich:"
    $fbLabel.Location = New-Object System.Drawing.Point(15, 24)
    $fbLabel.Size = New-Object System.Drawing.Size(115, 20)
    $fbGroup.Controls.Add($fbLabel)
    $fbCombo = New-Object System.Windows.Forms.ComboBox
    $fbCombo.Location = New-Object System.Drawing.Point(135, 22)
    $fbCombo.Size = New-Object System.Drawing.Size(215, 20)
    $fbCombo.DropDownStyle = "DropDownList"
    $fbCombo.Items.Add("Alle")
    foreach ($g in $groupsWithData) { $fbCombo.Items.Add($g) }
    $fbCombo.Items.Add("Custom")
    $fbGroup.Controls.Add($fbCombo)

    # --- Checkbox-Panel (initial unsichtbar, innerhalb fbGroup) ---
    $checkPanel = New-Object System.Windows.Forms.Panel
    $checkPanel.Location = New-Object System.Drawing.Point(15, 52)
    $checkPanel.Size = New-Object System.Drawing.Size(340, $checkboxAreaHeight)
    $checkPanel.AutoScroll = $true
    $checkPanel.Visible = $false
    $fbGroup.Controls.Add($checkPanel)

    $checkBoxes = New-Object System.Collections.ArrayList
    $yPos = 5
    foreach ($g in $groupsWithData) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $g
        $cb.Location = New-Object System.Drawing.Point(10, $yPos)
        $cb.Size = New-Object System.Drawing.Size(310, 20)
        $cb.Checked = ($script:cockpitGroupFilter -eq "Alle") -or ($script:cockpitGroupFilterCustom -contains $g)
        $checkPanel.Controls.Add($cb)
        $checkBoxes.Add($cb) | Out-Null
        $yPos += 24
    }

    # Aktuellen Funktionsbereich-Filter setzen
    if ($script:cockpitGroupFilter -eq "Custom") {
        $fbCombo.SelectedItem = "Custom"
    } elseif ($fbCombo.Items.Contains($script:cockpitGroupFilter)) {
        $fbCombo.SelectedItem = $script:cockpitGroupFilter
    } else {
        $fbCombo.SelectedItem = "Alle"
    }

    # --- Sub-GroupBox: Bewertung ---
    $bewertungGroup = New-Object System.Windows.Forms.GroupBox
    $bewertungGroup.Text = "Bewertung"
    $bewertungGroup.Location = New-Object System.Drawing.Point(15, $bewertungBaseY)
    $bewertungGroup.Size = New-Object System.Drawing.Size(370, 60)
    $filterGroup.Controls.Add($bewertungGroup)
    $bewertungLabel = New-Object System.Windows.Forms.Label
    $bewertungLabel.Text = "Bewertung:"
    $bewertungLabel.Location = New-Object System.Drawing.Point(15, 24)
    $bewertungLabel.Size = New-Object System.Drawing.Size(80, 20)
    $bewertungGroup.Controls.Add($bewertungLabel)
    $bewertungCombo = New-Object System.Windows.Forms.ComboBox
    $bewertungCombo.Location = New-Object System.Drawing.Point(100, 22)
    $bewertungCombo.Size = New-Object System.Drawing.Size(250, 20)
    $bewertungCombo.DropDownStyle = "DropDownList"
    $bewertungCombo.Items.AddRange(@("Alle", "Optimal", "Akzeptabel", "Außerhalb der Norm", "Keine Bewertung"))
    if ($bewertungCombo.Items.Contains($script:cockpitRatingFilter)) {
        $bewertungCombo.SelectedItem = $script:cockpitRatingFilter
    } else {
        $bewertungCombo.SelectedItem = "Alle"
    }
    $bewertungGroup.Controls.Add($bewertungCombo)
    $bewertungCombo.Add_SelectedIndexChanged({ $script:cockpitRatingFilter = $bewertungCombo.SelectedItem })

    # --- Hilfsfunktion: Layout aktualisieren ---
    $updateLayout = {
        param($isExpanded)
        if ($isExpanded) {
            $checkPanel.Visible = $true
            $fbGroup.Size = New-Object System.Drawing.Size(370, (60 + $checkboxAreaHeight))
            $bewertungGroup.Location = New-Object System.Drawing.Point(15, ($bewertungBaseY + $checkboxAreaHeight))
            $filterGroup.Size = New-Object System.Drawing.Size(400, ($filterGroupBaseH + $checkboxAreaHeight))
            $closeBtn.Location = New-Object System.Drawing.Point(20, ($filterGroupBaseH + $checkboxAreaHeight + 30))
            $resetBtn.Location = New-Object System.Drawing.Point(280, ($filterGroupBaseH + $checkboxAreaHeight + 30))
            $popupForm.Size = New-Object System.Drawing.Size(450, $expandedHeight)
        } else {
            $checkPanel.Visible = $false
            $fbGroup.Size = New-Object System.Drawing.Size(370, 60)
            $bewertungGroup.Location = New-Object System.Drawing.Point(15, $bewertungBaseY)
            $filterGroup.Size = New-Object System.Drawing.Size(400, $filterGroupBaseH)
            $closeBtn.Location = New-Object System.Drawing.Point(20, ($filterGroupBaseH + 30))
            $resetBtn.Location = New-Object System.Drawing.Point(280, ($filterGroupBaseH + 30))
            $popupForm.Size = New-Object System.Drawing.Size(450, $collapsedHeight)
        }
    }

    # --- Event: Funktionsbereich-Dropdown ---
    $fbCombo.Add_SelectedIndexChanged({
        $selected = $fbCombo.SelectedItem
        & $updateLayout ($selected -eq "Custom")
        $script:cockpitGroupFilter = $selected
    })

    # --- Schließen-Button ---
    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Schließen"
    $closeBtn.Location = New-Object System.Drawing.Point(20, ($filterGroupBaseH + 30))
    $closeBtn.Size = New-Object System.Drawing.Size(100, 30)
    $popupForm.Controls.Add($closeBtn)
    $popupForm.AcceptButton = $closeBtn
    $closeBtn.Add_Click({
        if ($fbCombo.SelectedItem -eq "Custom") {
            $script:cockpitGroupFilterCustom.Clear()
            foreach ($cb in $checkBoxes) {
                if ($cb.Checked) { $script:cockpitGroupFilterCustom.Add($cb.Text) | Out-Null }
            }
        }
        $script:cockpitGroupFilter = $fbCombo.SelectedItem
        $script:cockpitRatingFilter = $bewertungCombo.SelectedItem
        $popupForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $popupForm.Close()
    })

    # --- Filter zurücksetzen-Button ---
    $resetBtn = New-Object System.Windows.Forms.Button
    $resetBtn.Text = "Filter zurücksetzen"
    $resetBtn.Location = New-Object System.Drawing.Point(280, ($filterGroupBaseH + 30))
    $resetBtn.Size = New-Object System.Drawing.Size(140, 30)
    $popupForm.Controls.Add($resetBtn)
    $resetBtn.Add_Click({
        $zeitraumCombo.SelectedItem = "Alle Daten"
        $fbCombo.SelectedItem = "Alle"
        $bewertungCombo.SelectedItem = "Alle"
        foreach ($cb in $checkBoxes) { $cb.Checked = $true }
        $script:cockpitTimeFilter = "Alle Daten"
        $script:cockpitGroupFilter = "Alle"
        $script:cockpitGroupFilterCustom.Clear()
        $script:cockpitRatingFilter = "Alle"
    })

    # Falls bereits Custom aktiv, Popup direkt expanded öffnen
    if ($script:cockpitGroupFilter -eq "Custom") {
        & $updateLayout $true
    }

    $popupForm.ShowDialog()
}

function Show-DataMgmtSettingsPopup {
    $popupForm = New-Object System.Windows.Forms.Form
    $popupForm.Size = New-Object System.Drawing.Size(450, 250)
    $popupForm.Text = "Einstellungen"
    $popupForm.StartPosition = "CenterParent"
    $popupForm.FormBorderStyle = "FixedDialog"
    $popupForm.MaximizeBox = $false
    $popupForm.MinimizeBox = $false
    $filterGroup = New-Object System.Windows.Forms.GroupBox
    $filterGroup.Text = "Filter"
    $filterGroup.Location = New-Object System.Drawing.Point(20, 20)
    $filterGroup.Size = New-Object System.Drawing.Size(400, 110)
    $popupForm.Controls.Add($filterGroup)
    # --- Sub-GroupBox: Zeitraum ---
    $zeitraumGroup = New-Object System.Windows.Forms.GroupBox
    $zeitraumGroup.Text = "Zeitraum"
    $zeitraumGroup.Location = New-Object System.Drawing.Point(15, 25)
    $zeitraumGroup.Size = New-Object System.Drawing.Size(370, 60)
    $filterGroup.Controls.Add($zeitraumGroup)
    $zeitraumLabel = New-Object System.Windows.Forms.Label
    $zeitraumLabel.Text = "Zeitraum:"
    $zeitraumLabel.Location = New-Object System.Drawing.Point(15, 24)
    $zeitraumLabel.Size = New-Object System.Drawing.Size(70, 20)
    $zeitraumGroup.Controls.Add($zeitraumLabel)
    $zeitraumCombo = New-Object System.Windows.Forms.ComboBox
    $zeitraumCombo.Location = New-Object System.Drawing.Point(90, 22)
    $zeitraumCombo.Size = New-Object System.Drawing.Size(260, 20)
    $zeitraumCombo.DropDownStyle = "DropDownList"
    $zeitraumCombo.Items.AddRange(@("Aktuelles Jahr", "Letzte 3 Jahre", "Letzte 5 Jahre", "Letzte 10 Jahre", "Alle Daten"))
    $zeitraumCombo.SelectedItem = $script:dataMgmtTimeFilter
    $zeitraumGroup.Controls.Add($zeitraumCombo)
    $zeitraumCombo.Add_SelectedIndexChanged({ $script:dataMgmtTimeFilter = $zeitraumCombo.SelectedItem })
    # --- Buttons ---
    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Schließen"
    $closeBtn.Location = New-Object System.Drawing.Point(20, 140)
    $closeBtn.Size = New-Object System.Drawing.Size(100, 30)
    $closeBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $popupForm.Controls.Add($closeBtn)
    $popupForm.AcceptButton = $closeBtn
    $resetBtn = New-Object System.Windows.Forms.Button
    $resetBtn.Text = "Filter zurücksetzen"
    $resetBtn.Location = New-Object System.Drawing.Point(280, 140)
    $resetBtn.Size = New-Object System.Drawing.Size(140, 30)
    $popupForm.Controls.Add($resetBtn)
    $resetBtn.Add_Click({
        $zeitraumCombo.SelectedItem = "Alle Daten"
        $script:dataMgmtTimeFilter = "Alle Daten"
    })
    $popupForm.ShowDialog()
}

function Show-LongevitySettingsPopup {
    $popupForm = New-Object System.Windows.Forms.Form
    $popupForm.Size = New-Object System.Drawing.Size(450, 250)
    $popupForm.Text = "Einstellungen"
    $popupForm.StartPosition = "CenterParent"
    $popupForm.FormBorderStyle = "FixedDialog"
    $popupForm.MaximizeBox = $false
    $popupForm.MinimizeBox = $false
    $filterGroup = New-Object System.Windows.Forms.GroupBox
    $filterGroup.Text = "Filter"
    $filterGroup.Location = New-Object System.Drawing.Point(20, 20)
    $filterGroup.Size = New-Object System.Drawing.Size(400, 110)
    $popupForm.Controls.Add($filterGroup)
    $zeitraumGroup = New-Object System.Windows.Forms.GroupBox
    $zeitraumGroup.Text = "Zeitraum"
    $zeitraumGroup.Location = New-Object System.Drawing.Point(15, 25)
    $zeitraumGroup.Size = New-Object System.Drawing.Size(370, 60)
    $filterGroup.Controls.Add($zeitraumGroup)
    $zeitraumLabel = New-Object System.Windows.Forms.Label
    $zeitraumLabel.Text = "Zeitraum:"
    $zeitraumLabel.Location = New-Object System.Drawing.Point(15, 24)
    $zeitraumLabel.Size = New-Object System.Drawing.Size(70, 20)
    $zeitraumGroup.Controls.Add($zeitraumLabel)
    $zeitraumCombo = New-Object System.Windows.Forms.ComboBox
    $zeitraumCombo.Location = New-Object System.Drawing.Point(90, 22)
    $zeitraumCombo.Size = New-Object System.Drawing.Size(260, 20)
    $zeitraumCombo.DropDownStyle = "DropDownList"
    $zeitraumCombo.Items.AddRange(@("Aktuelles Jahr", "Letzte 3 Jahre", "Letzte 5 Jahre", "Letzte 10 Jahre", "Alle Daten"))
    $zeitraumCombo.SelectedItem = $script:longevityTimeFilter
    $zeitraumGroup.Controls.Add($zeitraumCombo)
    $zeitraumCombo.Add_SelectedIndexChanged({ $script:longevityTimeFilter = $zeitraumCombo.SelectedItem })
    $closeBtn = New-Object System.Windows.Forms.Button
    $closeBtn.Text = "Schließen"
    $closeBtn.Location = New-Object System.Drawing.Point(20, 140)
    $closeBtn.Size = New-Object System.Drawing.Size(100, 30)
    $closeBtn.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $popupForm.Controls.Add($closeBtn)
    $popupForm.AcceptButton = $closeBtn
    $resetBtn = New-Object System.Windows.Forms.Button
    $resetBtn.Text = "Filter zurücksetzen"
    $resetBtn.Location = New-Object System.Drawing.Point(280, 140)
    $resetBtn.Size = New-Object System.Drawing.Size(140, 30)
    $popupForm.Controls.Add($resetBtn)
    $resetBtn.Add_Click({
        $zeitraumCombo.SelectedItem = "Alle Daten"
        $script:longevityTimeFilter = "Alle Daten"
    })
    $popupForm.ShowDialog()
}

function Show-DataEntryForm {
    $form = New-Object System.Windows.Forms.Form; $form.Size = New-Object System.Drawing.Size(1200, 850); $form.StartPosition = "CenterScreen"; $form.Text = "Blutwert-Tracker - Version $Version"; $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 10)
    $tabControl.Size = New-Object System.Drawing.Size(1160, 790)
    $tabControl.Anchor = 'Top, Bottom, Left, Right'
    $form.Controls.Add($tabControl)

    # v2.15.0/2.15.1: Form-globale Buttons in der Tab-Header-Zeile (aus JEDEM Tab sichtbar).
    # Y=8 damit die Buttons die obere Linie des TabControl-Rahmens nicht minimal überlappen.
    # BringToFront() damit sie optisch über dem TabControl liegen.
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Beenden"
    $closeButton.Size = New-Object System.Drawing.Size(90, 24)
    $closeButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - $closeButton.Width - 20), 8)
    $closeButton.Anchor = "Top, Right"
    $closeButton.ForeColor = [System.Drawing.Color]::DarkRed
    $closeButton.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)
    $closeButton.BringToFront()

    # Globaler Einstellungen-Button direkt links neben Beenden (v2.15.1).
    # Ehemals in Tab "Einzelmarker-Analyse" als "Einstellungen" - jetzt form-weit verfügbar.
    $globalSettingsButton = New-Object System.Windows.Forms.Button
    $globalSettingsButton.Text = "globale Einstellungen"
    $globalSettingsButton.Size = New-Object System.Drawing.Size(140, 24)
    $globalSettingsButton.Location = New-Object System.Drawing.Point(($closeButton.Left - $globalSettingsButton.Width - 8), 8)
    $globalSettingsButton.Anchor = "Top, Right"
    $globalSettingsButton.Add_Click({
        Show-SettingsPopup -mainForm $form
        # v2.20.0: Tooltip-Verhalten aktualisieren nach Settings-Änderung
        $script:persistentTooltips = $script:data.Config.Personal.PersistentTooltips
        if ($script:persistentTooltips) {
            $toolTip.AutoPopDelay = 30000
            $toolTip.InitialDelay = 400
        } else {
            $toolTip.AutoPopDelay = 800
            $toolTip.InitialDelay = 300
            $toolTip.ReshowDelay = 100
        }
        # Tab-spezifische Refreshes nach Settings-Änderung (Funktionen sind im gleichen
        # Show-DataEntryForm-Scope definiert und zur Laufzeit verfügbar).
        try { Update-FilterDropdown } catch { Write-Warning "Update-FilterDropdown: $($_.Exception.Message)" }
        try { Update-Chart }          catch { Write-Warning "Update-Chart: $($_.Exception.Message)" }
    })
    $form.Controls.Add($globalSettingsButton)
    $globalSettingsButton.BringToFront()

    $tabPage1 = New-Object System.Windows.Forms.TabPage; $tabPage1.Text = "Einzelmarker-Analyse"
    $tabPage1.Tag = "Hier können Sie einzelne Blutwerte über die Zeit analysieren, neue Werte eintragen und die Marker-Konfiguration verwalten."
    $tabControl.TabPages.Add($tabPage1)

    $tabPage2 = New-Object System.Windows.Forms.TabPage; $tabPage2.Text = "Risiko-Cockpit"
    $tabPage2.Tag = "Dieses Cockpit bietet eine schnelle Übersicht aller wichtigen Marker, bewertet nach Referenz- und Optimalbereichen, um Risiken auf einen Blick zu erkennen."
    $tabControl.TabPages.Add($tabPage2)

    $tabPage3 = New-Object System.Windows.Forms.TabPage; $tabPage3.Text = "Korrelationen"
    $tabPage3.Tag = "Analysieren Sie hier den Zusammenhang zwischen zwei verschiedenen Blutmarkern in einem Streudiagramm."
    $tabControl.TabPages.Add($tabPage3)

    $tabPage4 = New-Object System.Windows.Forms.TabPage; $tabPage4.Text = "Daten nach Bluttests"
    $tabPage4.Tag = "Verwalten Sie hier alle erfassten Bluttests. Sie können komplette Testtage einsehen, Werte bearbeiten und Einträge löschen."
    $tabControl.TabPages.Add($tabPage4)

    $tabPage5 = New-Object System.Windows.Forms.TabPage; $tabPage5.Text = "Longevity-Indizes"
    $tabPage5.Tag = "Dieses Dashboard bewertet Schlüsselmarker für ein langes, gesundes Leben und berechnet einen Gesamt-Score."
    $tabControl.TabPages.Add($tabPage5)

    $tabPageCustom = New-Object System.Windows.Forms.TabPage; $tabPageCustom.Text = "Custom Report"
    $tabPageCustom.Tag = "Frei konfigurierbarer Report: wähle beliebige Blutmarker und einen Zeitraum, drucke oder exportiere das Ergebnis."
    $tabControl.TabPages.Add($tabPageCustom)

    $tabPage6 = New-Object System.Windows.Forms.TabPage; $tabPage6.Text = "Persönliche Metriken"
    $tabPage6.Tag = "Verwalten Sie hier Ihre persönlichen Basisdaten, die für einige Berechnungen (z.B. BMI, PREVENT-Score) benötigt werden."
    $tabControl.TabPages.Add($tabPage6)
    
    # ---------- TAB 2: RISIKO-COCKPIT ----------
    $cockpitGrid = New-Object System.Windows.Forms.DataGridView
    $cockpitGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $cockpitGrid.AllowUserToAddRows = $false; $cockpitGrid.AllowUserToDeleteRows = $false; $cockpitGrid.ReadOnly = $true
    $cockpitGrid.AutoSizeColumnsMode = "Fill"
    $cockpitGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $cockpitGrid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $cockpitGrid.RowTemplate.Height = 30
    $cockpitGrid.Columns.Add("Group", "Funktionsbereich"); $cockpitGrid.Columns.Add("Marker", "Marker"); $cockpitGrid.Columns.Add("Value", "Letzter Wert"); $cockpitGrid.Columns.Add("Date", "Datum"); $cockpitGrid.Columns.Add("Risk", "10Y-Risiko (%)"); $cockpitGrid.Columns.Add("Unit", "Einheit"); $cockpitGrid.Columns.Add("Rating", "Bewertung")
    $sortOrderCol = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $sortOrderCol.Name = "SortOrder"; $sortOrderCol.Visible = $false
    $cockpitGrid.Columns.Add($sortOrderCol)
    $cockpitGrid.Columns["Group"].FillWeight = 25; $cockpitGrid.Columns["Marker"].FillWeight = 40; $cockpitGrid.Columns["Value"].FillWeight = 15; $cockpitGrid.Columns["Date"].FillWeight = 15; $cockpitGrid.Columns["Risk"].FillWeight = 12; $cockpitGrid.Columns["Unit"].FillWeight = 10; $cockpitGrid.Columns["Rating"].FillWeight = 10
    $cockpitTopPanel = New-Object System.Windows.Forms.Panel; $cockpitTopPanel.Height = 40; $cockpitTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $cockpitSettingsButton = New-Object System.Windows.Forms.Button; $cockpitSettingsButton.Text = "Einstellungen"; $cockpitSettingsButton.Size = New-Object System.Drawing.Size(100, 24); $cockpitSettingsButton.Location = New-Object System.Drawing.Point(10, 8)
    $cockpitTopPanel.Controls.Add($cockpitSettingsButton)
    $cockpitPrintButton = New-Object System.Windows.Forms.Button; $cockpitPrintButton.Text = "Drucken"; $cockpitPrintButton.Size = New-Object System.Drawing.Size(80, 24); $cockpitPrintButton.Location = New-Object System.Drawing.Point(120, 8)
    $cockpitTopPanel.Controls.Add($cockpitPrintButton)
    $cockpitExportButton = New-Object System.Windows.Forms.Button; $cockpitExportButton.Text = "Exportieren..."; $cockpitExportButton.Size = New-Object System.Drawing.Size(100, 24); $cockpitExportButton.Location = New-Object System.Drawing.Point(($cockpitPrintButton.Right + 8), 8)
    $cockpitTopPanel.Controls.Add($cockpitExportButton)
    $cockpitFilterHintLabel = New-Object System.Windows.Forms.Label; $cockpitFilterHintLabel.Location = New-Object System.Drawing.Point(320, 12); $cockpitFilterHintLabel.Size = New-Object System.Drawing.Size(820, 20); $cockpitFilterHintLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $cockpitFilterHintLabel.ForeColor = [System.Drawing.Color]::Red; $cockpitFilterHintLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter; $cockpitFilterHintLabel.Text = ""
    $cockpitTopPanel.Controls.Add($cockpitFilterHintLabel)
    $tabPage2.Controls.Add($cockpitGrid)
    $tabPage2.Controls.Add($cockpitTopPanel)

    # ---------- TAB 1: EINZELMARKER-ANALYSE ----------
    $toolTip = New-Object System.Windows.Forms.ToolTip
    # v2.20.0: Tooltip-Verhalten (hover-only vs. dauerhaft)
    $script:persistentTooltips = $script:data.Config.Personal.PersistentTooltips
    if ($script:persistentTooltips) {
        $toolTip.AutoPopDelay = 30000  # 30 Sekunden – dauerhaft sichtbar
        $toolTip.InitialDelay = 400
    } else {
        $toolTip.AutoPopDelay = 800    # 0.8 Sekunden – verschwindet schnell
        $toolTip.InitialDelay = 300
        $toolTip.ReshowDelay = 100
    }

    $entryGroup = New-Object System.Windows.Forms.GroupBox; $entryGroup.Location = New-Object System.Drawing.Point(20, 20); $entryGroup.Size = New-Object System.Drawing.Size(420, 220); $entryGroup.Text = "Wert eintragen"; $tabPage1.Controls.Add($entryGroup)
    $labelDate = New-Object System.Windows.Forms.Label; $labelDate.Text = "Datum:"; $labelDate.Location = New-Object System.Drawing.Point(10, 30); $labelDate.Size = New-Object System.Drawing.Size(100, 20); $entryGroup.Controls.Add($labelDate)
    $datePicker = New-Object System.Windows.Forms.DateTimePicker; $datePicker.Location = New-Object System.Drawing.Point(110, 30); $datePicker.Size = New-Object System.Drawing.Size(300, 20); $entryGroup.Controls.Add($datePicker)
    $labelMarker = New-Object System.Windows.Forms.Label; $labelMarker.Text = "Blutmarker:"; $labelMarker.Location = New-Object System.Drawing.Point(10, 60); $labelMarker.Size = New-Object System.Drawing.Size(100, 20); $entryGroup.Controls.Add($labelMarker)
    # v2.23.0: Blutmarker-Dropdown mit Alias-Live-Suche
    # Auswahl aus der Liste wie bisher ODER Direkteingabe von Labor-Synonymen
    # ("WBC", "Leukozyten", "HbA1c", "GOT" ...). Die Liste filtert live mit.
    $markerComboBox = New-Object System.Windows.Forms.ComboBox
    $markerComboBox.Location = New-Object System.Drawing.Point(110, 60)
    $markerComboBox.Size = New-Object System.Drawing.Size(300, 20)
    $markerComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $markerComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::None
    $markerComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::None
    $markerComboBox.MaxDropDownItems = 12
    $entryGroup.Controls.Add($markerComboBox)
    $toolTip.SetToolTip($markerComboBox, "Marker aus der Liste waehlen ODER direkt tippen - auch Labor-Kuerzel wie 'WBC', 'GOT', 'HbA1c', 'hs-CRP'. Enter uebernimmt den Treffer, Esc setzt zurueck.")
    $labelValue = New-Object System.Windows.Forms.Label; $labelValue.Text = "Wert:"; $labelValue.Location = New-Object System.Drawing.Point(10, 90); $labelValue.Size = New-Object System.Drawing.Size(100, 20); $entryGroup.Controls.Add($labelValue)
    $valueTextBox = New-Object System.Windows.Forms.TextBox; $valueTextBox.Location = New-Object System.Drawing.Point(110, 90); $valueTextBox.Size = New-Object System.Drawing.Size(100, 20); $entryGroup.Controls.Add($valueTextBox)
    # v2.20.0: HIV-Dropdown (nicht reaktiv / reaktiv) - ersetzt TextBox bei HIV-Auswahl
    $hivComboBox = New-Object System.Windows.Forms.ComboBox
    $hivComboBox.Location = New-Object System.Drawing.Point(110, 90)
    $hivComboBox.Size = New-Object System.Drawing.Size(160, 20)
    $hivComboBox.DropDownStyle = "DropDownList"
    $hivComboBox.Items.AddRange(@("nicht reaktiv", "reaktiv"))
    $hivComboBox.SelectedIndex = 0
    $hivComboBox.Visible = $false
    $entryGroup.Controls.Add($hivComboBox)
    # v2.26.0: APOE-Genotyp-Dropdown - ersetzt die TextBox bei Auswahl des
    # Markers "Apolipoprotein E-Genotyp (APOE)". Reihenfolge = Code 1-6.
    $apoeComboBox = New-Object System.Windows.Forms.ComboBox
    $apoeComboBox.Location = New-Object System.Drawing.Point(110, 90)
    $apoeComboBox.Size = New-Object System.Drawing.Size(160, 20)
    $apoeComboBox.DropDownStyle = "DropDownList"
    $apoeComboBox.Items.AddRange([string[]]$script:ApoeGenotypeOptions)
    $apoeComboBox.SelectedIndex = 2   # Vorbelegung E3/E3 (haeufigster Genotyp)
    $apoeComboBox.Visible = $false
    $entryGroup.Controls.Add($apoeComboBox)
    $toolTip.SetToolTip($apoeComboBox, "APOE-Genotyp exakt so waehlen, wie er im Befund steht. Sortierung nach aufsteigendem Risiko: E2/E2, E2/E3, E3/E3 (kein E4-Allel) - E2/E4, E3/E4, E4/E4 (mit E4-Allel).")
    # v2.20.0: Testosteron-Einheiten-Dropdown (ng/dl oder nmol/l)
    $testoUnitCombo = New-Object System.Windows.Forms.ComboBox
    $testoUnitCombo.Location = New-Object System.Drawing.Point(215, 90)
    $testoUnitCombo.Size = New-Object System.Drawing.Size(80, 20)
    $testoUnitCombo.DropDownStyle = "DropDownList"
    $testoUnitCombo.Items.AddRange(@("ng/dl", "nmol/l"))
    $testoUnitCombo.SelectedIndex = 0
    $testoUnitCombo.Visible = $false
    $entryGroup.Controls.Add($testoUnitCombo)
    # v2.23.0: Info-Zeile rechts neben dem Wert-Feld (Einheit + Referenzbereich)
    $markerInfoLabel = New-Object System.Windows.Forms.Label
    $markerInfoLabel.Location = New-Object System.Drawing.Point(216, 93)
    $markerInfoLabel.Size = New-Object System.Drawing.Size(194, 16)
    $markerInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 7.5)
    $markerInfoLabel.ForeColor = [System.Drawing.Color]::DimGray
    $markerInfoLabel.Text = ""
    $markerInfoLabel.Visible = $false
    $entryGroup.Controls.Add($markerInfoLabel)

    # ---- v2.23.0: Zustand + Helfer der Alias-Live-Suche -------------------
    $script:markerFilterBusy   = $false   # verhindert Event-Rekursion
    $script:markerFullList     = @()      # vollstaendige, sortierte Markerliste
    $script:currentMarkerName  = $null    # aktuell aufgeloester Marker

    function Set-MarkerComboItems {
        param([string[]]$Items)
        $prev = $script:markerFilterBusy
        $script:markerFilterBusy = $true
        $markerComboBox.BeginUpdate()
        $markerComboBox.Items.Clear()
        if ($Items -and $Items.Count -gt 0) { $markerComboBox.Items.AddRange([string[]]$Items) }
        $markerComboBox.EndUpdate()
        $script:markerFilterBusy = $prev
    }

    function Update-MarkerInputMode {
        # Steuert HIV-/APOE-/Testosteron-Sonderfelder und die Info-Zeile.
        param([string]$MarkerName)
        $script:currentMarkerName = $MarkerName
        $isHIV   = ($MarkerName -eq "HIV (Anti-HIV-1/2)")
        $isTesto = ($MarkerName -eq "Testosteron, gesamt")
        $isApoe  = ($MarkerName -eq $script:ApoeMarkerName)   # v2.26.0
        $hivComboBox.Visible    = $isHIV
        $apoeComboBox.Visible   = $isApoe
        $valueTextBox.Visible   = -not ($isHIV -or $isApoe)
        $testoUnitCombo.Visible = $isTesto
        # v2.26.0: Beim APOE-Genotyp die Referenz-Logik als Hilfetext einblenden
        if ($isApoe) {
            $markerInfoLabel.Text = "Genotyp waehlen  |  Referenz: kein E4-Allel (E2/E2, E2/E3, E3/E3)"
            $markerInfoLabel.Visible = $true
            return
        }
        if ($isHIV -or $isTesto -or [string]::IsNullOrWhiteSpace($MarkerName)) {
            $markerInfoLabel.Text = ""
            $markerInfoLabel.Visible = $false
            return
        }
        $cfg = $script:data.Config.Markers | Where-Object { $_.Name -eq $MarkerName } | Select-Object -First 1
        if (-not $cfg) { $markerInfoLabel.Text = ""; $markerInfoLabel.Visible = $false; return }
        $parts = @()
        if ("$($cfg.Unit)".Trim()) { $parts += "$($cfg.Unit)".Trim() }
        if ($null -ne $cfg.RefMin -and $null -ne $cfg.RefMax -and "$($cfg.RefMin)" -ne "" -and "$($cfg.RefMax)" -ne "") {
            $parts += "Ref $($cfg.RefMin)-$($cfg.RefMax)"
        }
        if ($parts.Count -gt 0) {
            $markerInfoLabel.Text = ($parts -join "  |  ")
            $markerInfoLabel.Visible = $true
        } else {
            $markerInfoLabel.Text = ""
            $markerInfoLabel.Visible = $false
        }
    }

    function Commit-MarkerSelection {
        # Loest die aktuelle Eingabe auf, stellt die volle Liste wieder her
        # und markiert nicht aufloesbare Eingaben farblich.
        # Ergebnis liegt danach in $script:currentMarkerName.
        $typed = "$($markerComboBox.Text)".Trim()
        $resolved = Resolve-MarkerName -Query $typed
        $prev = $script:markerFilterBusy
        $script:markerFilterBusy = $true
        if ($markerComboBox.DroppedDown) { $markerComboBox.DroppedDown = $false }
        Set-MarkerComboItems -Items $script:markerFullList
        if ($resolved) {
            $idx = $markerComboBox.Items.IndexOf($resolved)
            if ($idx -ge 0) { $markerComboBox.SelectedIndex = $idx } else { $markerComboBox.Text = $resolved }
            $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
        } else {
            $markerComboBox.SelectedIndex = -1
            $markerComboBox.Text = $typed
            if ($typed) {
                $markerComboBox.BackColor = [System.Drawing.Color]::MistyRose
            } else {
                $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
            }
        }
        $script:markerFilterBusy = $prev
        Update-MarkerInputMode -MarkerName $resolved
    }

    # Live-Filter waehrend der Eingabe (TextUpdate feuert nur bei Nutzer-Tippen)
    $markerComboBox.Add_TextUpdate({
        if ($script:markerFilterBusy) { return }
        $typed = "$($markerComboBox.Text)"
        if ([string]::IsNullOrWhiteSpace($typed)) {
            Set-MarkerComboItems -Items $script:markerFullList
            $script:markerFilterBusy = $true
            $markerComboBox.SelectedIndex = -1
            $markerComboBox.Text = ""
            $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
            $script:markerFilterBusy = $false
            Update-MarkerInputMode -MarkerName $null
            return
        }
        $hits = @(Get-MarkerMatches -Query $typed)
        $noHit = ($hits.Count -eq 0)
        if ($noHit) { $hits = @($script:markerFullList) }
        Set-MarkerComboItems -Items $hits
        $script:markerFilterBusy = $true
        $markerComboBox.SelectedIndex = -1
        $markerComboBox.Text = $typed
        if ($noHit) {
            $markerComboBox.BackColor = [System.Drawing.Color]::MistyRose
        } else {
            $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
            if (-not $markerComboBox.DroppedDown) { $markerComboBox.DroppedDown = $true }
            [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
        }
        $markerComboBox.SelectionStart = $typed.Length
        $markerComboBox.SelectionLength = 0
        $script:markerFilterBusy = $false
        # Live-Vorschau nur bei eindeutigem Treffer. Wird die Eingabe wieder
        # mehrdeutig, muessen HIV-/Testo-Sonderfelder zurueckgesetzt werden,
        # damit kein falsches Eingabefeld stehen bleibt.
        $preview = Resolve-MarkerName -Query $typed
        if ($preview) {
            Update-MarkerInputMode -MarkerName $preview
        } elseif ($hivComboBox.Visible -or $apoeComboBox.Visible -or $testoUnitCombo.Visible) {
            Update-MarkerInputMode -MarkerName $null
        }
    })

    $markerComboBox.Add_KeyDown({
        param($eventSender, $eventArgs)
        if ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $eventArgs.Handled = $true
            $eventArgs.SuppressKeyPress = $true
            Commit-MarkerSelection
        } elseif ($eventArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $eventArgs.Handled = $true
            $eventArgs.SuppressKeyPress = $true
            $script:markerFilterBusy = $true
            if ($markerComboBox.DroppedDown) { $markerComboBox.DroppedDown = $false }
            Set-MarkerComboItems -Items $script:markerFullList
            $markerComboBox.SelectedIndex = -1
            $markerComboBox.Text = ""
            $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
            $script:markerFilterBusy = $false
            Update-MarkerInputMode -MarkerName $null
        }
    })

    $markerComboBox.Add_Leave({
        if ($script:markerFilterBusy) { return }
        Commit-MarkerSelection
    })

    # v2.20.0/v2.23.0: Marker-Auswahl steuert HIV/Testo-UI und Info-Zeile
    $markerComboBox.Add_SelectedIndexChanged({
        if ($script:markerFilterBusy) { return }
        $sel = [string]$markerComboBox.SelectedItem
        $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
        Update-MarkerInputMode -MarkerName $sel
    })
    $labelNote = New-Object System.Windows.Forms.Label; $labelNote.Text = "Notiz:"; $labelNote.Location = New-Object System.Drawing.Point(10, 120); $labelNote.Size = New-Object System.Drawing.Size(100, 20); $entryGroup.Controls.Add($labelNote)
    $noteTextBox = New-Object System.Windows.Forms.TextBox; $noteTextBox.Location = New-Object System.Drawing.Point(110, 120); $noteTextBox.Size = New-Object System.Drawing.Size(300, 50); $noteTextBox.Multiline = $true; $noteTextBox.ScrollBars = "Vertical"; $entryGroup.Controls.Add($noteTextBox)
    $addButton = New-Object System.Windows.Forms.Button; $addButton.Location = New-Object System.Drawing.Point(110, 180); $addButton.Size = New-Object System.Drawing.Size(300, 30); $addButton.Text = "Hinzufügen"; $entryGroup.Controls.Add($addButton)
    $markerMgmtGroup = New-Object System.Windows.Forms.GroupBox; $markerMgmtGroup.Location = New-Object System.Drawing.Point(20, 250); $markerMgmtGroup.Size = New-Object System.Drawing.Size(420, 220); $markerMgmtGroup.Text = "Blutmarker verwalten"; $tabPage1.Controls.Add($markerMgmtGroup)
    $labelEditMarker = New-Object System.Windows.Forms.Label; $labelEditMarker.Text = "Bestehenden Marker bearbeiten oder löschen:"; $labelEditMarker.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $labelEditMarker.Location = New-Object System.Drawing.Point(10, 20); $labelEditMarker.AutoSize = $true; $markerMgmtGroup.Controls.Add($labelEditMarker)
    $deleteMarkerComboBox = New-Object System.Windows.Forms.ComboBox; $deleteMarkerComboBox.Location = New-Object System.Drawing.Point(10, 45); $deleteMarkerComboBox.Size = New-Object System.Drawing.Size(250, 20); $markerMgmtGroup.Controls.Add($deleteMarkerComboBox)
    $loadMarkerForEditButton = New-Object System.Windows.Forms.Button; $loadMarkerForEditButton.Location = New-Object System.Drawing.Point(265, 42); $loadMarkerForEditButton.Size = New-Object System.Drawing.Size(145, 26); $loadMarkerForEditButton.Text = "Zum Bearbeiten laden"; $markerMgmtGroup.Controls.Add($loadMarkerForEditButton)
    $separatorLabel = New-Object System.Windows.Forms.Label; $separatorLabel.Location = New-Object System.Drawing.Point(10, 75); $separatorLabel.Size = New-Object System.Drawing.Size(400, 2); $separatorLabel.BorderStyle = 'Fixed3D'; $markerMgmtGroup.Controls.Add($separatorLabel)
    $labelCreateMarker = New-Object System.Windows.Forms.Label; $labelCreateMarker.Text = "Marken anlegen / Details bearbeiten:"; $labelCreateMarker.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $labelCreateMarker.Location = New-Object System.Drawing.Point(10, 90); $labelCreateMarker.AutoSize = $true; $markerMgmtGroup.Controls.Add($labelCreateMarker)
    $advancedViewCheckBox = New-Object System.Windows.Forms.CheckBox; $advancedViewCheckBox.Text = "Erweiterte Einstellungen"; $advancedViewCheckBox.Location = New-Object System.Drawing.Point(245, 88); $advancedViewCheckBox.AutoSize = $true; $markerMgmtGroup.Controls.Add($advancedViewCheckBox)
    $labelNewName = New-Object System.Windows.Forms.Label; $labelNewName.Text = "Name:"; $labelNewName.Location = New-Object System.Drawing.Point(10, 115); $labelNewName.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewName)
    $textNewName = New-Object System.Windows.Forms.TextBox; $textNewName.Location = New-Object System.Drawing.Point(110, 115); $textNewName.Size = New-Object System.Drawing.Size(300, 20); $markerMgmtGroup.Controls.Add($textNewName)
    $labelNewUnit = New-Object System.Windows.Forms.Label; $labelNewUnit.Text = "Einheit:"; $labelNewUnit.Location = New-Object System.Drawing.Point(10, 145); $labelNewUnit.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewUnit)
    $textNewUnit = New-Object System.Windows.Forms.TextBox; $textNewUnit.Location = New-Object System.Drawing.Point(110, 145); $textNewUnit.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($textNewUnit)
    $labelNewGroup = New-Object System.Windows.Forms.Label; $labelNewGroup.Text = "Gruppe:"; $labelNewGroup.Location = New-Object System.Drawing.Point(10, 175); $labelNewGroup.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewGroup)
    $textNewGroup = New-Object System.Windows.Forms.TextBox; $textNewGroup.Location = New-Object System.Drawing.Point(110, 175); $textNewGroup.Size = New-Object System.Drawing.Size(150, 20); $markerMgmtGroup.Controls.Add($textNewGroup)
    $labelNewMin = New-Object System.Windows.Forms.Label; $labelNewMin.Text = "Ref. Min:"; $labelNewMin.Location = New-Object System.Drawing.Point(10, 205); $labelNewMin.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewMin)
    $textNewMin = New-Object System.Windows.Forms.TextBox; $textNewMin.Location = New-Object System.Drawing.Point(110, 205); $textNewMin.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($textNewMin)
    $labelNewMax = New-Object System.Windows.Forms.Label; $labelNewMax.Text = "Ref. Max:"; $labelNewMax.Location = New-Object System.Drawing.Point(10, 235); $labelNewMax.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewMax)
    $textNewMax = New-Object System.Windows.Forms.TextBox; $textNewMax.Location = New-Object System.Drawing.Point(110, 235); $textNewMax.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($textNewMax)
    $labelNewOptMin = New-Object System.Windows.Forms.Label; $labelNewOptMin.Text = "Optimal Min:"; $labelNewOptMin.Location = New-Object System.Drawing.Point(10, 265); $labelNewOptMin.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewOptMin)
    $textNewOptMin = New-Object System.Windows.Forms.TextBox; $textNewOptMin.Location = New-Object System.Drawing.Point(110, 265); $textNewOptMin.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($textNewOptMin)
    $labelNewOptMax = New-Object System.Windows.Forms.Label; $labelNewOptMax.Text = "Optimal Max:"; $labelNewOptMax.Location = New-Object System.Drawing.Point(10, 295); $labelNewOptMax.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($labelNewOptMax)
    $textNewOptMax = New-Object System.Windows.Forms.TextBox; $textNewOptMax.Location = New-Object System.Drawing.Point(110, 295); $textNewOptMax.Size = New-Object System.Drawing.Size(100, 20); $markerMgmtGroup.Controls.Add($textNewOptMax)
    $addNewMarkerButton = New-Object System.Windows.Forms.Button; $addNewMarkerButton.Location = New-Object System.Drawing.Point(110, 175); $addNewMarkerButton.Size = New-Object System.Drawing.Size(150, 30); $addNewMarkerButton.Text = "Speichern"; $markerMgmtGroup.Controls.Add($addNewMarkerButton)
    $deleteExistingMarkerButton = New-Object System.Windows.Forms.Button; $deleteExistingMarkerButton.Location = New-Object System.Drawing.Point(265, 175); $deleteExistingMarkerButton.Size = New-Object System.Drawing.Size(145, 30); $deleteExistingMarkerButton.Text = "Auswahl löschen"; $markerMgmtGroup.Controls.Add($deleteExistingMarkerButton)
    $warningThresholdGroup = New-Object System.Windows.Forms.GroupBox; $warningThresholdGroup.Location = New-Object System.Drawing.Point(20, 480); $warningThresholdGroup.Size = New-Object System.Drawing.Size(420, 70); $warningThresholdGroup.Text = "Schwellenwert für Warnungen"; $tabPage1.Controls.Add($warningThresholdGroup)
    $labelThreshold = New-Object System.Windows.Forms.Label; $labelThreshold.Text = "Warnen bei >"; $labelThreshold.Location = New-Object System.Drawing.Point(10, 30); $labelThreshold.Size = New-Object System.Drawing.Size(100, 20); $warningThresholdGroup.Controls.Add($labelThreshold)
    $thresholdTextBox = New-Object System.Windows.Forms.TextBox; $thresholdTextBox.Location = New-Object System.Drawing.Point(110, 28); $thresholdTextBox.Size = New-Object System.Drawing.Size(50, 20); $warningThresholdGroup.Controls.Add($thresholdTextBox)
    $labelPercent = New-Object System.Windows.Forms.Label; $labelPercent.Text = "% Abweichung / Jahr"; $labelPercent.Location = New-Object System.Drawing.Point(165, 30); $labelPercent.AutoSize = $true; $warningThresholdGroup.Controls.Add($labelPercent)
    $saveSettingsButton = New-Object System.Windows.Forms.Button; $saveSettingsButton.Location = New-Object System.Drawing.Point(290, 26); $saveSettingsButton.Size = New-Object System.Drawing.Size(110, 28); $saveSettingsButton.Text = "Speichern"; $warningThresholdGroup.Controls.Add($saveSettingsButton)
    # Beenden-Button wird auf $form-Ebene platziert (siehe Form-Setup unten, v2.15.0)
    $filterLabel = New-Object System.Windows.Forms.Label; $filterLabel.Text = "Verlauf anzeigen für:"; $filterLabel.Location = New-Object System.Drawing.Point(460, 20); $filterLabel.AutoSize = $true; $tabPage1.Controls.Add($filterLabel)
    $filterComboBox = New-Object System.Windows.Forms.ComboBox; $filterComboBox.Location = New-Object System.Drawing.Point(600, 18); $filterComboBox.Size = New-Object System.Drawing.Size(250, 20); $tabPage1.Controls.Add($filterComboBox)
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart; $chart.Location = New-Object System.Drawing.Point(460, 50); $chart.Size = New-Object System.Drawing.Size(680, 570); $chart.Anchor = "Top, Left"; $tabPage1.Controls.Add($chart)
    # v2.15.1: $settingsButton entfernt - wurde zu globalem "globale Einstellungen"-Button
    # auf Form-Ebene (siehe Show-DataEntryForm). Bearbeiten/Drucken rücken an die frei
    # gewordene Position rechts im Chart-Bereich.
    $editButton = New-Object System.Windows.Forms.Button; $editButton.Text = "Bearbeiten"; $editButton.Size = New-Object System.Drawing.Size(90, 24); $editButton.Location = New-Object System.Drawing.Point(($chart.Right - $editButton.Width), 17); $tabPage1.Controls.Add($editButton)
    $chartPrintButton = New-Object System.Windows.Forms.Button; $chartPrintButton.Text = "Drucken"; $chartPrintButton.Size = New-Object System.Drawing.Size(90, 24); $chartPrintButton.Location = New-Object System.Drawing.Point(($editButton.Left - $chartPrintButton.Width - 8), 17); $tabPage1.Controls.Add($chartPrintButton)
    $chartExportButton = New-Object System.Windows.Forms.Button; $chartExportButton.Text = "Exportieren..."; $chartExportButton.Size = New-Object System.Drawing.Size(110, 28); $chartExportButton.Location = New-Object System.Drawing.Point($chart.Left, ($chart.Bottom + 4)); $tabPage1.Controls.Add($chartExportButton)
    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea; $chartArea.AxisX.Title = "Datum"; $chartArea.AxisX.LabelStyle.Format = "dd.MM.yy"; $chartArea.AxisY.Title = "Wert"; $chartArea.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray; $chart.ChartAreas.Add($chartArea); $chart.Legends.Clear()
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series; $series.Name = "Verlauf"; $series.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime; $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line; $series.BorderWidth = 3; $series.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle; $series.MarkerSize = 8; $chart.Series.Add($series)
    $trendSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series; $trendSeries.Name = "Trend"; $trendSeries.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime; $trendSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line; $trendSeries.BorderWidth = 2; $trendSeries.Color = [System.Drawing.Color]::MediumPurple; $trendSeries.BorderDashStyle = "Dash"; $chart.Series.Add($trendSeries)
    $warningDisplayGroup = New-Object System.Windows.Forms.GroupBox; $warningDisplayGroup.Location = New-Object System.Drawing.Point(580, 628); $warningDisplayGroup.Size = New-Object System.Drawing.Size(560, 122); $warningDisplayGroup.Text = "Wichtige Hinweise & Details"; $warningDisplayGroup.Anchor = "Top, Left"; $warningDisplayGroup.Visible = $false; $tabPage1.Controls.Add($warningDisplayGroup)
    $warningLabel = New-Object System.Windows.Forms.Label; $warningLabel.Location = New-Object System.Drawing.Point(10, 20); $warningLabel.Size = New-Object System.Drawing.Size(660, 90); $warningLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $warningLabel.ForeColor = [System.Drawing.Color]::DarkRed; $warningDisplayGroup.Controls.Add($warningLabel)
    $advancedControls = @($labelNewGroup, $textNewGroup, $labelNewMin, $textNewMin, $labelNewMax, $textNewMax, $labelNewOptMin, $textNewOptMin, $labelNewOptMax, $textNewOptMax)
    $lastTooltipIndex = -1

    # ---------- TAB 3: KORRELATIONEN ----------
    $corrGroup = New-Object System.Windows.Forms.GroupBox; $corrGroup.Text = "Marker auswählen"; $corrGroup.Location = New-Object System.Drawing.Point(20, 20); $corrGroup.Size = New-Object System.Drawing.Size(1120, 80); $corrGroup.Anchor = "Top, Left, Right"; $tabPage3.Controls.Add($corrGroup)
    $labelCorrMarker1 = New-Object System.Windows.Forms.Label; $labelCorrMarker1.Text = "Marker 1 (X-Achse):"; $labelCorrMarker1.Location = New-Object System.Drawing.Point(20, 35); $labelCorrMarker1.AutoSize = $true; $corrGroup.Controls.Add($labelCorrMarker1)
    $corrMarker1ComboBox = New-Object System.Windows.Forms.ComboBox; $corrMarker1ComboBox.Location = New-Object System.Drawing.Point(160, 32); $corrMarker1ComboBox.Size = New-Object System.Drawing.Size(250, 20); $corrGroup.Controls.Add($corrMarker1ComboBox)
    $labelCorrMarker2 = New-Object System.Windows.Forms.Label; $labelCorrMarker2.Text = "Marker 2 (Y-Achse):"; $labelCorrMarker2.Location = New-Object System.Drawing.Point(430, 35); $labelCorrMarker2.AutoSize = $true; $corrGroup.Controls.Add($labelCorrMarker2)
    $corrMarker2ComboBox = New-Object System.Windows.Forms.ComboBox; $corrMarker2ComboBox.Location = New-Object System.Drawing.Point(570, 32); $corrMarker2ComboBox.Size = New-Object System.Drawing.Size(250, 20); $corrGroup.Controls.Add($corrMarker2ComboBox)
    
    $runCorrelationButton = New-Object System.Windows.Forms.Button; $runCorrelationButton.Text = "Analyse durchführen"; $runCorrelationButton.Location = New-Object System.Drawing.Point(840, 28); $runCorrelationButton.Size = New-Object System.Drawing.Size(130, 30); $corrGroup.Controls.Add($runCorrelationButton)
    $toolTip.SetToolTip($runCorrelationButton, "Startet die Korrelationsanalyse für die beiden ausgewählten Blutmarker.")

    $suggestPairsButton = New-Object System.Windows.Forms.Button; $suggestPairsButton.Text = "Übersicht rel. Paare"; $suggestPairsButton.Location = New-Object System.Drawing.Point(980, 28); $suggestPairsButton.Size = New-Object System.Drawing.Size(130, 30); $corrGroup.Controls.Add($suggestPairsButton)
    $toolTip.SetToolTip($suggestPairsButton, "Analysiert alle für Ihr Risikoprofil relevanten Markerpaare und schlägt die mit der stärksten Korrelation vor.")

    $correlationResultLabel = New-Object System.Windows.Forms.Label; $correlationResultLabel.Text = "Bitte zwei Marker auswählen und Analyse starten."; $correlationResultLabel.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold); $correlationResultLabel.Location = New-Object System.Drawing.Point(20, 110); $correlationResultLabel.Size = New-Object System.Drawing.Size(800, 30); $tabPage3.Controls.Add($correlationResultLabel)
    $corrPrintButton = New-Object System.Windows.Forms.Button; $corrPrintButton.Text = "Drucken"; $corrPrintButton.Size = New-Object System.Drawing.Size(90, 24); $corrPrintButton.Location = New-Object System.Drawing.Point(830, 113); $tabPage3.Controls.Add($corrPrintButton)
    $corrExportButton = New-Object System.Windows.Forms.Button; $corrExportButton.Text = "Exportieren..."; $corrExportButton.Size = New-Object System.Drawing.Size(110, 24); $corrExportButton.Location = New-Object System.Drawing.Point(($corrPrintButton.Right + 8), 113); $corrExportButton.Enabled = $false; $tabPage3.Controls.Add($corrExportButton)
    $correlationChart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart; $correlationChart.Location = New-Object System.Drawing.Point(20, 150); $correlationChart.Size = New-Object System.Drawing.Size(1020, 570); $correlationChart.Anchor = "Top, Left"; $tabPage3.Controls.Add($correlationChart)
    $corrChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea; $correlationChart.ChartAreas.Add($corrChartArea); $correlationChart.Legends.Clear()
    $corrSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series; $corrSeries.Name = "Correlation"; $corrSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Point; $corrSeries.MarkerStyle = "Circle"; $corrSeries.MarkerSize = 10; $correlationChart.Series.Add($corrSeries)

    # ---------- TAB 4: DATEN NACH BLUTTESTS (ROBUSTES LAYOUT) ----------
    $buttonPanel = New-Object System.Windows.Forms.Panel
    $buttonPanel.Height = 40
    $buttonPanel.Dock = [System.Windows.Forms.DockStyle]::Bottom
    $tabPage4.Controls.Add($buttonPanel)

    $deleteTestButton = New-Object System.Windows.Forms.Button
    $deleteTestButton.Location = New-Object System.Drawing.Point(830, 5)
    $deleteTestButton.Size = New-Object System.Drawing.Size(240, 30)
    $deleteTestButton.Text = "Tag löschen"
    $buttonPanel.Controls.Add($deleteTestButton)

    $saveChangesButton = New-Object System.Windows.Forms.Button
    $saveChangesButton.Location = New-Object System.Drawing.Point(320, 5)
    $saveChangesButton.Size = New-Object System.Drawing.Size(240, 30)
    $saveChangesButton.Text = "Änderungen speichern"
    $buttonPanel.Controls.Add($saveChangesButton)
    
    $deleteSelectedRowsButton = New-Object System.Windows.Forms.Button
    $deleteSelectedRowsButton.Location = New-Object System.Drawing.Point(575, 5)
    $deleteSelectedRowsButton.Size = New-Object System.Drawing.Size(240, 30)
    $deleteSelectedRowsButton.Text = "Ausgewählte Einträge löschen"
    $buttonPanel.Controls.Add($deleteSelectedRowsButton)

    $dataMgmtTopPanel = New-Object System.Windows.Forms.Panel; $dataMgmtTopPanel.Height = 75; $dataMgmtTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $dataMgmtSettingsButton = New-Object System.Windows.Forms.Button; $dataMgmtSettingsButton.Text = "Einstellungen"; $dataMgmtSettingsButton.Size = New-Object System.Drawing.Size(100, 24); $dataMgmtSettingsButton.Location = New-Object System.Drawing.Point(10, 8)
    $dataMgmtTopPanel.Controls.Add($dataMgmtSettingsButton)
    $dataMgmtPrintButton = New-Object System.Windows.Forms.Button; $dataMgmtPrintButton.Text = "Drucken"; $dataMgmtPrintButton.Size = New-Object System.Drawing.Size(80, 24); $dataMgmtPrintButton.Location = New-Object System.Drawing.Point(118, 8)
    $dataMgmtTopPanel.Controls.Add($dataMgmtPrintButton)
    $dataMgmtExportButton = New-Object System.Windows.Forms.Button; $dataMgmtExportButton.Text = "Exportieren..."; $dataMgmtExportButton.Size = New-Object System.Drawing.Size(110, 24); $dataMgmtExportButton.Location = New-Object System.Drawing.Point(($dataMgmtPrintButton.Right + 8), 8); $dataMgmtExportButton.Enabled = $false
    $dataMgmtTopPanel.Controls.Add($dataMgmtExportButton)
    # v2.17.0: Button "Dokument hochladen"
    $dataMgmtUploadButton = New-Object System.Windows.Forms.Button; $dataMgmtUploadButton.Text = "Dokument hochladen"; $dataMgmtUploadButton.Size = New-Object System.Drawing.Size(150, 24); $dataMgmtUploadButton.Location = New-Object System.Drawing.Point(($dataMgmtExportButton.Right + 16), 8); $dataMgmtUploadButton.Enabled = $false
    $dataMgmtTopPanel.Controls.Add($dataMgmtUploadButton)
    # v2.19.0: PDF-Import Button (Beta Version)
    $pdfImportButton = New-Object System.Windows.Forms.Button
    $pdfImportButton.Text = "PDF importieren (Beta)"
    $pdfImportButton.Size = New-Object System.Drawing.Size(170, 24)
    $pdfImportButton.Location = New-Object System.Drawing.Point(($dataMgmtUploadButton.Right + 8), 8)
    $pdfImportButton.ForeColor = [System.Drawing.Color]::DarkOrange
    $dataMgmtTopPanel.Controls.Add($pdfImportButton)
    # v2.22.0: Button "Dokument anzeigen" (nur sichtbar wenn Dokument vorhanden)
    $showDocButton = New-Object System.Windows.Forms.Button
    $showDocButton.Text = "Dokument`nanzeigen"
    $showDocButton.Size = New-Object System.Drawing.Size(100, 65)
    $showDocButton.Location = New-Object System.Drawing.Point(($pdfImportButton.Right + 40), 5)
    $showDocButton.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Bold)
    $showDocButton.ForeColor = [System.Drawing.Color]::FromArgb(57, 255, 20)
    $showDocButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $showDocButton.Visible = $false
    $dataMgmtTopPanel.Controls.Add($showDocButton)
    # v2.24.1: Tooltip listet alle Dokumente des gewählten Bluttests
    $showDocToolTip = New-Object System.Windows.Forms.ToolTip
    $showDocToolTip.AutoPopDelay = 15000
    $showDocToolTip.InitialDelay = 400
    $showDocToolTip.ReshowDelay = 200

    # v2.19.0: Verschlüsselungs-Buttons (zweite Zeile)
    $encSeparator = New-Object System.Windows.Forms.Label
    $encSeparator.Text = "🔒"
    $encSeparator.Location = New-Object System.Drawing.Point(8, 42)
    $encSeparator.Size = New-Object System.Drawing.Size(20, 20)
    $dataMgmtTopPanel.Controls.Add($encSeparator)

    $encExportButton = New-Object System.Windows.Forms.Button
    $encExportButton.Text = "Verschlüsseltes Backup exportieren"
    $encExportButton.Size = New-Object System.Drawing.Size(230, 24)
    $encExportButton.Location = New-Object System.Drawing.Point(30, 40)
    $dataMgmtTopPanel.Controls.Add($encExportButton)

    $encImportButton = New-Object System.Windows.Forms.Button
    $encImportButton.Text = "Backup importieren"
    $encImportButton.Size = New-Object System.Drawing.Size(150, 24)
    $encImportButton.Location = New-Object System.Drawing.Point(268, 40)
    $dataMgmtTopPanel.Controls.Add($encImportButton)

    # v2.21.1: Unverschluesselter Migrations-Export (Klartext-ZIP)
    $plainExportButton = New-Object System.Windows.Forms.Button
    $plainExportButton.Text = "Unverschlüsselt exportieren (Migration)"
    $plainExportButton.Size = New-Object System.Drawing.Size(250, 24)
    $plainExportButton.Location = New-Object System.Drawing.Point(($encImportButton.Right + 16), 40)
    $plainExportButton.ForeColor = [System.Drawing.Color]::Firebrick
    $dataMgmtTopPanel.Controls.Add($plainExportButton)

    $tabPage4.Controls.Add($dataMgmtTopPanel)

    $dataMgmtTreeView = New-Object System.Windows.Forms.TreeView
    $dataMgmtTreeView.Location = New-Object System.Drawing.Point(10, 85)
    $dataMgmtTreeView.Anchor = 'Top, Bottom, Left'
    $dataMgmtTreeView.Size = New-Object System.Drawing.Size(300, 625)
    $tabPage4.Controls.Add($dataMgmtTreeView)

    $dataMgmtDataGridView = New-Object System.Windows.Forms.DataGridView
    $dataMgmtDataGridView.Location = New-Object System.Drawing.Point(320, 85)
    $dataMgmtDataGridView.Anchor = 'Top, Bottom, Left' 
    $dataMgmtDataGridView.Size = New-Object System.Drawing.Size(750, 625) 
    $dataMgmtDataGridView.AutoSizeColumnsMode = "Fill"
    $dataMgmtDataGridView.AllowUserToAddRows = $false
    $dataMgmtDataGridView.ReadOnly = $false
    $dataMgmtDataGridView.Columns.Add("Marker", "Marker")
    $dataMgmtDataGridView.Columns.Add("Value", "Wert")
    $dataMgmtDataGridView.Columns.Add("Unit", "Einheit")
    $dataMgmtDataGridView.Columns.Add("Note", "Notiz")
    $dataMgmtDataGridView.Columns["Marker"].FillWeight = 30
    $dataMgmtDataGridView.Columns["Note"].FillWeight = 40
    $dataMgmtDataGridView.Columns["Value"].FillWeight = 20
    $dataMgmtDataGridView.Columns["Unit"].FillWeight = 15
    $dataMgmtDataGridView.Columns["Marker"].ReadOnly = $true
    $dataMgmtDataGridView.Columns["Unit"].ReadOnly = $true
    $tabPage4.Controls.Add($dataMgmtDataGridView)

    # ---------- NEU: TAB 5 - PERSÖNLICHE METRIKEN (VOLLSTÄNDIG KORRIGIERT) ----------
    $personalMetricsGroup = New-Object System.Windows.Forms.GroupBox; $personalMetricsGroup.Text = "Persönliche & biometrische Daten"; $personalMetricsGroup.Location = New-Object System.Drawing.Point(20, 20); $personalMetricsGroup.Size = New-Object System.Drawing.Size(500, 455); $personalMetricsGroup.Anchor = 'Top, Left'; $tabPage6.Controls.Add($personalMetricsGroup)
    
    $controlsYPos = 30; $labelWidth = 200; $controlWidth = 250; $yStep = 35
    
    # --- Name ---
    $labelName = New-Object System.Windows.Forms.Label; $labelName.Text = "Vor- und Nachname:"; $labelName.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelName.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelName)
    $textName = New-Object System.Windows.Forms.TextBox; $textName.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textName.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textName); $controlsYPos += $yStep

    # --- Alter ---
    $labelAlter = New-Object System.Windows.Forms.Label; $labelAlter.Text = "Alter (Jahre):"; $labelAlter.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelAlter.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelAlter)
    $textAlter = New-Object System.Windows.Forms.TextBox; $textAlter.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textAlter.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textAlter); $controlsYPos += $yStep

    # --- Geschlecht ---
    $labelGeschlecht = New-Object System.Windows.Forms.Label; $labelGeschlecht.Text = "Geschlecht:"; $labelGeschlecht.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelGeschlecht.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelGeschlecht)
    $comboGeschlecht = New-Object System.Windows.Forms.ComboBox; $comboGeschlecht.Location = New-Object System.Drawing.Point(230, $controlsYPos); $comboGeschlecht.Size = New-Object System.Drawing.Size($controlWidth, 20); $comboGeschlecht.Items.AddRange(@("männlich", "weiblich")); $personalMetricsGroup.Controls.Add($comboGeschlecht); $controlsYPos += $yStep

    # --- Größe ---
    $labelGroesse = New-Object System.Windows.Forms.Label; $labelGroesse.Text = "Größe (cm):"; $labelGroesse.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelGroesse.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelGroesse)
    $textGroesse = New-Object System.Windows.Forms.TextBox; $textGroesse.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textGroesse.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textGroesse); $controlsYPos += $yStep

    # --- Gewicht ---
    $labelGewicht = New-Object System.Windows.Forms.Label; $labelGewicht.Text = "Aktuelles Gewicht (kg):"; $labelGewicht.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelGewicht.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelGewicht)
    $textGewicht = New-Object System.Windows.Forms.TextBox; $textGewicht.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textGewicht.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textGewicht); $controlsYPos += $yStep
    
    # --- KFA ---
    $labelKFA = New-Object System.Windows.Forms.Label; $labelKFA.Text = "Aktueller Körperfettanteil (%):"; $labelKFA.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelKFA.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelKFA)
    $textKFA = New-Object System.Windows.Forms.TextBox; $textKFA.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textKFA.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textKFA); $controlsYPos += $yStep
    
    # --- Grundumsatz ---
    $labelGrundumsatz = New-Object System.Windows.Forms.Label; $labelGrundumsatz.Text = "Kalorien-Grundumsatz (kcal):"; $labelGrundumsatz.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelGrundumsatz.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelGrundumsatz)
    $textGrundumsatz = New-Object System.Windows.Forms.TextBox; $textGrundumsatz.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textGrundumsatz.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textGrundumsatz); $controlsYPos += $yStep
    
    # --- Workout ---
    $labelWorkout = New-Object System.Windows.Forms.Label; $labelWorkout.Text = "Workout-Frequenz:"; $labelWorkout.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelWorkout.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelWorkout)
    $comboWorkout = New-Object System.Windows.Forms.ComboBox; $comboWorkout.Location = New-Object System.Drawing.Point(230, $controlsYPos); $comboWorkout.Size = New-Object System.Drawing.Size($controlWidth, 20); $comboWorkout.Items.AddRange(@("1-2x/Woche", "3-4x/Woche", "5-6x/Woche", "7+x/Woche")); $personalMetricsGroup.Controls.Add($comboWorkout); $controlsYPos += $yStep

    # --- Blutdruck ---
    $labelBlutdruck = New-Object System.Windows.Forms.Label; $labelBlutdruck.Text = "typ. Blutdruck (SYST):"; $labelBlutdruck.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelBlutdruck.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelBlutdruck)
    $textBlutdruck = New-Object System.Windows.Forms.TextBox; $textBlutdruck.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textBlutdruck.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textBlutdruck); $controlsYPos += $yStep

    # --- Raucher ---
    $checkRaucher = New-Object System.Windows.Forms.CheckBox; $checkRaucher.Text = "Raucher"; $checkRaucher.Location = New-Object System.Drawing.Point(230, $controlsYPos); $checkRaucher.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($checkRaucher); $controlsYPos += ($yStep + 20)
    
    # --- Speicherbutton ---
    $savePersonalButton = New-Object System.Windows.Forms.Button; $savePersonalButton.Text = "Persönliche Daten speichern"; $savePersonalButton.Location = New-Object System.Drawing.Point(230, $controlsYPos); $savePersonalButton.Size = New-Object System.Drawing.Size($controlWidth, 40); $personalMetricsGroup.Controls.Add($savePersonalButton)
    
    # --- Tooltips ---
    $toolTip.SetToolTip($textGroesse, "Ihre Körpergröße in Zentimetern. Wird zur Berechnung des BMI benötigt.")
    $toolTip.SetToolTip($textBlutdruck, "Geben Sie hier Ihren typischen systolischen Blutdruck (oberer Wert) in mmHg ein. Dieser Wert ist für die Risikoberechnung (PREVENT-Score) erforderlich.")

    # ---------- v2.17.0: GENETISCHE VORBELASTUNGEN (Tab Persönliche Metriken) ----------
    $geneticGroup = New-Object System.Windows.Forms.GroupBox
    $geneticGroup.Text = "Genetische Vorbelastungen"
    $geneticGroup.Location = New-Object System.Drawing.Point(540, 20)
    $geneticGroup.Size = New-Object System.Drawing.Size(600, 760)
    $geneticGroup.Anchor = 'Top, Left'
    $tabPage6.Controls.Add($geneticGroup)

    # --- Sub-GroupBox: Vorbelastungen berücksichtigen ---
    $geneticEnableGroup = New-Object System.Windows.Forms.GroupBox
    $geneticEnableGroup.Text = "Vorbelastungen berücksichtigen"
    $geneticEnableGroup.Location = New-Object System.Drawing.Point(15, 25)
    $geneticEnableGroup.Size = New-Object System.Drawing.Size(570, 360)
    $geneticGroup.Controls.Add($geneticEnableGroup)

    $geneticMasterCheck = New-Object System.Windows.Forms.CheckBox
    $geneticMasterCheck.Text = "Vorbelastungen berücksichtigen (aktiviert Hinweise im Blutbild)"
    $geneticMasterCheck.Location = New-Object System.Drawing.Point(15, 25)
    $geneticMasterCheck.Size = New-Object System.Drawing.Size(540, 20)
    $geneticMasterCheck.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $geneticEnableGroup.Controls.Add($geneticMasterCheck)

    $script:geneticBuiltInChecks = @()
    $gpY = 55
    $builtInLabels = @(
        "Chronische Veneninsuffizienz (Venenleiden)",
        "Familiäre Hypertonie / KHK-Prädisposition (Herz-Kreislauf-Probleme)",
        "Familiäre Hypercholesterinämie (Cholesterin)",
        "Typ-2-Diabetes-Mellitus-Prädisposition (Diabetes)"
    )
    foreach ($label in $builtInLabels) {
        $cb = New-Object System.Windows.Forms.CheckBox
        $cb.Text = $label
        $cb.Location = New-Object System.Drawing.Point(35, $gpY)
        $cb.Size = New-Object System.Drawing.Size(520, 20)
        $cb.Enabled = $false
        $cb.Tag = $label
        $geneticEnableGroup.Controls.Add($cb)
        $script:geneticBuiltInChecks += $cb
        $gpY += 30
    }

    # Bereich für benutzerdefinierte Vorbelastungen (dynamisch befüllt)
    $geneticCustomPanel = New-Object System.Windows.Forms.Panel
    $geneticCustomPanel.Location = New-Object System.Drawing.Point(15, ($gpY + 10))
    $geneticCustomPanel.Size = New-Object System.Drawing.Size(540, (360 - $gpY - 60))
    $geneticCustomPanel.AutoScroll = $true
    $geneticEnableGroup.Controls.Add($geneticCustomPanel)
    $script:geneticCustomChecks = @()

    # Buttons: Löschen + Bearbeiten (unten in der GroupBox)
    $geneticDeleteButton = New-Object System.Windows.Forms.Button
    $geneticDeleteButton.Text = "ausgewählte Vorbelastung löschen"
    $geneticDeleteButton.Location = New-Object System.Drawing.Point(15, 322)
    $geneticDeleteButton.Size = New-Object System.Drawing.Size(250, 28)
    $geneticEnableGroup.Controls.Add($geneticDeleteButton)

    $geneticEditButton = New-Object System.Windows.Forms.Button
    $geneticEditButton.Text = "Vorbelastung bearbeiten"
    $geneticEditButton.Location = New-Object System.Drawing.Point(275, 322)
    $geneticEditButton.Size = New-Object System.Drawing.Size(200, 28)
    $geneticEnableGroup.Controls.Add($geneticEditButton)

    # Master-Checkbox steuert Aktivierung der Unter-Checkboxen
    $geneticMasterCheck.Add_CheckedChanged({
        $enabled = $geneticMasterCheck.Checked
        foreach ($cb in $script:geneticBuiltInChecks) { $cb.Enabled = $enabled }
        foreach ($cb in $script:geneticCustomChecks) { $cb.Enabled = $enabled }
    })

    # --- Sub-GroupBox: Neue genetische Vorbelastung/en definieren ---
    $geneticDefineGroup = New-Object System.Windows.Forms.GroupBox
    $geneticDefineGroup.Text = "Neue genetische Vorbelastung/en definieren"
    $geneticDefineGroup.Location = New-Object System.Drawing.Point(15, 395)
    $geneticDefineGroup.Size = New-Object System.Drawing.Size(570, 350)
    $geneticGroup.Controls.Add($geneticDefineGroup)

    $geneticNewNameLabel = New-Object System.Windows.Forms.Label
    $geneticNewNameLabel.Text = "Name der Vorbelastung:"
    $geneticNewNameLabel.Location = New-Object System.Drawing.Point(15, 25)
    $geneticNewNameLabel.Size = New-Object System.Drawing.Size(160, 20)
    $geneticDefineGroup.Controls.Add($geneticNewNameLabel)

    $geneticNewNameText = New-Object System.Windows.Forms.TextBox
    $geneticNewNameText.Location = New-Object System.Drawing.Point(180, 22)
    $geneticNewNameText.Size = New-Object System.Drawing.Size(370, 20)
    $geneticDefineGroup.Controls.Add($geneticNewNameText)

    $geneticNewMarkerLabel = New-Object System.Windows.Forms.Label
    $geneticNewMarkerLabel.Text = "Zugehörige Blutmarker auswählen:"
    $geneticNewMarkerLabel.Location = New-Object System.Drawing.Point(15, 52)
    $geneticNewMarkerLabel.Size = New-Object System.Drawing.Size(250, 20)
    $geneticDefineGroup.Controls.Add($geneticNewMarkerLabel)

    $geneticNewMarkerList = New-Object System.Windows.Forms.CheckedListBox
    $geneticNewMarkerList.Location = New-Object System.Drawing.Point(15, 75)
    $geneticNewMarkerList.Size = New-Object System.Drawing.Size(535, 230)
    $geneticNewMarkerList.CheckOnClick = $true
    $geneticNewMarkerList.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $geneticDefineGroup.Controls.Add($geneticNewMarkerList)

    $geneticNewAddButton = New-Object System.Windows.Forms.Button
    $geneticNewAddButton.Text = "Vorbelastung hinzufügen"
    $geneticNewAddButton.Location = New-Object System.Drawing.Point(15, 313)
    $geneticNewAddButton.Size = New-Object System.Drawing.Size(200, 30)
    $geneticDefineGroup.Controls.Add($geneticNewAddButton)

    # ---------- NEU: TAB 6 - LONGEVITY-INDIZES ----------
    $longevityGrid = New-Object System.Windows.Forms.DataGridView
    $longevityGrid.Dock = [System.Windows.Forms.DockStyle]::Fill
    $longevityGrid.AllowUserToAddRows = $false; $longevityGrid.AllowUserToDeleteRows = $false; $longevityGrid.ReadOnly = $true
    $longevityGrid.AutoSizeColumnsMode = "Fill"
    $longevityGrid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $longevityGrid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $longevityGrid.RowTemplate.Height = 30
    $longevityGrid.Columns.Add("Marker", "Marker"); $longevityGrid.Columns.Add("Value", "Letzter Wert"); $longevityGrid.Columns.Add("Ideal", "Idealbereich"); $longevityGrid.Columns.Add("Score", "Score (0-100)"); $longevityGrid.Columns.Add("Interpretation", "Interpretation")
    $longevityGrid.Columns["Marker"].FillWeight = 25; $longevityGrid.Columns["Value"].FillWeight = 15; $longevityGrid.Columns["Ideal"].FillWeight = 15; $longevityGrid.Columns["Score"].FillWeight = 10; $longevityGrid.Columns["Interpretation"].FillWeight = 35
    $longevityTopPanel = New-Object System.Windows.Forms.Panel; $longevityTopPanel.Height = 40; $longevityTopPanel.Dock = [System.Windows.Forms.DockStyle]::Top
    $longevitySettingsButton = New-Object System.Windows.Forms.Button; $longevitySettingsButton.Text = "Einstellungen"; $longevitySettingsButton.Size = New-Object System.Drawing.Size(100, 24); $longevitySettingsButton.Location = New-Object System.Drawing.Point(10, 8)
    $longevityTopPanel.Controls.Add($longevitySettingsButton)
    $longevityFilterHintLabel = New-Object System.Windows.Forms.Label; $longevityFilterHintLabel.Location = New-Object System.Drawing.Point(120, 12); $longevityFilterHintLabel.Size = New-Object System.Drawing.Size(1020, 20); $longevityFilterHintLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $longevityFilterHintLabel.ForeColor = [System.Drawing.Color]::Red; $longevityFilterHintLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter; $longevityFilterHintLabel.Text = ""
    $longevityTopPanel.Controls.Add($longevityFilterHintLabel)
    $tabPage5.Controls.Add($longevityGrid)
    $tabPage5.Controls.Add($longevityTopPanel)

    # ---------- TAB 6: CUSTOM REPORT ----------
    $customHintLabel = New-Object System.Windows.Forms.Label
    $customHintLabel.Text = "Über die Custom-Report Funktion können beliebige Blutmarker zu einem beliebigen Zeitraum ausgewählt und als Report exportiert werden."
    $customHintLabel.Location = New-Object System.Drawing.Point(20, 15)
    $customHintLabel.Size = New-Object System.Drawing.Size(1120, 35)
    $customHintLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Italic)
    $customHintLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $tabPageCustom.Controls.Add($customHintLabel)

    # --- Links: Marker-Auswahl (GroupBox mit CheckedListBox) ---
    $customMarkerGroup = New-Object System.Windows.Forms.GroupBox
    $customMarkerGroup.Text = "Blutmarker auswählen"
    $customMarkerGroup.Location = New-Object System.Drawing.Point(20, 60)
    $customMarkerGroup.Size = New-Object System.Drawing.Size(560, 580)
    $tabPageCustom.Controls.Add($customMarkerGroup)

    $customMarkerSearchLabel = New-Object System.Windows.Forms.Label
    $customMarkerSearchLabel.Text = "Filter:"
    $customMarkerSearchLabel.Location = New-Object System.Drawing.Point(15, 25)
    $customMarkerSearchLabel.Size = New-Object System.Drawing.Size(50, 20)
    $customMarkerGroup.Controls.Add($customMarkerSearchLabel)

    $customMarkerSearchBox = New-Object System.Windows.Forms.TextBox
    $customMarkerSearchBox.Location = New-Object System.Drawing.Point(70, 22)
    $customMarkerSearchBox.Size = New-Object System.Drawing.Size(280, 20)
    $customMarkerGroup.Controls.Add($customMarkerSearchBox)

    $customSelectAllButton = New-Object System.Windows.Forms.Button
    $customSelectAllButton.Text = "Alle auswählen"
    $customSelectAllButton.Location = New-Object System.Drawing.Point(360, 20)
    $customSelectAllButton.Size = New-Object System.Drawing.Size(95, 24)
    $customMarkerGroup.Controls.Add($customSelectAllButton)

    $customDeselectAllButton = New-Object System.Windows.Forms.Button
    $customDeselectAllButton.Text = "Alle abwählen"
    $customDeselectAllButton.Location = New-Object System.Drawing.Point(460, 20)
    $customDeselectAllButton.Size = New-Object System.Drawing.Size(90, 24)
    $customMarkerGroup.Controls.Add($customDeselectAllButton)

    $customMarkerList = New-Object System.Windows.Forms.CheckedListBox
    $customMarkerList.Location = New-Object System.Drawing.Point(15, 55)
    $customMarkerList.Size = New-Object System.Drawing.Size(535, 475)
    $customMarkerList.CheckOnClick = $true
    $customMarkerList.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $customMarkerGroup.Controls.Add($customMarkerList)

    $customMarkerStatsLabel = New-Object System.Windows.Forms.Label
    $customMarkerStatsLabel.Location = New-Object System.Drawing.Point(15, 538)
    $customMarkerStatsLabel.Size = New-Object System.Drawing.Size(535, 30)
    $customMarkerStatsLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8, [System.Drawing.FontStyle]::Italic)
    $customMarkerStatsLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $customMarkerStatsLabel.Text = "Ausgewählt: 0 Marker – Deaktivierte Marker haben keine Datensätze."
    $customMarkerGroup.Controls.Add($customMarkerStatsLabel)

    # --- Rechts: Zeitraum (orientiert an Show-DataMgmtSettingsPopup) ---
    $customTimeGroup = New-Object System.Windows.Forms.GroupBox
    $customTimeGroup.Text = "Zeitraum"
    $customTimeGroup.Location = New-Object System.Drawing.Point(600, 60)
    $customTimeGroup.Size = New-Object System.Drawing.Size(540, 580)
    $tabPageCustom.Controls.Add($customTimeGroup)

    $customTimeLabel = New-Object System.Windows.Forms.Label
    $customTimeLabel.Text = "Zeitraum auswählen:"
    $customTimeLabel.Location = New-Object System.Drawing.Point(20, 35)
    $customTimeLabel.Size = New-Object System.Drawing.Size(140, 20)
    $customTimeGroup.Controls.Add($customTimeLabel)

    $customTimeCombo = New-Object System.Windows.Forms.ComboBox
    $customTimeCombo.Location = New-Object System.Drawing.Point(170, 32)
    $customTimeCombo.Size = New-Object System.Drawing.Size(340, 20)
    $customTimeCombo.DropDownStyle = "DropDownList"
    $customTimeCombo.Items.AddRange(@("Aktuelles Jahr", "Letzte 3 Jahre", "Letzte 5 Jahre", "Letzte 10 Jahre", "Alle Daten"))
    $customTimeCombo.SelectedItem = $script:customReportTimeFilter
    $customTimeGroup.Controls.Add($customTimeCombo)

    # v2.17.0: Vorbelastungs-Report Schnellauswahl
    $customGeneticLabel = New-Object System.Windows.Forms.Label
    $customGeneticLabel.Text = "Vorbelastungs-Report:"
    $customGeneticLabel.Location = New-Object System.Drawing.Point(20, 62)
    $customGeneticLabel.Size = New-Object System.Drawing.Size(140, 20)
    $customTimeGroup.Controls.Add($customGeneticLabel)

    $customGeneticCombo = New-Object System.Windows.Forms.ComboBox
    $customGeneticCombo.Location = New-Object System.Drawing.Point(170, 59)
    $customGeneticCombo.Size = New-Object System.Drawing.Size(260, 20)
    $customGeneticCombo.DropDownStyle = "DropDownList"
    $customGeneticCombo.Items.Add("– Keine Vorbelastung –")
    $customGeneticCombo.SelectedIndex = 0
    $customTimeGroup.Controls.Add($customGeneticCombo)

    $customGeneticApplyButton = New-Object System.Windows.Forms.Button
    $customGeneticApplyButton.Text = "übernehmen"
    $customGeneticApplyButton.Location = New-Object System.Drawing.Point(435, 57)
    $customGeneticApplyButton.Size = New-Object System.Drawing.Size(90, 24)
    $customTimeGroup.Controls.Add($customGeneticApplyButton)

    # Vorschau-Zusammenfassung
    $customPreviewLabel = New-Object System.Windows.Forms.Label
    $customPreviewLabel.Text = "Report-Vorschau"
    $customPreviewLabel.Location = New-Object System.Drawing.Point(20, 92)
    $customPreviewLabel.Size = New-Object System.Drawing.Size(500, 20)
    $customPreviewLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $customTimeGroup.Controls.Add($customPreviewLabel)

    $customPreviewGrid = New-Object System.Windows.Forms.DataGridView
    $customPreviewGrid.Location = New-Object System.Drawing.Point(20, 115)
    $customPreviewGrid.Size = New-Object System.Drawing.Size(500, 445)
    $customPreviewGrid.AllowUserToAddRows = $false
    $customPreviewGrid.AllowUserToDeleteRows = $false
    $customPreviewGrid.ReadOnly = $true
    $customPreviewGrid.AutoSizeColumnsMode = "Fill"
    $customPreviewGrid.RowHeadersVisible = $false
    $customPreviewGrid.SelectionMode = "FullRowSelect"
    $customPreviewGrid.Columns.Add("Datum", "Datum")
    $customPreviewGrid.Columns.Add("Marker", "Marker")
    $customPreviewGrid.Columns.Add("Wert", "Wert")
    $customPreviewGrid.Columns.Add("Einheit", "Einheit")
    $customPreviewGrid.Columns["Datum"].FillWeight = 18
    $customPreviewGrid.Columns["Marker"].FillWeight = 45
    $customPreviewGrid.Columns["Wert"].FillWeight = 15
    $customPreviewGrid.Columns["Einheit"].FillWeight = 15
    $customTimeGroup.Controls.Add($customPreviewGrid)

    # --- Unten: Drucken + Exportieren ---
    $customPrintButton = New-Object System.Windows.Forms.Button
    $customPrintButton.Text = "Drucken"
    $customPrintButton.Location = New-Object System.Drawing.Point(20, 660)
    $customPrintButton.Size = New-Object System.Drawing.Size(120, 32)
    $customPrintButton.Enabled = $false
    $tabPageCustom.Controls.Add($customPrintButton)

    $customExportButton = New-Object System.Windows.Forms.Button
    $customExportButton.Text = "Exportieren..."
    $customExportButton.Location = New-Object System.Drawing.Point(150, 660)
    $customExportButton.Size = New-Object System.Drawing.Size(140, 32)
    $customExportButton.Enabled = $false
    $tabPageCustom.Controls.Add($customExportButton)

    $customStatusLabel = New-Object System.Windows.Forms.Label
    $customStatusLabel.Text = "Bitte mindestens einen Marker auswählen, um Drucken/Exportieren zu aktivieren."
    $customStatusLabel.Location = New-Object System.Drawing.Point(310, 667)
    $customStatusLabel.Size = New-Object System.Drawing.Size(830, 20)
    $customStatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
    $customStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
    $tabPageCustom.Controls.Add($customStatusLabel)


    # ---------- Funktionen für Logik und Events ----------
    function Update-DataManagementTab {
        $dataMgmtTreeView.Nodes.Clear()
        $dataMgmtDataGridView.Rows.Clear()

        # Zeitfilter anwenden
        $cutoffDate = $null
        $now = Get-Date
        switch ($script:dataMgmtTimeFilter) {
            "Aktuelles Jahr"   { $cutoffDate = [datetime]::new($now.Year, 1, 1) }
            "Letzte 3 Jahre"   { $cutoffDate = $now.AddYears(-3) }
            "Letzte 5 Jahre"   { $cutoffDate = $now.AddYears(-5) }
            "Letzte 10 Jahre"  { $cutoffDate = $now.AddYears(-10) }
            default            { $cutoffDate = $null }
        }

        $filteredItems = $script:allHistoricalItems
        if ($cutoffDate) {
            $filteredItems = @($script:allHistoricalItems | Where-Object { [datetime]$_.Date -ge $cutoffDate })
        }

        $groupedData = $filteredItems | Group-Object -Property Date | Sort-Object @{Expression={[datetime]$_.Name}} -Descending
        foreach ($group in $groupedData) {
            $date = [datetime]$group.Name
            $nodeText = "$($date.ToString('dd. MMMM yyyy')) ($($group.Count) Marker)"
            $node = New-Object System.Windows.Forms.TreeNode($nodeText)
            $node.Tag = $group.Name
            $dataMgmtTreeView.Nodes.Add($node) | Out-Null
        }
    }
    function Get-PairedData {
        param($marker1Name, $marker2Name)
        $marker1Data = if($script:data.Config.CalculatedMarkers.Name -contains $marker1Name) { Get-CalculatedValuesForMarker -markerName $marker1Name } else { $script:allHistoricalItems | Where-Object { $_.Name -eq $marker1Name } }
        $marker2Data = if($script:data.Config.CalculatedMarkers.Name -contains $marker2Name) { Get-CalculatedValuesForMarker -markerName $marker2Name } else { $script:allHistoricalItems | Where-Object { $_.Name -eq $marker2Name } }
        $marker1DataDict = @{}
        $marker1Data | ForEach-Object { $marker1DataDict[$_.Date] = $_.Value }
        $pairedData = New-Object System.Collections.ArrayList
        $marker2Data | ForEach-Object {
            if ($marker1DataDict.ContainsKey($_.Date)) {
                $null = $pairedData.Add([PSCustomObject]@{ Value1 = $marker1DataDict[$_.Date]; Value2 = $_.Value })
            }
        }
        return $pairedData
    }
    function Show-CorrelationSuggestionsPopup {
        $popup = New-Object System.Windows.Forms.Form; $popup.Text = "Vorschläge für relevante Korrelationen"; $popup.Size = New-Object System.Drawing.Size(800, 600); $popup.StartPosition = "CenterParent"
        
        $grid = New-Object System.Windows.Forms.DataGridView; $grid.Dock = "Fill"; $grid.AllowUserToAddRows = $false; $grid.ReadOnly = $true; $grid.SelectionMode = "FullRowSelect"; $grid.MultiSelect = $false
        $grid.Columns.Add("Marker1", "Marker 1"); $grid.Columns.Add("Marker2", "Marker 2"); $grid.Columns.Add("Correlation", "Korrelation (r)"); $grid.Columns.Add("DataPoints", "Datenpunkte"); $grid.Columns.Add("Interpretation", "Interpretation")
        $grid.Columns["Marker1"].FillWeight = 30; $grid.Columns["Marker2"].FillWeight = 30; $grid.Columns["Correlation"].FillWeight = 15; $grid.Columns["DataPoints"].FillWeight = 10; $grid.Columns["Interpretation"].FillWeight = 15
        
        $okButton = New-Object System.Windows.Forms.Button; $okButton.Text = "Auswahl übernehmen"; $okButton.Dock = "Bottom"; $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $popup.Controls.Add($grid); $popup.Controls.Add($okButton)
        $popup.AcceptButton = $okButton

        $popup.Add_Shown({
            $popup.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            $healthGroups = @('Herz-Kreislauf', 'Venen', 'Blutfette', 'Stoffwechsel', 'Entzündung', 'Hormone')
            $lifestyleGroup = @('Lebensstil')
            $allConfigMarkers = $script:data.Config.Markers + $script:data.Config.CalculatedMarkers
            $healthMarkers = $allConfigMarkers | Where-Object { $healthGroups -contains $_.Group } | Select-Object -ExpandProperty Name -Unique
            $lifestyleMarkers = $allConfigMarkers | Where-Object { $lifestyleGroup -contains $_.Group } | Select-Object -ExpandProperty Name -Unique
            
            $correlationResults = New-Object System.Collections.ArrayList
            
            foreach ($lifestyleMarker in $lifestyleMarkers) {
                foreach ($healthMarker in $healthMarkers) {
                    $pairedData = Get-PairedData -marker1Name $lifestyleMarker -marker2Name $healthMarker
                    if ($pairedData.Count -gt 3) {
                        $correlation = Calculate-Correlation -values1 $pairedData.Value1 -values2 $pairedData.Value2
                        $absCorr = [Math]::Abs($correlation)
                        if ($absCorr -ge 0.3) {
                           $interpretation = ""
                            if ($absCorr -ge 0.7) { $interpretation = "Sehr starker " }
                            elseif ($absCorr -ge 0.5) { $interpretation = "Starker " }
                            else { $interpretation = "Moderater " }
                            if($correlation -ne 0) {
                                if ($correlation -gt 0) { $interpretation += "positiver Zusammenhang" } else { $interpretation += "negativer Zusammenhang" }
                            } else { $interpretation = "Kein Zusammenhang"}
                            $null = $correlationResults.Add([PSCustomObject]@{ Marker1 = $lifestyleMarker; Marker2 = $healthMarker; Correlation = $correlation; DataPoints = $pairedData.Count; Interpretation = $interpretation; AbsCorr = $absCorr })
                        }
                    }
                }
            }
            
            $sortedResults = $correlationResults | Sort-Object AbsCorr -Descending
            foreach ($res in $sortedResults) {
                $grid.Rows.Add($res.Marker1, $res.Marker2, [Math]::Round($res.Correlation, 4), $res.DataPoints, $res.Interpretation)
            }
            $popup.Cursor = [System.Windows.Forms.Cursors]::Default
        })
        
        if ($popup.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK -and $grid.SelectedRows.Count -gt 0) {
            $selectedRow = $grid.SelectedRows[0]
            return @{ Marker1 = $selectedRow.Cells["Marker1"].Value; Marker2 = $selectedRow.Cells["Marker2"].Value }
        }
        return $null
    }
    function Update-CorrelationDropdowns {
        $corrMarker1ComboBox.Items.Clear(); $corrMarker2ComboBox.Items.Clear()
        $allMarkerNames = ($script:data.Config.Markers.Name + $script:data.Config.CalculatedMarkers.Name) | Sort-Object -Unique
        foreach ($name in $allMarkerNames) {
            $corrMarker1ComboBox.Items.Add($name)
            $corrMarker2ComboBox.Items.Add($name)
        }
    }
    function Update-CorrelationAnalysis {
        $marker1Name = $corrMarker1ComboBox.SelectedItem
        $marker2Name = $corrMarker2ComboBox.SelectedItem
        $corrSeries.Points.Clear()
        $corrChartArea.AxisX.Title = ""
        $corrChartArea.AxisY.Title = ""
        
        if ([string]::IsNullOrEmpty($marker1Name) -or [string]::IsNullOrEmpty($marker2Name)) {
            $correlationResultLabel.Text = "Bitte zwei Marker auswählen und Analyse starten."
            $corrExportButton.Enabled = $false
            return
        }
        $corrChartArea.AxisX.Title = $marker1Name
        $corrChartArea.AxisY.Title = $marker2Name
        
        $pairedData = Get-PairedData -marker1Name $marker1Name -marker2Name $marker2Name

        if ($pairedData.Count -lt 2) {
            $correlationResultLabel.Text = "Nicht genügend gemeinsame Datenpunkte ($($pairedData.Count)) für eine Analyse."
            $corrExportButton.Enabled = $false
            return
        }
        foreach ($pair in $pairedData) {
            $corrSeries.Points.AddXY($pair.Value1, $pair.Value2)
        }
        $correlation = Calculate-Correlation -values1 $pairedData.Value1 -values2 $pairedData.Value2
        $correlationResultLabel.Text = "Analyse mit $($pairedData.Count) Datenpunkten. Pearson-Korrelationskoeffizient: $([Math]::Round($correlation, 4))"
        $corrExportButton.Enabled = $true
    }
    function Update-AllMarkerDropdowns {
        # v2.23.0: baut zusaetzlich den Alias-Suchindex neu auf, damit auch
        # selbst angelegte Marker sofort per Direkteingabe gefunden werden.
        $script:markerFilterBusy = $true
        Build-MarkerSearchIndex
        $currentMarkers = [System.Collections.ArrayList]@($script:data.Config.Markers | ForEach-Object { [PSCustomObject]$_ })
        $script:markerFullList = @($currentMarkers | Sort-Object Name | ForEach-Object { [string]$_.Name })
        $markerComboBox.Items.Clear(); $deleteMarkerComboBox.Items.Clear()
        foreach ($n in $script:markerFullList) { [void]$markerComboBox.Items.Add($n); [void]$deleteMarkerComboBox.Items.Add($n) }
        $markerComboBox.BackColor = [System.Drawing.SystemColors]::Window
        if ($markerComboBox.Items.Count -gt 0) { $markerComboBox.SelectedIndex = 0 }
        if ($deleteMarkerComboBox.Items.Count -gt 0) { $deleteMarkerComboBox.SelectedIndex = 0 }
        $script:markerFilterBusy = $false
        Update-MarkerInputMode -MarkerName ([string]$markerComboBox.SelectedItem)
    }
    function Update-FilterDropdown {
        $currentSelection = $filterComboBox.SelectedItem
        $filterComboBox.Items.Clear(); $filterComboBox.Items.Add("--- Bitte auswählen ---")
        $uniqueNames = $script:allHistoricalItems.Name | Sort-Object -Unique
        foreach ($name in $uniqueNames) { $filterComboBox.Items.Add($name) }
        $currentCalculatedMarkers = [System.Collections.ArrayList]@($script:data.Config.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ })
        foreach ($calcMarker in $currentCalculatedMarkers) {
            $required = $calcMarker.RequiredMarkers
            $requiredMarkerNames = [System.Collections.ArrayList]@($required)
            $datesWithAllRequired = $script:allHistoricalItems | Where-Object { $requiredMarkerNames -contains $_.Name } | Group-Object Date | Where-Object { 
                $markerNamesOnDate = [System.Collections.ArrayList]@($_.Group.Name | Select-Object -Unique)
                $foundCount = ($requiredMarkerNames | ForEach-Object { $markerName = $_; $markerNamesOnDate -contains $markerName } | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
                $requiredMarkerNames.Count -gt 0 -and $foundCount -eq $requiredMarkerNames.Count
            }
            
            if ($datesWithAllRequired.Count -gt 0 -and -not $filterComboBox.Items.Contains($calcMarker.Name)) {
                $filterComboBox.Items.Add($calcMarker.Name)
            }
        }
        if ($currentSelection -and $filterComboBox.Items.Contains($currentSelection)) {
            $filterComboBox.SelectedItem = $currentSelection
        } else {
            $filterComboBox.SelectedIndex = 0
        }
    }
    
    function Update-CockpitFilterHint {
        $parts = New-Object System.Collections.ArrayList
        if ($script:cockpitTimeFilter -ne "Alle Daten") { $parts.Add("Zeitraum: $($script:cockpitTimeFilter)") | Out-Null }
        if ($script:cockpitGroupFilter -eq "Custom" -and $script:cockpitGroupFilterCustom.Count -gt 0) {
            $parts.Add("Bereich: $($script:cockpitGroupFilterCustom -join ', ')") | Out-Null
        } elseif ($script:cockpitGroupFilter -ne "Alle" -and $script:cockpitGroupFilter -ne "Custom") {
            $parts.Add("Bereich: $($script:cockpitGroupFilter)") | Out-Null
        }
        if ($script:cockpitRatingFilter -ne "Alle") { $parts.Add("Bewertung: $($script:cockpitRatingFilter)") | Out-Null }
        if ($parts.Count -gt 0) {
            $cockpitFilterHintLabel.Text = "Filter aktiv: $($parts -join ' | ')"
        } else {
            $cockpitFilterHintLabel.Text = ""
        }
    }
    function Update-Dashboard {
        $cockpitGrid.Rows.Clear()
        $allMarkers = @([System.Collections.ArrayList]@($script:data.Config.Markers | ForEach-Object { [PSCustomObject]$_ }) + [System.Collections.ArrayList]@($script:data.Config.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ }))
        $allMarkers = $allMarkers | Sort-Object Group, Name

        # v2.17.0: Aktive genetische Vorbelastungs-Marker ermitteln
        $geneticMarkerHints = Get-AllActiveGeneticMarkers

        # Gruppenfilter anwenden
        if ($script:cockpitGroupFilter -eq "Custom" -and $script:cockpitGroupFilterCustom.Count -gt 0) {
            $allMarkers = @($allMarkers | Where-Object { $script:cockpitGroupFilterCustom -contains $_.Group })
        } elseif ($script:cockpitGroupFilter -ne "Alle" -and $script:cockpitGroupFilter -ne "Custom") {
            $allMarkers = @($allMarkers | Where-Object { $_.Group -eq $script:cockpitGroupFilter })
        }

        # Zeitfilter anwenden
        $cutoffDate = $null
        $now = Get-Date
        switch ($script:cockpitTimeFilter) {
            "Aktuelles Jahr"   { $cutoffDate = [datetime]::new($now.Year, 1, 1) }
            "Letzte 3 Jahre"   { $cutoffDate = $now.AddYears(-3) }
            "Letzte 5 Jahre"   { $cutoffDate = $now.AddYears(-5) }
            "Letzte 10 Jahre"  { $cutoffDate = $now.AddYears(-10) }
            default            { $cutoffDate = $null }
        }

        $filteredItems = $script:allHistoricalItems
        if ($cutoffDate) {
            $filteredItems = @($script:allHistoricalItems | Where-Object { [datetime]$_.Date -ge $cutoffDate })
        }

        $preventScoreData = Get-CalculatedValuesForMarker -markerName "PREVENT-ASCVD-10Y"
        if ($cutoffDate) { $preventScoreData = @($preventScoreData | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
        $preventScoreData = $preventScoreData | Sort-Object Date -Descending | Select-Object -First 1

        foreach ($marker in $allMarkers) {
            $latestItem = $null
            $isCalculated = $marker.PSObject.Properties.Name -contains "RequiredMarkers"
            
            if ($isCalculated) {
                if ($marker.Name -eq "PREVENT-ASCVD-10Y") { $latestItem = $preventScoreData }
                else {
                    $calculatedData = Get-CalculatedValuesForMarker -markerName $marker.Name
                    if ($cutoffDate) { $calculatedData = @($calculatedData | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
                    if ($calculatedData.Count -gt 0) {
                        $latestItem = $calculatedData | Sort-Object Date -Descending | Select-Object -First 1
                    }
                }
            } else {
                $latestItem = $filteredItems | Where-Object { $_.Name -eq $marker.Name } | Sort-Object Date -Descending | Select-Object -First 1
            }

            if ($latestItem) {
                # Bewertung vorab berechnen
                $val = $latestItem.Value
                $color = [System.Drawing.Color]::White
                $sortValue = 4
                $ratingName = "Keine Bewertung"
                if ($marker.OptimalMin -ne $null -and $marker.OptimalMax -ne $null -and $val -ge $marker.OptimalMin -and $val -le $marker.OptimalMax) {
                    $color = [System.Drawing.Color]::LightGreen; $sortValue = 3; $ratingName = "Optimal"
                } elseif ($marker.RefMin -ne $null -and $marker.RefMax -ne $null -and $val -ge $marker.RefMin -and $val -le $marker.RefMax) {
                    $color = [System.Drawing.Color]::LightYellow; $sortValue = 2; $ratingName = "Akzeptabel"
                } elseif (($marker.RefMin -ne $null -and $val -lt $marker.RefMin) -or ($marker.RefMax -ne $null -and $val -gt $marker.RefMax)) {
                    $color = [System.Drawing.Color]::LightCoral; $sortValue = 1; $ratingName = "Außerhalb der Norm"
                }

                # Bewertungsfilter anwenden
                if ($script:cockpitRatingFilter -ne "Alle" -and $ratingName -ne $script:cockpitRatingFilter) { continue }

                $rowIndex = $cockpitGrid.Rows.Add()
                $row = $cockpitGrid.Rows[$rowIndex]
                $row.Cells["Group"].Value = $marker.Group
                $row.Cells["Marker"].Value = $marker.Name
                # v2.20.0: HIV als Text, Testosteron mit doppelter Einheit
                if ($marker.Name -eq "HIV (Anti-HIV-1/2)") {
                    $row.Cells["Value"].Value = if ([double]$latestItem.Value -ge 1) { "reaktiv" } else { "nicht reaktiv" }
                } elseif ($marker.Name -eq $script:ApoeMarkerName) {
                    # v2.26.0: Genotyp statt Zahlencode
                    $row.Cells["Value"].Value = Get-ApoeGenotypeText -Value $latestItem.Value
                } elseif ($marker.Name -eq "Testosteron, gesamt") {
                    $ngdl = [double]$latestItem.Value
                    $nmol = [Math]::Round($ngdl * 0.0347, 2)
                    $row.Cells["Value"].Value = "$ngdl ($nmol nmol/l)"
                } else {
                    $row.Cells["Value"].Value = $latestItem.Value
                }
                $row.Cells["Date"].Value = ([datetime]$latestItem.Date).ToString("dd.MM.yyyy")
                $row.Cells["Unit"].Value = $marker.Unit
                
                if ($marker.Name -eq "PREVENT-ASCVD-10Y" -and $latestItem) {
                    $row.Cells["Risk"].Value = $latestItem.Value
                    $riskVal = $latestItem.Value
                    $riskColor = [System.Drawing.Color]::White
                    if ($riskVal -lt 5) { $riskColor = [System.Drawing.Color]::LightGreen }
                    elseif ($riskVal -lt 7.5) { $riskColor = [System.Drawing.Color]::LightYellow }
                    elseif ($riskVal -lt 20) { $riskColor = [System.Drawing.Color]::Orange }
                    else { $riskColor = [System.Drawing.Color]::LightCoral }
                    $row.Cells["Risk"].Style.BackColor = $riskColor
                } else {
                    $row.Cells["Risk"].Value = "-"
                }

                $row.Cells["Rating"].Style.BackColor = $color
                $row.Cells["SortOrder"].Value = $sortValue

                # v2.17.0: Genetische Vorbelastungs-Hinweise
                if ($geneticMarkerHints.ContainsKey($marker.Name)) {
                    $hints = $geneticMarkerHints[$marker.Name] -join "`n"
                    $row.Cells["Marker"].ToolTipText = $hints
                    $row.Cells["Marker"].Value = "⚠ " + $marker.Name
                }
            }
        }

        $cockpitGrid.Sort($cockpitGrid.Columns["SortOrder"], [System.ComponentModel.ListSortDirection]::Ascending)

        for ($i = 1; $i -lt $cockpitGrid.Rows.Count; $i++) {
            if ($cockpitGrid.Rows[$i].Cells["Group"].Value -eq $cockpitGrid.Rows[$i - 1].Cells["Group"].Value) {
                if ($cockpitGrid.Rows[$i - 1].Cells["Group"].Value -ne "") {
                    $cockpitGrid.Rows[$i].Cells["Group"].Value = ""
                }
            }
        }
        Update-CockpitFilterHint
    }

    function Update-LongevityTab {
        $longevityGrid.Rows.Clear()

        # Zeitfilter berechnen
        $cutoffDate = $null
        $now = Get-Date
        switch ($script:longevityTimeFilter) {
            "Aktuelles Jahr"   { $cutoffDate = [datetime]::new($now.Year, 1, 1) }
            "Letzte 3 Jahre"   { $cutoffDate = $now.AddYears(-3) }
            "Letzte 5 Jahre"   { $cutoffDate = $now.AddYears(-5) }
            "Letzte 10 Jahre"  { $cutoffDate = $now.AddYears(-10) }
            default            { $cutoffDate = $null }
        }

        $filteredItems = $script:allHistoricalItems
        if ($cutoffDate) {
            $filteredItems = @($script:allHistoricalItems | Where-Object { [datetime]$_.Date -ge $cutoffDate })
        }

        $keyMarkers = @(
            @{ Name = "Langzeitzucker (HbA1c)"; Ideal = "< 5.2 %"; Interpretation = "Indikator für den durchschnittlichen Blutzucker der letzten 3 Monate." },
            @{ Name = "Nüchterninsulin"; Ideal = "< 5 µU/ml"; Interpretation = "Zeigt die basale Insulinproduktion und Sensitivität an." },
            @{ Name = "Vitamin D (25-OH)"; Ideal = "40-60 ng/ml"; Interpretation = "Wichtig für Immunsystem, Knochengesundheit und Zellregulation." },
            @{ Name = "Homocystein"; Ideal = "< 7 µmol/l"; Interpretation = "Ein erhöhter Wert kann auf ein erhöhtes Risiko für Herz-Kreislauf-Erkrankungen hindeuten." },
            @{ Name = "C-reaktives Protein (CRP)"; Ideal = "< 1 mg/l"; Interpretation = "Ein Marker für systemische Entzündungen im Körper." },
            @{ Name = "Hochsensitives CRP (hs-CRP)"; Ideal = "< 1 mg/l"; Interpretation = "Stille Entzündung: < 1 = niedriges, 1-3 = mittleres, > 3 mg/l = hohes kardiovaskuläres Risiko (AHA/CDC, Pearson 2003). > 10 mg/l = akuter Infekt, Kontrolle nach 2 Wochen." },
            @{ Name = "NAD+"; Ideal = "> 30 µM"; Interpretation = "Ein entscheidendes Coenzym für den zellulären Energiestoffwechsel und die DNA-Reparatur." },
            @{ Name = "Albumin"; Ideal = "> 4 g/dl"; Interpretation = "Wichtigstes Bluteiweiß, Indikator für Ernährungsstatus und Leberfunktion." },
            @{ Name = "Longevity-Score"; Ideal = "> 80 Punkte"; Interpretation = "Aggregierter Score, der die Gesamt-Performance der Longevity-Marker zusammenfasst." },
            @{ Name = "PhenoAge (biologisch)"; Ideal = "< chronol. Alter"; Interpretation = "Biologisches Alter (Levine 2018). Vergleich mit deinem Kalenderalter zeigt epigenetisches Altern." },
            @{ Name = "PhenoAge-Accel"; Ideal = "< 0 Jahre"; Interpretation = "Alters-Akzeleration. Negativ = jünger als Kalenderalter (sehr gut). Positiv = beschleunigtes Altern." },
            @{ Name = "InflammAging-Score"; Ideal = "> 70 Punkte"; Interpretation = "Systemische Entzündung (hsCRP + NLR + PLR). Chronische niedriggradige Inflammation treibt biologisches Altern." },
            @{ Name = "Neutrophile/Lymphozyten-Ratio (NLR)"; Ideal = "1.0 - 2.0"; Interpretation = "NLR > 3 assoziiert mit erhöhter Gesamtmortalität und kardiovaskulärem Risiko (Zahorec 2021)." },
            @{ Name = "TyG-Index"; Ideal = "< 4.68"; Interpretation = "Surrogatmarker für Insulinresistenz. Werte > 4.68 prädizieren Typ-2-Diabetes (Simental-Mendia 2008)." },
            @{ Name = "Triglyceride/HDL-Ratio"; Ideal = "< 2.0"; Interpretation = "Metabolischer Gesundheitsmarker. Hohe Ratio = small-dense LDL, Insulinresistenz, CVD-Risiko." }
        )
        
        $longevityScoreData = Get-CalculatedValuesForMarker -markerName "Longevity-Score"
        if ($cutoffDate) { $longevityScoreData = @($longevityScoreData | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
        $longevityScoreLatest = $longevityScoreData | Sort-Object Date -Descending | Select-Object -First 1
        
        $longevityMarkerConfig = [System.Collections.ArrayList]@($script:data.Config.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ }) | Where-Object { $_.Name -eq "Longevity-Score" } | Select-Object -First 1

        # Neu (v2.10.0): Liste der Marker, die eigenständige CalculatedMarkers sind
        # und nicht aus $filteredItems geholt werden können, sondern via Get-CalculatedValuesForMarker
        $calculatedOnlyMarkers = @(
            "Longevity-Score",
            "PhenoAge (biologisch)",
            "PhenoAge-Accel",
            "InflammAging-Score",
            "Neutrophile/Lymphozyten-Ratio (NLR)",
            "TyG-Index",
            "Triglyceride/HDL-Ratio"
        )

        foreach ($markerInfo in $keyMarkers) {
            $markerName = $markerInfo.Name
            $latestItem = $null
            if ($markerName -eq "Longevity-Score") {
                $latestItem = $longevityScoreLatest
            } elseif ($calculatedOnlyMarkers -contains $markerName) {
                # Berechnete Werte für diesen Marker holen
                $calcData = Get-CalculatedValuesForMarker -markerName $markerName
                if ($cutoffDate) { $calcData = @($calcData | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
                $latestItem = $calcData | Sort-Object Date -Descending | Select-Object -First 1
            } else {
                $latestItem = $filteredItems | Where-Object { $_.Name -eq $markerName } | Sort-Object Date -Descending | Select-Object -First 1
            }

            if ($latestItem) {
                $score = "-"
                if ($markerName -eq "Longevity-Score") { 
                    $score = [Math]::Round($latestItem.Value, 0)
                } elseif ($calculatedOnlyMarkers -contains $markerName) {
                    # Eigene Score-Berechnung für neue Longevity-Marker (0-100 Skala)
                    $score = Convert-ToLongevityScore -markerName $markerName -value $latestItem.Value
                } else {
                    $requiredMarkers = $longevityMarkerConfig.RequiredMarkers
                    $index = [Array]::IndexOf($requiredMarkers, $markerName)
                    # v2.25.0: Ersatz-Marker auf den Slot des Original-Markers abbilden
                    # (hs-CRP wird wie CRP bewertet - gleiche Einheit mg/l)
                    if ($index -eq -1) {
                        foreach ($origName in $script:MarkerFallbacks.Keys) {
                            if ($script:MarkerFallbacks[$origName] -contains $markerName) {
                                $index = [Array]::IndexOf($requiredMarkers, $origName)
                                if ($index -ne -1) { break }
                            }
                        }
                    }

                    if ($index -ne -1) {
                        # Array korrekt mit $null-Werten initialisieren
                        $valuesForSingleScore = @($null) * $requiredMarkers.Count
                        $valuesForSingleScore[$index] = $latestItem.Value
                        
                        $scoreValue = Calculate-LongevityScore -values $valuesForSingleScore
                        $score = [Math]::Round($scoreValue, 0)
                    }
                }
                
                $rowIndex = $longevityGrid.Rows.Add($markerName, "$($latestItem.Value) $($latestItem.Unit)", $markerInfo.Ideal, $score, $markerInfo.Interpretation)
                $row = $longevityGrid.Rows[$rowIndex]
                $scoreCell = $row.Cells["Score"]
                $interpretationCell = $row.Cells["Interpretation"]
                
                $tooltipText = ""
                if ($score -is [double] -or $score -is [int]) {
                    if ($score -ge 80) { 
                        $scoreCell.Style.BackColor = [System.Drawing.Color]::LightGreen
                        $tooltipText = "OPTIMAL: Dein Wert ist hervorragend. Behalte deine aktuelle Ernährung und deinen Lebensstil bei, um diesen Status zu halten."
                    } elseif ($score -ge 50) { 
                        $scoreCell.Style.BackColor = [System.Drawing.Color]::LightYellow
                        $tooltipText = "AKZEPTABEL: Dein Wert ist im soliden Mittelfeld. Es gibt jedoch Optimierungspotenzial. Prüfe, ob gezielte Anpassungen in Ernährung oder Lebensstil möglich sind."
                    } else { 
                        $scoreCell.Style.BackColor = [System.Drawing.Color]::LightCoral
                        $tooltipText = "OPTIMIERUNGSBEDARF: Dein Wert liegt deutlich außerhalb des optimalen Bereichs. Eine genauere Analyse und gezielte Maßnahmen (z.B. Ernährungsumstellung, Supplementierung nach ärztlicher Rücksprache) sind empfehlenswert."
                    }
                }
                switch ($markerName) {
                    "Langzeitzucker (HbA1c)" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Reduziere einfache Kohlenhydrate und Zucker. Priorisiere ballaststoffreiche Lebensmittel und regelmäßige Bewegung." } }
                    "Nüchterninsulin" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Ähnlich wie bei HbA1c, zusätzlich kann Intervallfasten hilfreich sein." } }
                    "Vitamin D (25-OH)" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Mehr Sonnenexposition (Mittagszeit) oder Supplementierung mit Vitamin D3/K2 nach ärztlicher Absprache." } }
                    "Homocystein" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Stelle eine ausreichende Versorgung mit B-Vitaminen (B6, B12, Folsäure) sicher, z.B. durch grünes Blattgemüse und hochwertige tierische Produkte." } }
                    "C-reaktives Protein (CRP)" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Konzentriere dich auf eine entzündungshemmende Ernährung (Omega-3-Fettsäuren, Antioxidantien) und ausreichend Schlaf." } }
                    "NAD+" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Hochintensives Intervalltraining (HIIT), Kalorienrestriktion und potenzielle Vorstufen wie NMN oder NR können den Spiegel erhöhen." } }
                    "Albumin" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Achte auf eine ausreichende Proteinzufuhr (ca. 1.6-2.0g/kg Körpergewicht bei deinem Trainingspensum) und die allgemeine Lebergesundheit." } }
                    "PhenoAge (biologisch)" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Fokus auf die 4 Haupttreiber: (1) Inflammation senken (hsCRP, Omega-3), (2) Glukosestoffwechsel optimieren (Nüchternglukose), (3) Nierenfunktion schützen (Hydration, Kreatinin), (4) Anti-oxidatives Training (HIIT + Krafttraining)." } }
                    "PhenoAge-Accel" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Accel > 0 bedeutet beschleunigtes Altern. Priorisiere: Schlaf (7-9h), Krafttraining (deine 5-6x/Woche ist optimal), proteinreiche Ernährung, Stressmanagement." } }
                    "InflammAging-Score" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Chronische Entzündung adressieren: Omega-3 (EPA+DHA >2g/d), Mittelmeer-Diät, Schlafhygiene, ggf. Kurkuma/Curcumin. hsCRP, NLR und Thrombozyten einzeln prüfen." } }
                    "Neutrophile/Lymphozyten-Ratio (NLR)" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: NLR spiegelt das Gleichgewicht zwischen akuter (Neutrophile) und adaptiver (Lymphozyten) Immunantwort. Chronischer Stress, Infektionen, schlechter Schlaf erhöhen NLR." } }
                    "TyG-Index" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Insulinresistenz angehen: (1) Kohlenhydrate reduzieren, (2) Krafttraining für Muskelmasse (Glukose-Senke), (3) Bauchfett reduzieren, (4) ggf. Metformin-Evaluation durch Arzt." } }
                    "Triglyceride/HDL-Ratio" { if($score -lt 80) { $tooltipText += "`nEmpfehlung: Hohe Ratio = Zeichen für small-dense LDL. Maßnahmen: Reduzierter Zucker/Fruktose, mehr Omega-3, Intervallfasten, weniger Alkohol." } }
                }
                $scoreCell.ToolTipText = $tooltipText
                $interpretationCell.ToolTipText = $tooltipText
            }
        }

        # Filter-Hint aktualisieren
        if ($script:longevityTimeFilter -ne "Alle Daten") {
            $longevityFilterHintLabel.Text = "Filter aktiv: Zeitraum: $($script:longevityTimeFilter)"
        } else {
            $longevityFilterHintLabel.Text = ""
        }
    }

    # ---------- Custom Report (v2.12.0) ----------
    # $script:customReportAvailableMarkers enthält die Liste der Marker, die aktuell
    # Datensätze im gewählten Zeitraum haben. Disabled-Marker bleiben in der Liste,
    # lassen sich aber nicht anhaken (ItemCheck-Handler blockt).
    $script:customReportAvailableMarkers = @()
    $script:customReportAllMarkers = @()

    function Update-CustomReportTab {
        # 1. Zeitraum-Cutoff bestimmen
        $cutoffDate = $null
        $now = Get-Date
        switch ($script:customReportTimeFilter) {
            "Aktuelles Jahr"   { $cutoffDate = [datetime]::new($now.Year, 1, 1) }
            "Letzte 3 Jahre"   { $cutoffDate = $now.AddYears(-3) }
            "Letzte 5 Jahre"   { $cutoffDate = $now.AddYears(-5) }
            "Letzte 10 Jahre"  { $cutoffDate = $now.AddYears(-10) }
            default            { $cutoffDate = $null }
        }

        # 2. Welche Marker haben Daten im Zeitraum? (Rohdaten + berechnete Marker)
        $rawItems = $script:allHistoricalItems
        if ($cutoffDate) {
            $rawItems = @($rawItems | Where-Object { [datetime]$_.Date -ge $cutoffDate })
        }
        $markersWithRawData = $rawItems.Name | Sort-Object -Unique

        $markersWithCalcData = New-Object System.Collections.ArrayList
        foreach ($calcMarker in $script:data.Config.CalculatedMarkers) {
            $calcData = Get-CalculatedValuesForMarker -markerName $calcMarker.Name
            if ($cutoffDate) { $calcData = @($calcData | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
            if (($calcData | Measure-Object).Count -gt 0) {
                [void]$markersWithCalcData.Add($calcMarker.Name)
            }
        }
        $script:customReportAvailableMarkers = @($markersWithRawData) + @($markersWithCalcData) | Sort-Object -Unique

        # 3. Alle Marker (Rohe + Berechnete) für die Anzeige
        $allNames = @($script:data.Config.Markers.Name) + @($script:data.Config.CalculatedMarkers.Name) | Sort-Object -Unique
        $script:customReportAllMarkers = $allNames

        # 4. CheckedListBox befüllen (vorher aktuelle Auswahl merken)
        $previouslyChecked = @()
        for ($i = 0; $i -lt $customMarkerList.Items.Count; $i++) {
            if ($customMarkerList.GetItemChecked($i)) {
                $previouslyChecked += $customMarkerList.Items[$i]
            }
        }

        $filterText = $customMarkerSearchBox.Text.Trim().ToLower()

        $customMarkerList.BeginUpdate()
        $customMarkerList.Items.Clear()
        foreach ($name in $allNames) {
            if ($filterText -and (-not $name.ToLower().Contains($filterText))) { continue }
            $hasData = $script:customReportAvailableMarkers -contains $name
            $display = if ($hasData) { $name } else { "$name  [keine Daten]" }
            $idx = $customMarkerList.Items.Add($display)
            # Auswahl wiederherstellen
            if (($previouslyChecked -contains $display) -or ($previouslyChecked -contains $name)) {
                if ($hasData) { $customMarkerList.SetItemChecked($idx, $true) }
            }
        }
        $customMarkerList.EndUpdate()

        # v2.17.0: Vorbelastungs-Combo aktualisieren
        $currentSelection = $customGeneticCombo.SelectedItem
        $customGeneticCombo.Items.Clear()
        $customGeneticCombo.Items.Add("– Keine Vorbelastung –")
        $allPreds = Get-AllGeneticPredispositions
        foreach ($pred in $allPreds) {
            if ($pred['Active']) { $customGeneticCombo.Items.Add($pred['Name']) }
        }
        if ($currentSelection -and $customGeneticCombo.Items.Contains($currentSelection)) {
            $customGeneticCombo.SelectedItem = $currentSelection
        } else {
            $customGeneticCombo.SelectedIndex = 0
        }

        Update-CustomReportPreview
        Update-CustomReportButtonState
    }

    function Get-CustomReportSelectedMarkers {
        $selected = New-Object System.Collections.ArrayList
        for ($i = 0; $i -lt $customMarkerList.Items.Count; $i++) {
            if ($customMarkerList.GetItemChecked($i)) {
                # Original-Namen aus Display extrahieren (Suffix "  [keine Daten]" entfernen)
                $raw = [string]$customMarkerList.Items[$i] -replace '\s+\[keine Daten\]$', ''
                # Disabled-Items dürfen technisch nicht ankreuzbar sein - Sicherheitscheck
                if ($script:customReportAvailableMarkers -contains $raw) {
                    [void]$selected.Add($raw)
                }
            }
        }
        return $selected
    }

    function Get-CustomReportData {
        # Liefert alle Datensätze der ausgewählten Marker im gewählten Zeitraum
        $selectedMarkers = Get-CustomReportSelectedMarkers
        if ($selectedMarkers.Count -eq 0) { return @() }

        $cutoffDate = $null
        $now = Get-Date
        switch ($script:customReportTimeFilter) {
            "Aktuelles Jahr"  { $cutoffDate = [datetime]::new($now.Year, 1, 1) }
            "Letzte 3 Jahre"  { $cutoffDate = $now.AddYears(-3) }
            "Letzte 5 Jahre"  { $cutoffDate = $now.AddYears(-5) }
            "Letzte 10 Jahre" { $cutoffDate = $now.AddYears(-10) }
            default           { $cutoffDate = $null }
        }

        $calcMarkerNames = @($script:data.Config.CalculatedMarkers.Name)
        $result = New-Object System.Collections.ArrayList

        foreach ($markerName in $selectedMarkers) {
            $items = $null
            if ($calcMarkerNames -contains $markerName) {
                $items = Get-CalculatedValuesForMarker -markerName $markerName
            } else {
                $items = $script:allHistoricalItems | Where-Object { $_.Name -eq $markerName }
            }
            if ($cutoffDate) {
                $items = @($items | Where-Object { [datetime]$_.Date -ge $cutoffDate })
            }
            foreach ($it in $items) {
                # v2.26.0: APOE als Genotyp-Klartext (Vorschau, CSV/XLSX-Export)
                $wertAnzeige = $it.Value
                $einheitAnzeige = $it.Unit
                if ($it.Name -eq $script:ApoeMarkerName) {
                    $wertAnzeige = Get-ApoeGenotypeText -Value $it.Value
                    $einheitAnzeige = ""
                } elseif ($it.Name -eq "HIV (Anti-HIV-1/2)") {
                    $wertAnzeige = if ([double]$it.Value -ge 1) { "reaktiv" } else { "nicht reaktiv" }
                    $einheitAnzeige = ""
                }
                [void]$result.Add([PSCustomObject]@{
                    Datum   = $it.Date
                    Marker  = $it.Name
                    Wert    = $wertAnzeige
                    Einheit = $einheitAnzeige
                    Notiz   = if ($it.PSObject.Properties.Name -contains "Note") { [string]$it.Note } else { "" }
                })
            }
        }
        return @($result | Sort-Object @{ Expression = "Datum"; Descending = $true }, Marker)
    }

    function Update-CustomReportPreview {
        $customPreviewGrid.Rows.Clear()
        $data = Get-CustomReportData
        foreach ($row in $data) {
            # v2.13.0: Datumsformat "yyyy-MM-dd" → "dd.MM.yyyy" für Anzeige (interner Speicher bleibt ISO)
            $dateDisplay = $row.Datum
            try {
                $parsed = [datetime]::ParseExact($row.Datum, "yyyy-MM-dd", $null)
                $dateDisplay = $parsed.ToString("dd.MM.yyyy")
            } catch { <# Fallback: Original anzeigen #> }
            [void]$customPreviewGrid.Rows.Add($dateDisplay, $row.Marker, $row.Wert, $row.Einheit)
        }
        $countSel = (Get-CustomReportSelectedMarkers).Count
        $countRows = ($data | Measure-Object).Count
        $customMarkerStatsLabel.Text = "Ausgewählt: $countSel Marker – Datensätze im Zeitraum: $countRows"
    }

    function Update-CustomReportButtonState {
        $hasSel = (Get-CustomReportSelectedMarkers).Count -gt 0
        $customPrintButton.Enabled = $hasSel
        $customExportButton.Enabled = $hasSel
        if ($hasSel) {
            $customStatusLabel.Text = "Bereit. Report kann gedruckt oder exportiert werden."
            $customStatusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
        } else {
            $customStatusLabel.Text = "Bitte mindestens einen Marker auswählen, um Drucken/Exportieren zu aktivieren."
            $customStatusLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray
        }
    }

    function Get-CustomReportMarkerSections {
        # Liefert pro ausgewähltem Marker ein Section-Objekt mit Items, MarkerConfig
        $selected = Get-CustomReportSelectedMarkers
        $sections = New-Object System.Collections.ArrayList
        $cutoffDate = $null
        $now = Get-Date
        switch ($script:customReportTimeFilter) {
            "Aktuelles Jahr"  { $cutoffDate = [datetime]::new($now.Year, 1, 1) }
            "Letzte 3 Jahre"  { $cutoffDate = $now.AddYears(-3) }
            "Letzte 5 Jahre"  { $cutoffDate = $now.AddYears(-5) }
            "Letzte 10 Jahre" { $cutoffDate = $now.AddYears(-10) }
            default           { $cutoffDate = $null }
        }
        $calcNames = @($script:data.Config.CalculatedMarkers.Name)
        foreach ($markerName in $selected) {
            $items = $null
            $cfg = $null
            if ($calcNames -contains $markerName) {
                $items = Get-CalculatedValuesForMarker -markerName $markerName
                $cfg = $script:data.Config.CalculatedMarkers | Where-Object { $_.Name -eq $markerName } | Select-Object -First 1
            } else {
                $items = $script:allHistoricalItems | Where-Object { $_.Name -eq $markerName }
                $cfg = $script:data.Config.Markers | Where-Object { $_.Name -eq $markerName } | Select-Object -First 1
            }
            if ($cutoffDate) { $items = @($items | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
            $items = @($items | Sort-Object { [datetime]::ParseExact($_.Date, "yyyy-MM-dd", $null) })
            if ($items.Count -eq 0) { continue }

            [void]$sections.Add([PSCustomObject]@{
                MarkerName = $markerName
                Unit       = if ($cfg) { $cfg.Unit } else { "" }
                RefMin     = if ($cfg -and $cfg.PSObject.Properties.Name -contains "RefMin") { $cfg.RefMin } else { $null }
                RefMax     = if ($cfg -and $cfg.PSObject.Properties.Name -contains "RefMax") { $cfg.RefMax } else { $null }
                OptMin     = if ($cfg -and $cfg.PSObject.Properties.Name -contains "OptimalMin") { $cfg.OptimalMin } else { $null }
                OptMax     = if ($cfg -and $cfg.PSObject.Properties.Name -contains "OptimalMax") { $cfg.OptimalMax } else { $null }
                Items      = $items
            })
        }
        return $sections
    }

    function Invoke-CustomReportPrint {
        <#
        .SYNOPSIS
            Druckt bzw. erstellt PDF des Custom Reports.
            Pro Marker: Überschrift + Liniengrafik + Tabelle der Werte.
        .PARAMETER TargetPdfPath
            Wenn gesetzt: PDF-Export-Modus (nutzt Microsoft Print to PDF).
            Wenn $null: normaler Druck mit PrintDialog.
        #>
        param([string]$TargetPdfPath)

        $sections = Get-CustomReportMarkerSections
        if (($sections | Measure-Object).Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Erstellen des Reports vorhanden.", "Report", "OK", "Information")
            return $false
        }

        Add-Type -AssemblyName System.Drawing

        # Script-Scope-State für asynchrone Print-Page-Handler
        $script:reportSections         = @($sections)
        $script:reportSectionIndex     = 0   # welche Section wird gerade gedruckt
        $script:reportSectionRowIndex  = 0   # nächste Tabellenzeile innerhalb der Section
        $script:reportSectionPhase     = 0   # 0 = Heading+Chart, 1 = Tabellen-Header, 2 = Rows
        $script:reportPageNum          = 0
        $script:reportSelectedMarkers  = @(Get-CustomReportSelectedMarkers)
        $script:reportTimeFilter       = $script:customReportTimeFilter
        $script:reportChartBmp         = $null

        $printDoc = New-Object System.Drawing.Printing.PrintDocument
        $printDoc.DefaultPageSettings.Landscape = $true
        $printDoc.DocumentName = "Custom Report"

        # v2.26.0 (Aufgabe #3): Papierformat fest auf DIN A4 setzen, sofern der
        # gewaehlte Drucker es anbietet. Ohne diese Zuweisung gilt das
        # Standardformat des Druckers (z. B. Letter) - der erzwungene
        # Seitenumbruch waere dann nicht auf A4 bezogen.
        # Querformat bleibt bewusst aktiv: die Verlaufsgrafik nutzt die volle
        # Seitenbreite. Ergebnis = DIN A4 quer (297 x 210 mm).
        $applyA4Paper = {
            param($doc)
            try {
                $a4 = $doc.PrinterSettings.PaperSizes | Where-Object { $_.Kind -eq [System.Drawing.Printing.PaperKind]::A4 } | Select-Object -First 1
                if ($a4) { $doc.DefaultPageSettings.PaperSize = $a4 }
            } catch { }
        }

        if ($TargetPdfPath) {
            # PDF-Drucker finden
            $installed = [System.Drawing.Printing.PrinterSettings]::InstalledPrinters
            $pdfPrinterName = $null
            foreach ($p in $installed) {
                if ($p -match "Microsoft Print to PDF|Microsoft PDF|PDF|Foxit|Adobe PDF") {
                    $pdfPrinterName = $p
                    if ($p -eq "Microsoft Print to PDF") { break }
                }
            }
            if (-not $pdfPrinterName) {
                [System.Windows.Forms.MessageBox]::Show("Kein PDF-Drucker gefunden. Aktiviere 'Microsoft Print to PDF'.", "PDF-Drucker fehlt", "OK", "Warning")
                return $false
            }
            $printDoc.PrinterSettings.PrinterName = $pdfPrinterName
            $printDoc.PrinterSettings.PrintToFile = $true
            $printDoc.PrinterSettings.PrintFileName = $TargetPdfPath
            # v2.26.0: A4 erst NACH der Druckerzuweisung setzen - PaperSizes
            # haengt am konkreten Drucker.
            & $applyA4Paper $printDoc
        }

        $printDoc.Add_PrintPage({
            param($sender, $ev)
            $g = $ev.Graphics
            $margins = $ev.MarginBounds

            $titleFont    = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
            $sectionFont  = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
            $headerFont   = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $cellFont     = New-Object System.Drawing.Font("Segoe UI", 9)
            $smallFont    = New-Object System.Drawing.Font("Segoe UI", 8)
            $pen          = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 0.5)
            $yPos         = $margins.Top
            $rowHeight    = 22

            # v2.26.0 (BUGFIX): Fusszeile auf JEDER Seite. Bisher stand der
            # DrawString-Aufruf am Ende des Handlers - jeder Seitenumbruch
            # verlaesst den Handler aber vorher per "return", die Fusszeile
            # erschien deshalb nur auf der letzten Seite.
            $drawFooter = {
                param($gfx, $fnt, $mgn, $pageNo)
                $gfx.DrawString("Seite $pageNo – Blood-Tracker Custom Report", $fnt, [System.Drawing.Brushes]::DarkSlateGray, $mgn.Left, ($mgn.Top + $mgn.Height - 15))
            }

            # Seitenkopf nur auf Seite 1
            if ($script:reportPageNum -eq 0) {
                $g.DrawString("Custom Report", $titleFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                $yPos += 28
                $selStr = ($script:reportSelectedMarkers) -join ", "
                if ($selStr.Length -gt 200) { $selStr = $selStr.Substring(0, 197) + "..." }
                $g.DrawString("Zeitraum: $($script:reportTimeFilter)", $smallFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, $yPos)
                $yPos += 14
                $g.DrawString("Marker: $selStr", $smallFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, $yPos)
                $yPos += 14
                $g.DrawString("Druckdatum: $(Get-Date -Format 'dd.MM.yyyy HH:mm')", $smallFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, $yPos)
                $yPos += 22
            }
            $script:reportPageNum++

            $bottomLimit = $margins.Top + $margins.Height - 25

            # Sections durchlaufen
            while ($script:reportSectionIndex -lt $script:reportSections.Count) {
                $sec   = $script:reportSections[$script:reportSectionIndex]
                $items = @($sec.Items)

                # Phase 0: Überschrift + Grafik
                if ($script:reportSectionPhase -eq 0) {
                    $headingHeight = 26
                    $chartTargetWidth  = $margins.Width
                    $chartTargetHeight = 280
                    $neededSpace = $headingHeight + $chartTargetHeight + 10

                    # Passt die Überschrift + Chart noch auf die Seite?
                    if (($yPos + $headingHeight + 40) -gt $bottomLimit) {
                        & $drawFooter $g $smallFont $margins $script:reportPageNum
                        $ev.HasMorePages = $true
                        return
                    }

                    # Überschrift
                    $g.DrawString("$($sec.MarkerName)  ($($sec.Unit))", $sectionFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += $headingHeight

                    # Chart rendern (nur 1x pro Section)
                    if (-not $script:reportChartBmp) {
                        $script:reportChartBmp = Render-MarkerChartToImage -items $items `
                            -markerName $sec.MarkerName -unit $sec.Unit `
                            -refMin $sec.RefMin -refMax $sec.RefMax `
                            -optMin $sec.OptMin -optMax $sec.OptMax `
                            -width 1600 -height 560
                    }

                    $chartAvail = $bottomLimit - $yPos - 15
                    if ($chartAvail -lt 100) {
                        # zu wenig Platz → Seitenumbruch, Phase 0 wiederholen
                        & $drawFooter $g $smallFont $margins $script:reportPageNum
                        $ev.HasMorePages = $true
                        return
                    }
                    $chartH = [Math]::Min($chartTargetHeight, $chartAvail)
                    $g.DrawImage($script:reportChartBmp, $margins.Left, $yPos, $chartTargetWidth, $chartH)
                    $yPos += $chartH + 10

                    $script:reportSectionPhase = 1
                }

                # Phase 1: Tabellen-Header
                if ($script:reportSectionPhase -eq 1) {
                    if (($yPos + $rowHeight * 2) -gt $bottomLimit) {
                        & $drawFooter $g $smallFont $margins $script:reportPageNum
                        $ev.HasMorePages = $true
                        return
                    }
                    $props   = @("Datum", "Marker", "Wert", "Einheit", "Notiz")
                    $weights = @(0.15, 0.35, 0.10, 0.10, 0.30)
                    $xPos = $margins.Left
                    for ($i = 0; $i -lt $props.Count; $i++) {
                        $w = [int]($margins.Width * $weights[$i])
                        $g.FillRectangle([System.Drawing.Brushes]::LightGray, $xPos, $yPos, $w, $rowHeight)
                        $g.DrawRectangle($pen, $xPos, $yPos, $w, $rowHeight)
                        $g.DrawString($props[$i], $headerFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 4))
                        $xPos += $w
                    }
                    $yPos += $rowHeight
                    $script:reportSectionPhase = 2
                }

                # Phase 2: Tabellen-Zeilen
                if ($script:reportSectionPhase -eq 2) {
                    $weights = @(0.15, 0.35, 0.10, 0.10, 0.30)
                    while ($script:reportSectionRowIndex -lt $items.Count) {
                        if (($yPos + $rowHeight) -gt $bottomLimit) {
                            & $drawFooter $g $smallFont $margins $script:reportPageNum
                            $ev.HasMorePages = $true
                            return
                        }
                        $it = $items[$script:reportSectionRowIndex]
                        $dateDisp = $it.Date
                        try { $dateDisp = ([datetime]::ParseExact($it.Date, "yyyy-MM-dd", $null)).ToString("dd.MM.yyyy") } catch { }
                        $notiz = if ($it.PSObject.Properties.Name -contains "Note") { [string]$it.Note } else { "" }
                        # v2.26.0: qualitative Marker im PDF als Klartext ausgeben
                        $wertPdf    = [string]$it.Value
                        $einheitPdf = [string]$it.Unit
                        if ($it.Name -eq $script:ApoeMarkerName) {
                            $wertPdf = Get-ApoeGenotypeText -Value $it.Value
                            $einheitPdf = ""
                        } elseif ($it.Name -eq "HIV (Anti-HIV-1/2)") {
                            $wertPdf = if ([double]$it.Value -ge 1) { "reaktiv" } else { "nicht reaktiv" }
                            $einheitPdf = ""
                        }
                        $vals = @([string]$dateDisp, [string]$it.Name, $wertPdf, $einheitPdf, $notiz)
                        $xPos = $margins.Left
                        for ($i = 0; $i -lt 5; $i++) {
                            $w = [int]($margins.Width * $weights[$i])
                            $val = $vals[$i]
                            if ($val.Length -gt 80) { $val = $val.Substring(0, 77) + "..." }
                            $g.DrawRectangle($pen, $xPos, $yPos, $w, $rowHeight)
                            $g.DrawString($val, $cellFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 4))
                            $xPos += $w
                        }
                        $yPos += $rowHeight
                        $script:reportSectionRowIndex++
                    }
                    # Section fertig, Aufräumen
                    if ($script:reportChartBmp) {
                        try { $script:reportChartBmp.Dispose() } catch { }
                        $script:reportChartBmp = $null
                    }
                    $script:reportSectionIndex++
                    $script:reportSectionPhase    = 0
                    $script:reportSectionRowIndex = 0
                    $yPos += 18  # Abstand zwischen Sections

                    # v2.26.0 (Aufgabe #3): Jeder Blutmarker beginnt zwingend auf
                    # einer NEUEN Seite. Ohne diesen Umbruch wurde die naechste
                    # Marker-Section direkt unter der vorherigen Tabelle
                    # weitergedruckt, sobald noch Platz war.
                    if ($script:reportSectionIndex -lt $script:reportSections.Count) {
                        & $drawFooter $g $smallFont $margins $script:reportPageNum
                        $ev.HasMorePages = $true
                        return
                    }
                }
            }

            # Footer der letzten Seite
            & $drawFooter $g $smallFont $margins $script:reportPageNum
            $ev.HasMorePages = $false
        })

        if ($TargetPdfPath) {
            $printDoc.Print()
            return $true
        } else {
            $printDialog = New-Object System.Windows.Forms.PrintDialog
            $printDialog.Document = $printDoc
            if ($printDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $printDoc.Print()
                return $true
            }
            return $false
        }
    }

    function Load-PersonalMetrics {
        $personal = $script:data.Config.Personal
        $textName.Text = if ($personal.Name) { $personal.Name } else { "" }
        $textAlter.Text = $personal.Age
        $comboGeschlecht.SelectedItem = $personal.Geschlecht
        $textGroesse.Text = $personal.Groesse
        $textGewicht.Text = $personal.Gewicht
        $textKFA.Text = $personal.KFA
        $textGrundumsatz.Text = $personal.Grundumsatz
        $comboWorkout.SelectedItem = $personal.WorkoutFreq
        $textBlutdruck.Text = $personal.SystBlutdruck
        $checkRaucher.Checked = $personal.Raucher

        # v2.17.0: Genetische Vorbelastungen laden
        $gp = $script:data.Config.GeneticPredispositions
        $geneticMasterCheck.Checked = $gp.Enabled
        for ($i = 0; $i -lt $script:geneticBuiltInChecks.Count; $i++) {
            $cb = $script:geneticBuiltInChecks[$i]
            $cb.Enabled = $gp.Enabled
            $match = $gp.BuiltIn | Where-Object { $_['Name'] -eq $cb.Tag }
            if ($match) { $cb.Checked = $match['Active'] }
        }
        # Benutzerdefinierte Vorbelastungen in Panel laden
        Update-GeneticCustomPanel
        # Marker-CheckedListBox für "Neue Vorbelastung" befüllen
        Update-GeneticNewMarkerList
    }
    function Save-PersonalMetrics {
        try {
            $script:data.Config.Personal.Name = $textName.Text.Trim()
            $age = [int]$textAlter.Text
            if ($age -lt 18) { throw "Das Alter muss mindestens 18 sein." }
            $script:data.Config.Personal.Age = $age
            $script:data.Config.Personal.Geschlecht = $comboGeschlecht.SelectedItem
            $script:data.Config.Personal.Groesse = Parse-Number $textGroesse.Text
            $script:data.Config.Personal.Gewicht = Parse-Number $textGewicht.Text
            $script:data.Config.Personal.KFA = Parse-Number $textKFA.Text
            $script:data.Config.Personal.Grundumsatz = Parse-Number $textGrundumsatz.Text
            $script:data.Config.Personal.WorkoutFreq = $comboWorkout.SelectedItem
            $script:data.Config.Personal.SystBlutdruck = Parse-Number $textBlutdruck.Text
            $script:data.Config.Personal.Raucher = $checkRaucher.Checked

            # v2.17.0: Genetische Vorbelastungen speichern
            $script:data.Config.GeneticPredispositions['Enabled'] = $geneticMasterCheck.Checked
            for ($i = 0; $i -lt $script:geneticBuiltInChecks.Count; $i++) {
                $cb = $script:geneticBuiltInChecks[$i]
                $match = $script:data.Config.GeneticPredispositions['BuiltIn'] | Where-Object { $_['Name'] -eq $cb.Tag }
                if ($match) { $match['Active'] = $cb.Checked }
            }
            # Benutzerdefinierte Vorbelastungen: Active-Status aus Checkboxen übernehmen
            foreach ($cb in $script:geneticCustomChecks) {
                $match = $script:data.Config.GeneticPredispositions['Custom'] | Where-Object { $_['Name'] -eq $cb.Tag }
                if ($match) { $match['Active'] = $cb.Checked }
            }
            
            Save-Data -data $script:data -type "Config"
            [System.Windows.Forms.MessageBox]::Show("Persönliche Daten erfolgreich gespeichert.", "Erfolg", "OK", "Information")
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern der persönlichen Daten: $($_.Exception.Message)", "Validierungsfehler", "OK", "Error")
        }
    }

    # v2.17.0: Hilfsfunktionen für genetische Vorbelastungen
    function Update-GeneticCustomPanel {
        $geneticCustomPanel.Controls.Clear()
        $script:geneticCustomChecks = @()
        $gp = $script:data.Config.GeneticPredispositions
        $yPos = 0
        if ($gp.Custom -and $gp.Custom.Count -gt 0) {
            foreach ($entry in $gp.Custom) {
                $cb = New-Object System.Windows.Forms.CheckBox
                $cb.Text = $entry['Name']
                $cb.Location = New-Object System.Drawing.Point(20, $yPos)
                $cb.Size = New-Object System.Drawing.Size(500, 20)
                $cb.Enabled = $gp.Enabled
                $cb.Checked = $entry['Active']
                $cb.Tag = $entry['Name']
                $geneticCustomPanel.Controls.Add($cb)
                $script:geneticCustomChecks += $cb
                $yPos += 28
            }
        }
    }

    function Update-GeneticNewMarkerList {
        $geneticNewMarkerList.Items.Clear()
        $allNames = @($script:data.Config.Markers.Name) + @($script:data.Config.CalculatedMarkers.Name) | Sort-Object -Unique
        foreach ($name in $allNames) {
            $geneticNewMarkerList.Items.Add($name) | Out-Null
        }
    }

    function Get-AllActiveGeneticMarkers {
        # Gibt ein Hashtable zurück: MarkerName -> @(Hinweis-Texte)
        $result = @{}
        $gp = $script:data.Config.GeneticPredispositions
        if (-not $gp -or -not $gp['Enabled']) { return $result }
        $allEntries = @()
        if ($gp['BuiltIn']) { $allEntries += @($gp['BuiltIn']) }
        if ($gp['Custom'])  { $allEntries += @($gp['Custom']) }
        foreach ($entry in $allEntries) {
            if ($entry['Active']) {
                foreach ($marker in $entry['Markers']) {
                    if (-not $result.ContainsKey($marker)) { $result[$marker] = @() }
                    $shortName = ($entry['Name'] -split '\(')[0].Trim()
                    $result[$marker] += "⚠ Vorbelastung: $shortName"
                }
            }
        }
        return $result
    }

    function Get-AllGeneticPredispositions {
        # Liefert alle (BuiltIn + Custom) Vorbelastungen als flache Liste
        $gp = $script:data.Config.GeneticPredispositions
        $all = @()
        if ($gp['BuiltIn']) { $all += @($gp['BuiltIn']) }
        if ($gp['Custom'])  { $all += @($gp['Custom']) }
        return $all
    }

    function Check-Warnings {
        param($markerData, $markerName)
        $warningDisplayGroup.Visible = $false
        $warningLabel.Text = ""
        $warningMessages = New-Object System.Collections.ArrayList

        if ($markerName -eq "Atherogener Index (AIP)") {
            $msg = "INFO: Der AIP wird mit Triglyceriden und HDL in mg/dl berechnet. Das Ergebnis ist ein dimensionsloser Index."
            $warningMessages.Add($msg) | Out-Null
        }
        
        $regression = Calculate-LinearRegression -dataPoints $markerData
        if ($regression) {
            $meanValue = ($markerData.Value | Measure-Object -Average).Average
            if ($meanValue -ne 0) {
                $yearlyChange = $regression.Slope * 365.25
                $relativeChange = $yearlyChange / $meanValue
                if ([Math]::Abs($relativeChange) -gt 0.1) {
                    $direction = if ($regression.Slope -gt 0) { "steigender" } else { "fallender" }
                    $msg = "WARNUNG (Trend): Es wird ein signifikant $direction Trend von ca. $('{0:P0}' -f $relativeChange) pro Jahr prognostiziert."
                    $warningMessages.Add($msg) | Out-Null
                }
            }
        }

        if ($markerData.Count -lt 2) { 
             if ($warningMessages.Count -gt 0) { $warningLabel.Text = $warningMessages -join "`r`n"; $warningDisplayGroup.Visible = $true }
             return 
        }

        try {
            $threshold = (Parse-Number $thresholdTextBox.Text) / 100
        } catch {
            $warningMessages.Add("Ungültiger Schwellenwert.")
            $warningLabel.Text = $warningMessages -join "`r`n"
            $warningDisplayGroup.Visible = $true
            return
        }

        $markerDataWithDates = $markerData | Select-Object *, @{N = 'DateObject'; E = { [datetime]$_.Date } } | Sort-Object DateObject
        $latestEntry = $markerDataWithDates[-1]

        $targetDate1Y = $latestEntry.DateObject.AddYears(-1)
        $comparisonEntry1Y = $markerDataWithDates | Where-Object { $_.DateObject -le $targetDate1Y } | Select-Object -Last 1
        if ($comparisonEntry1Y -and $comparisonEntry1Y.Value -ne 0) {
            $percentChange = (($latestEntry.Value - $comparisonEntry1Y.Value) / $comparisonEntry1Y.Value)
            if ([Math]::Abs($percentChange) -gt $threshold) {
                $changeText = "{0:P0}" -f $percentChange
                $msg = "WARNUNG (1 Jahr): Wert hat sich um $changeText verändert (Aktuell: $($latestEntry.Value) am $($latestEntry.DateObject.ToString('dd.MM.yyyy')) vs. $($comparisonEntry1Y.Value) am $($comparisonEntry1Y.DateObject.ToString('dd.MM.yyyy')))."
                $warningMessages.Add($msg) | Out-Null
            }
        }

        $targetDate3Y = $latestEntry.DateObject.AddYears(-3)
        $comparisonEntry3Y = $markerDataWithDates | Where-Object { $_.DateObject -le $targetDate3Y } | Select-Object -Last 1
        if ($comparisonEntry3Y -and $comparisonEntry3Y.Value -ne 0) {
            $percentChange = (($latestEntry.Value - $comparisonEntry3Y.Value) / $comparisonEntry3Y.Value)
            if ([Math]::Abs($percentChange) -gt $threshold) {
                $changeText = "{0:P0}" -f $percentChange
                $msg = "WARNUNG (3 Jahre): Wert hat sich um $changeText verändert (Aktuell: $($latestEntry.Value) am $($latestEntry.DateObject.ToString('dd.MM.yyyy')) vs. $($comparisonEntry3Y.Value) am $($comparisonEntry3Y.DateObject.ToString('dd.MM.yyyy')))."
                $warningMessages.Add($msg) | Out-Null
            }
        }

        if ($warningMessages.Count -gt 0) {
            $warningLabel.Text = $warningMessages -join "`r`n"
            $warningDisplayGroup.Visible = $true
        }
    }
    function Update-Chart {
        $series.Points.Clear(); $trendSeries.Points.Clear(); $chartArea.AxisY.StripLines.Clear(); $chart.Annotations.Clear()
        $selectedMarker = $filterComboBox.SelectedItem
        if ($selectedMarker -eq "--- Bitte auswählen ---" -or [string]::IsNullOrEmpty($selectedMarker)) {
            $chart.Titles.Clear(); $chart.Titles.Add("Bitte einen Marker aus dem Filter auswählen."); Check-Warnings -markerData @() -markerName ""
            return
        }
        $chart.Titles.Clear(); $chart.Titles.Add("Verlauf für '$selectedMarker'")
        
        $markerConfig = $script:data.Config.Markers | Where-Object { $_.Name -eq $selectedMarker } | Select-Object -First 1
        $isCalculated = $false
        if (-not $markerConfig) { 
            $markerConfig = $script:data.Config.CalculatedMarkers | Where-Object { $_.Name -eq $selectedMarker } | Select-Object -First 1
            $isCalculated = $true
        }

        $markerData = if($isCalculated) { Get-CalculatedValuesForMarker -markerName $selectedMarker } else { $script:allHistoricalItems | Where-Object { $_.Name -eq $selectedMarker } }
        $markerData = $markerData | Sort-Object Date

        foreach ($item in $markerData) { $series.Points.AddXY([datetime]$item.Date, $item.Value) }
        
        $regression = Calculate-LinearRegression -dataPoints $markerData
        if ($regression -and $markerData.Count > 1) {
            $firstDate = ([datetime]$markerData[0].Date).ToOADate()
            $lastDate = ([datetime]$markerData[-1].Date).ToOADate()
            $firstValue = $regression.Intercept + $regression.Slope * $firstDate
            $lastValue = $regression.Intercept + $regression.Slope * $lastDate
            $trendSeries.Points.AddXY([datetime]::FromOADate($firstDate), $firstValue)
            $trendSeries.Points.AddXY([datetime]::FromOADate($lastDate), $lastValue)
        }

        if ($markerConfig) {
            if ($markerConfig.OptimalMin -ne $null -and $markerConfig.OptimalMax -ne $null) {
                $optimalStrip = New-Object System.Windows.Forms.DataVisualization.Charting.StripLine
                $optimalStrip.Interval = 0; $optimalStrip.IntervalOffset = $markerConfig.OptimalMin; $optimalStrip.StripWidth = $markerConfig.OptimalMax - $markerConfig.OptimalMin
                $optimalStrip.BackColor = [System.Drawing.Color]::FromArgb(64, [System.Drawing.Color]::LightGreen)
                $chartArea.AxisY.StripLines.Add($optimalStrip)
            }
            if ($markerConfig.RefMin -ne $null) {
                $stripLineMin = New-Object System.Windows.Forms.DataVisualization.Charting.StripLine; $stripLineMin.IntervalOffset = $markerConfig.RefMin; $stripLineMin.BorderColor = [System.Drawing.Color]::Orange; $stripLineMin.BorderDashStyle = "Dash"; $stripLineMin.BorderWidth = 2; $chartArea.AxisY.StripLines.Add($stripLineMin)
            }
            if ($markerConfig.RefMax -ne $null) {
                $stripLineMax = New-Object System.Windows.Forms.DataVisualization.Charting.StripLine; $stripLineMax.IntervalOffset = $markerConfig.RefMax; $stripLineMax.BorderColor = [System.Drawing.Color]::Orange; $stripLineMax.BorderDashStyle = "Dash"; $stripLineMax.BorderWidth = 2; $chartArea.AxisY.StripLines.Add($stripLineMax)
            }
        }
        # v2.22.0: Letzten Wert + Datum im Graph hervorheben (schnelles Ablesen)
        if ($markerData.Count -gt 0) {
            $lastItem = $markerData[-1]
            $lastUnit = if ($lastItem.PSObject.Properties['Unit'] -and $lastItem.Unit) { $lastItem.Unit } elseif ($markerConfig) { $markerConfig.Unit } else { "" }
            $lastDateStr = ([datetime]$lastItem.Date).ToString("dd.MM.yyyy")
            $lastSubTitle = New-Object System.Windows.Forms.DataVisualization.Charting.Title
            $lastSubTitle.Text = "Letzter Wert: $($lastItem.Value) $lastUnit  (am $lastDateStr)"
            $lastSubTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
            $lastSubTitle.ForeColor = [System.Drawing.Color]::FromArgb(0, 70, 140)
            $chart.Titles.Add($lastSubTitle)
            try {
                if ($series.Points.Count -gt 0) {
                    $lastAnnotation = New-Object System.Windows.Forms.DataVisualization.Charting.TextAnnotation
                    $lastAnnotation.Text = "$($lastItem.Value) $lastUnit`n$lastDateStr"
                    $lastAnnotation.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                    $lastAnnotation.ForeColor = [System.Drawing.Color]::Black
                    $lastAnnotation.LineColor = [System.Drawing.Color]::DarkGray
                    $lastAnnotation.AnchorDataPoint = $series.Points[$series.Points.Count - 1]
                    $lastAnnotation.AnchorAlignment = 'BottomCenter'
                    $lastAnnotation.AllowMoving = $true
                    $chart.Annotations.Add($lastAnnotation)
                }
            } catch { Write-Warning "Annotation letzter Wert: $($_.Exception.Message)" }
        }
        Check-Warnings -markerData $markerData -markerName $selectedMarker
    }

    $cockpitGrid.Add_CellMouseEnter({
        param($sender, $e)
        if ($e.RowIndex -lt 0) { return }
    
        $row = $sender.Rows[$e.RowIndex]
        $cell = $row.Cells[$e.ColumnIndex]
    
        if ($e.ColumnIndex -eq $sender.Columns["Rating"].Index) {
            if ($cell.Style.BackColor.IsEmpty) { $cell.ToolTipText = ""; return }
            $colorName = $cell.Style.BackColor.Name
            $tooltipText = ""
            if ($colorName -eq ([System.Drawing.Color]::LightGreen).Name) { $tooltipText = "Optimal: Der Wert liegt im definierten Optimalbereich." } 
            elseif ($colorName -eq ([System.Drawing.Color]::LightYellow).Name) { $tooltipText = "Akzeptabel: Der Wert liegt im allgemeinen Referenzbereich, aber außerhalb des Optimalbereichs." }
            elseif ($colorName -eq ([System.Drawing.Color]::LightCoral).Name) { $tooltipText = "Außerhalb der Norm: Der Wert liegt außerhalb des Referenzbereichs. Eine ärztliche Abklärung könnte sinnvoll sein." }
            $cell.ToolTipText = $tooltipText
        }
        elseif ($e.ColumnIndex -eq $sender.Columns["Marker"].Index) {
            $markerName = $cell.Value
            if ($markerName) {
                $allMarkers = [System.Collections.ArrayList]@($script:data.Config.Markers + $script:data.Config.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ })
                $markerConfig = $allMarkers | Where-Object { $_.Name -eq $markerName } | Select-Object -First 1
    
                if ($markerConfig -and $markerConfig.PSObject.Properties['Description'] -and -not [string]::IsNullOrWhiteSpace($markerConfig.Description)) {
                    $cell.ToolTipText = $markerConfig.Description
                } else {
                    $cell.ToolTipText = ""
                }
            }
        }
    })
    $runCorrelationButton.Add_Click({ Update-CorrelationAnalysis })
    $suggestPairsButton.Add_Click({
        $selectedPair = Show-CorrelationSuggestionsPopup
        if ($selectedPair) {
            $corrMarker1ComboBox.SelectedItem = $selectedPair.Marker1
            $corrMarker2ComboBox.SelectedItem = $selectedPair.Marker2
            Update-CorrelationAnalysis
        }
    })
    $corrExportButton.Add_Click({
        $dataProvider = {
            $m1 = $corrMarker1ComboBox.SelectedItem
            $m2 = $corrMarker2ComboBox.SelectedItem
            if ([string]::IsNullOrEmpty($m1) -or [string]::IsNullOrEmpty($m2)) { return @() }

            $m1Data = if ($script:data.Config.CalculatedMarkers.Name -contains $m1) { Get-CalculatedValuesForMarker -markerName $m1 } else { $script:allHistoricalItems | Where-Object { $_.Name -eq $m1 } }
            $m2Data = if ($script:data.Config.CalculatedMarkers.Name -contains $m2) { Get-CalculatedValuesForMarker -markerName $m2 } else { $script:allHistoricalItems | Where-Object { $_.Name -eq $m2 } }

            $m1Dict = @{}
            $m1Data | ForEach-Object { $m1Dict[$_.Date] = $_.Value }

            $rows = New-Object System.Collections.ArrayList
            $m2Data | Sort-Object Date | ForEach-Object {
                if ($m1Dict.ContainsKey($_.Date)) {
                    $null = $rows.Add([PSCustomObject]@{
                        Datum  = $_.Date
                        Marker1 = $m1
                        Wert1   = $m1Dict[$_.Date]
                        Marker2 = $m2
                        Wert2   = $_.Value
                    })
                }
            }
            return $rows
        }
        $m1Clean = if ($corrMarker1ComboBox.SelectedItem) { ($corrMarker1ComboBox.SelectedItem -replace '[^\w\-]', '_') } else { "Marker1" }
        $m2Clean = if ($corrMarker2ComboBox.SelectedItem) { ($corrMarker2ComboBox.SelectedItem -replace '[^\w\-]', '_') } else { "Marker2" }
        Show-ExportPopup -source "Korrelationen" -dataProvider $dataProvider -defaultFileName "Korrelation_${m1Clean}_vs_${m2Clean}_$(Get-Date -Format 'yyyy-MM-dd')"
    })

    $corrPrintButton.Add_Click({
        try {
            $marker1Name = $corrMarker1ComboBox.SelectedItem
            $marker2Name = $corrMarker2ComboBox.SelectedItem
            if ([string]::IsNullOrEmpty($marker1Name) -or [string]::IsNullOrEmpty($marker2Name)) {
                [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst zwei Marker aus.", "Keine Auswahl", "OK", "Information")
                return
            }

            $printDoc = New-Object System.Drawing.Printing.PrintDocument
            $printDoc.DefaultPageSettings.Landscape = $true
            $printDoc.Add_PrintPage({
                param($sender, $ev)
                $g = $ev.Graphics
                $margins = $ev.MarginBounds
                $titleFont = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
                $subFont = New-Object System.Drawing.Font("Segoe UI", 10)
                $smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
                $yPos = $margins.Top

                $g.DrawString("Korrelationsanalyse", $titleFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                $yPos += 30
                $personalName = $script:data.Config.Personal.Name
                if (-not [string]::IsNullOrWhiteSpace($personalName)) {
                    $g.DrawString("Patient: $personalName", $subFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += 20
                }
                $g.DrawString("$marker1Name  vs.  $marker2Name", $subFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, $yPos)
                $yPos += 20
                $g.DrawString("Druckdatum: $(Get-Date -Format 'dd.MM.yyyy HH:mm')", $smallFont, [System.Drawing.Brushes]::Gray, $margins.Left, $yPos)
                $yPos += 16
                if (-not [string]::IsNullOrWhiteSpace($correlationResultLabel.Text) -and $correlationResultLabel.Text -ne "Bitte zwei Marker auswählen und Analyse starten.") {
                    $g.DrawString($correlationResultLabel.Text, $subFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += 22
                }
                $yPos += 5

                $chartWidth = $margins.Width
                $chartHeight = [Math]::Min(($margins.Bottom - $yPos - 20), 480)
                $chartBitmap = New-Object System.Drawing.Bitmap($correlationChart.Width, $correlationChart.Height)
                $correlationChart.DrawToBitmap($chartBitmap, (New-Object System.Drawing.Rectangle(0, 0, $correlationChart.Width, $correlationChart.Height)))
                $destRect = New-Object System.Drawing.Rectangle($margins.Left, $yPos, $chartWidth, $chartHeight)
                $g.DrawImage($chartBitmap, $destRect)
                $chartBitmap.Dispose()
                $ev.HasMorePages = $false
            })

            $preview = New-Object System.Windows.Forms.PrintPreviewDialog
            $preview.Document = $printDoc
            $preview.Width = 1000
            $preview.Height = 700
            $preview.ShowDialog()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Drucken: $($_.Exception.Message)", "Druckfehler", "OK", "Error")
        }
    })
    $advancedViewCheckBox.Add_CheckedChanged({
        $isExpanded = $advancedViewCheckBox.Checked
        $advancedControls | ForEach-Object { $_.Visible = $isExpanded }
        if ($isExpanded) {
            $markerMgmtGroup.Height = 370
            $addNewMarkerButton.Location = New-Object System.Drawing.Point(110, 325)
            $deleteExistingMarkerButton.Location = New-Object System.Drawing.Point(265, 325)
            $warningThresholdGroup.Location = New-Object System.Drawing.Point(20, 630)
        } else {
            $markerMgmtGroup.Height = 220
            $addNewMarkerButton.Location = New-Object System.Drawing.Point(110, 175)
            $deleteExistingMarkerButton.Location = New-Object System.Drawing.Point(265, 175)
            $warningThresholdGroup.Location = New-Object System.Drawing.Point(20, 480)
        }
    })
    $tabControl.Add_MouseMove({
        param($sender, $e)
        $currentIndex = -1
        for ($i = 0; $i -lt $tabControl.TabCount; $i++) {
            if ($tabControl.GetTabRect($i).Contains($e.Location)) { $currentIndex = $i; break }
        }
        if ($currentIndex -ne $lastTooltipIndex) {
            if ($currentIndex -ne -1) { $toolTip.SetToolTip($tabControl, $tabControl.TabPages[$currentIndex].Tag) } 
            else { $toolTip.SetToolTip($tabControl, "") }
            $lastTooltipIndex = $currentIndex
        }
    })
    $tabControl.Add_SelectedIndexChanged({ 
        if ($tabControl.SelectedTab.Text -eq "Einzelmarker-Analyse") { Update-FilterDropdown; Update-Chart }
        if ($tabControl.SelectedTab.Text -eq "Risiko-Cockpit") { Update-Dashboard }
        if ($tabControl.SelectedTab.Text -eq "Korrelationen") { Update-CorrelationDropdowns }
        if ($tabControl.SelectedTab.Text -eq "Daten nach Bluttests") { Update-DataManagementTab }
        if ($tabControl.SelectedTab.Text -eq "Persönliche Metriken") { Load-PersonalMetrics }
        if ($tabControl.SelectedTab.Text -eq "Longevity-Indizes") { Update-LongevityTab }
        if ($tabControl.SelectedTab.Text -eq "Custom Report") { Update-CustomReportTab }
    })
    # v2.15.1: Alter $settingsButton.Add_Click-Handler entfernt - globaler "globale Einstellungen"-
    # Button auf Form-Ebene übernimmt die Funktion (siehe Show-DataEntryForm, inkl. Update-FilterDropdown
    # und Update-Chart nach Popup-Schließung).
    $cockpitSettingsButton.Add_Click({ Show-CockpitSettingsPopup; Update-Dashboard })
    $longevitySettingsButton.Add_Click({ Show-LongevitySettingsPopup; Update-LongevityTab })

    # ---------- Custom Report Events (v2.12.0) ----------
    # ItemCheck: blockiert Klicks auf deaktivierte Marker (ohne Daten)
    # und aktualisiert Preview/Button-State nach erfolgreicher Änderung.
    $customMarkerList.Add_ItemCheck({
        param($sender, $e)
        $display = [string]$customMarkerList.Items[$e.Index]
        if ($display -match '\s+\[keine Daten\]$') {
            # Disabled-Item darf nicht angehakt werden
            $e.NewValue = [System.Windows.Forms.CheckState]::Unchecked
            return
        }
        # Preview/Button-State erst nach Übernahme des neuen Wertes aktualisieren
        $customMarkerList.BeginInvoke([Action]{
            Update-CustomReportPreview
            Update-CustomReportButtonState
        }) | Out-Null
    })

    # DrawItem: graut Marker ohne Daten optisch aus (Owner-Drawn)
    $customMarkerList.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawFixed
    $customMarkerList.ItemHeight = 20
    $customMarkerList.Add_DrawItem({
        param($sender, $e)
        if ($e.Index -lt 0) { return }
        $itemText = [string]$customMarkerList.Items[$e.Index]
        $isDisabled = $itemText -match '\s+\[keine Daten\]$'

        $e.DrawBackground()

        # Check-Box zeichnen
        $checkSize = 14
        $checkRect = New-Object System.Drawing.Rectangle(
            ($e.Bounds.X + 2), ($e.Bounds.Y + 3), $checkSize, $checkSize
        )
        $checked = $customMarkerList.GetItemChecked($e.Index)
        $state = if ($isDisabled) {
            [System.Windows.Forms.VisualStyles.CheckBoxState]::UncheckedDisabled
        } elseif ($checked) {
            [System.Windows.Forms.VisualStyles.CheckBoxState]::CheckedNormal
        } else {
            [System.Windows.Forms.VisualStyles.CheckBoxState]::UncheckedNormal
        }
        [System.Windows.Forms.CheckBoxRenderer]::DrawCheckBox($e.Graphics, $checkRect.Location, $state)

        # Text zeichnen
        $textBrush = if ($isDisabled) {
            [System.Drawing.Brushes]::Gray
        } elseif (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -ne 0) {
            [System.Drawing.Brushes]::White
        } else {
            [System.Drawing.Brushes]::Black
        }
        $textRect = New-Object System.Drawing.Rectangle(
            ($e.Bounds.X + $checkSize + 8), $e.Bounds.Y,
            ($e.Bounds.Width - $checkSize - 10), $e.Bounds.Height
        )
        $e.Graphics.DrawString($itemText, $customMarkerList.Font, $textBrush, $textRect)

        $e.DrawFocusRectangle()
    })

    $customMarkerSearchBox.Add_TextChanged({ Update-CustomReportTab })

    $customSelectAllButton.Add_Click({
        for ($i = 0; $i -lt $customMarkerList.Items.Count; $i++) {
            $display = [string]$customMarkerList.Items[$i]
            if (-not ($display -match '\s+\[keine Daten\]$')) {
                $customMarkerList.SetItemChecked($i, $true)
            }
        }
        Update-CustomReportPreview
        Update-CustomReportButtonState
    })
    $customDeselectAllButton.Add_Click({
        for ($i = 0; $i -lt $customMarkerList.Items.Count; $i++) {
            $customMarkerList.SetItemChecked($i, $false)
        }
        Update-CustomReportPreview
        Update-CustomReportButtonState
    })

    $customTimeCombo.Add_SelectedIndexChanged({
        $script:customReportTimeFilter = $customTimeCombo.SelectedItem
        Update-CustomReportTab
    })

    # v2.17.0: Vorbelastungs-Marker automatisch auswählen
    $customGeneticApplyButton.Add_Click({
        $selectedPred = $customGeneticCombo.SelectedItem
        if (-not $selectedPred -or $selectedPred -eq "– Keine Vorbelastung –") {
            # Alle abwählen
            for ($i = 0; $i -lt $customMarkerList.Items.Count; $i++) {
                $customMarkerList.SetItemChecked($i, $false)
            }
        } else {
            # Vorbelastung finden und deren Marker auswählen
            $allPreds = Get-AllGeneticPredispositions
            $pred = $allPreds | Where-Object { $_['Name'] -eq $selectedPred } | Select-Object -First 1
            if ($pred) {
                $predMarkers = @($pred['Markers'])
                for ($i = 0; $i -lt $customMarkerList.Items.Count; $i++) {
                    $raw = [string]$customMarkerList.Items[$i] -replace '\s+\[keine Daten\]$', ''
                    $shouldCheck = $predMarkers -contains $raw
                    $hasData = $script:customReportAvailableMarkers -contains $raw
                    $customMarkerList.SetItemChecked($i, ($shouldCheck -and $hasData))
                }
            }
        }
        Update-CustomReportPreview
        Update-CustomReportButtonState
    })

    $customExportButton.Add_Click({
        # Das Exportieren-Popup öffnen. JSON/CSV nutzen die Default-Tabellendaten,
        # PDF hingegen wird für Custom-Report über eine spezielle "report-mit-grafik"-Variante
        # ersetzt (siehe Invoke-CustomReportPrint).
        $dataProvider = { Get-CustomReportData }
        $zeit = ($script:customReportTimeFilter -replace '\s+', '_')

        # Eigenes Popup statt Show-ExportPopup, weil PDF hier Grafiken enthält
        $popup = New-Object System.Windows.Forms.Form
        $popup.Size = New-Object System.Drawing.Size(450, 310)
        $popup.Text = "Exportieren – Custom Report"
        $popup.StartPosition = "CenterParent"
        $popup.FormBorderStyle = "FixedDialog"
        $popup.MaximizeBox = $false; $popup.MinimizeBox = $false

        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = "Wähle ein Format zum Exportieren des Custom Reports."
        $lbl.Location = New-Object System.Drawing.Point(20, 20)
        $lbl.Size = New-Object System.Drawing.Size(400, 20)
        $popup.Controls.Add($lbl)

        $jsonBtn = New-Object System.Windows.Forms.Button
        $jsonBtn.Text = "Als JSON exportieren (.json)..."
        $jsonBtn.Location = New-Object System.Drawing.Point(20, 60); $jsonBtn.Size = New-Object System.Drawing.Size(400, 35)
        $popup.Controls.Add($jsonBtn)

        $csvBtn = New-Object System.Windows.Forms.Button
        $csvBtn.Text = "Als CSV exportieren (.csv)..."
        $csvBtn.Location = New-Object System.Drawing.Point(20, 105); $csvBtn.Size = New-Object System.Drawing.Size(400, 35)
        $popup.Controls.Add($csvBtn)

        $pdfBtn = New-Object System.Windows.Forms.Button
        $pdfBtn.Text = "Als PDF exportieren (inkl. Grafik pro Marker)..."
        $pdfBtn.Location = New-Object System.Drawing.Point(20, 150); $pdfBtn.Size = New-Object System.Drawing.Size(400, 35)
        $popup.Controls.Add($pdfBtn)

        $closeBtn = New-Object System.Windows.Forms.Button
        $closeBtn.Text = "Schließen"
        $closeBtn.Location = New-Object System.Drawing.Point(170, 220); $closeBtn.Size = New-Object System.Drawing.Size(100, 30)
        $closeBtn.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $popup.Controls.Add($closeBtn); $popup.CancelButton = $closeBtn

        $jsonBtn.Add_Click({
            try {
                $data = @(Get-CustomReportData)
                if ($data.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", "OK", "Information"); return }
                $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
                $saveDlg.Filter = "JSON-Datei (*.json)|*.json"
                $saveDlg.FileName = "CustomReport_${zeit}_$(Get-Date -Format 'yyyy-MM-dd').json"
                if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $data | ConvertTo-Json -Depth 5 | Out-File -FilePath $saveDlg.FileName -Encoding UTF8
                    $popup.Close()
                    Invoke-PostExportAction -FilePath $saveDlg.FileName -SuccessTitle "Export" -SuccessMessage "JSON erfolgreich gespeichert:`n$($saveDlg.FileName)"
                }
            } catch { [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Exportfehler", "OK", "Error") }
        })
        $csvBtn.Add_Click({
            try {
                $data = @(Get-CustomReportData)
                if ($data.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Keine Daten zum Exportieren vorhanden.", "Export", "OK", "Information"); return }
                $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
                $saveDlg.Filter = "CSV-Datei (*.csv)|*.csv"
                $saveDlg.FileName = "CustomReport_${zeit}_$(Get-Date -Format 'yyyy-MM-dd').csv"
                if ($saveDlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                    $data | Export-Csv -Path $saveDlg.FileName -Delimiter ";" -Encoding UTF8 -NoTypeInformation
                    $popup.Close()
                    Invoke-PostExportAction -FilePath $saveDlg.FileName -SuccessTitle "Export" -SuccessMessage "CSV erfolgreich gespeichert:`n$($saveDlg.FileName)"
                }
            } catch { [System.Windows.Forms.MessageBox]::Show("Fehler: $($_.Exception.Message)", "Exportfehler", "OK", "Error") }
        })
        $pdfBtn.Add_Click({
            try {
                $saveDlg = New-Object System.Windows.Forms.SaveFileDialog
                $saveDlg.Filter = "PDF-Datei (*.pdf)|*.pdf"
                $saveDlg.FileName = "CustomReport_${zeit}_$(Get-Date -Format 'yyyy-MM-dd').pdf"
                if ($saveDlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
                $result = Invoke-CustomReportPrint -TargetPdfPath $saveDlg.FileName
                if ($result) {
                    $popup.Close()
                    Invoke-PostExportAction -FilePath $saveDlg.FileName -SuccessTitle "PDF-Export" -SuccessMessage "PDF erfolgreich erstellt:`n$($saveDlg.FileName)"
                }
            } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim PDF-Export: $($_.Exception.Message)", "Exportfehler", "OK", "Error") }
        })

        $popup.ShowDialog() | Out-Null
    })

    $customPrintButton.Add_Click({
        try {
            Invoke-CustomReportPrint -TargetPdfPath $null | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Drucken: $($_.Exception.Message)", "Druckfehler", "OK", "Error")
        }
    })
    $cockpitExportButton.Add_Click({
        $dataProvider = {
            $rows = New-Object System.Collections.ArrayList
            foreach ($row in $cockpitGrid.Rows) {
                if ($row.IsNewRow) { continue }
                $obj = [PSCustomObject]@{
                    Gruppe  = [string]$row.Cells["Group"].Value
                    Marker  = [string]$row.Cells["Marker"].Value
                    Wert    = [string]$row.Cells["Value"].Value
                    Einheit = [string]$row.Cells["Unit"].Value
                    Datum   = [string]$row.Cells["Date"].Value
                    Risiko  = [string]$row.Cells["Risk"].Value
                    Bewertung = [string]$row.Cells["Rating"].Value
                }
                $rows.Add($obj) | Out-Null
            }
            return $rows
        }
        Show-ExportPopup -source "Risiko-Cockpit" -dataProvider $dataProvider -defaultFileName "Cockpit_$(Get-Date -Format 'yyyy-MM-dd')"
    })

    $cockpitPrintButton.Add_Click({
        try {
            Add-Type -AssemblyName System.Drawing
            $printDoc = New-Object System.Drawing.Printing.PrintDocument
            $printDoc.DefaultPageSettings.Landscape = $true
            $script:printRowIndex = 0
            $script:printPageNum = 0

            $printDoc.Add_PrintPage({
                param($sender, $ev)
                $g = $ev.Graphics
                $margins = $ev.MarginBounds
                $headerFont = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
                $cellFont = New-Object System.Drawing.Font("Segoe UI", 9)
                $titleFont = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
                $smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
                $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 0.5)
                $yPos = $margins.Top
                $rowHeight = 26

                # Sichtbare Spalten
                $visibleCols = @($cockpitGrid.Columns | Where-Object { $_.Visible })
                $totalWeight = ($visibleCols | Measure-Object -Property FillWeight -Sum).Sum
                $colWidths = @{}
                foreach ($col in $visibleCols) {
                    $colWidths[$col.Name] = [int](($col.FillWeight / $totalWeight) * $margins.Width)
                }

                # Titel und Filterinfo auf erster Seite
                if ($script:printPageNum -eq 0) {
                    $g.DrawString("Risiko-Cockpit", $titleFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += 30
                    $personalName = $script:data.Config.Personal.Name
                    if (-not [string]::IsNullOrWhiteSpace($personalName)) {
                        $g.DrawString("Patient: $personalName", $cellFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                        $yPos += 18
                    }
                    $dateStr = "Druckdatum: $(Get-Date -Format 'dd.MM.yyyy HH:mm')"
                    $g.DrawString($dateStr, $smallFont, [System.Drawing.Brushes]::Gray, $margins.Left, $yPos)
                    $yPos += 16
                    if ($cockpitFilterHintLabel.Text -ne "") {
                        $g.DrawString($cockpitFilterHintLabel.Text, $smallFont, [System.Drawing.Brushes]::Red, $margins.Left, $yPos)
                        $yPos += 16
                    }
                    $yPos += 10
                }
                $script:printPageNum++

                # Spaltenheader
                $xPos = $margins.Left
                foreach ($col in $visibleCols) {
                    $rect = New-Object System.Drawing.RectangleF($xPos, $yPos, $colWidths[$col.Name], $rowHeight)
                    $g.FillRectangle([System.Drawing.Brushes]::LightSteelBlue, $rect)
                    $g.DrawRectangle($pen, [int]$rect.X, [int]$rect.Y, [int]$rect.Width, [int]$rect.Height)
                    $g.DrawString($col.HeaderText, $headerFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 4))
                    $xPos += $colWidths[$col.Name]
                }
                $yPos += $rowHeight

                # Datenzeilen
                while ($script:printRowIndex -lt $cockpitGrid.Rows.Count) {
                    if (($yPos + $rowHeight) -gt $margins.Bottom) {
                        $ev.HasMorePages = $true
                        return
                    }
                    $row = $cockpitGrid.Rows[$script:printRowIndex]
                    $xPos = $margins.Left
                    foreach ($col in $visibleCols) {
                        $cellValue = if ($row.Cells[$col.Name].Value) { $row.Cells[$col.Name].Value.ToString() } else { "" }
                        $rect = New-Object System.Drawing.RectangleF($xPos, $yPos, $colWidths[$col.Name], $rowHeight)
                        # Bewertungs-Farbe übernehmen
                        if ($col.Name -eq "Rating") {
                            $bgColor = $row.Cells[$col.Name].Style.BackColor
                            if (-not $bgColor.IsEmpty) {
                                $brush = New-Object System.Drawing.SolidBrush($bgColor)
                                $g.FillRectangle($brush, $rect)
                                $brush.Dispose()
                            }
                        }
                        $g.DrawRectangle($pen, [int]$rect.X, [int]$rect.Y, [int]$rect.Width, [int]$rect.Height)
                        $g.DrawString($cellValue, $cellFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 5))
                        $xPos += $colWidths[$col.Name]
                    }
                    $yPos += $rowHeight
                    $script:printRowIndex++
                }
                $ev.HasMorePages = $false
            })

            $preview = New-Object System.Windows.Forms.PrintPreviewDialog
            $preview.Document = $printDoc
            $preview.Width = 1000
            $preview.Height = 700
            $preview.ShowDialog()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Drucken: $($_.Exception.Message)", "Druckfehler", "OK", "Error")
        }
    })
    $dataMgmtSettingsButton.Add_Click({ Show-DataMgmtSettingsPopup; Update-DataManagementTab })
    $dataMgmtExportButton.Add_Click({
        $dataProvider = {
            $rows = New-Object System.Collections.ArrayList
            $selectedDate = "unknown"
            if ($dataMgmtTreeView.SelectedNode -and $dataMgmtTreeView.SelectedNode.Tag) {
                $selectedDate = [string]$dataMgmtTreeView.SelectedNode.Tag
            }
            foreach ($row in $dataMgmtDataGridView.Rows) {
                if ($row.IsNewRow) { continue }
                $null = $rows.Add([PSCustomObject]@{
                    Datum   = $selectedDate
                    Marker  = [string]$row.Cells["Marker"].Value
                    Wert    = [string]$row.Cells["Value"].Value
                    Einheit = [string]$row.Cells["Unit"].Value
                    Notiz   = [string]$row.Cells["Note"].Value
                })
            }
            return $rows
        }
        $dateSlug = "Auswahl"
        if ($dataMgmtTreeView.SelectedNode -and $dataMgmtTreeView.SelectedNode.Tag) {
            $dateSlug = [string]$dataMgmtTreeView.SelectedNode.Tag
        }
        Show-ExportPopup -source "Daten nach Bluttests" -dataProvider $dataProvider -defaultFileName "Bluttest_${dateSlug}"
    })

    $dataMgmtPrintButton.Add_Click({
        try {
            $selectedNode = $dataMgmtTreeView.SelectedNode
            if (-not $selectedNode) {
                [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst einen Test-Tag aus der Liste aus.", "Keine Auswahl", "OK", "Information")
                return
            }
            $selectedDate = $selectedNode.Tag
            $dateFormatted = ([datetime]$selectedDate).ToString("dd. MMMM yyyy")
            $itemsForDate = $script:allHistoricalItems | Where-Object { $_.Date -eq $selectedDate } | Sort-Object Name

            $printDoc = New-Object System.Drawing.Printing.PrintDocument
            $printDoc.DefaultPageSettings.Landscape = $false
            $script:printRowIdx = 0
            $script:printPageIdx = 0
            $script:printItems = @($itemsForDate)

            $printDoc.Add_PrintPage({
                param($sender, $ev)
                $g = $ev.Graphics
                $margins = $ev.MarginBounds
                $titleFont = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
                $subFont = New-Object System.Drawing.Font("Segoe UI", 10)
                $smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
                $headerFont = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
                $cellFont = New-Object System.Drawing.Font("Segoe UI", 9)
                $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::Gray, 0.5)
                $yPos = $margins.Top
                $rowHeight = 26

                # Spaltenbreiten
                $colMarkerW = [int]($margins.Width * 0.40)
                $colValueW = [int]($margins.Width * 0.15)
                $colUnitW = [int]($margins.Width * 0.12)
                $colRatingW = [int]($margins.Width * 0.13)
                $colNoteW = $margins.Width - $colMarkerW - $colValueW - $colUnitW - $colRatingW

                # Header auf erster Seite
                if ($script:printPageIdx -eq 0) {
                    $g.DrawString("Blutbild vom $dateFormatted", $titleFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += 30
                    $personalName = $script:data.Config.Personal.Name
                    if (-not [string]::IsNullOrWhiteSpace($personalName)) {
                        $g.DrawString("Patient: $personalName", $subFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                        $yPos += 20
                    }
                    $g.DrawString("Druckdatum: $(Get-Date -Format 'dd.MM.yyyy HH:mm')  |  Anzahl Marker: $($script:printItems.Count)", $smallFont, [System.Drawing.Brushes]::Gray, $margins.Left, $yPos)
                    $yPos += 25
                }
                $script:printPageIdx++

                # Spaltenheader
                $xPos = $margins.Left
                $headers = @(@{Text="Marker";W=$colMarkerW}, @{Text="Wert";W=$colValueW}, @{Text="Einheit";W=$colUnitW}, @{Text="Bewertung";W=$colRatingW}, @{Text="Notiz";W=$colNoteW})
                foreach ($h in $headers) {
                    $rect = New-Object System.Drawing.RectangleF($xPos, $yPos, $h.W, $rowHeight)
                    $g.FillRectangle([System.Drawing.Brushes]::LightSteelBlue, $rect)
                    $g.DrawRectangle($pen, [int]$rect.X, [int]$rect.Y, [int]$rect.Width, [int]$rect.Height)
                    $g.DrawString($h.Text, $headerFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 4))
                    $xPos += $h.W
                }
                $yPos += $rowHeight

                # Datenzeilen
                while ($script:printRowIdx -lt $script:printItems.Count) {
                    if (($yPos + $rowHeight) -gt $margins.Bottom) {
                        $ev.HasMorePages = $true
                        return
                    }
                    $item = $script:printItems[$script:printRowIdx]
                    $markerConfig = $script:data.Config.Markers | Where-Object { $_.Name -eq $item.Name } | Select-Object -First 1

                    # Bewertung berechnen
                    $ratingText = "-"
                    $ratingColor = [System.Drawing.Color]::White
                    if ($markerConfig) {
                        $val = $item.Value
                        if ($markerConfig.OptimalMin -ne $null -and $markerConfig.OptimalMax -ne $null -and $val -ge $markerConfig.OptimalMin -and $val -le $markerConfig.OptimalMax) {
                            $ratingText = "Optimal"; $ratingColor = [System.Drawing.Color]::LightGreen
                        } elseif ($markerConfig.RefMin -ne $null -and $markerConfig.RefMax -ne $null -and $val -ge $markerConfig.RefMin -and $val -le $markerConfig.RefMax) {
                            $ratingText = "Akzeptabel"; $ratingColor = [System.Drawing.Color]::LightYellow
                        } elseif (($markerConfig.RefMin -ne $null -and $val -lt $markerConfig.RefMin) -or ($markerConfig.RefMax -ne $null -and $val -gt $markerConfig.RefMax)) {
                            $ratingText = "Außerhalb"; $ratingColor = [System.Drawing.Color]::LightCoral
                        }
                    }

                    $noteText = if ($item.PSObject.Properties['Note'] -and $item.Note) { $item.Note } else { "" }
                    # v2.20.0: HIV/Testosteron Anzeige
                    $printValue = $item.Value.ToString()
                    $printUnit = $item.Unit
                    if ($item.Name -eq "HIV (Anti-HIV-1/2)") {
                        $printValue = if ([double]$item.Value -ge 1) { "reaktiv" } else { "nicht reaktiv" }
                        $printUnit = ""
                    } elseif ($item.Name -eq $script:ApoeMarkerName) {
                        # v2.26.0: Genotyp statt Zahlencode, Einheit entfaellt
                        $printValue = Get-ApoeGenotypeText -Value $item.Value
                        $printUnit = ""
                    } elseif ($item.Name -eq "Testosteron, gesamt") {
                        $nmol = [Math]::Round([double]$item.Value * 0.0347, 2)
                        $printValue = "$($item.Value) ($nmol nmol/l)"
                    }
                    $cellValues = @($item.Name, $printValue, $printUnit, $ratingText, $noteText)
                    $cellWidths = @($colMarkerW, $colValueW, $colUnitW, $colRatingW, $colNoteW)

                    $xPos = $margins.Left
                    for ($c = 0; $c -lt $cellValues.Count; $c++) {
                        $rect = New-Object System.Drawing.RectangleF($xPos, $yPos, $cellWidths[$c], $rowHeight)
                        if ($c -eq 3 -and -not $ratingColor.IsEmpty -and $ratingColor.Name -ne "White") {
                            $brush = New-Object System.Drawing.SolidBrush($ratingColor)
                            $g.FillRectangle($brush, $rect)
                            $brush.Dispose()
                        }
                        $g.DrawRectangle($pen, [int]$rect.X, [int]$rect.Y, [int]$rect.Width, [int]$rect.Height)
                        $g.DrawString($cellValues[$c], $cellFont, [System.Drawing.Brushes]::Black, ($xPos + 4), ($yPos + 5))
                        $xPos += $cellWidths[$c]
                    }
                    $yPos += $rowHeight
                    $script:printRowIdx++
                }
                $ev.HasMorePages = $false
            })

            $preview = New-Object System.Windows.Forms.PrintPreviewDialog
            $preview.Document = $printDoc
            $preview.Width = 900
            $preview.Height = 700
            $preview.ShowDialog()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Drucken: $($_.Exception.Message)", "Druckfehler", "OK", "Error")
        }
    })
    $chartExportButton.Add_Click({
        $dataProvider = {
            $markerName = $filterComboBox.SelectedItem
            if (-not $markerName) { return @() }
            $markerData = $null
            if ($script:data.Config.CalculatedMarkers.Name -contains $markerName) {
                $markerData = Get-CalculatedValuesForMarker -markerName $markerName
            } else {
                $markerData = $script:allHistoricalItems | Where-Object { $_.Name -eq $markerName }
            }
            if (-not $markerData) { return @() }
            $rows = $markerData | Sort-Object Date | ForEach-Object {
                [PSCustomObject]@{
                    Datum   = $_.Date
                    Marker  = $_.Name
                    Wert    = $_.Value
                    Einheit = $_.Unit
                    Notiz   = if ($_.PSObject.Properties.Name -contains "Note") { $_.Note } else { "" }
                }
            }
            return $rows
        }
        $selected = if ($filterComboBox.SelectedItem) { ($filterComboBox.SelectedItem -replace '[^\w\-]', '_') } else { "Einzelmarker" }
        Show-ExportPopup -source "Einzelmarker-Analyse" -dataProvider $dataProvider -defaultFileName "Marker_${selected}_$(Get-Date -Format 'yyyy-MM-dd')"
    })

    $chartPrintButton.Add_Click({
        try {
            $selectedMarker = $filterComboBox.SelectedItem
            if ($selectedMarker -eq "--- Bitte auswählen ---" -or [string]::IsNullOrEmpty($selectedMarker)) {
                [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst einen Marker aus.", "Kein Marker ausgewählt", "OK", "Information")
                return
            }

            $printDoc = New-Object System.Drawing.Printing.PrintDocument
            $printDoc.DefaultPageSettings.Landscape = $true
            $printDoc.Add_PrintPage({
                param($sender, $ev)
                $g = $ev.Graphics
                $margins = $ev.MarginBounds
                $titleFont = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
                $subFont = New-Object System.Drawing.Font("Segoe UI", 10)
                $smallFont = New-Object System.Drawing.Font("Segoe UI", 8)
                $warnFont = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
                $yPos = $margins.Top

                # Titel
                $g.DrawString("Einzelmarker-Analyse: $selectedMarker", $titleFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                $yPos += 30

                # Patient
                $personalName = $script:data.Config.Personal.Name
                if (-not [string]::IsNullOrWhiteSpace($personalName)) {
                    $g.DrawString("Patient: $personalName", $subFont, [System.Drawing.Brushes]::Black, $margins.Left, $yPos)
                    $yPos += 20
                }

                # Druckdatum
                $g.DrawString("Druckdatum: $(Get-Date -Format 'dd.MM.yyyy HH:mm')", $smallFont, [System.Drawing.Brushes]::Gray, $margins.Left, $yPos)
                $yPos += 20

                # Marker-Info
                $markerConfig = $script:data.Config.Markers | Where-Object { $_.Name -eq $selectedMarker } | Select-Object -First 1
                if (-not $markerConfig) { $markerConfig = $script:data.Config.CalculatedMarkers | Where-Object { $_.Name -eq $selectedMarker } | Select-Object -First 1 }
                if ($markerConfig) {
                    $infoText = "Einheit: $($markerConfig.Unit)"
                    if ($markerConfig.RefMin -ne $null -and $markerConfig.RefMax -ne $null) { $infoText += "  |  Referenz: $($markerConfig.RefMin) - $($markerConfig.RefMax)" }
                    if ($markerConfig.OptimalMin -ne $null -and $markerConfig.OptimalMax -ne $null) { $infoText += "  |  Optimal: $($markerConfig.OptimalMin) - $($markerConfig.OptimalMax)" }
                    $g.DrawString($infoText, $smallFont, [System.Drawing.Brushes]::DarkSlateGray, $margins.Left, $yPos)
                    $yPos += 18
                }
                $yPos += 5

                # Chart als Bild rendern
                $chartWidth = $margins.Width
                $chartHeight = [Math]::Min(($margins.Bottom - $yPos - 80), 500)
                $chartBitmap = New-Object System.Drawing.Bitmap($chart.Width, $chart.Height)
                $chart.DrawToBitmap($chartBitmap, (New-Object System.Drawing.Rectangle(0, 0, $chart.Width, $chart.Height)))
                $destRect = New-Object System.Drawing.Rectangle($margins.Left, $yPos, $chartWidth, $chartHeight)
                $g.DrawImage($chartBitmap, $destRect)
                $yPos += $chartHeight + 10
                $chartBitmap.Dispose()

                # Warnungen
                if ($warningDisplayGroup.Visible -and -not [string]::IsNullOrWhiteSpace($warningLabel.Text)) {
                    $warnLines = $warningLabel.Text -split "`r`n|`n"
                    foreach ($line in $warnLines) {
                        if (-not [string]::IsNullOrWhiteSpace($line)) {
                            $g.DrawString($line, $warnFont, [System.Drawing.Brushes]::DarkRed, $margins.Left, $yPos)
                            $yPos += 16
                        }
                    }
                }

                $ev.HasMorePages = $false
            })

            $preview = New-Object System.Windows.Forms.PrintPreviewDialog
            $preview.Document = $printDoc
            $preview.Width = 1000
            $preview.Height = 700
            $preview.ShowDialog()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Drucken: $($_.Exception.Message)", "Druckfehler", "OK", "Error")
        }
    })
    $editButton.Add_Click({
        $selectedMarker = $filterComboBox.SelectedItem
        $isCalculated = ([System.Collections.ArrayList]@($script:data.Config.CalculatedMarkers.Name) -contains $selectedMarker)
        if($isCalculated){ [System.Windows.Forms.MessageBox]::Show("Berechnete Marker können nicht direkt bearbeitet werden.", "Hinweis", "OK", "Information"); return }
        if ($selectedMarker -eq "--- Bitte auswählen ---" -or [string]::IsNullOrEmpty($selectedMarker)) { [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst einen Marker aus.", "Kein Marker ausgewählt", "OK", "Information") } 
        else {
            if ((Show-EditPopup -markerName $selectedMarker -historicalData $script:allHistoricalItems) -eq [System.Windows.Forms.DialogResult]::OK) {
                Save-AllHistoricalData -AllItems $script:allHistoricalItems; Update-FilterDropdown
                $filterComboBox.SelectedItem = $selectedMarker; Update-Chart
            }
        }
    })
    $filterComboBox.Add_SelectedIndexChanged({ Update-Chart })
    $chart.Add_MouseMove({
        param($sender, $e)
        $result = $chart.HitTest($e.X, $e.Y)
        if ($result.ChartElementType -eq 'DataPoint') {
            $dataPoint = $result.Series.Points[$result.PointIndex]
            $date = [datetime]::FromOADate($dataPoint.XValue)
            $selectedMarker = $filterComboBox.SelectedItem
            $historicalItem = $script:allHistoricalItems | Where-Object { $_.Name -eq $selectedMarker -and ([datetime]$_.Date) -eq $date } | Select -First 1
            # v2.20.0: Wertanzeige für HIV/Testosteron
            $displayValue = $dataPoint.YValues[0]
            if ($selectedMarker -eq "HIV (Anti-HIV-1/2)") {
                $displayValue = if ([double]$dataPoint.YValues[0] -ge 1) { "reaktiv" } else { "nicht reaktiv" }
            } elseif ($selectedMarker -eq $script:ApoeMarkerName) {
                # v2.26.0: Code 1-6 als Genotyp anzeigen
                $displayValue = Get-ApoeGenotypeText -Value $dataPoint.YValues[0]
            } elseif ($selectedMarker -eq "Testosteron, gesamt") {
                $nmol = [Math]::Round([double]$dataPoint.YValues[0] * 0.0347, 2)
                $displayValue = "$($dataPoint.YValues[0]) ng/dl ($nmol nmol/l)"
            }
            if ($historicalItem -and $historicalItem.PSObject.Properties['Note'] -and -not [string]::IsNullOrWhiteSpace($historicalItem.Note)) {
                $toolTipText = "$($date.ToString('dd.MM.yyyy')): $displayValue`nNotiz: $($historicalItem.Note)"
                $toolTip.SetToolTip($chart, $toolTipText)
            } else { $toolTip.SetToolTip($chart, "$($date.ToString('dd.MM.yyyy')): $displayValue") }
        } else { $toolTip.SetToolTip($chart, "") }
    })

    $addButton.Add_Click({ 
        try { 
            $selectedDateString = $datePicker.Value.ToString("yyyy-MM-dd"); 
            # v2.23.0: Eingabe zuerst aufloesen (Liste ODER Alias-Direkteingabe).
            # Behebt zugleich das Risiko, dass SelectedItem noch auf einem alten
            # Marker stand, waehrend im Textfeld bereits ein anderer stand.
            Commit-MarkerSelection
            $markerName = $script:currentMarkerName
            if ([string]::IsNullOrWhiteSpace($markerName)) {
                $typedRaw = "$($markerComboBox.Text)".Trim()
                if ($typedRaw) {
                    $cand = @(Get-MarkerMatches -Query $typedRaw)
                    if ($cand.Count -gt 1) {
                        throw "Die Eingabe '$typedRaw' ist nicht eindeutig ($($cand.Count) Treffer, z. B. '$($cand[0])'). Bitte den gewuenschten Marker aus der Liste waehlen."
                    }
                    throw "Der Blutmarker '$typedRaw' ist unbekannt. Bitte aus der Liste waehlen oder ein gaengiges Kuerzel eingeben (z. B. WBC, HbA1c, GOT)."
                }
                throw "Bitte wählen Sie einen Blutmarker aus."
            }
            
            # v2.20.0: HIV → Wert aus ComboBox (0 = nicht reaktiv, 1 = reaktiv)
            if ($markerName -eq "HIV (Anti-HIV-1/2)") {
                $value = if ($hivComboBox.SelectedItem -eq "reaktiv") { 1 } else { 0 }
            }
            # v2.26.0: APOE → Genotyp aus ComboBox als Code 1-6 speichern
            elseif ($markerName -eq $script:ApoeMarkerName) {
                if ($apoeComboBox.SelectedIndex -lt 0) { throw "Bitte einen APOE-Genotyp aus der Liste wählen." }
                $value = $apoeComboBox.SelectedIndex + 1
            }
            # v2.20.0: Testosteron → bei nmol/l nach ng/dl umrechnen
            elseif ($markerName -eq "Testosteron, gesamt" -and $testoUnitCombo.SelectedItem -eq "nmol/l") {
                $rawValue = Parse-Number $valueTextBox.Text
                if ($null -eq $rawValue) { throw "Der Wert darf nicht leer sein." }
                $value = [Math]::Round($rawValue * 28.818, 1) # nmol/l → ng/dl
            }
            else {
                $value = Parse-Number $valueTextBox.Text
                if ($null -eq $value) { throw "Der Wert darf nicht leer sein." }
            }
            
            $markerConfig = $script:data.Config.Markers | Where-Object { $_.Name -eq $markerName } | Select-Object -First 1
            if (-not $markerConfig) { throw "Fehler: Die Konfiguration für den Marker '$markerName' konnte nicht gefunden werden." }
            
            # KORREKTUR: Robuste Duplikatsprüfung mit korrekter Where-Object Syntax
            $existingEntries = @($script:allHistoricalItems | Where-Object { ($_.Date -eq $selectedDateString) -and ($_.Name -eq $markerName) })
            if ($existingEntries.Count -gt 0) { 
                throw "Für '$markerName' existiert am $selectedDateString bereits ein Eintrag. Bitte in Tab 'Daten nach Bluttests' oder über 'bearbeiten' ändern."
            }
            
            # v2.20.0: Note ergänzen bei Testosteron-Umrechnung
            $note = $noteTextBox.Text
            if ($markerName -eq "Testosteron, gesamt" -and $testoUnitCombo.SelectedItem -eq "nmol/l") {
                $note = if ($note) { "$note | Eingabe: $($valueTextBox.Text) nmol/l" } else { "Eingabe: $($valueTextBox.Text) nmol/l" }
            }
            if ($markerName -eq "HIV (Anti-HIV-1/2)") {
                $note = if ($note) { "$note | $($hivComboBox.SelectedItem)" } else { $hivComboBox.SelectedItem }
            }
            # v2.26.0: Genotyp zusätzlich im Klartext in der Notiz sichern
            if ($markerName -eq $script:ApoeMarkerName) {
                $apoeText = "APOE-Genotyp: $($apoeComboBox.SelectedItem)"
                $note = if ($note) { "$note | $apoeText" } else { $apoeText }
            }
            
            $newItem = [PSCustomObject]@{ Date = $selectedDateString; Name = $markerName; Value = $value; Unit = $markerConfig.Unit; Note = $note }
            
            Save-Data -data @{ NewItem = $newItem } -type "Daily" -dateString $selectedDateString
            
            $script:allHistoricalItems = [System.Collections.ArrayList]@(Load-AllHistoricalItems)
            
            Update-FilterDropdown
            $filterComboBox.SelectedItem = $newItem.Name; $valueTextBox.Text = ""; $noteTextBox.Text = ""
            Update-Chart
            
        } catch { 
            [System.Windows.Forms.MessageBox]::Show("Fehler bei der Eingabe: $($_.Exception.Message)") 
        } 
    })
    
    $addNewMarkerButton.Add_Click({ 
        try { 
            $name = $textNewName.Text.Trim(); if ([string]::IsNullOrWhiteSpace($name)) { throw "Der Marker-Name darf nicht leer sein." }
            
            $currentMarkers = [System.Collections.ArrayList]@($script:data.Config.Markers | ForEach-Object { [PSCustomObject]$_ })
            $existingMarker = $currentMarkers | Where-Object { $_.Name -eq $name } | Select-Object -First 1

            if ($existingMarker) {
                $existingMarker.Unit = $textNewUnit.Text; 
                $existingMarker.Group = $textNewGroup.Text; 
                $existingMarker.RefMin = Parse-Number $textNewMin.Text; 
                $existingMarker.RefMax = Parse-Number $textNewMax.Text; 
                $existingMarker.OptimalMin = Parse-Number $textNewOptMin.Text; 
                $existingMarker.OptimalMax = Parse-Number $textNewOptMax.Text
                $msg = "Marker '$name' wurde aktualisiert."
            } else {
                $newMarker = [PSCustomObject]@{ Name = $name; Unit = $textNewUnit.Text; Group = $textNewGroup.Text; RefMin = Parse-Number $textNewMin.Text; RefMax = Parse-Number $textNewMax.Text; OptimalMin = Parse-Number $textNewOptMin.Text; OptimalMax = Parse-Number $textNewOptMax.Text; Description = "Vom Nutzer definierter Marker." }
                $script:data.Config.Markers.Add($newMarker); $msg = "Marker '$name' wurde hinzugefügt."
            }
            
            Save-Data -data $script:data -type "Config"; Update-AllMarkerDropdowns; [System.Windows.Forms.MessageBox]::Show($msg)
            $textNewName.Text = ""; $textNewUnit.Text = ""; $textNewGroup.Text = ""; $textNewMin.Text = ""; $textNewMax.Text = ""; $textNewOptMin.Text = ""; $textNewOptMax.Text = ""
        } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern des Markers: $($_.Exception.Message)") } 
    })
    $loadMarkerForEditButton.Add_Click({
        # v2.26.0 (Aufgabe #1): Nach dem Laden wird "Erweiterte Einstellungen"
        # automatisch aktiviert. Bisher wurden Gruppe, Ref-Min/Max und
        # Optimal-Min/Max zwar befuellt, lagen aber in ausgeblendeten Feldern.
        $selectedName = [string]$deleteMarkerComboBox.SelectedItem
        if ([string]::IsNullOrWhiteSpace($selectedName)) { $selectedName = "$($deleteMarkerComboBox.Text)".Trim() }
        if ([string]::IsNullOrWhiteSpace($selectedName)) {
            [System.Windows.Forms.MessageBox]::Show("Bitte zuerst einen Blutmarker aus der Liste auswählen.", "Kein Marker ausgewählt", "OK", "Information")
            return
        }
        $markerToLoad = $script:data.Config.Markers | Where-Object { $_.Name -eq $selectedName } | Select-Object -First 1
        if (-not $markerToLoad) {
            [System.Windows.Forms.MessageBox]::Show("Der Marker '$selectedName' wurde in der Konfiguration nicht gefunden.", "Marker nicht gefunden", "OK", "Warning")
            return
        }
        $textNewName.Text = $markerToLoad.Name; $textNewUnit.Text = $markerToLoad.Unit; $textNewGroup.Text = $markerToLoad.Group; $textNewMin.Text = $markerToLoad.RefMin; $textNewMax.Text = $markerToLoad.RefMax; $textNewOptMin.Text = $markerToLoad.OptimalMin; $textNewOptMax.Text = $markerToLoad.OptimalMax

        # Erweiterte Ansicht erzwingen. Add_CheckedChanged feuert nur bei einer
        # echten Zustandsaenderung - ist die Checkbox bereits gesetzt, ist das
        # Layout ohnehin schon aufgeklappt.
        if (-not $advancedViewCheckBox.Checked) { $advancedViewCheckBox.Checked = $true }
        $textNewName.Focus()
    })
    $deleteExistingMarkerButton.Add_Click({ 
        try { 
            $markerToDelete = $deleteMarkerComboBox.SelectedItem; if ([string]::IsNullOrEmpty($markerToDelete)) { throw "Bitte einen Marker zum Löschen auswählen."}
            if (([System.Windows.Forms.MessageBox]::Show("Möchtest du den Marker '$markerToDelete' und alle zugehörigen Daten wirklich dauerhaft löschen?", "Löschen bestätigen", "YesNo", "Warning")) -eq "Yes") { 
                $markerObject = $script:data.Config.Markers | Where-Object { $_.Name -eq $markerToDelete } | Select-Object -First 1
                $script:data.Config.Markers.Remove($markerObject)
                
                $itemsToRemove = $script:allHistoricalItems | Where-Object { $_.Name -eq $markerToDelete }
                foreach($item in $itemsToRemove){ $script:allHistoricalItems.Remove($item) }
                
                Save-Data -data $script:data -type "Config"; Save-AllHistoricalData -AllItems $script:allHistoricalItems
                Update-AllMarkerDropdowns; Update-FilterDropdown; Update-Chart
                [System.Windows.Forms.MessageBox]::Show("Marker '$markerToDelete' und alle seine Daten wurden gelöscht.") 
            } 
        } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Löschen des Markers: $($_.Exception.Message)") } 
    })
    $saveSettingsButton.Add_Click({
        try {
            $thresholdValue = Parse-Number $thresholdTextBox.Text
            $script:data.Config.WarningThreshold = $thresholdValue
            Save-Data -data $script:data -type "Config"
            [System.Windows.Forms.MessageBox]::Show("Der Schwellenwert wurde erfolgreich gespeichert.", "Gespeichert", "OK", "Information")
            Update-Chart
        } catch { [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern: $($_.Exception.Message)", "Fehler", "OK", "Error") }
    })
    $thresholdTextBox.Add_TextChanged({ Update-Chart })

    # ---------- Event-Handler für Tab 4 ----------
    # v2.24.1: Zentrale Aktualisierung des Buttons "Dokument anzeigen".
    # Ersetzt die bisher dreifach duplizierte (und fehlerhafte) Inline-Logik.
    # Tag = PSCustomObject mit Date + Files (vollständige Pfade, Upload-Reihenfolge).
    $UpdateShowDocButton = {
        param($dateString)
        $showDocButton.Visible = $false
        $showDocButton.Tag     = $null
        $showDocButton.Text    = "Dokument`nanzeigen"
        $showDocToolTip.SetToolTip($showDocButton, "")
        if ([string]::IsNullOrWhiteSpace($dateString)) { return }
        try {
            $docs = @(Get-BloodTestDocuments -DateString $dateString)
            if ($docs.Count -eq 0) { return }

            $showDocButton.Tag = [PSCustomObject]@{
                Date  = $dateString
                Files = @($docs | ForEach-Object { $_.FullName })
            }
            $showDocButton.Text = if ($docs.Count -gt 1) { "Dokument`nanzeigen ($($docs.Count))" } else { "Dokument`nanzeigen" }

            $dateFmt = try { ([datetime]::ParseExact($dateString, 'yyyy-MM-dd', $null)).ToString('dd.MM.yyyy') } catch { $dateString }
            $tipList = ($docs | ForEach-Object { " • $($_.Name)  ($([Math]::Round($_.Length / 1KB, 0)) KB)" }) -join "`r`n"
            $showDocToolTip.SetToolTip($showDocButton, "Bluttest $dateFmt – $($docs.Count) Dokument(e):`r`n$tipList")

            $showDocButton.Visible = $true
        } catch {
            Write-Warning "Aktualisierung Dokument-Button: $($_.Exception.Message)"
        }
    }
    $dataMgmtTreeView.Add_AfterSelect({
        $dataMgmtDataGridView.Rows.Clear()
        $selectedNode = $dataMgmtTreeView.SelectedNode
        if (-not $selectedNode) { $dataMgmtExportButton.Enabled = $false; $dataMgmtUploadButton.Enabled = $false; & $UpdateShowDocButton $null; return }
        $selectedDate = $selectedNode.Tag
        $itemsForDate = $script:allHistoricalItems | Where-Object { $_.Date -eq $selectedDate } | Sort-Object Name
        foreach ($item in $itemsForDate) {
            # v2.20.0: HIV als Text, Testosteron mit doppelter Einheit
            $displayValue = $item.Value
            $displayUnit = $item.Unit
            if ($item.Name -eq "HIV (Anti-HIV-1/2)") {
                $displayValue = if ([double]$item.Value -ge 1) { "reaktiv" } else { "nicht reaktiv" }
                $displayUnit = ""
            } elseif ($item.Name -eq $script:ApoeMarkerName) {
                # v2.26.0: Genotyp statt Zahlencode
                $displayValue = Get-ApoeGenotypeText -Value $item.Value
                $displayUnit = ""
            } elseif ($item.Name -eq "Testosteron, gesamt") {
                $nmol = [Math]::Round([double]$item.Value * 0.0347, 2)
                $displayValue = "$($item.Value) ($nmol nmol/l)"
            }
            $rowIndex = $dataMgmtDataGridView.Rows.Add($item.Name, $displayValue, $displayUnit, $item.Note)
            $dataMgmtDataGridView.Rows[$rowIndex].Tag = $item
        }
        # Export-Button nur aktivieren wenn tatsächlich Daten vorliegen
        $dataMgmtExportButton.Enabled = ($itemsForDate | Measure-Object).Count -gt 0
        # v2.17.0: Upload-Button aktivieren wenn Datum ausgewählt
        $dataMgmtUploadButton.Enabled = $true

        # v2.24.1: "Dokument anzeigen" strikt für das gewählte Bluttest-DATUM
        # (bisher wurden alle Dokumente des Monatsordners gezählt/geöffnet)
        & $UpdateShowDocButton $selectedDate
    })
    # v2.17.0 / v2.24.1: Dokument-Upload Handler (jetzt mit Mehrfachauswahl)
    $dataMgmtUploadButton.Add_Click({
        $selectedNode = $dataMgmtTreeView.SelectedNode
        if (-not $selectedNode) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst einen Test-Tag aus der Liste aus.", "Keine Auswahl", "OK", "Information")
            return
        }
        $selectedDate = $selectedNode.Tag
        $dateFmt = try { ([datetime]::ParseExact($selectedDate, 'yyyy-MM-dd', $null)).ToString('dd.MM.yyyy') } catch { $selectedDate }

        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = "Dokument(e) zum Bluttest vom $dateFmt hochladen"
        $ofd.Filter = "Alle Dateien (*.*)|*.*"
        $ofd.Multiselect = $true   # v2.24.1: mehrere Dokumente je Bluttest
        if ($ofd.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $copied = New-Object System.Collections.ArrayList
            $failed = New-Object System.Collections.ArrayList
            try {
                $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $selectedDate
                if (-not (Test-Path $dailyDataDir)) { New-Item -Path $dailyDataDir -ItemType Directory -Force | Out-Null }

                foreach ($sourceFile in @($ofd.FileNames)) {
                    try {
                        $ext = [System.IO.Path]::GetExtension($sourceFile)
                        # Schema: YYYY-MM-dd_BLUTTEST[_n].ext  -> Datum bleibt eindeutiger Schlüssel
                        $targetName = "${selectedDate}_BLUTTEST${ext}"
                        $targetPath = Join-Path -Path $dailyDataDir -ChildPath $targetName
                        $counter = 1
                        while (Test-Path $targetPath) {
                            $targetName = "${selectedDate}_BLUTTEST_${counter}${ext}"
                            $targetPath = Join-Path -Path $dailyDataDir -ChildPath $targetName
                            $counter++
                        }
                        Copy-Item -LiteralPath $sourceFile -Destination $targetPath -Force
                        [void]$copied.Add($targetName)
                    } catch {
                        [void]$failed.Add("$([System.IO.Path]::GetFileName($sourceFile)): $($_.Exception.Message)")
                    }
                }

                # v2.24.1: Button/Zähler/Tooltip zentral neu aufbauen
                & $UpdateShowDocButton $selectedDate

                $totalDocs = @(Get-BloodTestDocuments -DateString $selectedDate).Count
                if ($copied.Count -gt 0) {
                    $msg = "$($copied.Count) Dokument(e) dem Bluttest vom $dateFmt zugeordnet:`n`n" +
                           (($copied | ForEach-Object { " • $_" }) -join "`n") +
                           "`n`nAblage: $dailyDataDir" +
                           "`nDokumente zu diesem Bluttest insgesamt: $totalDocs"
                    if ($failed.Count -gt 0) { $msg += "`n`nFehlgeschlagen:`n" + (($failed | ForEach-Object { " • $_" }) -join "`n") }
                    [System.Windows.Forms.MessageBox]::Show($msg, "Upload erfolgreich", "OK", "Information")
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Es konnte kein Dokument archiviert werden.`n`n" + (($failed | ForEach-Object { " • $_" }) -join "`n"), "Fehler", "OK", "Error")
                }
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Hochladen des Dokuments: $($_.Exception.Message)", "Fehler", "OK", "Error")
            }
        }
    })
    # v2.22.0 / v2.24.1: "Dokument anzeigen" – öffnet NUR Dokumente des gewählten Bluttests.
    # 1 Dokument  -> direkt öffnen
    # n Dokumente -> Auswahlmenü (Name, Größe, Zeitstempel) + "Ordner öffnen"
    $showDocButton.Add_Click({
        $docInfo = $showDocButton.Tag
        $docDate = if ($docInfo -and $docInfo.Date) { $docInfo.Date } else { $null }

        # Immer frisch vom Dateisystem lesen (Dokument könnte extern gelöscht worden sein)
        $docs = @()
        if ($docDate) { $docs = @(Get-BloodTestDocuments -DateString $docDate) }

        if ($docs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Für diesen Bluttest wurde kein Dokument gefunden.", "Kein Dokument", "OK", "Information")
            & $UpdateShowDocButton $docDate
            return
        }

        if ($docs.Count -eq 1) {
            try { Start-Process -FilePath $docs[0].FullName }
            catch { [System.Windows.Forms.MessageBox]::Show("Dokument konnte nicht geöffnet werden: $($_.Exception.Message)", "Fehler", "OK", "Error") }
            return
        }

        $docMenu = New-Object System.Windows.Forms.ContextMenuStrip
        $idx = 0
        foreach ($doc in $docs) {
            $idx++
            $sizeKb = [Math]::Round($doc.Length / 1KB, 0)
            $menuItem = New-Object System.Windows.Forms.ToolStripMenuItem
            $menuItem.Text = "$idx.  $($doc.Name)   —   $sizeKb KB, $($doc.LastWriteTime.ToString('dd.MM.yyyy HH:mm'))"
            $menuItem.Tag  = $doc.FullName
            $menuItem.Add_Click({
                param($menuSender, $menuArgs)
                $path = $menuSender.Tag
                if ($path -and (Test-Path $path)) {
                    try { Start-Process -FilePath $path }
                    catch { [System.Windows.Forms.MessageBox]::Show("Dokument konnte nicht geöffnet werden: $($_.Exception.Message)", "Fehler", "OK", "Error") }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Datei nicht gefunden:`n$path", "Fehler", "OK", "Warning")
                }
            })
            [void]$docMenu.Items.Add($menuItem)
        }
        [void]$docMenu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator))
        $folderItem = New-Object System.Windows.Forms.ToolStripMenuItem
        $folderItem.Text = "Ordner öffnen"
        $folderItem.Tag  = (Split-Path -Parent $docs[0].FullName)
        $folderItem.Add_Click({
            param($menuSender, $menuArgs)
            $folder = $menuSender.Tag
            if ($folder -and (Test-Path $folder)) { Start-Process -FilePath "explorer.exe" -ArgumentList "`"$folder`"" }
        })
        [void]$docMenu.Items.Add($folderItem)

        $docMenu.Show($showDocButton, (New-Object System.Drawing.Point(0, $showDocButton.Height)))
    })
    # v2.19.0: PDF-Import Handler (Beta Version) – Nachtrag für Template-Script
    $pdfImportButton.Add_Click({
        Import-PdfBloodValues
        # Tabs aktualisieren
        Update-DataManagementTab
        Update-Chart
        # v2.24.1: Dokument-Button nach dem Import (PDF wird mit archiviert) neu aufbauen
        if ($dataMgmtTreeView.SelectedNode) { & $UpdateShowDocButton $dataMgmtTreeView.SelectedNode.Tag }
    })
    # v2.19.0: Verschlüsseltes Backup exportieren
    $encExportButton.Add_Click({
        $passphrase = Show-PassphraseDialog -Title "Passwort für verschlüsseltes Backup" -Confirm $true
        if (-not $passphrase) { return }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "Verschlüsseltes Backup speichern"
        $sfd.Filter = "Blood-Tracker Backup (*.btbackup)|*.btbackup"
        $sfd.FileName = "BloodTracker_Backup_$(Get-Date -Format 'yyyy-MM-dd').btbackup"
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $cursor = $form.Cursor
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Export-AesBackup -OutputPath $sfd.FileName -Passphrase $passphrase
            $form.Cursor = $cursor
            [System.Windows.Forms.MessageBox]::Show(
                "Verschlüsseltes Backup erfolgreich erstellt:`n$($sfd.FileName)`n`nDieses Backup kann auf jedem PC mit dem Passwort importiert werden.",
                "Export erfolgreich", "OK", "Information"
            )
        } catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Erstellen des Backups: $($_.Exception.Message)", "Export-Fehler", "OK", "Error")
        }
    })
    # v2.21.1: Unverschluesselter Export fuer PC-Migration (Klartext-ZIP)
    $plainExportButton.Add_Click({
        $warn = [System.Windows.Forms.MessageBox]::Show(
            "Dieser Export erstellt ein UNVERSCHLÜSSELTES ZIP mit allen Blutwerten und persönlichen Daten im Klartext.`n`nEs ist für die Migration auf einen anderen PC gedacht und sollte nach der Übertragung sicher gelöscht werden.`n`nFortfahren?",
            "Unverschlüsselter Export", "YesNo", "Warning"
        )
        if ($warn -ne "Yes") { return }

        $sfd = New-Object System.Windows.Forms.SaveFileDialog
        $sfd.Title = "Unverschlüsselten Export speichern"
        $sfd.Filter = "ZIP-Archiv (*.zip)|*.zip"
        $sfd.FileName = "BloodTracker_Migration_$(Get-Date -Format 'yyyy-MM-dd').zip"
        if ($sfd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        try {
            $cursor = $form.Cursor
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Export-PlaintextBackup -OutputPath $sfd.FileName
            $form.Cursor = $cursor
            [System.Windows.Forms.MessageBox]::Show(
                "Unverschlüsselter Export erstellt:`n$($sfd.FileName)`n`nMigration: Config.json und den Ordner 'data' auf dem neuen PC in den Datenordner (%UserProfile%\PSC.Blood-Tracker) entpacken. Details siehe LIESMICH_Migration.txt im ZIP.`n`nBitte das ZIP nach der Übertragung sicher löschen.",
                "Export erfolgreich", "OK", "Information"
            )
        } catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Export: $($_.Exception.Message)", "Export-Fehler", "OK", "Error")
        }
    })
    # v2.19.0: Backup importieren
    $encImportButton.Add_Click({
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "ACHTUNG: Der Import überschreibt alle vorhandenen Daten!`n`nVorher ein Backup erstellen, falls Sie unsicher sind.`n`nFortfahren?",
            "Import bestätigen", "YesNo", "Warning"
        )
        if ($confirm -ne "Yes") { return }

        $ofd = New-Object System.Windows.Forms.OpenFileDialog
        $ofd.Title = "Verschlüsseltes Backup öffnen"
        $ofd.Filter = "Blood-Tracker Backup (*.btbackup)|*.btbackup"
        if ($ofd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

        $passphrase = Show-PassphraseDialog -Title "Backup-Passwort eingeben" -Confirm $false
        if (-not $passphrase) { return }

        try {
            $cursor = $form.Cursor
            $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            Import-AesBackup -InputPath $ofd.FileName -Passphrase $passphrase
            $form.Cursor = $cursor

            # Daten neu laden
            $script:data = @{ Config = Load-Config }
            $script:allHistoricalItems = [System.Collections.ArrayList]@(Load-AllHistoricalItems)
            Load-PersonalMetrics
            Update-AllMarkerDropdowns
            Update-FilterDropdown
            Update-CorrelationDropdowns
            Update-Chart
            Update-DataManagementTab

            [System.Windows.Forms.MessageBox]::Show(
                "Backup erfolgreich importiert!`n`nAlle Daten wurden mit Ihrem lokalen Profil verschlüsselt.`nAb jetzt ist keine Passworteingabe mehr nötig.",
                "Import erfolgreich", "OK", "Information"
            )
        } catch {
            $form.Cursor = [System.Windows.Forms.Cursors]::Default
            $errMsg = $_.Exception.Message
            if ($errMsg -match "Padding|PKCS7|Auffüllung|Decrypt") {
                [System.Windows.Forms.MessageBox]::Show("Das Passwort ist falsch oder die Backup-Datei ist beschädigt.", "Import-Fehler", "OK", "Error")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Importieren: $errMsg", "Import-Fehler", "OK", "Error")
            }
        }
    })
    $saveChangesButton.Add_Click({
        $selectedNode = $dataMgmtTreeView.SelectedNode
        if (-not $selectedNode) { 
            [System.Windows.Forms.MessageBox]::Show("Es ist kein Test-Tag zum Speichern ausgewählt.", "Keine Auswahl", "OK", "Information")
            return 
        }
        $changesMade = $false
        try {
            foreach ($row in $dataMgmtDataGridView.Rows) {
                $originalItem = $row.Tag
                if (-not $originalItem) { continue } 

                $newValueStr = $row.Cells["Value"].Value
                $newValue = Parse-Number $newValueStr
                if ($newValue -eq $null) { throw "Ungültiger oder leerer Wert in Zeile für '$($originalItem.Name)'." }
                
                $newNote = if ($row.Cells["Note"].Value) { $row.Cells["Note"].Value } else { "" }

                if ($originalItem.Value -ne $newValue) {
                    $originalItem.Value = $newValue
                    $changesMade = $true
                }
                if ($originalItem.Note -ne $newNote) {
                    $originalItem.Note = $newNote
                    $changesMade = $true
                }
            }

            if ($changesMade) {
                Save-AllHistoricalData -AllItems $script:allHistoricalItems
                Update-FilterDropdown
                Update-Chart
                [System.Windows.Forms.MessageBox]::Show("Alle Änderungen für den ausgewählten Tag wurden erfolgreich gespeichert.", "Erfolg", "OK", "Information")
            } else {
                [System.Windows.Forms.MessageBox]::Show("Es wurden keine Änderungen zum Speichern festgestellt.", "Information", "OK", "Information")
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern der Änderungen: $($_.Exception.Message)", "Validierungsfehler", "OK", "Error")
        }
    })
    $deleteSelectedRowsButton.Add_Click({
        if ($dataMgmtDataGridView.SelectedRows.Count -eq 0) {
             [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst eine oder mehrere Zeilen zum Löschen aus.", "Keine Auswahl", "OK", "Information")
            return
        }
        $count = $dataMgmtDataGridView.SelectedRows.Count
        $confirmation = [System.Windows.Forms.MessageBox]::Show("Möchten Sie wirklich die ausgewählten $count Einträge dauerhaft löschen?", "Löschen bestätigen", "YesNo", "Warning")
        
        if ($confirmation -eq "Yes") {
            # Betroffene Tage sammeln und Einträge aus Speicher entfernen
            $affectedDates = New-Object System.Collections.ArrayList
            foreach ($row in $dataMgmtDataGridView.SelectedRows) {
                $item = $row.Tag
                if ($item -and -not $affectedDates.Contains($item.Date)) {
                    $affectedDates.Add($item.Date) | Out-Null
                }
                $script:allHistoricalItems.Remove($item)
            }
            
            # Nur die betroffenen Tages-Dateien gezielt aktualisieren oder löschen
            foreach ($dateString in $affectedDates) {
                $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $dateString
                $remainingItems = @($script:allHistoricalItems | Where-Object { $_.Date -eq $dateString })
                
                if ($remainingItems.Count -gt 0) {
                    # Datei mit verbleibenden Einträgen neu schreiben
                    $savableItems = $remainingItems | ForEach-Object {
                        $obj = [PSCustomObject]@{ Date = $_.Date; Name = $_.Name; Value = [double]$_.Value; Unit = $_.Unit }
                        if ($_.PSObject.Properties['Note']) { $obj | Add-Member -NotePropertyName 'Note' -NotePropertyValue $_.Note -Force }
                        $obj
                    }
                    # v2.19.0: DPAPI-verschlüsselt schreiben
                    $jsonString = @{ Items = $savableItems } | ConvertTo-Json -Depth 5
                    Write-ProtectedJsonFile -Path $dailyDataFile -JsonString $jsonString
                } else {
                    # Keine Einträge mehr → Datei löschen
                    if (Test-Path $dailyDataFile) { Remove-Item -Path $dailyDataFile -Force }
                }
            }
            
            $selectedNode = $dataMgmtTreeView.SelectedNode
            Update-DataManagementTab
            if ($selectedNode) {
                $nodeToSelect = $dataMgmtTreeView.Nodes | Where-Object { $_.Tag -eq $selectedNode.Tag } | Select-Object -First 1
                if ($nodeToSelect) { $dataMgmtTreeView.SelectedNode = $nodeToSelect }
            }
            Update-FilterDropdown
            Update-Chart
            [System.Windows.Forms.MessageBox]::Show("$count Einträge wurden gelöscht.", "Erfolg", "OK", "Information")
        }
    })
    $deleteTestButton.Add_Click({
        $selectedNode = $dataMgmtTreeView.SelectedNode
        if (-not $selectedNode) { 
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie zuerst einen Test-Tag aus der Liste aus.", "Keine Auswahl", "OK", "Information")
            return 
        }
        $dateToDelete = $selectedNode.Tag
        $itemsToDelete = @($script:allHistoricalItems | Where-Object { $_.Date -eq $dateToDelete })
        $dateFormatted = ([datetime]$dateToDelete).ToString("dd.MM.yyyy")

        $confirmation = [System.Windows.Forms.MessageBox]::Show("Möchten Sie wirklich alle $($itemsToDelete.Count) Einträge vom $dateFormatted dauerhaft löschen?", "Löschen bestätigen", "YesNo", "Warning")
        if ($confirmation -eq "Yes") {
            foreach ($item in $itemsToDelete) {
                $script:allHistoricalItems.Remove($item)
            }
            # Tages-Datei direkt löschen
            $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $dateToDelete
            if (Test-Path $dailyDataFile) { Remove-Item -Path $dailyDataFile -Force }
            Save-AllHistoricalData -AllItems $script:allHistoricalItems
            Update-DataManagementTab
            Update-FilterDropdown
            Update-Chart
            [System.Windows.Forms.MessageBox]::Show("Alle Einträge vom $dateFormatted wurden gelöscht.", "Erfolg", "OK", "Information")
        }
    })

    # --- NEU: Event-Handler für Tab 5 ---
    $savePersonalButton.Add_Click({ Save-PersonalMetrics })

    # v2.17.0: Genetische Vorbelastung hinzufügen
    $geneticNewAddButton.Add_Click({
        $name = $geneticNewNameText.Text.Trim()
        if (-not $name) {
            [System.Windows.Forms.MessageBox]::Show("Bitte geben Sie einen Namen für die Vorbelastung ein.", "Fehlende Eingabe", "OK", "Warning")
            return
        }
        # Prüfen ob Name bereits existiert
        $allPreds = Get-AllGeneticPredispositions
        if ($allPreds | Where-Object { $_['Name'] -eq $name }) {
            [System.Windows.Forms.MessageBox]::Show("Eine Vorbelastung mit diesem Namen existiert bereits.", "Duplikat", "OK", "Warning")
            return
        }
        # Ausgewählte Marker sammeln
        $selectedMarkers = @()
        for ($i = 0; $i -lt $geneticNewMarkerList.Items.Count; $i++) {
            if ($geneticNewMarkerList.GetItemChecked($i)) {
                $selectedMarkers += $geneticNewMarkerList.Items[$i]
            }
        }
        if ($selectedMarkers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie mindestens einen Blutmarker aus.", "Fehlende Auswahl", "OK", "Warning")
            return
        }
        # Neue Vorbelastung erstellen
        $newEntry = @{
            Name    = $name
            Active  = $false
            Markers = $selectedMarkers
            Hint    = "Vorbelastung ($name): Achten Sie besonders auf diesen Marker."
        }
        $script:data.Config.GeneticPredispositions['Custom'] += @($newEntry)
        Save-Data -data $script:data -type "Config"
        # UI aktualisieren
        $geneticNewNameText.Text = ""
        for ($i = 0; $i -lt $geneticNewMarkerList.Items.Count; $i++) { $geneticNewMarkerList.SetItemChecked($i, $false) }
        Update-GeneticCustomPanel
        [System.Windows.Forms.MessageBox]::Show("Vorbelastung '$name' wurde hinzugefügt.", "Erfolg", "OK", "Information")
    })

    # v2.17.0: Benutzerdefinierte Vorbelastung löschen
    $geneticDeleteButton.Add_Click({
        # Welche benutzerdefinierten Checkboxen sind aktiv (angehakt)?
        $toDelete = @($script:geneticCustomChecks | Where-Object { $_.Checked })
        if ($toDelete.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Bitte aktivieren Sie die Checkbox einer benutzerdefinierten Vorbelastung, um sie zu löschen.", "Keine Auswahl", "OK", "Information")
            return
        }
        $names = ($toDelete | ForEach-Object { $_.Tag }) -join ', '
        $confirm = [System.Windows.Forms.MessageBox]::Show("Möchten Sie die folgenden benutzerdefinierten Vorbelastungen löschen?`n$names", "Löschen bestätigen", "YesNo", "Warning")
        if ($confirm -eq "Yes") {
            $deleteNames = @($toDelete | ForEach-Object { $_.Tag })
            $script:data.Config.GeneticPredispositions['Custom'] = @($script:data.Config.GeneticPredispositions['Custom'] | Where-Object { $deleteNames -notcontains $_['Name'] })
            Save-Data -data $script:data -type "Config"
            Update-GeneticCustomPanel
            [System.Windows.Forms.MessageBox]::Show("$($toDelete.Count) Vorbelastung(en) gelöscht.", "Erfolg", "OK", "Information")
        }
    })

    # v2.17.0: Benutzerdefinierte Vorbelastung bearbeiten (Popup)
    $geneticEditButton.Add_Click({
        $customPreds = @($script:data.Config.GeneticPredispositions['Custom'])
        if ($customPreds.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Es sind keine benutzerdefinierten Vorbelastungen vorhanden, die bearbeitet werden können.", "Keine Vorbelastungen", "OK", "Information")
            return
        }
        # --- Popup-Fenster erstellen ---
        $editForm = New-Object System.Windows.Forms.Form
        $editForm.Text = "Benutzerdefinierte Vorbelastung bearbeiten"
        $editForm.Size = New-Object System.Drawing.Size(580, 560)
        $editForm.StartPosition = "CenterParent"
        $editForm.FormBorderStyle = "FixedDialog"
        $editForm.MaximizeBox = $false
        $editForm.MinimizeBox = $false

        # Auswahl-Dropdown
        $editSelectLabel = New-Object System.Windows.Forms.Label
        $editSelectLabel.Text = "Vorbelastung auswählen:"
        $editSelectLabel.Location = New-Object System.Drawing.Point(15, 15)
        $editSelectLabel.Size = New-Object System.Drawing.Size(160, 20)
        $editForm.Controls.Add($editSelectLabel)

        $editSelectCombo = New-Object System.Windows.Forms.ComboBox
        $editSelectCombo.Location = New-Object System.Drawing.Point(180, 12)
        $editSelectCombo.Size = New-Object System.Drawing.Size(365, 22)
        $editSelectCombo.DropDownStyle = "DropDownList"
        foreach ($pred in $customPreds) { $editSelectCombo.Items.Add($pred['Name']) }
        $editForm.Controls.Add($editSelectCombo)

        # Name-Feld
        $editNameLabel = New-Object System.Windows.Forms.Label
        $editNameLabel.Text = "Name:"
        $editNameLabel.Location = New-Object System.Drawing.Point(15, 50)
        $editNameLabel.Size = New-Object System.Drawing.Size(50, 20)
        $editForm.Controls.Add($editNameLabel)

        $editNameText = New-Object System.Windows.Forms.TextBox
        $editNameText.Location = New-Object System.Drawing.Point(70, 47)
        $editNameText.Size = New-Object System.Drawing.Size(475, 22)
        $editForm.Controls.Add($editNameText)

        # Marker-Liste
        $editMarkerLabel = New-Object System.Windows.Forms.Label
        $editMarkerLabel.Text = "Zugehörige Blutmarker:"
        $editMarkerLabel.Location = New-Object System.Drawing.Point(15, 80)
        $editMarkerLabel.Size = New-Object System.Drawing.Size(200, 20)
        $editForm.Controls.Add($editMarkerLabel)

        $editMarkerList = New-Object System.Windows.Forms.CheckedListBox
        $editMarkerList.Location = New-Object System.Drawing.Point(15, 103)
        $editMarkerList.Size = New-Object System.Drawing.Size(535, 360)
        $editMarkerList.CheckOnClick = $true
        $editMarkerList.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
        $editForm.Controls.Add($editMarkerList)

        # Alle Marker in die Liste laden
        $allMarkerNames = @($script:data.Config.Markers.Name) + @($script:data.Config.CalculatedMarkers.Name) | Sort-Object -Unique
        foreach ($m in $allMarkerNames) { $editMarkerList.Items.Add($m) | Out-Null }

        # Buttons: Speichern + Abbrechen
        $editSaveButton = New-Object System.Windows.Forms.Button
        $editSaveButton.Text = "Speichern"
        $editSaveButton.Location = New-Object System.Drawing.Point(355, 480)
        $editSaveButton.Size = New-Object System.Drawing.Size(90, 30)
        $editForm.Controls.Add($editSaveButton)

        $editCancelButton = New-Object System.Windows.Forms.Button
        $editCancelButton.Text = "Abbrechen"
        $editCancelButton.Location = New-Object System.Drawing.Point(455, 480)
        $editCancelButton.Size = New-Object System.Drawing.Size(90, 30)
        $editForm.Controls.Add($editCancelButton)
        $editForm.CancelButton = $editCancelButton

        # --- Logik: Auswahl → Felder befüllen ---
        $editSelectCombo.Add_SelectedIndexChanged({
            $selectedName = $editSelectCombo.SelectedItem
            $pred = $script:data.Config.GeneticPredispositions['Custom'] | Where-Object { $_['Name'] -eq $selectedName } | Select-Object -First 1
            if ($pred) {
                $editNameText.Text = $pred['Name']
                $predMarkers = @($pred['Markers'])
                for ($i = 0; $i -lt $editMarkerList.Items.Count; $i++) {
                    $editMarkerList.SetItemChecked($i, ($predMarkers -contains $editMarkerList.Items[$i]))
                }
            }
        })
        # Erste Vorbelastung vorab auswählen
        $editSelectCombo.SelectedIndex = 0

        # --- Speichern ---
        $editSaveButton.Add_Click({
            $originalName = $editSelectCombo.SelectedItem
            $newName = $editNameText.Text.Trim()
            if (-not $newName) {
                [System.Windows.Forms.MessageBox]::Show("Der Name darf nicht leer sein.", "Fehlende Eingabe", "OK", "Warning")
                return
            }
            # Prüfen ob neuer Name bereits vergeben (bei anderer Vorbelastung)
            if ($newName -ne $originalName) {
                $allPreds = Get-AllGeneticPredispositions
                $conflict = $allPreds | Where-Object { $_['Name'] -eq $newName }
                if ($conflict) {
                    [System.Windows.Forms.MessageBox]::Show("Eine Vorbelastung mit dem Namen '$newName' existiert bereits.", "Duplikat", "OK", "Warning")
                    return
                }
            }
            # Marker sammeln
            $selectedMarkers = @()
            for ($i = 0; $i -lt $editMarkerList.Items.Count; $i++) {
                if ($editMarkerList.GetItemChecked($i)) { $selectedMarkers += $editMarkerList.Items[$i] }
            }
            if ($selectedMarkers.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Bitte wählen Sie mindestens einen Blutmarker aus.", "Fehlende Auswahl", "OK", "Warning")
                return
            }
            # Vorbelastung aktualisieren
            $pred = $script:data.Config.GeneticPredispositions['Custom'] | Where-Object { $_['Name'] -eq $originalName } | Select-Object -First 1
            if ($pred) {
                $pred['Name'] = $newName
                $pred['Markers'] = $selectedMarkers
                $pred['Hint'] = "Vorbelastung ($newName): Achten Sie besonders auf diesen Marker."
                Save-Data -data $script:data -type "Config"
                Update-GeneticCustomPanel
                [System.Windows.Forms.MessageBox]::Show("Vorbelastung '$newName' wurde gespeichert.", "Erfolg", "OK", "Information")
                # Combo aktualisieren (falls Name geändert)
                $currentIndex = $editSelectCombo.SelectedIndex
                $editSelectCombo.Items[$currentIndex] = $newName
            }
        })

        # --- Abbrechen ---
        $editCancelButton.Add_Click({
            $editForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $editForm.Close()
        })

        $editForm.ShowDialog()
    })
    
    $form.Add_Load({
        $script:data = @{ Config = Load-Config }
        $script:allHistoricalItems = [System.Collections.ArrayList]@(Load-AllHistoricalItems)
        $thresholdTextBox.Text = $script:data.Config.WarningThreshold
        Update-AllMarkerDropdowns
        Update-FilterDropdown
        Update-CorrelationDropdowns
        Load-PersonalMetrics
        Update-Chart
        $advancedControls | ForEach-Object { $_.Visible = $false }

        # AutoBackup: Start-Check + periodischer Scheduler (v2.14.0)
        $script:autoBackupRanThisSession = $false
        Invoke-AutoBackupIfDue

        # Timer alle 15 Minuten prüfen, ob Backup fällig ist
        $script:autoBackupTimer = New-Object System.Windows.Forms.Timer
        $script:autoBackupTimer.Interval = 15 * 60 * 1000  # 15 Minuten
        $script:autoBackupTimer.Add_Tick({ Invoke-AutoBackupIfDue })
        $script:autoBackupTimer.Start()
    })

    $form.Add_FormClosing({
        # Timer beim Schließen sauber stoppen (verhindert Handle-Leak)
        if ($script:autoBackupTimer) {
            try { $script:autoBackupTimer.Stop(); $script:autoBackupTimer.Dispose() } catch { }
        }
    })

    $form.ShowDialog()
}

# ---------- Skript-Einstiegspunkt ----------
Show-DataEntryForm