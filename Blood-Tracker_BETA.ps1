# PowerShell Script für Blutwert-Tracking mit GUI
# Version 3.0.0 (09. Oktober 2026)
# Zweck: Grafische Analyse, Erfassung und proaktive Überwachung von Blutwerten.
#
# ╔══════════════════════════════════════════════════════════════════════════════════════╗
# ║ PFLICHT BEI JEDER AENDERUNG - gilt fuer Menschen UND KI-Modelle                      ║
# ║                                                                                      ║
# ║ 1. Version erhoehen (Bugfix x.y.Z | neue Funktion x.Y.0) - an DREI Stellen gleich:   ║
# ║    a) Kopfzeile 2 dieses Scripts: "# Version x.y.z (TT. Monat JJJJ)"                 ║
# ║    b) Variable $Version (Abschnitt "Globale Konfiguration") - genau EINE Zuweisung,  ║
# ║       der BETA-Starter ersetzt sie per Textsuche                                     ║
# ║    c) neuer Eintrag GANZ OBEN in $script:ReleaseNotes (direkt unter $Version)        ║
# ║ 2. Release-Notes-Eintrag: Version, Datum JJJJ-MM-TT (= Auslieferungsdatum), Titel,   ║
# ║    Eintraege mit Typ (Neu | Bugfix | Anpassung | Sicherheit | Hinweis) und Text aus  ║
# ║    Benutzersicht. Texte in einfachen Anfuehrungszeichen, keine Platzhalter/TODO.     ║
# ║ 3. Technische Details zusaetzlich als "# CHANGELOG vX.Y.Z (...)" direkt unter        ║
# ║    diesem Block (neueste oben).                                                      ║
# ║ 4. Veroeffentlichte Release-Notes-Eintraege NIE aendern, umformulieren, loeschen.    ║
# ║ 5. Neuer Dateiname: Blood-Tracker_vX.Y.Z_JJJJ-MM-TT_HHMM.ps1 (UTF-8 mit BOM).        ║
# ║ 6. MEHRERE RELEASES AM SELBEN TAG: trotzdem JE VERSION ein eigener Eintrag mit       ║
# ║    demselben Datum - nie mehrere Versionen in einen Eintrag schreiben. Der Reiter    ║
# ║    fasst Versionen gleichen Datums automatisch zu EINEM Listeneintrag zusammen       ║
# ║    (neueste Version, "TT.MM.JJJJ · n Versionen"); rechts stehen alle Versionen des   ║
# ║    Tages, neueste zuerst (Get-ReleaseNotesGroups). Nichts von Hand gruppieren.       ║
# ║ 7. Datum nicht bekannt (nur Altbestand): Datum = '' - angezeigt wird dann nur die    ║
# ║    Versionsnummer (kein "Datum unbekannt"). Neue Versionen haben IMMER ein Datum.    ║
# ║ Pruefung: Test-ReleaseNotesConsistency - der BETA-Start warnt bei Abweichungen.      ║
# ║ Weitere Regeln fuer KI-Modelle: AGENTS.md im Repository (sofern vorhanden).          ║
# ╚══════════════════════════════════════════════════════════════════════════════════════╝
#
# CHANGELOG v3.0.0 (GUI: Ausrichtung und Bedienelemente optimiert):
# - ANSICHTEN: Marker-hinzufuegen oberhalb der ersten Markerzeile; Tab-Reihenfolge
#   passend zur neuen Position, bisherige Listenhoehe bleibt erhalten.
# - RELEASE-NOTES: aktuelle Version statt Laeuft; Kopfzeile mit ausreichend Platz
#   fuer den rechtsbuendigen Versionshinweis an der Kante der Detailflaeche.
# - EINZELMARKER: Exportieren links von Drucken, gleiche Groesse und Abstaende.
#   Hinweis/e ersetzt Wichtige Hinweise & Details und nutzt die volle Chart-Breite.
# - VERSION: Auf Wunsch des Nutzers Sprung auf 3.0.0; keine Datenformat-Aenderung.
#
# CHANGELOG v2.49.0 (GUI: Kategorien und zugehoerige Unterreiter):
# - NEU: Uebersicht, Auswertung, Daten & Profil, Info in der oberen Zeile;
#   darunter nur die zugehoerigen bisherigen Reiter. Start auf Dashboard.
# - NAVIGATION: Letzten Unterreiter je Kategorie nur fuer die Sitzung merken;
#   Dashboard-Spruenge wechseln Kategorie und Zielreiter gemeinsam.
# - ERHALT: Vorhandene Reiterseiten umhaengen, nicht neu erstellen. Persoenliche
#   Eingaben beim Reiterwechsel nicht erneut aus gespeicherten Daten laden.
# - DARSTELLUNG: Eigene Schriftressourcen je Tab-Leiste; aktive Auswahl rot/fett.
#   Fenster 30 Pixel hoeher, bei kleinem Bildschirm begrenzen und Inhalt scrollbar.
# - QS: Aktualisierung nur fuer den fertigen Zielreiter; versteckte Seiten beim
#   Schliessen freigeben. Keine neuen Config-Schluessel oder Datenformate.
#
# CHANGELOG v2.48.2 (Navigation: Gemeinsame interne Reiterwechsel):
# - VORBEREITUNG: Select-BtNavigationPage buendelt interne Reiterwechsel.
#   Dashboard-Kacheln und der Link zu persoenlichen Metriken nutzen diesen Weg.
# - ROBUSTHEIT: Aktualisierung bei voruebergehend fehlender Reiterauswahl aussetzen.
#   Keine Aenderung an sichtbarer Navigation, Datenformat oder Konfiguration.
#
# CHANGELOG v2.48.1 (GUI: Schaltflaeche fuer Blutbild-Eingabe umbenannt):
# - TEXT: Formular-Schaltflaeche heisst jetzt Werte nach Blutbild speichern.
#
# CHANGELOG v2.48.0 (GUI: Mehrere Blutwerte im Formular erfassen):
# - NEU: Zeilenformular mit gemeinsamem Datum, Alias-Suche, HIV/APOE-Auswahl,
#   Testosteron-Einheiten, Einzelnotizen, Plus/Entfernen und Scrollbereich.
# - PRUEFUNG: Gemeinsame Eingabelogik und Referenzbewertung; separate Pruefliste,
#   danach ein atomarer Speichervorgang. Fehler erhalten Eingaben und Notizen.
# - BEDIENUNG: Tab/Shift+Tab, Enter fuer Markerauswahl, kein AcceptButton;
#   Rueckfrage beim Verwerfen ungespeicherter Eingaben, keine HIV/APOE-Vorbelegung.
# - LAYOUT A: Werte erfassen mit Einzelwert speichern / Mehrere Werte; Verwaltung
#   und Warnschwelle auch im erweiterten Zustand passend nach unten verschieben.
# - TEXT: Genotyp-Hinweis mit deutschem Umlaut. Keine neuen Daten-/Config-Formate.
#
# CHANGELOG v2.47.0 (Eingabe: Gemeinsam speichern und APOE-Layout):
# - NEU: Save-BtBloodValueItems mit frischer Bestandspruefung, Schreibsperre und
#   einem verschluesselten Schreibvorgang fuer alle vorbereiteten Werte eines Tages.
# - SPEICHERN: Save-Data Daily/NewItems; Einzeleingabe nutzt den neuen Weg mit
#   ThrowOnError. Eingaben erst nach erfolgreichem Speichern leeren.
# - SICHERHEIT: Write-ProtectedJsonFile -RequireAtomic verhindert Kopier-Fallback;
#   unlesbare/unerwartete Tagesformate und Duplikate blockieren ohne Datenverlust.
# - BUGFIX: APOE-Auswahlliste 90 statt 160 Pixel breit; Abstand zum Hinweistext.
# - Keine Formular-GUI und keine Aenderung am gespeicherten Datenformat.
#
# CHANGELOG v2.46.0 (Eingabe: Gemeinsame Pruefung und Umrechnung):
# - NEU: New-BtBloodValueItem fuer Alias-Aufloesung, Zahl-/Auswahlpruefung,
#   HIV/APOE, Testosteron, Zusatznotizen und Duplikate in Bestand/Eingabeliste.
# - ANBINDUNG: Bestehende Einzeleingabe nutzt denselben Helfer; kein Formular,
#   keine Aenderung an Speicherfunktionen, Dateiformat oder Layout.
# - RUNDUNG: Testosteron nmol/l -> ng/dl mit Faktor 28.818, eine Nachkommastelle,
#   explizit AwayFromZero; nicht endliche Zahlen und Umrechnungsueberlauf ablehnen.
#
# CHANGELOG v2.45.0 (Engine: YTD und letzte sechs Monate):
# - NEU: Get-/Measure-ActivitySummary und Get-ActivitySummaryWindow; nur lesen/rechnen.
# - TAGESWERTE: Nur abgeschlossene Tage, echte Nullwerte behalten, Luecken nicht fuellen.
# - CARDIO: Training/Tagesmittel, Intervalle, Verfahren, Quellen und Geraete trennen;
#   unbekannte Intervalle/Verfahren nicht automatisch aggregieren, negative Werte erhalten.
# - QUELLEN: Exakte Zeitraum-Mittel haben Vorrang und bleiben neben Berechnungen erhalten.
#   Abweichungen/Abdeckung melden; andere Zeitraeume niemals aufteilen oder hineinrechnen.
# - QS: Ladefehler/Duplikate blockieren; keine Dateien, GUI oder Datenformate aendern.
# - KONFIG: ActivitySummaryDecimals nur fuer Anzeige, Rundung AwayFromZero.
#
# CHANGELOG v2.44.0 (GUI: Aktiven Hauptreiter fett beschriften):
# - NEU: Ausgewaehlten Reiternamen zusaetzlich zur Farblinie fett zeichnen.
# - KONFIG: ActiveTabBoldEnabled im USER-Block; Standard true.
# - LAYOUT: Native Reiterbreiten anhand der fetten Schrift bestimmen; inaktive
#   Namen und Inhalte behalten die bisherige Schrift. Eigene Schrift freigeben.
#
# CHANGELOG v2.43.0 (GUI: Aktiven Hauptreiter hervorheben):
# - NEU: Dezente rote Linie ueber der Beschriftung des ausgewaehlten Hauptreiters.
# - KONFIG: ActiveTabIndicatorEnabled und ActiveTabIndicatorColor im USER-Block.
# - DARSTELLUNG: OwnerDrawFixed nur fuer die Hauptreiter; vorhandene Schrift,
#   Reitergroessen, Tooltip-Texte und Auswahl-/Aktualisierungsereignisse bleiben erhalten.
#   Zeichnen ueber args, GDI-Ressourcen in finally freigeben, Tastaturfokus anzeigen.
# - Keine EnableVisualStyles-/DPI-Aenderung, keine Daten-/Engine-Aenderung.
#
# CHANGELOG v2.42.0 (Engine: Uebernommene Zeitraum-Mittel):
# - NEU (ohne GUI): New-/Save-/Get-/Remove-ActivityPeriodRecord(s) fuer genau eine
#   Messgroesse je Zeitraum, Herkunft Quelle, Von/Bis inklusive und optionale Anzahl.
# - BASIS: Tageswerte fuer Energie/Schritte; Trainingswerte/Tagesmittel/Unbekannt
#   fuer Cardio. Intervalle und Verfahren bleiben erhalten; Schrittmittel dezimal erlaubt.
# - ABLAGE: Separate DPAPI-Datei _messungen\_zeitraeume\Aktivitaet.json (Schema 1).
#   Schreibsperre, explizites Ersetzen per Id, Duplikatpruefung, strikter Formatschutz.
# - FILTER: Ueberlappungen komplett liefern; ExactPeriod filtert exakt. Keine Aufteilung,
#   Addition, Durchschnittsberechnung oder Rueckschreibung in Tages-/Trainingswerte.
# - Bestehende Funktionen, GUI, Tagesdateien und veroeffentlichte Release-Notes unveraendert.
#
# CHANGELOG v2.41.0 (Engine: Aktivitaet und Cardio-Erholung):
# - NEU (ohne GUI): Kategorie Aktivitaet fuer konsolidierte Tageswerte (kcal, Schritte,
#   optional uebernommenes Cardio-Tagesmittel); Fitness um Cardio-Trainingswerte erweitert.
# - HERKUNFT: Quelle, Geraet, Tagesabschluss, Datenart, Erholungsintervall und bekannte
#   Stichprobengroesse erhalten; unbekannt bleibt null, keine Hersteller-Vorbelegung.
# - PRUEFUNG: Endliche Zahlen, Energie/Schritte >= 0, Tagesschritte ganzzahlig.
#   Negative Cardio-Werte bleiben mit Hinweis; optionale Warnschwellen im USER-Block.
# - ABLAGE: Vorhandenes Schema 2 / DPAPI; ein Aktivitaets-Eintrag je Tag ohne Uhrzeit.
#   Duplikate nur explizit ersetzen; Tagesmittel und Trainingswerte bleiben getrennt.
# - Keine medizinischen Bereiche, keine GUI und noch keine Zeitraum-/Durchschnittsrechnung.
#
# CHANGELOG v2.40.0 (Dashboard: Eigene Ansichten):
# - NEU: Eigene Ansichten erstellen, bearbeiten und nach Bestaetigung loeschen.
#   Suchbare Markerzeilen mit Alias-Suche, Reihenfolge, Duplikat- und Namenspruefung.
# - GUI: Auswahl ueber stabile Kennungen, vier Kacheln je Zeile, Scrollbereich;
#   fehlende Marker bleiben sichtbar. Start immer Standard; Loeschen -> Standard.
# - SPEICHERN: Save-Data meldet Fehler auf Wunsch per -ThrowOnError an den neuen
#   Dialog. Bestehende Aufrufer unveraendert. Speicher erst nach Dateierfolg ersetzen.
#   Abbrechen schreibt nichts; ungueltige Bestandsansichten werden nicht verworfen.
# - Unveraendert: Messwert-Engine, Eingabewege, Score-Bewertung und alte Release-Notes.
#
# CHANGELOG v2.39.0 (DashboardViews: Konfiguration vorbereitet):
# - FEATURE (ohne Dialog): DashboardViews als Liste aus Id (GUID), Name und Markers.
#   Neuer Schluessel an allen vier Stellen: Defaults, Lade-Merge, finalConfig, configToSave.
# - NEU: ConvertTo-BtDashboardViews prueft Struktur, eindeutige Kennungen/Namen,
#   reservierte Ansichtsnamen und doppelte Marker; Reihenfolge und unbekannte Marker bleiben.
#   Leere/eine/mehrere Ansichten bleiben Listen; ein einzelnes Objekt wird normalisiert.
# - DATENSCHUTZ: Fehlerhafter Ansichtsblock bleibt im Speicher erhalten; Load-Config
#   schreibt die Config dann nicht automatisch zurueck. Explizites Config-Speichern
#   lehnt ungueltige Ansichten ab. Keine Kennungen oder Inhalte in Warnmeldungen.
# - CACHE: DashboardViews in der vorhandenen kanonischen Dashboard-Pruefsumme.
# - Unveraendert: GUI, Eingabewege, Messdaten, Verschluesselung und alte Release-Notes.
#
# CHANGELOG v2.38.0 (Dashboard: Aktuelle Auffaelligkeiten):
# - FEATURE: Dynamische Ansicht "Aktuelle Auffaelligkeiten (n)" ab einem Treffer;
#   eindeutige gemessene/berechnete Marker, vier Kacheln je Zeile, Klick zur Analyse.
# - REGEL: Letzter Wert bis zum Stichtag, hoechstens 365 Tage alt, gueltig/konfliktfrei,
#   vollstaendiger gueltiger Referenzbereich, Wert ausserhalb UND Level Critical nach
#   bestehender Bewertung. Gruppe Genetik ausgeschlossen; kein Rueckgriff bei Konflikten.
# - SORTIERUNG: Bestehender Bereichsscore aufsteigend, dann Markername; keine neue Bewertung.
# - NEU: Get-HealthDashAbnormalities und Update-HdViewSelector. Vorhandener Kachelbereich
#   gezielt um Ansichtsparameter und Modell-/Ansichtsschluessel im Cache erweitert.
# - VERHALTEN: Zaehleraenderung behaelt die Auswahl; null Treffer -> Standard.
#   Modellfehler entfernt veraltete Treffer; bestehende Datenpruefsumme wiederverwendet.
# - Unveraendert: Standardmodell, Vorbelastungsregeln, Config, Messdaten und Eingabewege.
#
# CHANGELOG v2.37.0 (Dashboard-Ansichten: Standard und Vorbelastungen):
# - FEATURE: Auswahl "Ansicht" in der Kopfzeile; Start immer mit Standard.
#   Standardkarten unveraendert; Infozeile unter der Auswahl, Karten weiter ab y=72.
# - FEATURE: Vorbelastungen als eigene Ansicht: alle zugeordneten Marker je aktiver
#   Vorbelastung, vier Kacheln je Zeile, scrollbar, inklusive berechneter Marker und
#   Leer-/Fehlerzustaende. Hauptschalter wird beachtet; Leerzustand mit Einstellungslink.
# - NEU: Get-HealthDashPredispositionTiles, Update-HdPredispositionView und
#   Set-HealthDashboardView; Kachelzeichnung/Bewertung und Datenpruefsumme wiederverwendet.
#   Ansichtswechsel zeichnet ohne Neuberechnung; Datenwechsel erneuert die Ansicht.
# - ANPASSUNG: Kachel-Klick kann konfigurierte Marker ohne Messdaten gezielt oeffnen.
# - ANPASSUNG: Lokales $version in Get-ReleaseNotesView in $noteVersion umbenannt;
#   nur eine Versionszuweisung im gesamten Script. Keine Config-/Messdaten-Aenderung.
#
# CHANGELOG v2.36.0 (Neuer Reiter "Dashboard"):
# - FEATURE: Reiter "Dashboard" als ERSTER Reiter (Integration B; Modul BloodTracker_Dashboard
#   v0.1.1, isoliert getestet, unveraendert uebernommen; Spec v0.3.1). Inhalt: Gesundheitsscore
#   (Ring, Status, Vergleich zum vorherigen Testtag), Longevity-Score, 6 Longevity-Kacheln
#   (Klick/Enter/Leertaste -> Reiter "Einzelmarker-Analyse" mit diesem Marker), Vorbelastungen
#   (Score, "k von m Markern", 3 Fokusmarker; Link -> "Persoenliche Metriken"), Karte Testosteron
#   (gesamt, frei gemessen, SHBG + berechnet: freies T nach Vermeulen, FAI - Werte desselben
#   Testtags, Albumin sonst geschaetzt 4,3 g/dl), Dialog "Wie berechnet?".
# - Neu (oberste Ebene, Block "DASHBOARD (v2.36.0)"): Get-HealthDashboardModel, Get-HealthDashScore,
#   Get-HealthDashComparison, Get-HealthDashTiles, Get-HealthDashPredispositions,
#   Get-HealthDashTestosterone, Get-HdVermeulenFreeTestosterone, Get-HealthDashDataSignature,
#   Initialize-/Update-HealthDashboardTab u. a. (Praefixe Hd/HealthDash, keine Namenskollision).
# - Aufbau im Form-Load NACH dem Laden der Daten; ein Fehler zeigt einen Hinweis im Reiter, das
#   Programm laeuft weiter. Aktualisierung beim Wechsel auf den Reiter und nach "globale
#   Einstellungen"; neu gerechnet wird nur bei geaenderter Pruefsumme (SHA256 ueber Messwerte,
#   Marker-Konfiguration, persoenliche Metriken, Vorbelastungen, Datum).
# - Bewertung ausschliesslich ueber Get-MarkerRangeScore (wie Risiko-Cockpit), berechnete Werte
#   ueber Get-CalculatedValuesForMarker. Keine Aenderung an Daten, Config oder anderen Reitern;
#   Reiterkoepfe: Resolve-TabHeaderOverlap beruecksichtigt den zusaetzlichen Reiter.
#
# CHANGELOG v2.35.0 (Sortieren in Cockpit/Longevity + Messwert-Engine erweitert):
# - FEATURE: Reiter "Risiko-Cockpit" und "Longevity-Indizes" - jede Spalte per Klick auf die
#   Ueberschrift sortierbar: aufsteigend -> absteigend -> Standardreihenfolge (Pfeil im Kopf).
#   Zahlen nach Zahlenwert (auch "5.2 %", "612 (21.24 nmol/l)"), Datum chronologisch, Bewertung
#   nach Schwere (Hilfsspalte SortOrder), Text alphabetisch (de-DE); leere Werte und "-" immer
#   unten; Sortierung bleibt bei Filterwechsel/Aktualisierung erhalten. Neu: Register-BtGridSorting,
#   Invoke-BtGridApplySort, Compare-BtGridRows, Get-BtGridSortKey u. a. (Spalten SortMode
#   Programmatic, Vergleich im SortCompare-Ereignis).
# - BUGFIX: Sortieren per Spaltenkopf verglich Zellwerte direkt - Datum und "5.2 %" als Text,
#   gemischte Spalten (Zahl/"-", Zahl/Text) waren nicht zuverlaessig sortierbar.
# - BUGFIX Risiko-Cockpit: Das Ausblenden wiederholter Funktionsbereiche verglich mit der bereits
#   geleerten Vorzeile -> jede zweite Zeile einer Gruppe zeigte den Namen erneut. Jetzt echter
#   Name in Cell.Tag; ausgeblendet nur in der Standardansicht und bei Sortierung nach Bereich.
# - ANPASSUNG Risiko-Cockpit: Standardreihenfolge Bewertung, dann Bereich, dann Marker (bisher
#   nur Bewertung; Reihenfolge innerhalb gleicher Bewertung durch instabile Sortierung zufaellig).
# - FEATURE (vorbereitet, ohne Oberflaeche): Messwert-Engine um Vitalwerte (Herzfrequenz,
#   Blutdruck sys/dia, Koerpertemperatur), Koerper (Gewicht, BMI, Koerperfettanteil, fettfreie
#   Masse, FFMI), EKG (Rhythmus, Typ, Herzfrequenz, PQ, QRS, QT, QTc; Befund, Dokumentverweis)
#   und Fitness (VO2max) erweitert. MEHRERE Messungen je Tag mit Uhrzeit (Tagesdatei mit Liste
#   "Messungen", SchemaVersion 2); DEXA/Schlaf unveraendert (1 Datensatz je Tag).
#   Neu: New-VitalEntry, New-BodyEntry, New-EkgEntry, New-FitnessEntry, Add-/Remove-
#   MeasurementEntry, Get-MeasurementEntries, Get-BodyDerivedValues (BMI, FFM, KF-%, FFMI),
#   Get-EkgDerivedValues (QTc Bazett/Fridericia). Bewertungsbereiche nur mit Leitlinie:
#   Blutdruck ESC 2024 (Praxis), BMI WHO, Herzfrequenz < 50 / > 100 bpm (ACC/AHA/HRS 2018),
#   PQ > 200 ms, QRS ab 110 ms (ACC/AHA/HRS 2018); QTc geschlechtsabhaengig -> nur Beschreibung.
# - Bewusst NICHT: Stress-Werte (Herstellerformeln, nicht validiert - Entscheidung 08.10.2026).
# - Unveraendert: Blutmarker "Gewicht"/"Systolischer Blutdruck", BMI-Marker, PREVENT, Daten,
#   Config und alle uebrigen Reiter.
#
# CHANGELOG v2.34.1 (Release-Notes: Versionen eines Tages zusammengefasst):
# - ANPASSUNG Reiter "Release-Notes" (Modul BloodTracker_ReleaseNotes v0.2.0, vorab isoliert
#   getestet): Versionen mit GLEICHEM Datum bilden links EINEN Eintrag (neueste Version,
#   "TT.MM.JJJJ · n Versionen", Titel der neuesten Version); rechts stehen alle Versionen des
#   Tages untereinander (neueste zuerst, je "Version x · Datum", feine Trennlinie).
#   Neu: Get-ReleaseNotesGroups, Get-RnGroupItemHeight; Liste OwnerDrawVariable; die Auswahl
#   bleibt am Tag, auch wenn ein Filter dessen neueste Version ausblendet (-PreviousKey).
# - ANPASSUNG: Kein "Datum unbekannt" mehr - ohne Datum nur die Versionsnummer (Liste kompakt
#   mit 2 Zeilen / 42 px, Details "Version x", CHANGELOG.md "## [x]").
# - Release-Notes-Eintrag 2.34.0 auf Wunsch korrigiert (Satz zu "Datum unbekannt" entfernt;
#   v2.34.0 war noch nicht veroeffentlicht). Neuer Eintrag 2.34.1.
# - Regelblock oben um Regel 6 (mehrere Releases am selben Tag) und 7 (Datum unbekannt)
#   ergaenzt, ebenso der Hinweis ueber $script:ReleaseNotes.
# - Unveraendert: Datenmodell (ein Eintrag je Version), Filterregeln, Konsistenzpruefung,
#   BETA-Startpruefung, alle anderen Reiter, Daten, Config und Berechnungen.
#
# CHANGELOG v2.34.0 (Reiter "Release-Notes" + Pflicht-Regelblock, Integration A):
# - FEATURE: Neuer Reiter "Release-Notes" (letzter Reiter, nach "Persoenliche Metriken"):
#   Versionsliste (owner-drawn, numerisch sortiert, "aktuell" = laufende $Version),
#   schreibgeschuetzte Details nach Typ gruppiert, Suche + Typ-Filter, Leerzustand mit
#   "Filter zuruecksetzen" (Dashboard_Design-Spec v0.3.1, Abschnitt 8, Regeln 1-6).
# - NEU: $script:ReleaseNotes direkt unter $Version - 44 Versionen (2.7 bis 2.34.0) aus den
#   bisherigen CHANGELOG-Kommentaren, Texte aus Benutzersicht. Datum nur aus Quelldaten:
#   2.26.1-2.33.0 = 07.10.2026 (Kopfzeilen dieser Versionen), 2.26.0 = 20.08.2026,
#   2.7 = 04.02.2026 ("Datum:"-Zeile im Kopf), sonst leer = "Datum unbekannt".
# - NEU: Funktionen und Konstanten aus BloodTracker_ReleaseNotes v0.1.0 (isoliert getestet)
#   unveraendert auf oberster Ebene vor Show-DataEntryForm (u. a. Get-ReleaseNotesView,
#   Test-ReleaseNotesConsistency, Export-ReleaseNotesMarkdown, Initialize-ReleaseNotesTab;
#   $script:HealthUiTheme nutzt spaeter auch das Dashboard).
# - NEU: Invoke-ReleaseNotesStartCheck - nur im BETA-Testmodus ($Version enthaelt "BETA"):
#   prueft beim Anzeigen des Fensters Kopfzeile, $Version und obersten Release-Notes-
#   Eintrag; Abweichungen -> eine Warnmeldung, das Programm laeuft weiter.
# - NEU: Resolve-TabHeaderOverlap - schaetzt vor dem ersten Anzeigen die Breite aller
#   Reiterkoepfe; reichen sie unter "globale Einstellungen", wird der Innenabstand der
#   Reiterkoepfe verringert (6 -> 4 -> 2 px). Passt alles, bleibt alles unveraendert.
# - Regelblock "PFLICHT BEI JEDER AENDERUNG" oben im Kopf (auch fuer KI-Modelle).
# - Unveraendert: Daten, Config, Berechnungen, bestehende Reiter und Funktionen. Ein Fehler
#   beim Aufbau des Reiters wird im Reiter angezeigt und stoppt das Programm nicht.
#
#
# CHANGELOG v2.33.0 (Grundarchitektur DEXA + Schlaf, noch ohne Oberflaeche):
# - FEATURE (vorbereitet): Katalog aller ueblichen DEXA-Messwerte (61 Werte:
#   Koerperzusammensetzung, Fettverteilung, Viszeralfett, 6 Koerperregionen,
#   Knochendichte LWS/Schenkelhals/Gesamthuefte/Radius 33 %/Ganzkoerper mit
#   T- und Z-Score, TBS) und Schlaf (Schlafdauer in h, optional Schlafscore,
#   optional Schlafphasen Wach/Leicht/Tief/REM und einzelne Schlafzyklen -
#   Eingabe wahlweise in Stunden oder Minuten, gespeichert in Minuten).
# - Datenmodell + Funktionen: New-DexaRecord, New-SleepRecord (Pruefung inkl.
#   Plausibilitaet), Save-/Get-/Remove-MeasurementRecord, Get-DexaDerivedValues
#   (ALMI, LMI, FFMI, FMI, A/G aus Rohwerten), Get-SleepPhaseShares.
# - Ablage verschluesselt in data\_messungen\<Kategorie>\<JJJJ>\<Datum>.json -
#   damit automatisch in allen Sicherungen/Wiederherstellungen enthalten.
#   Blutwert-Laden und -Speichern ignorieren den Ordner (Test-IsMeasurementPath),
#   der Ordner wird beim Speichern der Blutwerte NIE geloescht.
# - Keine Aenderung an Oberflaeche, Config oder vorhandenen Daten.
#
# CHANGELOG v2.32.0 (Einzelmarker-Analyse: Suche in "Verlauf anzeigen für"):
# - FEATURE: Die Auswahl "Verlauf anzeigen für" laesst sich jetzt wie das Feld
#   "Blutmarker" per Eingabe durchsuchen: Name, Teil des Namens oder Kuerzel
#   (z. B. "HbA1c", "Leuk", "WBC", "PREVENT", "Pheno"), auch berechnete Marker.
#   Die Liste filtert beim Tippen; Enter uebernimmt einen eindeutigen Treffer,
#   Esc oder Verlassen ohne Treffer stellt die bisherige Auswahl wieder her.
#   Kein Treffer: Feld rot hinterlegt, die volle Liste bleibt sichtbar.
# - Neue Hilfsfunktionen Get-ListMatches / Resolve-ListSelection (gleiche
#   Regeln wie die Markersuche, aber nur innerhalb der angezeigten Liste).
#
# CHANGELOG v2.31.3 (Risiko-Cockpit: berechnete Marker mit nur einem Wert):
# - BUGFIX: Ein berechneter Marker mit genau EINEM Testtag (z. B. HOMA-IR nach der
#   ersten Insulin-Messung) fehlte im Risiko-Cockpit, und seine Gruppe fehlte in
#   der Gruppenauswahl der Cockpit-Einstellungen. Ursache wie v2.31.1: ein
#   Einzelergebnis hat in Windows PowerShell 5.1 kein .Count. Jetzt immer @().
#
# CHANGELOG v2.31.2 (Laden der Config: Listen mit einem Eintrag):
# - BUGFIX (gefunden mit der BETA-Test-Config): ConvertTo-Hashtable machte beim
#   Laden der Config.json aus jeder Liste mit genau EINEM Eintrag einen Einzelwert
#   ("$value = if (...) { $prop.Value }" gibt den Wert ueber die Pipeline aus).
#   Folge: RequiredMarkers @('Gewicht') wurde zu 'Gewicht', [0] = 'G' -> der BMI
#   (und jeder eigene berechnete Marker mit nur einem Pflichtmarker) wurde nach
#   dem ersten Neustart NIE berechnet - auch nach der BMI-Korrektur in v2.28.0.
#   Leere Listen wurden zu $null. Jetzt direkte Zuweisung - Listen bleiben Listen.
#
# CHANGELOG v2.31.1 (Einzelmarker-Analyse: letzter Wert auch bei nur einem Wert):
# - BUGFIX: Ueber dem Verlauf stand "Letzter Wert: ... (am ...)" (inkl. Wert-
#   Beschriftung am Messpunkt) nur bei mindestens zwei Werten. Ursache: Bei genau
#   einem Wert liefert Sort-Object ein Einzelobjekt; [pscustomobject] hat in
#   Windows PowerShell 5.1 keine .Count-Eigenschaft -> Bedingung "Count -gt 0"
#   war falsch. Jetzt immer als Array (@()); Trendlinie und Warnungen verhalten
#   sich bei einem Wert unveraendert (erst ab zwei Werten).
#
# CHANGELOG v2.31.0 (hs-CRP vor CRP in den Scores):
# - ANPASSUNG: PhenoAge, PhenoAge-Accel, InflammAging-Score und Longevity-Score
#   verwenden am Testtag vorrangig "Hochsensitives CRP (hs-CRP)" und nur ohne
#   hs-CRP das klassische CRP. Bisher galt die umgekehrte Reihenfolge - hs-CRP ist
#   im niedrigen Bereich (< 3 mg/l), in dem diese Scores unterscheiden, genauer.
#   Beide in mg/l; $script:MarkerPreferred (gilt nur fuer die Score-Formeln -
#   eigene Formeln und Anzeige der Einzelwerte bleiben unveraendert).
# - Tab "Longevity": Scores werden kaufmaennisch auf ganze Punkte gerundet
#   (64,5 -> 65; bisher Bankers Rounding 64,5 -> 64).
# - BUGFIX (neu gefunden beim BETA-Abgleich): Im Tab "Longevity" fehlten bei
#   PhenoAge, PhenoAge-Accel, NLR, TyG-Index und TG/HDL-Ratio Ampelfarbe und
#   Hinweistext der Score-Spalte. Ursache: Math.Round(Ganzzahl) liefert in
#   PowerShell [decimal]; die Farbpruefung erwartet double/int.
#
# CHANGELOG v2.30.0 (Longevity-Score an den eigenen Bereichen):
# - BUGFIX Longevity-Score: Die festen Formeln widersprachen den eigenen
#   Optimalbereichen (z. B. HbA1c 5,2 % = 0 Punkte, Insulin 5 = 0, Homocystein
#   7 = 0; ein gesunder Datensatz bekam ca. 33/100). Neu: jeder Wert wird an den
#   Referenz-/Optimalbereichen SEINES Markers bewertet - wie die Ampel im Cockpit:
#   optimal = 100, im Referenzbereich 79-50 (linear zur Grenze), ausserhalb 49-0.
#   Eigene Anpassungen der Bereiche wirken automatisch mit.
# - BUGFIX: Der Score wurde nur berechnet, wenn am selben Tag ALLE 7 Werte inkl.
#   NAD+ vorlagen (in der Praxis fast nie). Neu: Mittelwert der vorhandenen Werte,
#   mindestens 4 von 7 ($script:LongevityScoreMinComponents); NAD+ ist optional.
# - Tab "Longevity": Einzel-Scores der Laborwerte nach derselben Logik; Spalte
#   "Ideal" zeigt den Optimalbereich aus den Marker-Einstellungen (bisher feste
#   Texte, z. B. HbA1c "< 5.2 %" statt 4.5-5.3 %).
# - Beschreibung Longevity-Score aktualisiert (Korrektur-Migration nur bei
#   unveraendertem Standardtext).
#
# CHANGELOG v2.29.0 (Geburtsdatum statt festem Alter):
# - FEATURE: "Persoenliche Metriken" > Geburtsdatum (TT.MM.JJJJ) ersetzt das feste
#   Alter. Das Alter wird je Bluttest am Testtag berechnet (vollendete Lebensjahre,
#   wie in den offiziellen Rechnern) - fuer PhenoAge, PhenoAge-Accel, PREVENT,
#   eGFR aus Kreatinin (PREVENT) und die PhenoAge-Bewertung im Tab "Longevity".
#   Bisher wurden auch aeltere Tests mit dem heutigen Alter berechnet.
# - Neben dem Feld steht das heutige Alter. Beim Speichern wird Personal.Age
#   automatisch auf das heutige Alter gesetzt (Kompatibilitaet aelterer Versionen).
# - Migration: Ohne gespeichertes Geburtsdatum wird wie bisher mit dem festen Alter
#   gerechnet (Hinweis "bisher fest: .. J." neben dem Feld). Speichern der
#   persoenlichen Daten erfordert ab jetzt ein gueltiges Geburtsdatum.
# - PREVENT-ASCVD-10Y: Hinweis "Alter ist ein fester Wert" in der Beschreibung
#   ersetzt (Korrektur-Migration nur bei unveraendertem Standardtext).
#
# CHANGELOG v2.28.0 (Korrekturen der Rechenwerte und Referenzbereiche):
# - BUGFIX TyG-Index: Formel ln(TG x Glu / 2) (Werte ca. 8-9) passte nicht zum
#   Grenzwert 4,68 -> Wert nie "optimal". Neu ln(TG x Glu) / 2 (Simental-Mendia
#   2008 / Guerrero-Romero 2010); Ref 0-4,68, optimal < 4,5.
# - BUGFIX AIP: wurde mit mg/dl berechnet, die Grenzen gelten fuer mmol/l
#   (+0,36 Verzerrung). Neu log10((TG/88,57) / (HDL/38,67)); Ref <= 0,21,
#   optimal < 0,11 (Dobiasova). Hinweistext im Diagramm-Tab angepasst.
# - BUGFIX PLR: Thrombozyten / Lymphozyten-% (ca. 8 statt ca. 140). Neu mit
#   absoluten Lymphozyten (Lymphozyten-% x Leukozyten / 100); Leukozyten sind
#   jetzt Pflichtmarker. InflammAging-Score: PLR-Anteil ebenso korrigiert
#   (bisher immer 100); ohne Leukozyten wird nur CRP + NLR gewertet.
#   Zusaetzlich: Teilscores wurden durch [Math]::Min(100, ...) als Ganzzahl
#   gerundet (PowerShell waehlt Math.Min(int,int)) - jetzt exakt (bis 0,5 Pkt.).
# - BUGFIX Haematokrit: Einheit % mit Referenz 0,38-0,49 (l/l). Neu 38-49 %;
#   gespeicherte l/l-Werte (< 1,5) werden beim Laden in % umgerechnet.
# - ANPASSUNG LDL / Triglyceride: Optimalbereich ab 0 statt ab 50 mg/dl -
#   niedrige Werte gelten nicht mehr als suboptimal.
# - BUGFIX BMI (neu gefunden beim Erstellen der BETA-Testdaten): Die Standard-
#   Formel stand in doppelten Anfuehrungszeichen - $personal wurde beim Laden
#   leer eingesetzt ("[Math]::Pow(.Groesse / 100, 2)") -> Syntaxfehler, der BMI
#   wurde NIE berechnet. Jetzt einfache Anfuehrungszeichen + Korrektur der
#   gespeicherten Formel. Zusaetzlich: ungueltige Ergebnisse (NaN/Unendlich,
#   z. B. BMI ohne Koerpergroesse) werden nicht mehr angezeigt.
# - BUGFIX Aliase (Markersuche): 'BNP' (-> NT-proBNP), 'Troponin' (-> TnT) und
#   'Kreatinin-Clearance' (-> eGFR) entfernt - jeweils andere Laborwerte.
#   Zusaetzlich $script:MarkerSearchBlocklist: die Teilwort-Suche haette "BNP"
#   und "Troponin" sonst weiterhin automatisch zugeordnet.
# - Korrektur-Migration erweitert: auch normale Marker; Bedingung "OnlyIf" -
#   korrigiert wird nur, wenn noch der alte Standardwert gespeichert ist.
#   Eigene Anpassungen bleiben erhalten.
#
# CHANGELOG v2.27.2 (Sicherer Datenpfad-Umzug):
# - BUGFIX (Datenverlust): Einstellungen > "Datenpfad aendern..." nutzte
#   Move-Item "<Daten>\*" und loeschte danach den alten Ordner ohne Pruefung.
#   Move-Item verschiebt Ordner nur innerhalb desselben Laufwerks (Microsoft
#   Learn, Move-Item > Notes) - bei einem anderen Laufwerk blieb der Ordner
#   "data" zurueck und wurde anschliessend geloescht. Gesperrte Dateien ebenso.
#   Neu: Move-DataDirectorySafe - kopieren, jede Datei per Groesse + SHA256
#   pruefen, Registry umstellen, ERST DANACH alten Ordner loeschen. Fehler in
#   einem dieser Schritte: Kopien werden entfernt, alter Ordner bleibt aktiv.
#   Nicht loeschbare Reste des alten Ordners werden gemeldet.
# - SCHUTZ: Zielordner innerhalb des aktuellen Datenordners wird abgelehnt.
# - BUGFIX: Neustart nach dem Umzug nutzte $MyInvocation.MyCommand.Path (im
#   Event-Handler leer) - jetzt $PSCommandPath wie bei der Wiederherstellung;
#   ohne gueltigen Pfad Hinweis "bitte manuell neu starten".
#
# CHANGELOG v2.27.1 (Layout Bearbeiten-Popup):
# - UI: Tab "Einzelmarker-Analyse" > Button "Bearbeiten" > Popup "Eintraege fuer
#   '<Marker>' bearbeiten": Die drei Buttons lagen versetzt in zwei Reihen und
#   ueberlappten teils. Jetzt einheitlich in EINER Reihe unter der Tabelle,
#   gleiche Groesse (146 x 34 px), buendig zur Tabellenbreite:
#   [Auswahl loeschen] [Aenderungen uebernehmen] [Abbrechen].
# - UI: Button "Aenderungen uebernehmen & Schliessen" umbenannt in
#   "Aenderungen uebernehmen". Funktion unveraendert (speichert und schliesst).
#
# CHANGELOG v2.27.0 (PREVENT-Score nach AHA-Originalmodell):
# - BUGFIX (Rechenfehler): Calculate-PreventScore nutzte frei erfundene
#   Koeffizienten (gesunder 38-Jaehriger -> 35,5 % statt < 1 %). Ersetzt durch
#   das AHA-PREVENT-Basismodell (Khan et al., Circulation 2024;149:430-449):
#   neue Funktion Calculate-PreventRisk, validiert gegen das Referenzbeispiel
#   des R-Pakets preventr (Frau 50 J.: CVD 14,7 %, ASCVD 9,2 % - exakt).
#   Gueltigkeitsbereich wie offizieller Rechner (Alter 30-79, Gesamtchol.
#   130-320, HDL 20-100 mg/dl, SBP 90-180, eGFR 15-140), ausserhalb kein Wert.
# - NEU: Marker "PREVENT-CVD-10Y" (Herz-Kreislauf gesamt inkl. Herzinsuffizienz),
#   wird in bestehenden Installationen per Marker-Migration ergaenzt.
# - NEU: Eingaben je Testdatum mit Ersatzwerten - Pflicht nur Gesamt- und
#   HDL-Cholesterin; Blutdruck = Messwert am Testtag, sonst Standardwert;
#   eGFR = GFR-Marker, sonst aus Kreatinin (neue Funktion Get-EgfrCkdEpi2021);
#   Diabetes = Checkbox ODER HbA1c >= 6,5 % (>= 48 mmol/mol).
# - NEU: Tab "Persoenliche Metriken": Checkboxen "Blutdrucksenker",
#   "Lipidsenker (Statin)", "Diabetes diagnostiziert" und Dropdown
#   "PREVENT-Anzeige (Cockpit)": 10 Jahre ASCVD (Standard) / Herz-Kreislauf
#   gesamt. Defaults: nein / ASCVD - bestehende Configs laufen unveraendert.
# - ANPASSUNG Risiko-Cockpit: zeigt den gewaehlten PREVENT-Wert; Ampel nach
#   PREVENT-Kategorien (ASCVD 3/5/10 %, gesamt 5/7,5/20 %). Ref-/Optimalbereich
#   der Marker entsprechend (ASCVD optimal < 3, Ref < 5; gesamt < 5 / < 7,5).
# - NEU: Korrektur-Migration ($script:MarkerSetCorrections) - korrigiert
#   bestehende Marker-Definitionen in der Config.json einmalig (hier:
#   PREVENT-ASCVD-10Y). Wiederverwendbar fuer TyG/AIP/PLR.
# - BEKANNTE GRENZE: Alter ist ein fester Wert -> aeltere Testdaten werden mit
#   dem aktuellen Alter berechnet (Behebung mit Review-Punkt 5, Geburtsdatum).
#
# CHANGELOG v2.26.2 (Sichere Wiederherstellung aus ZIP-Backup):
# - BUGFIX (Datenverlust): Einstellungen > "Sicherung wiederherstellen (ZIP)"
#   loeschte im Modus "Ja = Alles loeschen" ZUERST den Datenordner und kopierte
#   DANACH das Backup. Neu: Invoke-SafeDataRestore - Zielzustand im Staging-Ordner
#   neben dem Datenordner aufbauen, Austausch per Umbenennen, Rueckweg bei jedem
#   Fehler, erst nach Erfolg aufraeumen. Gilt fuer alle drei Modi (clean/merge/
#   overwrite). Gesperrte Datei/Ordner = sauberer Abbruch ohne Aenderung.
# - NEU: Get-UnreadableJsonFiles - prueft VOR der Wiederherstellung, ob die
#   Backup-Dateien auf diesem PC/Benutzerprofil lesbar sind (DPAPI). Unlesbare
#   Dateien -> Warnung, Standard = Abbrechen (Hinweis auf portablen .btbackup).
# - SCHUTZ: Ordner "lib" (PDF-Bibliothek, ggf. geladen) bleibt unangetastet bzw.
#   wird nur ergaenzt, wenn er fehlt; das laufende Script wird nie verschoben;
#   Datenordner im Laufwerksstamm wird abgelehnt (bisher: "<Laufwerk>:\*" geloescht).
# - Unveraendert: Dialoge, Registry-Import, Neustart, Struktur-Erkennung
#   (flach / verschachtelt UserData\PSC.Blood-Tracker).
#
# CHANGELOG v2.26.1 (Datensicherheit + Trendlinie):
# - BUGFIX (Datenverlust): Write-ProtectedJsonFile schreibt jetzt atomar -
#   Inhalt zuerst in "<Datei>.bttmp", danach Austausch per [IO.File]::Replace
#   (Vorversion kurzzeitig als "<Datei>.btbak"; Fallback: Kopieren). Ein Absturz
#   oder Fehler beim Schreiben hinterlaesst keine halbe Datei mehr. Dateiformat
#   unveraendert. Fehler sind jetzt per try/catch abfangbar (vorher wurden
#   Schreibfehler von Set-Content stillschweigend uebergangen).
# - BUGFIX (Datenverlust): Save-AllHistoricalData loeschte ZUERST alle
#   Tagesdateien und schrieb sie danach neu. Ein Fehler mittendrin vernichtete
#   alle noch nicht geschriebenen Tage. Neu: vorbereiten -> atomar schreiben ->
#   erst danach verwaiste Dateien geloeschter Tage entfernen (Fix v2.7.3 bleibt).
# - BUGFIX (Datenverlust): Load-AllHistoricalItems brach bei der ersten
#   unlesbaren Tagesdatei komplett ab - alle danach folgenden Tage fehlten im
#   Speicher und wurden beim naechsten Speichern geloescht. Neu: Fehler je Datei,
#   unlesbare Dateien werden gemerkt, nie geloescht (vor einem Ueberschreiben
#   als "<Datei>.unlesbar_<Zeitstempel>.bak" gesichert) und einmalig gemeldet.
# - BUGFIX (Datenverlust): Import-AesBackup loeschte den Datenordner, BEVOR die
#   importierten Dateien geschrieben waren. Neu: Backup-Inhalt vorab pruefen,
#   Tagesdaten in Staging-Ordner aufbauen, Austausch per Umbenennen, Rollback
#   wenn Config.json nicht geschrieben werden kann. Gesperrter Datenordner =
#   sauberer Abbruch ohne Aenderung.
# - BUGFIX (Trendlinie): Update-Chart nutzte "$markerData.Count > 1" - in
#   PowerShell eine Datei-Umleitung, kein Vergleich. Die Trendlinie wurde nie
#   gezeichnet, stattdessen entstand eine Datei "1" im Arbeitsverzeichnis.
#   Jetzt "-gt 1" (Count per @() PS-5.1-sicher).
# - ANPASSUNG (Trendlinie): Trend-Serie liegt UNTER der Verlaufslinie
#   (Series.Insert(0)), Maus-Tooltip ignoriert die Trend-Serie - Messwert-
#   Tooltips bleiben unveraendert erreichbar.
# - NEU: Hilfsfunktion Get-BtPathKey (normalisierter Pfadvergleich).
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

# v2.41.0: Optionale obere technische Warnschwellen, KEINE medizinischen Grenzwerte.
# $null = keine obere Warnschwelle. Quellwerte werden bei Ueberschreitung NICHT verworfen.
# Einheiten: Aktivitaetsenergie kcal/Tag; Schritte Anzahl/Tag; CardioErholung bpm.
$script:ActivityWarningUpperBounds = @{ Aktivitaetsenergie = $null; Schritte = $null; CardioErholung = $null }

# v2.45.0: Nur AnzeigeMittelwert/AnzeigeDifferenz runden; Roh-/Quellwerte bleiben erhalten.
# Ganze Zahl 0..15. Keine medizinischen Bewertungsgrenzen.
$script:ActivitySummaryDecimals = 2

# v2.43.0: Aktiven Hauptreiter mit einer dezenten Linie oberhalb der Schrift markieren.
# $false = bisherige Windows-Darstellung. Farbe als #RRGGBB, z. B. '#B84040' (dezentes Rot).
$script:ActiveTabIndicatorEnabled = $true
$script:ActiveTabIndicatorColor = '#B84040'
# v2.44.0: Aktiven Namen auch fett zeichnen (wirksam bei aktivierter Reiter-Markierung).
$script:ActiveTabBoldEnabled = $true

# ---------- Globale Konfiguration ----------
Add-Type -AssemblyName System.Windows.Forms.DataVisualization
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System # Für ArrayList, falls benötigt
Add-Type -AssemblyName System.Security # v2.19.0: DPAPI (ProtectedData)

$Version = "3.0.0" # Aktuelle Versionierung
$script:BtScriptPath = $PSCommandPath # v2.34.0: Pfad dieses Scripts (Kopfzeilen-Pruefung der Release-Notes)

# ╔══════════════════════════════════════════════════════════════════════════════════╗
# ║ v2.34.0: RELEASE-NOTES (Reiter "Release-Notes") – Pflege: Regelblock ganz oben   ║
# ║ Neueste Version GANZ OBEN, numerisch absteigend. Datum JJJJ-MM-TT oder ''.       ║
# ║ Typen: Neu | Bugfix | Anpassung | Sicherheit | Hinweis. Texte aus Benutzersicht. ║
# ║ Mehrere Versionen am selben Tag: je Version ein Eintrag mit gleichem Datum -     ║
# ║ der Reiter fasst sie automatisch zusammen (Regel 6 im Regelblock oben).          ║
# ║ Veröffentlichte Einträge NIE ändern oder löschen.                                ║
# ╚══════════════════════════════════════════════════════════════════════════════════╝
$script:ReleaseNotes = @(
    [ordered]@{ Version = '3.0.0'; Datum = '2026-10-09'; Titel = 'Blood-Tracker 3: übersichtlichere Anordnung'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Beim Erstellen und Bearbeiten eigener Ansichten steht die Schaltfläche + Marker hinzufügen jetzt direkt über der Markerliste.' }
            @{ Typ = 'Anpassung'; Text = 'Die Release-Notes zeigen aktuelle Version: anstelle von Läuft:. Der Versionshinweis ist rechtsbündig zur Detailfläche ausgerichtet; Suche und Filter haben weiterhin Platz in derselben Kopfzeile.' }
            @{ Typ = 'Anpassung'; Text = 'In der Einzelmarker-Analyse steht Exportieren jetzt links neben Drucken. Beide Schaltflächen sind gleich groß. Der Bereich Hinweis/e liegt bündig unter der Grafik und nutzt deren volle Breite.' }
        ) }
    [ordered]@{ Version = '2.49.0'; Datum = '2026-10-09'; Titel = 'Mehr Platz durch Kategorien und Unterreiter'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Navigation hat jetzt zwei Zeilen: oben Übersicht, Auswertung, Daten & Profil und Info; darunter die Reiter der ausgewählten Kategorie. Einstellungen und Beenden bleiben oben rechts erreichbar.' }
            @{ Typ = 'Neu'; Text = 'Während einer Sitzung wird der zuletzt geöffnete Reiter je Kategorie gemerkt. Beim Programmstart erscheint immer das Dashboard. Dashboard-Verknüpfungen öffnen automatisch die passende Kategorie und den Zielreiter.' }
            @{ Typ = 'Anpassung'; Text = 'Die aktive Kategorie und der aktive Reiter werden hervorgehoben. Das Fenster bietet Platz für die zusätzliche Zeile; auf kleineren Bildschirmen lässt sich der Inhaltsbereich scrollen.' }
            @{ Typ = 'Bugfix'; Text = 'Noch nicht gespeicherte persönliche Eingaben bleiben beim Wechsel der Reiter erhalten. Gespeicherte Daten werden weiterhin beim Start und nach dem Datenimport geladen.' }
        ) }
    [ordered]@{ Version = '2.48.2'; Datum = '2026-10-09'; Titel = 'Reiterwechsel für die Kategorien-Navigation vorbereitet'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Die Verknüpfungen im Dashboard zur Einzelmarker-Analyse und zu den persönlichen Metriken wurden für die kommende Kategorien-Navigation vorbereitet. Die sichtbaren Reiter und ihre Bedienung bleiben in dieser Version gleich.' }
            @{ Typ = 'Bugfix'; Text = 'Ist vorübergehend kein Reiter ausgewählt, wird die Aktualisierung der Reiterinhalte sicher ausgesetzt.' }
        ) }
    [ordered]@{ Version = '2.48.1'; Datum = '2026-10-09'; Titel = 'Schaltfläche für die Blutbild-Eingabe umbenannt'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Die Schaltfläche zum Öffnen der Formulareingabe heißt jetzt Werte nach Blutbild speichern.' }
        ) }
    [ordered]@{ Version = '2.48.0'; Datum = '2026-10-09'; Titel = 'Mehrere Blutwerte bequem im Formular erfassen'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Unter Werte erfassen stehen Einzelwert speichern und Mehrere Werte als zwei Eingabewege zur Verfügung. Das Formular bietet ein gemeinsames Datum, beliebig ergänzbare Zeilen, Markersuche und eine Notiz je Wert.' }
            @{ Typ = 'Neu'; Text = 'HIV und APOE werden ausdrücklich ausgewählt, ohne vorbelegtes Ergebnis. Für Testosteron stehen ng/dl und nmol/l zur Wahl. Eine Prüfliste zeigt Eingaben, Umrechnungen und Referenzhinweise vor dem gemeinsamen Speichern.' }
            @{ Typ = 'Sicherheit'; Text = 'Unvollständige oder doppelte Werte verhindern das Speichern; vollständig leere Zeilen werden ignoriert. Bei Speicherfehlern bleiben die Eingaben erhalten. Beim Schließen ungespeicherter Eingaben wird nachgefragt.' }
            @{ Typ = 'Anpassung'; Text = 'Die Marker-Verwaltung und die Warnschwelle wurden unterhalb der gemeinsamen Eingabegruppe angeordnet. Auch die erweiterten Einstellungen werden berücksichtigt. Der APOE-Hinweis lautet jetzt Genotyp wählen.' }
        ) }
    [ordered]@{ Version = '2.47.0'; Datum = '2026-10-09'; Titel = 'Blutwerte gemeinsam und zuverlässig speichern'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Speicherfunktion kann mehrere vorbereitete Blutwerte für ein gemeinsames Datum in einem Vorgang übernehmen. Sie prüft alle Werte und den aktuellen Tagesbestand vor dem Schreiben.' }
            @{ Typ = 'Sicherheit'; Text = 'Bei doppelten Werten, unlesbaren Tagesdaten oder einem fehlgeschlagenen atomaren Ersetzen wird das gemeinsame Speichern abgebrochen. Die bisherige Tagesdatei bleibt erhalten.' }
            @{ Typ = 'Bugfix'; Text = 'Auch die Einzeleingabe verwendet den neuen Speicherweg. Bei einem Speicherfehler bleiben Wert und Notiz im Eingabebereich stehen.' }
            @{ Typ = 'Bugfix'; Text = 'Die APOE-Auswahlliste ist schmaler und überlappt den Hinweistext nicht mehr.' }
            @{ Typ = 'Hinweis'; Text = 'Das zusätzliche Formular folgt im nächsten Schritt. Das bisherige Dateiformat bleibt erhalten.' }
        ) }
    [ordered]@{ Version = '2.46.0'; Datum = '2026-10-09'; Titel = 'Gemeinsame Prüfung der Blutwerteingabe'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Die bisherige Einzeleingabe verwendet jetzt eine gemeinsame Prüfung für die spätere Formulareingabe. Markersuche, HIV, APOE, Notizen und die Prüfung auf doppelte Werte bleiben verfügbar.' }
            @{ Typ = 'Anpassung'; Text = 'Testosteron-Eingaben in nmol/l werden weiterhin in ng/dl umgerechnet. Exakte Rundungsmitten werden nun ausdrücklich von null weg gerundet; die ursprüngliche Eingabe bleibt in der Notiz erhalten.' }
            @{ Typ = 'Bugfix'; Text = 'Nicht endliche Zahlen und Werte, deren Umrechnung den Zahlenbereich überschreitet, werden mit einer Fehlermeldung abgelehnt.' }
            @{ Typ = 'Hinweis'; Text = 'Dieser Vorbereitungsschritt ergänzt noch kein Formular. Anordnung und Speicherung bleiben unverändert.' }
        ) }
    [ordered]@{ Version = '2.45.0'; Datum = '2026-10-09'; Titel = 'Aktivität seit Jahresbeginn und über sechs Monate auswerten'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Auswertungsfunktionen berechnen Durchschnitte für Aktivitätsenergie, Schritte und vergleichbare Cardio-Erholungswerte seit Jahresbeginn oder über die letzten sechs Kalendermonate. Dieser Schritt ergänzt die Engine, noch keine Bedienoberfläche.' }
            @{ Typ = 'Anpassung'; Text = 'Fehlende Tage zählen nicht als null. Unvollständige Tage bleiben ausgeschlossen. Anzahl, Datenabdeckung und Datum des letzten verwendeten Wertes machen die Grundlage der Berechnung sichtbar.' }
            @{ Typ = 'Hinweis'; Text = 'Cardio-Trainingswerte und Tagesmittel bleiben getrennt. Unbekannte Intervalle oder Verfahren werden nicht automatisch zusammengefasst. Negative Quellenwerte bleiben erhalten.' }
            @{ Typ = 'Hinweis'; Text = 'Übernommene Durchschnitte für genau dasselbe Zeitfenster haben Vorrang und bleiben unverändert neben der eigenen Berechnung erhalten. Unterschiede werden mit Hinweisen zur Datenabdeckung ausgewiesen; andere Zeiträume werden nicht anteilig umgerechnet.' }
        ) }
    [ordered]@{ Version = '2.44.0'; Datum = '2026-10-09'; Titel = 'Aktiven Reiter zusätzlich fett anzeigen'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Der ausgewählte Hauptreiter wird zusätzlich zur roten Linie fett beschriftet. Die Hervorhebung folgt der Auswahl per Maus und Tastatur.' }
            @{ Typ = 'Anpassung'; Text = 'Die Reiter erhalten Platz für die fette Beschriftung. Die Schrift der Inhalte bleibt unverändert. Fettschrift lässt sich in USER-KONFIGURATION separat ausschalten.' }
        ) }
    [ordered]@{ Version = '2.43.0'; Datum = '2026-10-09'; Titel = 'Aktiven Reiter leichter erkennen'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Der ausgewählte Hauptreiter erhält eine dezente rote Linie oberhalb seiner Beschriftung. Die Markierung folgt der Auswahl per Maus und Tastatur.' }
            @{ Typ = 'Anpassung'; Text = 'Die Markierung kann in USER-KONFIGURATION ausgeschaltet und ihre Farbe angepasst werden. Reiternamen, Inhalte und Bedienung bleiben erhalten.' }
        ) }
    [ordered]@{ Version = '2.42.0'; Datum = '2026-10-09'; Titel = 'Durchschnittswerte aus einer App für ganze Zeiträume speichern'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Engine kann übernommene Durchschnittswerte für Aktivitätsenergie, Schritte und Cardio-Erholung mit Anfangs- und Enddatum getrennt speichern, laden, ausdrücklich ersetzen und löschen.' }
            @{ Typ = 'Neu'; Text = 'Quelle, Gerät, Berechnungsbasis und bekannte Anzahl berücksichtigter Tage oder Trainingsmessungen bleiben erhalten. Unbekannte Angaben bleiben unbekannt. Ein Schritt-Durchschnitt darf Nachkommastellen haben.' }
            @{ Typ = 'Anpassung'; Text = 'Cardio-Mittel aus Trainingswerten und Tagesmitteln werden unterschieden; auch Erholungsintervall und Messverfahren bleiben festgehalten. Negative Quellenwerte werden mit Hinweis erhalten.' }
            @{ Typ = 'Sicherheit'; Text = 'Überlappende Zeiträume bleiben getrennt. Quellenmittel werden weder auf einzelne Tage verteilt noch miteinander verrechnet. Doppelte Einträge und unbeabsichtigtes Überschreiben werden verhindert.' }
            @{ Typ = 'Hinweis'; Text = 'Dieser Schritt enthält keine neue Oberfläche und noch keine YTD- oder Sechsmonats-Berechnung. Die bisherigen Tages- und Trainingswerte bleiben getrennt von den Zeitraum-Mitteln.' }
        ) }
    [ordered]@{ Version = '2.41.0'; Datum = '2026-10-08'; Titel = 'Aktivitätswerte und Cardio-Erholung: Speicherung vorbereitet'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Messwert-Engine kann Tagesgesamtwerte für Aktivitätsenergie und Schritte sowie übernommene Tagesmittel der Cardio-Erholung speichern und laden. Cardio-Erholung je Training wird getrennt in Fitness gespeichert.' }
            @{ Typ = 'Neu'; Text = 'Quelle und Gerätemodell können frei angegeben werden. Erholungsintervall, bekannte Anzahl der Messungen und der Abschluss eines Tages bleiben als Zusatzangaben erhalten; unbekannte Angaben werden nicht erfunden.' }
            @{ Typ = 'Anpassung'; Text = 'Energie und Schritte dürfen nicht negativ sein; Tagesschritte müssen ganzzahlig sein. Negative Cardio-Erholungswerte bleiben mit Hinweis erhalten. Obere technische Warnschwellen sind optional einstellbar und verändern keine Werte.' }
            @{ Typ = 'Sicherheit'; Text = 'Tageswerte werden nicht automatisch addiert oder überschrieben. Ein Ersetzen muss ausdrücklich angefordert werden und betrifft den gesamten Tages-Eintrag. Die Ablage bleibt verschlüsselt.' }
            @{ Typ = 'Hinweis'; Text = 'Dieser Schritt erweitert ausschließlich die Engine. Eingabefelder, Zeitraum-Durchschnitte, YTD- und Sechsmonats-Auswertungen sind noch nicht enthalten. Die neuen Werte erhalten keine medizinische Bewertung.' }
        ) }
    [ordered]@{ Version = '2.40.0'; Datum = '2026-10-08'; Titel = 'Eigene Dashboard-Ansichten erstellen und verwalten'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Über + Eigene Ansicht… lassen sich benannte Ansichten aus gemessenen und berechneten Markern zusammenstellen. Die Marker können per Name, Kürzel oder Alias gesucht und zeilenweise hinzugefügt oder entfernt werden.' }
            @{ Typ = 'Neu'; Text = 'Eigene Ansichten stehen im Dashboard zur Auswahl, zeigen vier Kacheln pro Zeile und lassen sich bearbeiten oder nach Rückfrage löschen. Ein Kachel-Klick öffnet die Einzelmarker-Analyse.' }
            @{ Typ = 'Anpassung'; Text = 'Doppelte Marker und bereits verwendete oder reservierte Ansichtsnamen werden verhindert. Marker ohne Daten und nicht mehr verfügbare Marker bleiben erkennbar.' }
            @{ Typ = 'Sicherheit'; Text = 'Abbrechen verändert nichts. Bei einem Speicherfehler bleibt der Dialog offen; Änderungen werden erst nach erfolgreichem Speichern übernommen.' }
            @{ Typ = 'Hinweis'; Text = 'Gespeicherte Ansichten bleiben nach einem Neustart erhalten. Das Dashboard startet weiterhin mit Standard; nach dem Löschen einer Ansicht wird ebenfalls Standard angezeigt.' }
        ) }
    [ordered]@{ Version = '2.39.0'; Datum = '2026-10-08'; Titel = 'Eigene Dashboard-Ansichten: Speicherung vorbereitet'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Konfiguration kann eigene Dashboard-Ansichten mit dauerhaft eindeutiger Kennung, Name und geordneter Markerliste speichern und wieder laden. Die Oberfläche zum Erstellen und Bearbeiten folgt in einem eigenen Schritt.' }
            @{ Typ = 'Anpassung'; Text = 'Vorhandene Konfigurationen ohne eigene Ansichten erhalten eine leere Liste. Auch einzelne Ansichten und einzelne Marker bleiben beim Speichern und Laden erhalten.' }
            @{ Typ = 'Sicherheit'; Text = 'Fehlerhafte Ansichtsdefinitionen werden gemeldet. Die Konfiguration wird damit beim Laden nicht automatisch überschrieben; auch ein ausdrückliches Speichern wird bis zur Korrektur abgelehnt.' }
            @{ Typ = 'Anpassung'; Text = 'Änderungen an gespeicherten Ansichten werden bei der Dashboard-Aktualisierung berücksichtigt. Verweise auf nicht mehr vorhandene Marker bleiben erhalten.' }
        ) }
    [ordered]@{ Version = '2.38.0'; Datum = '2026-10-08'; Titel = 'Dashboard: Aktuelle Auffälligkeiten'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Die Dashboard-Auswahl bietet Aktuelle Auffälligkeiten mit der Anzahl betroffener Marker an, sobald mindestens ein passender Wert vorliegt. Gemessene und berechnete Marker erscheinen als anklickbare Kacheln.' }
            @{ Typ = 'Hinweis'; Text = 'Es zählt der letzte Wert bis zum aktuellen Tag, höchstens 365 Tage alt: gültig, ohne widersprüchliche Werte am selben Tag und außerhalb eines vollständig hinterlegten Referenzbereichs, mit der bestehenden Bewertung Außerhalb der Norm. Marker der Gruppe Genetik werden nicht berücksichtigt.' }
            @{ Typ = 'Hinweis'; Text = 'Die Reihenfolge richtet sich nach dem bestehenden Bewertungsscore, niedrigere Werte zuerst, bei Gleichstand nach dem Markernamen. Die Sortierung ist keine medizinische Dringlichkeitseinstufung.' }
            @{ Typ = 'Neu'; Text = 'Ändert sich die Anzahl der Auffälligkeiten, bleibt die Ansicht ausgewählt. Gibt es keine passenden Werte mehr, verschwindet der Eintrag und das Dashboard wechselt zurück zu Standard.' }
        ) }
    [ordered]@{ Version = '2.37.0'; Datum = '2026-10-08'; Titel = 'Dashboard: Standard und vollständige Vorbelastungsansicht'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Im Dashboard lässt sich zwischen Standard und Vorbelastungen umschalten. Bei jedem Programmstart ist Standard ausgewählt.' }
            @{ Typ = 'Neu'; Text = 'Die Ansicht Vorbelastungen zeigt alle zugeordneten Marker je aktiver Vorbelastung als Kacheln: vier pro Zeile, mit Wert, Bewertung und Verlauf. Auch berechnete Marker und Marker ohne Messwerte werden angezeigt.' }
            @{ Typ = 'Neu'; Text = 'Sind Vorbelastungen ausgeschaltet oder keine aktiv, führt ein Hinweis zu Persönliche Metriken. Ein Kachel-Klick öffnet den zugehörigen konfigurierten Marker in der Einzelmarker-Analyse.' }
            @{ Typ = 'Anpassung'; Text = 'Die Infozeile steht unter der Ansichtsauswahl. Die bisherigen Karten und Kacheln der Standardansicht behalten ihre Anordnung.' }
        ) }
    [ordered]@{ Version = '2.36.0'; Datum = '2026-10-08'; Titel = 'Neuer Reiter "Dashboard": Gesundheitsstatus auf einen Blick'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Dashboard" als erster Reiter: Gesundheitsscore (0 bis 100) mit Status und Vergleich zum vorherigen Bluttest, darunter der Longevity-Score.' }
            @{ Typ = 'Neu'; Text = 'Sechs Longevity-Kacheln (ApoB, HbA1c, Nüchterninsulin, hs-CRP bzw. CRP, Homocystein, PhenoAge − Alter) mit aktuellem Wert, Bewertung, Verlauf und Trend. Klick, Enter oder Leertaste öffnet den Marker im Reiter "Einzelmarker-Analyse".' }
            @{ Typ = 'Neu'; Text = 'Vorbelastungen: je aktiver Vorbelastung ein Score, die Anzahl bewerteter Marker und die drei auffälligsten Marker. Ohne aktive Vorbelastung führt ein Link zu "Persönliche Metriken".' }
            @{ Typ = 'Neu'; Text = 'Karte Testosteron: Gesamt-Testosteron, freies Testosteron (gemessen) und SHBG mit Referenz- und Optimalbereich; zusätzlich berechnet: freies Testosteron nach Vermeulen und Freier Androgen-Index (FAI), ohne Bewertung.' }
            @{ Typ = 'Neu'; Text = '"Wie berechnet?" erklärt Rechenweg, Funktionsbereiche und nicht bewertete Marker.' }
            @{ Typ = 'Hinweis'; Text = 'Der Gesundheitsscore fasst die hinterlegten Markerbewertungen zusammen: Funktionsbereiche zählen gleich viel, nur Werte der letzten 365 Tage, mindestens 5 Marker aus 2 Bereichen. 100 Punkte bedeuten die beste Bewertung nach diesen Regeln, nicht vollständige Gesundheit.' }
            @{ Typ = 'Anpassung'; Text = 'Beim Start öffnet sich das Dashboard. Es rechnet nur neu, wenn sich Messwerte, Marker-Einstellungen, persönliche Metriken oder Vorbelastungen geändert haben; alle übrigen Reiter sind unverändert.' }
        ) }
    [ordered]@{ Version = '2.35.0'; Datum = '2026-10-08'; Titel = 'Sortieren in Cockpit und Longevity; weitere Messwerte vorbereitet'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Risiko-Cockpit und Longevity-Indizes: Jede Spalte lässt sich per Klick auf die Überschrift sortieren – erster Klick aufsteigend, zweiter absteigend, dritter wieder die Standardreihenfolge. Ein Pfeil zeigt die aktive Sortierung.' }
            @{ Typ = 'Anpassung'; Text = 'Zahlen werden nach Zahlenwert, das Datum chronologisch und die Bewertung nach Schwere sortiert; leere Werte und "-" stehen immer unten. Die gewählte Sortierung bleibt bei Filterwechsel und Aktualisierung erhalten.' }
            @{ Typ = 'Anpassung'; Text = 'Risiko-Cockpit: Die Standardreihenfolge ist jetzt fest nach Bewertung, dann Funktionsbereich und Marker.' }
            @{ Typ = 'Bugfix'; Text = 'Beim Sortieren wurden Datum und Werte bisher als Text verglichen (z. B. "10" vor "9"); Spalten mit gemischtem Inhalt (Zahl und "-") ließen sich nicht zuverlässig sortieren.' }
            @{ Typ = 'Bugfix'; Text = 'Risiko-Cockpit: Der Funktionsbereich erschien innerhalb einer Gruppe in jeder zweiten Zeile erneut. Er steht jetzt nur in der ersten Zeile der Gruppe (Standardansicht und Sortierung nach Funktionsbereich), bei anderer Sortierung in jeder Zeile.' }
            @{ Typ = 'Neu'; Text = 'Vorbereitet (noch ohne Eingabemaske): Speicherung von Herzfrequenz, Blutdruck, Körpertemperatur, EKG (Rhythmus, Herzfrequenz, PQ, QRS, QT, QTc, Befund), Gewicht, BMI, Körperfettanteil, fettfreier Masse und VO2max – mehrere Messungen pro Tag mit Uhrzeit, verschlüsselt und in allen Sicherungen enthalten.' }
            @{ Typ = 'Hinweis'; Text = 'Die bisherigen Blutmarker "Gewicht" und "Systolischer Blutdruck" sowie der PREVENT-Score bleiben unverändert.' }
        ) }
    [ordered]@{ Version = '2.34.1'; Datum = '2026-10-08'; Titel = 'Release-Notes: Versionen eines Tages zusammengefasst'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Mehrere Versionen vom selben Tag stehen in der Liste als EIN Eintrag: neueste Version, Datum und Anzahl der Versionen. Rechts erscheinen alle Versionen dieses Tages untereinander, neueste zuerst.' }
            @{ Typ = 'Anpassung'; Text = 'Versionen ohne bekanntes Datum zeigen nur die Versionsnummer und sind in der Liste kompakter.' }
            @{ Typ = 'Hinweis'; Text = 'Suche und Typ-Filter wirken wie bisher; ein Tag zeigt dann nur seine passenden Versionen, die Auswahl bleibt auf diesem Tag.' }
        ) }
    [ordered]@{ Version = '2.34.0'; Datum = '2026-10-08'; Titel = 'Neuer Reiter "Release-Notes"'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Release-Notes" (letzter Reiter): alle Versionen mit Datum, Titel und Änderungen. Neueste Version oben, die laufende Version ist mit "aktuell" gekennzeichnet.' }
            @{ Typ = 'Neu'; Text = 'Suche in Version, Titel und Text sowie Filter nach Typ (Neu, Bugfix, Anpassung, Sicherheit, Hinweis); beide wirken zusammen. Esc leert die Suche, "Filter zurücksetzen" stellt alles wieder her.' }
            @{ Typ = 'Neu'; Text = 'Änderungsverlauf ab Version 2.7 übernommen.' }
            @{ Typ = 'Anpassung'; Text = 'Reichen die Reiterköpfe bis unter die Schaltflächen "globale Einstellungen" und "Beenden", werden sie automatisch etwas schmaler.' }
            @{ Typ = 'Hinweis'; Text = 'Pflicht-Regelblock oben im Script: Jede Änderung (auch durch KI-Modelle) erhöht die Version und ergänzt hier einen Eintrag. Der BETA-Testmodus prüft das beim Start und warnt bei Abweichungen.' }
        ) }
    [ordered]@{ Version = '2.33.0'; Datum = '2026-10-07'; Titel = 'Grundlage für DEXA-Scan und Schlaf (noch ohne Oberfläche)'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Vorbereitet: Katalog aller üblichen DEXA-Messwerte (61 Werte: Körperzusammensetzung, Fettverteilung, Viszeralfett, Körperregionen, Knochendichte mit T- und Z-Score, TBS) und Schlafdaten (Dauer, Score, Schlafphasen, Zyklen).' }
            @{ Typ = 'Neu'; Text = 'Ablage verschlüsselt im Datenordner unter "_messungen" – damit automatisch in allen Sicherungen und Wiederherstellungen enthalten.' }
            @{ Typ = 'Hinweis'; Text = 'Noch keine Eingabemaske; keine Änderung an Oberfläche, Einstellungen oder vorhandenen Daten.' }
        ) }
    [ordered]@{ Version = '2.32.0'; Datum = '2026-10-07'; Titel = 'Suche in "Verlauf anzeigen für"'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Einzelmarker-Analyse": Die Auswahl "Verlauf anzeigen für" lässt sich wie das Feld "Blutmarker" per Eingabe durchsuchen – Name, Teil des Namens oder Kürzel (z. B. "HbA1c", "WBC", "PREVENT"), auch berechnete Marker.' }
            @{ Typ = 'Neu'; Text = 'Die Liste filtert beim Tippen; Enter übernimmt einen eindeutigen Treffer, Esc stellt die bisherige Auswahl wieder her. Kein Treffer: Feld rot hinterlegt.' }
        ) }
    [ordered]@{ Version = '2.31.3'; Datum = '2026-10-07'; Titel = 'Risiko-Cockpit: berechnete Marker mit nur einem Wert'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Ein berechneter Marker mit genau einem Testtag (z. B. HOMA-IR nach der ersten Insulin-Messung) fehlte im Risiko-Cockpit, seine Gruppe fehlte in den Cockpit-Einstellungen.' }
        ) }
    [ordered]@{ Version = '2.31.2'; Datum = '2026-10-07'; Titel = 'Einstellungen: Listen mit nur einem Eintrag'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Beim Laden der Einstellungen wurde jede Liste mit genau einem Eintrag zu einem Einzelwert. Folge: Der BMI (und jeder eigene berechnete Marker mit nur einem Pflichtmarker) wurde nach dem ersten Neustart nie berechnet. Leere Listen gingen verloren.' }
        ) }
    [ordered]@{ Version = '2.31.1'; Datum = '2026-10-07'; Titel = 'Einzelmarker-Analyse: letzter Wert bei nur einem Wert'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Die Zeile "Letzter Wert" über dem Verlauf und die Beschriftung am Messpunkt erschienen erst ab zwei Werten – jetzt auch bei genau einem Wert.' }
        ) }
    [ordered]@{ Version = '2.31.0'; Datum = '2026-10-07'; Titel = 'hs-CRP vor CRP in den Scores'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'PhenoAge, PhenoAge-Accel, InflammAging-Score und Longevity-Score verwenden am Testtag vorrangig hs-CRP und nur ohne hs-CRP das klassische CRP (genauer im niedrigen Bereich unter 3 mg/l).' }
            @{ Typ = 'Anpassung'; Text = 'Reiter "Longevity-Indizes": Scores werden kaufmännisch auf ganze Punkte gerundet (64,5 → 65).' }
            @{ Typ = 'Bugfix'; Text = 'Reiter "Longevity-Indizes": Bei PhenoAge, PhenoAge-Accel, NLR, TyG-Index und TG/HDL-Ratio fehlten Ampelfarbe und Hinweistext der Score-Spalte.' }
        ) }
    [ordered]@{ Version = '2.30.0'; Datum = '2026-10-07'; Titel = 'Longevity-Score an den eigenen Bereichen'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Longevity-Score: Jeder Wert wird an den Referenz- und Optimalbereichen seines Markers bewertet (optimal = 100, im Referenzbereich 79–50, außerhalb 49–0). Die bisherigen festen Formeln widersprachen den eigenen Bereichen (gesunder Datensatz ca. 33/100).' }
            @{ Typ = 'Bugfix'; Text = 'Der Score entstand nur, wenn am selben Tag alle 7 Werte inkl. NAD+ vorlagen. Neu: Mittelwert der vorhandenen Werte, mindestens 4 von 7; NAD+ ist optional.' }
            @{ Typ = 'Anpassung'; Text = 'Reiter "Longevity-Indizes": Einzel-Scores nach derselben Logik; Spalte "Ideal" zeigt den Optimalbereich aus den Marker-Einstellungen.' }
        ) }
    [ordered]@{ Version = '2.29.0'; Datum = '2026-10-07'; Titel = 'Geburtsdatum statt festem Alter'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Persönliche Metriken": Das Geburtsdatum (TT.MM.JJJJ) ersetzt das feste Alter. Das Alter wird je Bluttest am Testtag berechnet – für PhenoAge, PhenoAge-Accel, PREVENT und eGFR.' }
            @{ Typ = 'Anpassung'; Text = 'Neben dem Feld steht das heutige Alter. Das Speichern der persönlichen Daten erfordert ein gültiges Geburtsdatum.' }
            @{ Typ = 'Hinweis'; Text = 'Ohne gespeichertes Geburtsdatum wird wie bisher mit dem festen Alter gerechnet (Hinweis neben dem Feld).' }
        ) }
    [ordered]@{ Version = '2.28.0'; Datum = '2026-10-07'; Titel = 'Korrekturen der Rechenwerte und Referenzbereiche'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'TyG-Index: Formel korrigiert auf ln(TG × Glukose) / 2; Referenz bis 4,68, optimal unter 4,5.' }
            @{ Typ = 'Bugfix'; Text = 'AIP: Berechnung jetzt in mmol/l, passend zu den Grenzwerten (Referenz bis 0,21, optimal unter 0,11).' }
            @{ Typ = 'Bugfix'; Text = 'PLR mit absoluten Lymphozyten (Leukozyten sind jetzt Pflichtmarker); InflammAging-Score entsprechend korrigiert.' }
            @{ Typ = 'Bugfix'; Text = 'Hämatokrit: Referenz 38–49 % (bisher 0,38–0,49); gespeicherte Werte in l/l werden beim Laden in % umgerechnet.' }
            @{ Typ = 'Bugfix'; Text = 'BMI: wurde wegen einer fehlerhaften Standardformel nie berechnet. Ungültige Ergebnisse (z. B. ohne Körpergröße) werden nicht mehr angezeigt.' }
            @{ Typ = 'Bugfix'; Text = 'Markersuche: "BNP", "Troponin" und "Kreatinin-Clearance" werden nicht mehr falschen Markern zugeordnet.' }
            @{ Typ = 'Anpassung'; Text = 'LDL und Triglyceride: Optimalbereich ab 0 – niedrige Werte gelten nicht mehr als suboptimal.' }
            @{ Typ = 'Hinweis'; Text = 'Korrekturen werden nur übernommen, wenn noch der alte Standardwert gespeichert ist; eigene Anpassungen bleiben erhalten.' }
        ) }
    [ordered]@{ Version = '2.27.2'; Datum = '2026-10-07'; Titel = 'Sicherer Datenpfad-Umzug'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = '"Datenpfad ändern...": Beim Umzug auf ein anderes Laufwerk blieb der Ordner "data" zurück und wurde anschließend gelöscht (Datenverlust).' }
            @{ Typ = 'Sicherheit'; Text = 'Neuer Ablauf: kopieren, jede Datei per Größe und SHA256 prüfen, erst danach umstellen und den alten Ordner löschen. Bei einem Fehler bleibt der alte Ordner aktiv.' }
            @{ Typ = 'Sicherheit'; Text = 'Ein Zielordner innerhalb des aktuellen Datenordners wird abgelehnt.' }
            @{ Typ = 'Bugfix'; Text = 'Der Neustart nach dem Umzug funktioniert wieder (sonst Hinweis "bitte manuell neu starten").' }
        ) }
    [ordered]@{ Version = '2.27.1'; Datum = '2026-10-07'; Titel = 'Layout Bearbeiten-Fenster'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Reiter "Einzelmarker-Analyse" > "Bearbeiten": Die drei Schaltflächen liegen jetzt in einer Reihe unter der Tabelle (Auswahl löschen, Änderungen übernehmen, Abbrechen); bisher versetzt und teils überlappend.' }
            @{ Typ = 'Anpassung'; Text = '"Änderungen übernehmen & Schließen" heißt jetzt "Änderungen übernehmen" (Funktion unverändert).' }
        ) }
    [ordered]@{ Version = '2.27.0'; Datum = '2026-10-07'; Titel = 'PREVENT-Score nach AHA-Originalmodell'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Der PREVENT-Score rechnete mit falschen Koeffizienten (gesunder 38-Jähriger 35,5 % statt unter 1 %). Jetzt AHA-PREVENT-Basismodell (Khan et al., Circulation 2024), geprüft am Referenzbeispiel.' }
            @{ Typ = 'Neu'; Text = 'Marker "PREVENT-CVD-10Y" (Herz-Kreislauf gesamt inkl. Herzinsuffizienz).' }
            @{ Typ = 'Neu'; Text = 'Ersatzwerte je Testdatum: Pflicht sind nur Gesamt- und HDL-Cholesterin; Blutdruck, eGFR (auch aus Kreatinin) und Diabetes (Checkbox oder HbA1c ab 6,5 %) werden ergänzt.' }
            @{ Typ = 'Neu'; Text = 'Reiter "Persönliche Metriken": Blutdrucksenker, Lipidsenker (Statin), Diabetes diagnostiziert und Auswahl der PREVENT-Anzeige im Cockpit.' }
            @{ Typ = 'Anpassung'; Text = 'Risiko-Cockpit: Ampel nach PREVENT-Kategorien (ASCVD 3/5/10 %, gesamt 5/7,5/20 %).' }
            @{ Typ = 'Hinweis'; Text = 'Gültig für Alter 30–79 und die Wertebereiche des offiziellen Rechners; außerhalb wird kein Wert berechnet.' }
        ) }
    [ordered]@{ Version = '2.26.2'; Datum = '2026-10-07'; Titel = 'Sichere Wiederherstellung aus ZIP-Backup'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = '"Sicherung wiederherstellen (ZIP)" löschte im Modus "Alles löschen" zuerst den Datenordner und kopierte erst danach das Backup. Jetzt: Aufbau in einem Zwischenordner, Austausch erst nach Erfolg, Rückweg bei jedem Fehler – in allen drei Modi.' }
            @{ Typ = 'Sicherheit'; Text = 'Vor der Wiederherstellung wird geprüft, ob die Backup-Dateien auf diesem PC und Benutzerprofil lesbar sind; sonst Warnung mit der Vorauswahl "Abbrechen".' }
            @{ Typ = 'Sicherheit'; Text = 'PDF-Bibliothek und laufendes Script bleiben unangetastet; ein Datenordner im Laufwerksstamm wird abgelehnt.' }
        ) }
    [ordered]@{ Version = '2.26.1'; Datum = '2026-10-07'; Titel = 'Datensicherheit und Trendlinie'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Speichern schreibt jetzt atomar (erst Zwischendatei, dann Austausch) – ein Absturz hinterlässt keine halbe Datei mehr.' }
            @{ Typ = 'Bugfix'; Text = 'Beim Speichern aller Tage wurden zuerst alle Tagesdateien gelöscht; ein Fehler mittendrin konnte Daten vernichten. Jetzt erst schreiben, dann aufräumen.' }
            @{ Typ = 'Bugfix'; Text = 'Eine unlesbare Tagesdatei brach das Laden ab – alle folgenden Tage fehlten und wurden beim nächsten Speichern gelöscht. Unlesbare Dateien werden jetzt gemeldet und nie gelöscht.' }
            @{ Typ = 'Bugfix'; Text = '"Backup importieren" löschte den Datenordner, bevor die neuen Dateien geschrieben waren. Jetzt Prüfung vorab und Austausch erst nach Erfolg.' }
            @{ Typ = 'Bugfix'; Text = 'Reiter "Einzelmarker-Analyse": Die Trendlinie wurde nie gezeichnet (stattdessen entstand eine Datei "1" im Arbeitsordner).' }
            @{ Typ = 'Anpassung'; Text = 'Die Trendlinie liegt unter der Verlaufslinie; die Tooltips der Messwerte bleiben erreichbar.' }
        ) }
    [ordered]@{ Version = '2.26.0'; Datum = '2026-08-20'; Titel = 'Marker-Verwaltung, 3 neue Marker, Report-Seitenumbruch'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = '"Zum Bearbeiten laden" aktiviert automatisch "Erweiterte Einstellungen" – Gruppe und Bereiche sind sofort sichtbar.' }
            @{ Typ = 'Neu'; Text = 'Marker "Apolipoprotein E-Genotyp (APOE)" (neue Gruppe "Genetik"): Erfassung per Auswahlliste, Anzeige im Klartext (z. B. E3/E3).' }
            @{ Typ = 'Neu'; Text = 'Marker "Zink" (µg/dl) und "Magnesium" (mmol/l) in der Gruppe Mineralstoffe.' }
            @{ Typ = 'Neu'; Text = 'Custom Report (PDF): Jeder Blutmarker beginnt auf einer neuen Seite; Papierformat A4 quer.' }
            @{ Typ = 'Bugfix'; Text = 'Die Fußzeile des Custom Reports erschien nur auf der letzten Seite – jetzt auf jeder Seite.' }
            @{ Typ = 'Bugfix'; Text = 'Der PDF-Import nimmt qualitative Marker (HIV, APOE) aus – vorher waren Fehlzuordnungen möglich.' }
            @{ Typ = 'Hinweis'; Text = 'Referenzbereiche sind methoden- und laborabhängig; maßgeblich ist der Referenzbereich des eigenen Befunds.' }
        ) }
    [ordered]@{ Version = '2.25.0'; Datum = ''; Titel = 'Neue Blutmarker: hs-CRP und SHBG'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Marker "Hochsensitives CRP (hs-CRP)" (mg/l, Gruppe Entzündung): Referenz 0–3 mg/l, optimal 0–1 mg/l (AHA/CDC).' }
            @{ Typ = 'Neu'; Text = 'Marker "Sexualhormon-bindendes Globulin (SHBG)" (nmol/l, Gruppe Hormone): Referenz 18,3–54,1 nmol/l (Mann, 20–49 Jahre).' }
            @{ Typ = 'Bugfix'; Text = 'Neue Standard-Marker erscheinen jetzt auch in bestehenden Installationen; angepasste und bewusst gelöschte Marker bleiben unberührt.' }
            @{ Typ = 'Anpassung'; Text = 'Kürzel wie "hs-CRP" oder "hsCRP" führen zum neuen Marker, "CRP" weiterhin zum klassischen CRP.' }
            @{ Typ = 'Bugfix'; Text = 'Der PDF-Import ordnete "hs-CRP"-Zeilen dem klassischen CRP zu.' }
            @{ Typ = 'Neu'; Text = 'Fehlt in einer Score-Berechnung das klassische CRP, wird hs-CRP verwendet; hs-CRP erscheint als eigene Zeile im Reiter "Longevity-Indizes" und in zwei Vorbelastungen.' }
        ) }
    [ordered]@{ Version = '2.24.1'; Datum = ''; Titel = 'Dokument-Zuordnung pro Bluttest'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = '"Dokument anzeigen" zählte alle Dokumente des Monats und öffnete das zuletzt geänderte – jetzt nur die Dokumente des gewählten Bluttests.' }
            @{ Typ = 'Neu'; Text = 'Mehrere Dokumente zu einem Bluttest: Auswahlmenü (Dateiname, Größe, Zeitstempel) plus "Ordner öffnen".' }
            @{ Typ = 'Neu'; Text = '"Dokument hochladen" mit Mehrfachauswahl; der Tooltip am Button listet alle Dokumente des Bluttests.' }
            @{ Typ = 'Bugfix'; Text = 'Der Zähler am Button bleibt nach jedem Upload korrekt.' }
            @{ Typ = 'Hinweis'; Text = 'Versionsangaben im Script vereinheitlicht (bisher 2.19.0 bzw. 2.23.0).' }
        ) }
    [ordered]@{ Version = '2.23.0'; Datum = ''; Titel = 'Alias-Suche im Feld "Blutmarker"'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Einzelmarker-Analyse" > "Blutmarker": freie Eingabe mit Labor-Synonymen, z. B. "WBC" oder "Leukozyten", "GOT", "AST" oder "ASAT".' }
            @{ Typ = 'Neu'; Text = 'Live-Filter beim Tippen (exakter Treffer vor Wortanfang vor enthalten); Enter übernimmt, Esc setzt zurück, eine unbekannte Eingabe färbt das Feld rot.' }
            @{ Typ = 'Neu'; Text = 'Eine Info-Zeile neben dem Wert-Feld zeigt Einheit und Referenzbereich des erkannten Markers.' }
            @{ Typ = 'Bugfix'; Text = '"Hinzufügen" konnte einen Wert unter einem zuvor gewählten Marker speichern, obwohl ein anderer im Feld stand.' }
            @{ Typ = 'Hinweis'; Text = 'BUN, PCT und TG sind bewusst nicht als Kürzel hinterlegt (mehrdeutig).' }
        ) }
    [ordered]@{ Version = '2.22.1'; Datum = ''; Titel = 'Feinschliff Dokument-Button'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Button "Dokument anzeigen" quadratisch (60 × 60), Text zweizeilig und zentriert.' }
            @{ Typ = 'Bugfix'; Text = 'Die Beschriftung "letzter Wert" im Diagramm verursachte einen Fehler.' }
        ) }
    [ordered]@{ Version = '2.22.0'; Datum = ''; Titel = 'Dokument-Anzeige und letzter Wert im Diagramm'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Daten nach Bluttests": Button "Dokument anzeigen" öffnet das archivierte Dokument des gewählten Tests mit der Standard-App.' }
            @{ Typ = 'Neu'; Text = 'Reiter "Einzelmarker-Analyse": Das Diagramm zeigt den letzten Wert samt Datum (Untertitel und Beschriftung am letzten Punkt).' }
        ) }
    [ordered]@{ Version = '2.21.2'; Datum = ''; Titel = 'Export unter Windows PowerShell 5.1'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Unverschlüsselter Export und .btbackup-Export brachen unter Windows PowerShell 5.1 ab.' }
        ) }
    [ordered]@{ Version = '2.21.1'; Datum = ''; Titel = 'Unverschlüsselter Migrations-Export'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Daten nach Bluttests" > "Unverschlüsselt exportieren (Migration)": Klartext-ZIP für den Umzug auf einen anderen PC oder Windows-Benutzer; dort wird beim nächsten Speichern neu verschlüsselt.' }
            @{ Typ = 'Hinweis'; Text = 'Der Klartext-Export enthält sensible Gesundheitsdaten – nach der Migration löschen.' }
        ) }
    [ordered]@{ Version = '2.19.0'; Datum = ''; Titel = 'Verschlüsselung und portables Backup'
        Eintraege = @(
            @{ Typ = 'Sicherheit'; Text = 'Alle Daten (Einstellungen und Tagesdaten) werden mit DPAPI an das Windows-Benutzerprofil gebunden verschlüsselt; vorhandene Klartextdateien werden beim nächsten Speichern umgestellt.' }
            @{ Typ = 'Neu'; Text = '"Verschlüsseltes Backup exportieren" und "Backup importieren": portables .btbackup-Archiv (AES-256, eigenes Passwort) für einen PC- oder Profilwechsel.' }
            @{ Typ = 'Neu'; Text = 'Das automatische Backup kann zusätzlich ein .btbackup erstellen.' }
            @{ Typ = 'Neu'; Text = '"PDF importieren (Beta)": Blutbild aus einem PDF einlesen, mit Prüfdialog (Konfidenz, Werte bearbeitbar, Datum wählbar); das PDF wird mit archiviert.' }
        ) }
    [ordered]@{ Version = '2.17.0'; Datum = ''; Titel = 'Dokument-Upload und genetische Vorbelastungen'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Daten nach Bluttests" > "Dokument hochladen": Die Datei wird dem gewählten Bluttest zugeordnet und archiviert.' }
            @{ Typ = 'Neu'; Text = 'Reiter "Persönliche Metriken": genetische Vorbelastungen (4 vordefiniert, eigene definierbar); betroffene Marker erhalten Hinweise im Risiko-Cockpit.' }
            @{ Typ = 'Neu'; Text = 'Custom Report für eine Vorbelastung (nur die relevanten Marker).' }
            @{ Typ = 'Bugfix'; Text = 'Das Laden der Einstellungen brach bei leeren Feldern ab.' }
        ) }
    [ordered]@{ Version = '2.15.3'; Datum = ''; Titel = 'Layout Wochentage'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Automatisches Backup: Wochentage in zwei Zeilen, "Donnerstag" einzeilig, Hinweistext klarer abgesetzt.' }
        ) }
    [ordered]@{ Version = '2.15.2'; Datum = ''; Titel = 'Layout automatisches Backup'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = 'Automatisches Backup: Aktions-Schaltflächen bündig in einer Linie, Hinweistexte nicht mehr abgeschnitten, keine Überlappungen.' }
        ) }
    [ordered]@{ Version = '2.15.1'; Datum = ''; Titel = 'Feinschliff Schaltflächen'
        Eintraege = @(
            @{ Typ = 'Anpassung'; Text = '"Beenden" überlappt den Reiterrahmen nicht mehr.' }
            @{ Typ = 'Anpassung'; Text = '"globale Einstellungen" steht neben "Beenden" und ist aus jedem Reiter erreichbar (bisher nur im ersten Reiter).' }
        ) }
    [ordered]@{ Version = '2.15.0'; Datum = ''; Titel = '"Beenden" überall sichtbar und Backup-Bereinigung'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = '"Beenden" war nur im ersten Reiter sichtbar – jetzt in jedem Reiter.' }
            @{ Typ = 'Neu'; Text = 'Automatisches Backup: "Ältere Backups löschen" mit Auswahl, wie viele behalten werden (1, 2, 3 oder nur das neue).' }
            @{ Typ = 'Sicherheit'; Text = 'Bereinigung erst nach erfolgreichem Backup und nur für Dateien mit dem Präfix "BloodTracker_AutoBackup_"; ZIP- und Ordner-Backups werden getrennt gezählt.' }
        ) }
    [ordered]@{ Version = '2.14.3'; Datum = ''; Titel = 'Sichtbarkeit "Beenden"'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = '"Beenden" war nach Version 2.14.2 praktisch unsichtbar (von anderen Schaltflächen überdeckt).' }
        ) }
    [ordered]@{ Version = '2.14.2'; Datum = ''; Titel = 'Einstellungen des automatischen Backups'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Die Einstellungen des automatischen Backups gingen nach einem Neustart verloren (wurden nicht gespeichert).' }
            @{ Typ = 'Neu'; Text = 'Intervall "monatlich" mit Auswahl des Monatstags (1. bis 5. Tag, Monatsmitte, Monatsletzter).' }
            @{ Typ = 'Anpassung'; Text = '"Beenden" auf die Höhe der Reiter verschoben.' }
        ) }
    [ordered]@{ Version = '2.14.1'; Datum = ''; Titel = 'Wochentage für das automatische Backup'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Intervall "wöchentlich" mit Auswahl der Wochentage (Standard Montag); kein doppeltes Backup am selben Tag.' }
            @{ Typ = 'Hinweis'; Text = 'Es wird keine Windows-Aufgabe angelegt: Backups laufen nur bei geöffneter App (beim Start und alle 15 Minuten) und werden beim nächsten Start nachgeholt.' }
        ) }
    [ordered]@{ Version = '2.14.0'; Datum = ''; Titel = 'Automatisches Backup'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Automatisches Backup (in den Einstellungen): Zielordner, Format (ZIP, Ordner oder beides), Intervall und Uhrzeit; inklusive Sicherung des Registry-Eintrags zum Datenpfad.' }
            @{ Typ = 'Sicherheit'; Text = 'Ein Zielordner unterhalb des Datenordners wird abgelehnt.' }
        ) }
    [ordered]@{ Version = '2.13.1'; Datum = ''; Titel = 'Komfort nach dem Export'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Nach jedem Export wahlweise Ordner öffnen und Datei öffnen (Schalter oben im Script, Standard: beides an).' }
        ) }
    [ordered]@{ Version = '2.13.0'; Datum = ''; Titel = 'PDF-Export und Grafik-Reports'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Der PDF-Export brach bei kleinen Datensätzen ab.' }
            @{ Typ = 'Neu'; Text = 'Custom Report (PDF und Druck): je Marker Überschrift, Verlaufsgrafik und Wertetabelle (Grafik ab 2 Werten).' }
            @{ Typ = 'Anpassung'; Text = 'Datumsformat in der Report-Vorschau: TT.MM.JJJJ.' }
        ) }
    [ordered]@{ Version = '2.12.0'; Datum = ''; Titel = 'Custom Report und PDF-Export'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Reiter "Custom Report": Marker und Zeitraum frei wählen, drucken oder exportieren.' }
            @{ Typ = 'Neu'; Text = 'PDF als dritte Export-Option in allen Reitern (über "Microsoft Print to PDF" oder einen anderen PDF-Drucker); ohne PDF-Drucker erscheint eine klare Meldung.' }
        ) }
    [ordered]@{ Version = '2.11.0'; Datum = ''; Titel = 'Export und neuer Marker HIV'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Schaltfläche "Exportieren" (JSON, CSV) in Risiko-Cockpit, Einzelmarker-Analyse, Korrelationen und Daten nach Bluttests.' }
            @{ Typ = 'Neu'; Text = 'Marker "HIV (Anti-HIV-1/2)" in der Gruppe "Infektionskrankheiten".' }
        ) }
    [ordered]@{ Version = '2.10.0'; Datum = ''; Titel = 'Longevity Science Update'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'PhenoAge (biologisches Alter nach Levine 2018) und PhenoAge-Accel (biologisches minus chronologisches Alter).' }
            @{ Typ = 'Neu'; Text = 'InflammAging-Score, NLR, TyG-Index, Triglyceride/HDL-Ratio und HbA1c-Trend pro Jahr.' }
            @{ Typ = 'Neu'; Text = 'Neuer Longevity-Reiter mit Aufschlüsselung, Interpretationshinweisen und Literatur; Korrelationen um etablierte Paarungen erweitert.' }
            @{ Typ = 'Hinweis'; Text = 'Alle Scores sind Indikatoren, keine Diagnose.' }
        ) }
    [ordered]@{ Version = '2.9.2'; Datum = ''; Titel = 'Sicherung wiederherstellen'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Einstellungen > "Sicherung wiederherstellen (ZIP)...": Deinstallations-Backup (Daten und Registry) wiederherstellen – Daten ersetzen oder ergänzen, danach Neustart.' }
        ) }
    [ordered]@{ Version = '2.8'; Datum = ''; Titel = 'Longevity-Reiter'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Absturz (IndexOutOfRangeException) im Longevity-Reiter behoben.' }
        ) }
    [ordered]@{ Version = '2.7.4'; Datum = ''; Titel = 'Schaltfläche "Einstellungen"'
        Eintraege = @(
            @{ Typ = 'Neu'; Text = 'Schaltfläche "Einstellungen" in "Risiko-Cockpit" und "Daten nach Bluttests" (zunächst nur mit "Schließen").' }
        ) }
    [ordered]@{ Version = '2.7.3'; Datum = ''; Titel = 'Aktualisierung beim Reiterwechsel'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Änderungen in "Daten nach Bluttests" (Löschen, Bearbeiten) erscheinen sofort in den anderen Reitern, nicht erst nach einem Neustart.' }
            @{ Typ = 'Bugfix'; Text = 'Gelöschte Tage und Einträge kamen nach einem Neustart zurück – sie werden jetzt auch auf der Festplatte entfernt.' }
        ) }
    [ordered]@{ Version = '2.7.2'; Datum = ''; Titel = 'Löschen von Einträgen'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Fehler "Die Liste hatte eine feste Größe" beim Löschen, Importieren und Hinzufügen behoben.' }
        ) }
    [ordered]@{ Version = '2.7'; Datum = '2026-02-04'; Titel = 'Laden vorhandener Tagesdaten'
        Eintraege = @(
            @{ Typ = 'Bugfix'; Text = 'Fehler beim Laden vorhandener Tagesdaten behoben.' }
            @{ Typ = 'Bugfix'; Text = 'Duplikatsprüfung und Filter-Auswahl korrigiert.' }
        ) }
)

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
    # v2.28.0: 'Kreatinin-Clearance' entfernt - gemessene Clearance (ml/min) ist nicht die eGFR (ml/min/1,73 m²)
    'Glomeruläre Filtrationsrate (GFR)'         = @('GFR', 'eGFR', 'GFR CKD-EPI', 'CKD-EPI', 'glomeruläre Filtrationsrate')
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
    # v2.28.0: 'BNP' entfernt - BNP und NT-proBNP sind verschiedene Laborwerte mit eigenen Grenzwerten
    'NT-proBNP'                                 = @('NTproBNP', 'NT pro BNP', 'NT-pro-BNP', 'proBNP')
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
    # v2.28.0: 'Troponin' entfernt - hätte auch Troponin I (TnI, andere Grenzwerte) zugeordnet
    'TnT (Troponin T)'                          = @('TnT', 'Troponin T', 'hsTnT', 'hs-TnT', 'Trop T', 'Troponin T hs')
    'Triglyceride (Trig)'                       = @('TRIG', 'Trig', 'Triglyceride', 'Triglyzeride', 'Neutralfette')
    'TSH-basal (TSH)'                           = @('TSH', 'TSH basal', 'TSH-basal', 'Thyreotropin', 'TSH (basal)')
    'Vitamin B12'                               = @('B12', 'Vit B12', 'Vit. B12', 'Cobalamin', 'Vitamin-B12', 'Cyanocobalamin')
    'Vitamin D (25-OH)'                         = @('Vitamin D', 'Vit D', 'Vit. D', '25-OH-D', '25-OH Vitamin D', '25(OH)D', '25-Hydroxy-Vitamin D', 'Calcidiol', 'Vitamin D3', 'Vitamin-D')
    'Workout-Frequenz'                          = @('Workout', 'Training', 'Trainingsfrequenz', 'Sport', 'Trainingseinheiten')
}

# v2.28.0: Suchbegriffe, die NIE automatisch einem Marker zugeordnet werden duerfen.
# Die Teilwort-Suche (Get-MarkerMatches) haette sonst "BNP" -> NT-proBNP und
# "Troponin" -> TnT aufgeloest, obwohl es andere Laborwerte mit eigenen Grenzen sind.
# (Normalisierte Tokens wie Get-MarkerToken: klein, ohne Sonderzeichen.)
$script:MarkerSearchBlocklist = @('bnp', 'troponin', 'troponini', 'tni', 'hstni', 'kreatininclearance', 'creatininclearance')

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

$script:MarkerSetVersion = '2.30.0'

$script:MarkerSetAdditions = @(
    @{ Version = '2.25.0'; Markers = @('Hochsensitives CRP (hs-CRP)', 'Sexualhormon-bindendes Globulin (SHBG)'); CalculatedMarkers = @() },
    @{ Version = '2.26.0'; Markers = @('Apolipoprotein E-Genotyp (APOE)', 'Zink', 'Magnesium'); CalculatedMarkers = @() },
    @{ Version = '2.27.0'; Markers = @(); CalculatedMarkers = @('PREVENT-CVD-10Y') }
)

# ---------------------------------------------------------------------
# v2.27.0: KORREKTUR-MIGRATION fuer bereits vorhandene Marker-Definitionen.
# Merge-NewDefaultMarkers ergaenzt nur NEUE Marker. Wird die Definition eines
# bestehenden Markers fachlich korrigiert, muss sie in vorhandenen Config.json-
# Dateien einmalig ueberschrieben werden - genau die hier genannten Felder,
# nur wenn die gespeicherte MarkerSetVersion aelter ist. Wiederverwendbar fuer
# kuenftige Korrekturen (z. B. TyG, AIP, PLR).
# ---------------------------------------------------------------------
# v2.28.0: Optional "OnlyIf" - Korrektur nur, wenn der Marker in der Config.json noch
# den alten Standardwert hat (eigene Anpassungen des Nutzers bleiben unangetastet).
# "Markers" = normale Marker, "CalculatedMarkers" = berechnete Marker.
$script:MarkerSetCorrections = @(
    @{ Version = '2.27.0'; CalculatedMarkers = @('PREVENT-ASCVD-10Y'); Fields = @('RequiredMarkers', 'Formula', 'Unit', 'Group', 'RefMin', 'RefMax', 'OptimalMin', 'OptimalMax', 'Description') },
    @{ Version = '2.28.0'; CalculatedMarkers = @('TyG-Index'); Fields = @('Formula', 'RefMin', 'RefMax', 'OptimalMin', 'OptimalMax', 'Description'); OnlyIf = @{ Formula = '[Math]::Log(((%0%) * (%1%)) / 2)' } },
    @{ Version = '2.28.0'; CalculatedMarkers = @('Atherogener Index (AIP)'); Fields = @('Formula', 'RefMin', 'RefMax', 'OptimalMin', 'OptimalMax', 'Description'); OnlyIf = @{ Formula = '[Math]::Log10(%0% / %1%)' } },
    @{ Version = '2.28.0'; CalculatedMarkers = @('Thrombozyten/Lymphozyten-Ratio (PLR)'); Fields = @('RequiredMarkers', 'Formula', 'Description'); OnlyIf = @{ Formula = '%0% / %1%' } },
    @{ Version = '2.28.0'; CalculatedMarkers = @('InflammAging-Score'); Fields = @('RequiredMarkers'); OnlyIf = @{ Formula = '[INFLAMMAGING]' } },
    @{ Version = '2.28.0'; Markers = @('Hämatokrit (HCT/Hk)'); Fields = @('RefMin', 'RefMax', 'Description'); OnlyIf = @{ RefMin = 0.38; RefMax = 0.49 } },
    @{ Version = '2.28.0'; Markers = @('LDL-Cholesterin (LDL)', 'Triglyceride (Trig)'); Fields = @('OptimalMin'); OnlyIf = @{ OptimalMin = 50 } },
    @{ Version = '2.28.0'; CalculatedMarkers = @('BMI'); Fields = @('Formula'); OnlyIf = @{ Formula = '%0% / ([Math]::Pow(.Groesse / 100, 2))' } },
    # v2.29.0: Alter je Testtag aus dem Geburtsdatum - Hinweis "festes Alter" in der Beschreibung ersetzt
    # v2.30.0: Longevity-Score an den eigenen Bereichen der Marker - Beschreibung aktualisiert
    @{ Version = '2.30.0'; CalculatedMarkers = @('Longevity-Score'); Fields = @('Description'); OnlyIf = @{ Description = 'Ein zusammenfassender Score, der wichtige Biomarker für ein gesundes Altern bewertet.' } },
    @{ Version = '2.29.0'; CalculatedMarkers = @('PREVENT-ASCVD-10Y'); Fields = @('Description'); OnlyIf = @{ Description = "AHA-PREVENT (Khan et al., Circulation 2024): 10-Jahres-Risiko für atherosklerotische Herz-Kreislauf-Ereignisse (Herzinfarkt, Schlaganfall). Kategorien: < 3 % niedrig, 3 bis < 5 % grenzwertig, 5 bis < 10 % mittel, ab 10 % hoch. Eingaben: Alter, Geschlecht, Gesamt- und HDL-Cholesterin (Testtag), systolischer Blutdruck (Testtag, sonst Standardwert aus 'Persönliche Metriken'), eGFR (GFR-Marker, sonst aus Kreatinin nach CKD-EPI 2021), Diabetes (Checkbox oder HbA1c >= 6,5 %), Raucher, Blutdrucksenker, Lipidsenker. Gültig für 30-79 Jahre, Gesamtcholesterin 130-320 mg/dl, HDL 20-100 mg/dl, Blutdruck 90-180 mmHg, eGFR 15-140 - außerhalb wird kein Wert berechnet. Für Personen ohne bekannte Herz-Kreislauf-Erkrankung. Hinweis: Das Alter ist ein fester Wert - ältere Testdaten werden mit dem aktuellen Alter berechnet." } }
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
# v2.31.0: Vorrang-Marker fuer die Score-Formeln ([PHENOAGE], [PHENOAGE_ACCEL],
# [INFLAMMAGING], [LONGEVITY_SCORE] ...). Liegt am Testtag hs-CRP vor, wird es
# statt des klassischen CRP verwendet (genauer im niedrigen Bereich < 3 mg/l).
# Gleiche Einheit (mg/l). Eigene Formeln nutzen weiter zuerst den genannten
# Marker und erst danach $script:MarkerFallbacks.
# ---------------------------------------------------------------------
$script:MarkerPreferred = @{
    'C-reaktives Protein (CRP)' = @('Hochsensitives CRP (hs-CRP)')
}

# v2.30.0: Longevity-Score - Mindestanzahl bewerteter Werte am selben Testtag (von 7)
$script:LongevityScoreMinComponents = 4

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

        # v2.27.0: Korrekturen bestehender Definitionen (nur genannte Felder, nur einmalig)
        # v2.28.0: + normale Marker ("Markers") + Bedingung "OnlyIf" (nur alter Standardwert wird korrigiert)
        foreach ($correction in @($script:MarkerSetCorrections)) {
            if (-not $correction) { continue }
            $corrVer = $null
            if (-not [version]::TryParse($correction.Version, [ref]$corrVer)) { continue }
            if ($savedVer -and $corrVer -le $savedVer) { continue }
            $targets = @()
            foreach ($calcName in @($correction.CalculatedMarkers)) { if ($calcName) { $targets += @{ Name = $calcName; Defaults = $DefaultCalculatedMarkers; List = 'CalculatedMarkers' } } }
            foreach ($markerName in @($correction.Markers)) { if ($markerName) { $targets += @{ Name = $markerName; Defaults = $DefaultMarkers; List = 'Markers' } } }
            foreach ($target in $targets) {
                $default = $target.Defaults | Where-Object { $_.Name -eq $target.Name } | Select-Object -First 1
                if (-not $default) { Write-Warning "Korrektur: Default-Marker '$($target.Name)' nicht gefunden."; continue }
                foreach ($existing in @($ConfigData[$target.List] | Where-Object { $_.Name -eq $target.Name })) {
                    # Nur korrigieren, wenn der gespeicherte Stand noch dem alten Standard entspricht
                    if ($correction.OnlyIf) {
                        $matchesOld = $true
                        foreach ($condKey in $correction.OnlyIf.Keys) {
                            if ($existing.PSObject.Properties.Name -notcontains $condKey -or $existing.$condKey -ne $correction.OnlyIf[$condKey]) { $matchesOld = $false }
                        }
                        if (-not $matchesOld) { Write-Verbose "Korrektur '$($target.Name)' übersprungen (vom Nutzer angepasst)."; continue }
                    }
                    foreach ($field in @($correction.Fields)) {
                        if ($default.PSObject.Properties.Name -notcontains $field) { continue }
                        $existing | Add-Member -NotePropertyName $field -NotePropertyValue $default.$field -Force
                    }
                    $changed = $true
                }
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
    if ($script:MarkerSearchBlocklist -contains $q) { return @() }   # v2.28.0: mehrdeutige Laborbegriffe
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
    if ($script:MarkerSearchBlocklist -contains $q) { return $null }   # v2.28.0: mehrdeutige Laborbegriffe
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

function New-BtBloodValueItem {
    <#
    .SYNOPSIS
        v2.46.0: Gemeinsame Pruefung/Umrechnung fuer Einzel- und kuenftige Formulareingabe.
    .DESCRIPTION
        Liefert genau ein Item {Date; Name; Value; Unit; Note}, schreibt keine Datei.
        MarkerQuery verwendet die bestehende Alias-Aufloesung nur fuer gemessene Marker.
        HIV/APOE benoetigen eine ausdrueckliche Auswahl; keine medizinische Vorbelegung.
        Inaktive Sonderfelder werden ignoriert. Referenzabweichungen blockieren nicht.
        ExistingItems und PendingItems pruefen Datum + aufgeloesten Markernamen;
        PendingItems sind bereits vorbereitete Items des kuenftigen Formulars.
        Datum im bestehenden ISO-Format; keine neue Beschraenkung des Eingabedatums.
        Vorhandene UI-Vorbelegungen und Speicherung bleiben beim Aufrufer.
    #>
    param(
        [Parameter(Mandatory)][string]$DateString, [string]$MarkerQuery = '',
        $ValueText = $null, [string]$HivResult = '', [string]$ApoeGenotype = '',
        [string]$TestosteroneUnit = 'ng/dl', [string]$Note = '',
        [object[]]$ExistingItems = $script:allHistoricalItems, [object[]]$PendingItems = @()
    )
    $parsedDate = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($DateString, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) { throw 'Bitte ein gültiges Datum im Format JJJJ-MM-TT angeben.' }
    $query = $MarkerQuery.Trim()
    if (-not $query) { throw 'Bitte wählen Sie einen Blutmarker aus.' }
    $markerName = Resolve-MarkerName -Query $query
    if ([string]::IsNullOrWhiteSpace($markerName)) {
        $candidates = @(Get-MarkerMatches -Query $query)
        if ($candidates.Count -gt 1) { throw "Die Eingabe '$query' ist nicht eindeutig ($($candidates.Count) Treffer). Bitte den gewünschten Marker aus der Liste wählen." }
        throw "Der Blutmarker '$query' ist unbekannt oder nicht eindeutig zuordenbar. Bitte einen gemessenen Marker aus der Liste wählen."
    }
    $configs = @($script:data.Config.Markers | Where-Object { $_.Name -eq $markerName })
    if ($configs.Count -ne 1) { throw "Die Konfiguration für '$markerName' fehlt oder ist nicht eindeutig. Bitte die Marker-Verwaltung prüfen." }
    $markerConfig = $configs[0]
    $value = $null; $extraNote = ''
    if ($markerName -eq 'HIV (Anti-HIV-1/2)') {
        if ($HivResult -notin @('nicht reaktiv', 'reaktiv')) { throw 'Bitte für HIV ausdrücklich nicht reaktiv oder reaktiv auswählen.' }
        $value = if ($HivResult -eq 'reaktiv') { 1 } else { 0 }
        $extraNote = if ($value -eq 1) { 'reaktiv' } else { 'nicht reaktiv' }
    } elseif ($markerName -eq $script:ApoeMarkerName) {
        $index = [array]::IndexOf([string[]]$script:ApoeGenotypeOptions, $ApoeGenotype)
        if ($index -lt 0) { throw 'Bitte einen APOE-Genotyp aus der Liste wählen.' }
        $value = $index + 1
        $extraNote = 'APOE-Genotyp: ' + $script:ApoeGenotypeOptions[$index]
    } else {
        if ($ValueText -is [bool]) { throw 'Der Wert muss eine Zahl sein.' }
        $value = Parse-Number ([string]$ValueText)
        if ($null -eq $value) { throw 'Der Wert darf nicht leer sein.' }
        if ([double]::IsNaN($value) -or [double]::IsInfinity($value)) { throw 'Der Wert muss eine endliche Zahl sein.' }
        if ($markerName -eq 'Testosteron, gesamt') {
            if ($TestosteroneUnit -notin @('ng/dl', 'nmol/l')) { throw 'Bitte für Testosteron ng/dl oder nmol/l auswählen.' }
            if ($TestosteroneUnit -eq 'nmol/l') {
                $converted = $value * 28.818
                if ([double]::IsInfinity($converted) -or [double]::IsNaN($converted)) { throw 'Der Testosteronwert ist für die Umrechnung zu groß.' }
                $value = [Math]::Round($converted, 1, [MidpointRounding]::AwayFromZero)
                $extraNote = "Eingabe: $ValueText nmol/l"
            }
        }
    }
    if (@($ExistingItems | Where-Object { $_.Date -eq $DateString -and $_.Name -eq $markerName }).Count -gt 0) {
        throw "Für '$markerName' existiert am $DateString bereits ein Eintrag. Bitte in Tab 'Daten nach Bluttests' oder über 'bearbeiten' ändern."
    }
    if (@($PendingItems | Where-Object { $_.Date -eq $DateString -and $_.Name -eq $markerName }).Count -gt 0) { throw "Der Marker '$markerName' ist für dieses Datum bereits in der Eingabeliste enthalten." }
    $finalNote = $Note
    if ($extraNote) { $finalNote = if ($finalNote) { "$finalNote | $extraNote" } else { $extraNote } }
    return [PSCustomObject]@{ Date = $DateString; Name = $markerName; Value = $value; Unit = $markerConfig.Unit; Note = $finalNote }
}


function Get-ListItemTokens {
    <#
    .SYNOPSIS
        v2.32.0: Such-Tokens eines Listeneintrags (wie Build-MarkerSearchIndex):
        voller Name, Name ohne Klammerzusatz, Kürzel aus Klammern, Aliase.
    #>
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return @() }
    $candidates = @($Name)
    $baseName = ($Name -replace '\s*\([^)]*\)\s*', '').Trim()
    if ($baseName) { $candidates += $baseName }
    if ($Name -match '\(([^)]+)\)') {
        foreach ($part in ($Matches[1] -split '[/,]')) { $trimmedPart = $part.Trim(); if ($trimmedPart) { $candidates += $trimmedPart } }
    }
    if ($script:MarkerAliases -and $script:MarkerAliases.ContainsKey($Name)) {
        foreach ($alias in $script:MarkerAliases[$Name]) { if ($alias) { $candidates += [string]$alias } }
    }
    return @($candidates | ForEach-Object { Get-MarkerToken $_ } | Where-Object { $_ } | Select-Object -Unique)
}

function Get-ListMatches {
    <#
    .SYNOPSIS
        v2.32.0: Suche wie Get-MarkerMatches, aber nur in einer gegebenen Liste
        (z. B. "Verlauf anzeigen für" inkl. berechneter Marker).
    .DESCRIPTION
        Rang 0 = exakter Token-Treffer, 1 = Token beginnt mit der Eingabe,
        2 = Token enthält die Eingabe. Kürzel mit <= 3 Zeichen nur exakt.
    #>
    param([string]$Query, [string[]]$Items)
    $queryToken = Get-MarkerToken $Query
    if ([string]::IsNullOrEmpty($queryToken) -or -not $Items) { return @() }
    if ($script:MarkerSearchBlocklist -contains $queryToken) { return @() }
    $rank0 = @(); $rank1 = @(); $rank2 = @()
    foreach ($itemName in $Items) {
        if ([string]::IsNullOrWhiteSpace($itemName)) { continue }
        $best = 99
        foreach ($token in (Get-ListItemTokens -Name $itemName)) {
            if ($token -eq $queryToken) { $best = 0; break }
            if ($token.Length -le 3) { continue }
            if ($token.StartsWith($queryToken)) { if ($best -gt 1) { $best = 1 }; continue }
            if ($token.Contains($queryToken)) { if ($best -gt 2) { $best = 2 } }
        }
        switch ($best) {
            0 { $rank0 += $itemName }
            1 { $rank1 += $itemName }
            2 { $rank2 += $itemName }
        }
    }
    return @($rank0) + @($rank1) + @($rank2)
}

function Resolve-ListSelection {
    <#
    .SYNOPSIS
        v2.32.0: Löst eine Eingabe auf genau EINEN Eintrag der Liste auf.
    .OUTPUTS
        Listeneintrag oder $null (kein Treffer / mehrdeutig).
    #>
    param([string]$Query, [string[]]$Items)
    if ([string]::IsNullOrWhiteSpace($Query) -or -not $Items) { return $null }
    $trimmedQuery = $Query.Trim()
    foreach ($itemName in $Items) { if ($itemName -ieq $trimmedQuery) { return [string]$itemName } }
    $queryToken = Get-MarkerToken $trimmedQuery
    if (-not $queryToken) { return $null }
    if ($script:MarkerSearchBlocklist -contains $queryToken) { return $null }
    $exact = @($Items | Where-Object { $_ -and ((Get-ListItemTokens -Name $_) -contains $queryToken) })
    if ($exact.Count -eq 1) { return [string]$exact[0] }
    if ($exact.Count -gt 1) { return $null }
    $hits = @(Get-ListMatches -Query $trimmedQuery -Items $Items)
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

        v2.26.1: Atomares Schreiben. Der Inhalt wird zuerst vollständig in
        "<Datei>.bttmp" geschrieben und ersetzt erst danach die Zieldatei
        ([System.IO.File]::Replace, Vorversion kurzzeitig als "<Datei>.btbak").
        Ein Absturz oder Fehler während des Schreibens hinterlässt damit nie
        eine halb geschriebene Datei - die bisherige Version bleibt erhalten.
        Dateiformat und Encoding sind unverändert.
        Fehler werden als (abfangbarer) Fehler an den Aufrufer gemeldet; ohne
        try/catch beim Aufrufer läuft dieser - wie bisher - weiter.
    #>
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$JsonString,
        [switch]$RequireAtomic
    )
    $fullPath = $null; $tmpPath = $null; $bakPath = $null
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($JsonString)
        $encrypted = [System.Security.Cryptography.ProtectedData]::Protect(
            $bytes, $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser
        )
        $base64 = [Convert]::ToBase64String($encrypted)

        # v2.26.1: Absoluter Pfad - die .NET-Dateifunktionen kennen das PS-Arbeitsverzeichnis nicht
        $fullPath = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        $tmpPath  = "$fullPath.bttmp"
        $bakPath  = "$fullPath.btbak"

        # Schritt 1: vollständig in die Temp-Datei schreiben
        # (Marker-Prefix, damit Read-ProtectedJsonFile erkennen kann - Format unverändert)
        Set-Content -LiteralPath $tmpPath -Value "BTENC:$base64" -Encoding UTF8 -Force -ErrorAction Stop

        # Schritt 2a: Neue Datei -> einfach umbenennen
        if (-not [System.IO.File]::Exists($fullPath)) {
            [System.IO.File]::Move($tmpPath, $fullPath)
            return
        }

        # Schritt 2b: Bestehende Datei atomar ersetzen (3 Versuche, z. B. kurz gesperrt durch Virenscanner/Sync)
        # Veraltete Vorversion eines früheren Abbruchs entfernen (Zieldatei ist intakt)
        if ([System.IO.File]::Exists($bakPath)) {
            Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue
        }
        $replaced = $false
        for ($attempt = 1; ($attempt -le 3) -and (-not $replaced); $attempt++) {
            try {
                [System.IO.File]::Replace($tmpPath, $fullPath, $bakPath, $true)
                $replaced = $true
            } catch {
                if ($attempt -lt 3) { Start-Sleep -Milliseconds 150 }
            }
        }

        if (-not $replaced) {
            # v2.47.0: Gemeinsame Blutwerte duerfen nicht teilweise ueberschrieben werden.
            if ($RequireAtomic) { throw 'Die Tagesdatei konnte nicht atomar ersetzt werden. Die bisherige Datei bleibt erhalten; bitte erneut versuchen.' }
            # Fallback (z. B. Dateisystem ohne ReplaceFile-Unterstützung): überschreibend kopieren
            try {
                [System.IO.File]::Copy($tmpPath, $fullPath, $true)
                Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
            } catch {
                # Letzte Absicherung: Vorversion zurückholen, falls die Zieldatei fehlt
                if ((-not [System.IO.File]::Exists($fullPath)) -and [System.IO.File]::Exists($bakPath)) {
                    try { [System.IO.File]::Move($bakPath, $fullPath) } catch { Write-Warning "Vorversion konnte nicht zurückgeholt werden (liegt als '$bakPath' vor): $($_.Exception.Message)" }
                }
                throw
            }
        }

        # Schritt 3: Vorversion nach erfolgreichem Ersetzen entfernen
        if ([System.IO.File]::Exists($bakPath)) {
            Remove-Item -LiteralPath $bakPath -Force -ErrorAction SilentlyContinue
        }
    } catch {
        # Temp-Datei eines fehlgeschlagenen Schreibvorgangs aufräumen - außer in dem
        # Sonderfall "Zieldatei fehlt, Vorversion liegt als .btbak vor" (manuelle Rettung)
        if ($tmpPath -and [System.IO.File]::Exists($tmpPath) -and
            ([System.IO.File]::Exists($fullPath) -or -not [System.IO.File]::Exists($bakPath))) {
            Remove-Item -LiteralPath $tmpPath -Force -ErrorAction SilentlyContinue
        }
        $PSCmdlet.ThrowTerminatingError($_)
    }
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

        v2.26.1 BUGFIX (Datenverlust): Bisher wurde der Datenordner gelöscht, BEVOR
        die importierten Dateien geschrieben waren - ein Fehler mittendrin hinterließ
        einen leeren oder halben Datenbestand. Neu:
          a) Backup-Inhalt vorab prüfen (gültiges JSON) - vor jeder Änderung
          b) Tagesdaten vollständig in einem Staging-Ordner neben "data" aufbauen
          c) Austausch per Umbenennen (aktueller Ordner -> Sicherung, Staging -> "data")
          d) Config.json atomar schreiben; schlägt das fehl, wird der alte
             Datenordner zurückgeholt (Config und Daten bleiben konsistent)
          e) Erst nach vollständigem Erfolg wird der alte Stand entfernt
        Ist der Datenordner gesperrt (z. B. geöffnetes Dokument), bricht der Import
        ab, OHNE bestehende Daten zu verändern.
    #>
    param(
        [Parameter(Mandatory)][string]$InputPath,
        [Parameter(Mandatory)][System.Security.SecureString]$Passphrase
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $tempDir = Join-Path -Path $env:TEMP -ChildPath "BT_import_$timestamp"
    $tempZip = Join-Path -Path $env:TEMP -ChildPath "BT_import_$timestamp.zip"

    # v2.26.1: Staging-/Sicherungsordner liegen NEBEN dem Datenordner (gleiches Laufwerk
    # -> Austausch per Umbenennen, kein Kopieren)
    $dataBaseFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($script:dailyDataDirBase).TrimEnd('\', '/')
    $stagingDir   = "$dataBaseFull.import_$timestamp"
    $preImportDir = "$dataBaseFull.vorimport_$timestamp"
    $dataSwapped  = $false

    # Ordner umbenennen mit 3 Versuchen (kurzzeitige Sperren durch Virenscanner/Sync)
    $moveDirectory = {
        param([string]$Source, [string]$Destination)
        for ($moveAttempt = 1; $moveAttempt -le 3; $moveAttempt++) {
            try { [System.IO.Directory]::Move($Source, $Destination); return }
            catch { if ($moveAttempt -eq 3) { throw }; Start-Sleep -Milliseconds 200 }
        }
    }

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

        $importedConfig = Join-Path $tempDir "Config.json"
        $importedDataDir = Join-Path $tempDir "data"

        # 3. v2.26.1: Backup-Inhalt prüfen, BEVOR irgendetwas verändert wird
        $jsonFilesToCheck = @()
        if (Test-Path -LiteralPath $importedConfig) { $jsonFilesToCheck += @(Get-Item -LiteralPath $importedConfig) }
        if (Test-Path -LiteralPath $importedDataDir) { $jsonFilesToCheck += @(Get-ChildItem -LiteralPath $importedDataDir -Recurse -File -Filter "*.json" -ErrorAction Stop) }
        foreach ($jsonFile in $jsonFilesToCheck) {
            $checkOk = $false
            try {
                $checkContent = Get-Content -LiteralPath $jsonFile.FullName -Raw -Encoding UTF8
                if (-not [string]::IsNullOrWhiteSpace($checkContent)) {
                    $null = $checkContent | ConvertFrom-Json -ErrorAction Stop
                    $checkOk = $true
                }
            } catch { $checkOk = $false }
            if (-not $checkOk) {
                throw "Das Backup enthält eine ungültige Datei ($($jsonFile.Name)). Import abgebrochen - bestehende Daten wurden NICHT verändert."
            }
        }

        # 4. v2.26.1: Tagesdaten → DPAPI re-encrypt in den Staging-Ordner (bestehende Daten unberührt)
        if (Test-Path -LiteralPath $importedDataDir) {
            New-Item -ItemType Directory -Path $stagingDir -Force | Out-Null
            $importedFiles = Get-ChildItem -LiteralPath $importedDataDir -Recurse -File -ErrorAction Stop
            foreach ($file in $importedFiles) {
                $relativePath = $file.FullName.Substring($importedDataDir.Length).TrimStart('\','/')
                $targetPath = Join-Path $stagingDir $relativePath
                $targetDir = Split-Path $targetPath -Parent
                if (-not (Test-Path -LiteralPath $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }

                if ($file.Extension -eq ".json") {
                    $json = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
                    Write-ProtectedJsonFile -Path $targetPath -JsonString $json
                } else {
                    Copy-Item -LiteralPath $file.FullName -Destination $targetPath -Force -ErrorAction Stop
                }
            }

            # 5. Austausch: aktueller Datenordner -> Sicherung, Staging -> Datenordner
            if (Test-Path -LiteralPath $dataBaseFull) {
                try {
                    & $moveDirectory $dataBaseFull $preImportDir
                } catch {
                    throw "Der Datenordner ist gesperrt (z. B. geöffnetes Dokument, Explorer-Vorschau oder Synchronisierung). Import abgebrochen - bestehende Daten wurden NICHT verändert.`n`nDetails: $($_.Exception.Message)"
                }
            }
            try {
                & $moveDirectory $stagingDir $dataBaseFull
            } catch {
                # Rückweg: alten Datenordner wiederherstellen
                if (Test-Path -LiteralPath $preImportDir) { & $moveDirectory $preImportDir $dataBaseFull }
                throw
            }
            $dataSwapped = $true
        }

        # 6. Config.json → DPAPI re-encrypt (atomar)
        if (Test-Path -LiteralPath $importedConfig) {
            $json = Get-Content -LiteralPath $importedConfig -Raw -Encoding UTF8
            try {
                Write-ProtectedJsonFile -Path $script:dataFile -JsonString $json
            } catch {
                # Rollback der Tagesdaten, damit Config und Daten zusammenpassen
                if ($dataSwapped -and (Test-Path -LiteralPath $preImportDir)) {
                    $failedDir = "$dataBaseFull.fehlimport_$timestamp"
                    & $moveDirectory $dataBaseFull $failedDir
                    & $moveDirectory $preImportDir $dataBaseFull
                    Remove-Item -LiteralPath $failedDir -Recurse -Force -ErrorAction SilentlyContinue
                    $dataSwapped = $false
                }
                throw
            }
        }

        # 7. Erfolg: alten Stand entfernen (Import ersetzt alle Daten - wie bisher)
        if (Test-Path -LiteralPath $preImportDir) {
            Remove-Item -LiteralPath $preImportDir -Recurse -Force -ErrorAction SilentlyContinue
        }

        return $true
    } catch {
        throw $_
    } finally {
        if (Test-Path $tempDir)  { Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue }
        if (Test-Path $tempZip)  { Remove-Item $tempZip -Force -ErrorAction SilentlyContinue }
        # v2.26.1: Nicht übernommenen (unvollständigen) Staging-Ordner entfernen
        if (Test-Path -LiteralPath $stagingDir) { Remove-Item -LiteralPath $stagingDir -Recurse -Force -ErrorAction SilentlyContinue }
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
        # v2.31.2 BUGFIX: direkte Zuweisung. "$value = if (...) { $prop.Value }" gab den Wert über die
        # Pipeline aus - eine Liste mit EINEM Eintrag wurde dadurch zum Einzelwert (RequiredMarkers
        # @('Gewicht') -> 'Gewicht' -> [0] = 'G': BMI nach einem Neustart nie berechnet), leere Listen zu $null.
        if ($prop -is [System.Collections.DictionaryEntry]) { $name = $prop.Key; $value = $prop.Value } else { $name = $prop.Name; $value = $prop.Value }

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

function ConvertTo-BtDashboardViews {
    <#
        v2.39.0: Eigene Ansichten validieren/normalisieren, ohne Datenzugriff oder neue IDs.
        Ausgabe einzeln ueber Pipeline; Aufrufer verwenden @() (auch bei 0/1 Eintraegen).
        Id = GUID im D-Format; Name eindeutig; Markers = geordnete Liste von Namen.
        Unbekannte Marker bleiben erhalten. Leere Markerlisten sind speicherbar.
        Keine Teile einer ungueltigen Liste ausgeben; Fehler nennen keine IDs/Inhalte.
    #>
    param($Views)
    if ($null -eq $Views) { return }
    if ($Views -isnot [System.Collections.IDictionary] -and $Views -isnot [System.Management.Automation.PSCustomObject] -and $Views -isnot [System.Collections.IList]) {
        throw 'DashboardViews muss eine Liste von Ansichtsdefinitionen sein.'
    }
    $normalized = New-Object System.Collections.ArrayList
    $seenIds = @{}; $seenNames = @{}; $position = 0
    foreach ($view in @($Views)) {
        $position++
        if ($view -is [System.Collections.IDictionary]) { $keys = @($view.Keys) }
        elseif ($view -is [System.Management.Automation.PSCustomObject]) { $keys = @($view.PSObject.Properties.Name) }
        else { throw "DashboardViews: Eintrag $position ist keine Ansichtsdefinition." }
        if (@($keys | Where-Object { $_ -notin @('Id', 'Name', 'Markers') }).Count -gt 0 -or $keys -notcontains 'Id' -or $keys -notcontains 'Name' -or $keys -notcontains 'Markers') {
            throw "DashboardViews: Eintrag $position benoetigt genau die Felder Id, Name und Markers."
        }
        $rawId = $view.Id; $rawName = $view.Name; $rawMarkers = $view.Markers
        $parsedId = [guid]::Empty
        if ($rawId -isnot [string] -or -not [guid]::TryParseExact($rawId.Trim(), 'D', [ref]$parsedId) -or $parsedId -eq [guid]::Empty) {
            throw "DashboardViews: Eintrag $position hat keine gueltige Kennung."
        }
        $id = $parsedId.ToString('D')
        if ($seenIds.ContainsKey($id)) { throw "DashboardViews: Eintrag $position hat eine bereits verwendete Kennung." }
        if ($rawName -isnot [string] -or [string]::IsNullOrWhiteSpace($rawName)) { throw "DashboardViews: Eintrag $position hat keinen gueltigen Namen." }
        $name = $rawName.Trim()
        if ($name -match '^(Standard|Vorbelastungen|Aktuelle Auffälligkeiten(?:\s*\(\d+\))?)$') { throw "DashboardViews: Eintrag $position verwendet einen reservierten Namen." }
        if ($seenNames.ContainsKey($name)) { throw "DashboardViews: Eintrag $position hat einen bereits verwendeten Namen." }
        if ($null -ne $rawMarkers -and $rawMarkers -isnot [string] -and $rawMarkers -isnot [System.Collections.IList]) {
            throw "DashboardViews: Eintrag $position hat keine gueltige Markerliste."
        }
        $markers = @(); $seenMarkers = @{}
        if ($null -ne $rawMarkers) {
            foreach ($rawMarker in @($rawMarkers)) {
                if ($rawMarker -isnot [string] -or [string]::IsNullOrWhiteSpace($rawMarker)) { throw "DashboardViews: Eintrag $position enthaelt einen ungueltigen Markernamen." }
                $markerName = $rawMarker.Trim()
                if ($seenMarkers.ContainsKey($markerName)) { throw "DashboardViews: Eintrag $position enthaelt einen Marker mehrfach." }
                $seenMarkers[$markerName] = $true
                $markers += $markerName
            }
        }
        $seenIds[$id] = $true; $seenNames[$name] = $true
        [void]$normalized.Add([PSCustomObject]@{ Id = $id; Name = $name; Markers = @($markers) })
    }
    return $normalized.ToArray()
}

function Load-Config {
    $dashboardViewsInvalid = $false
    $defaultConfig = @{
        DashboardViews = @()   # v2.39.0: eigene Ansichten; Altbestand startet mit leerer Liste.
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
            # v2.28.0: Referenz in % (bisher 0,38-0,49 = l/l bei Einheit %). Werte < 1,5 werden beim Laden als l/l erkannt und x100 gerechnet.
            [PSCustomObject]@{ Name = 'Hämatokrit (HCT/Hk)'; Unit = '%'; RefMin = 38; RefMax = 49; Group = "Blutbild"; Description = "Anteil der festen Blutzellen (vor allem Erythrozyten) am Gesamtblutvolumen. Referenz Männer ca. 38-49 % (Frauen ca. 35-45 %; maßgeblich ist der Referenzbereich des eigenen Befundes). Laborangaben in l/l (z. B. 0,44) werden automatisch in % umgerechnet (0,44 l/l = 44 %)." },
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
            [PSCustomObject]@{ Name = 'LDL-Cholesterin (LDL)'; Unit = 'mg/dl'; RefMin = 0; RefMax = 130; OptimalMin = 0; OptimalMax = 100; Group = "Blutfette"; Description = "Low-Density Lipoprotein, oft als 'schlechtes' Cholesterin bezeichnet." },
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
            [PSCustomObject]@{ Name = 'Triglyceride (Trig)'; Unit = 'mg/dl'; RefMin = 0; RefMax = 150; OptimalMin = 0; OptimalMax = 100; Group = "Blutfette"; Description = "Neutralfette im Blut, die als Energiereserve dienen." },
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
            # v2.28.0: AIP in mmol/l (TG / 88,57; HDL / 38,67) - die Risikogrenzen nach Dobiášová gelten für mmol/l.
            [PSCustomObject]@{ Name = "Atherogener Index (AIP)"; Unit = ""; RequiredMarkers = @("Triglyceride (Trig)", "HDL-Cholesterin (HDL)"); Formula = "[Math]::Log10((%0% / 88.57) / (%1% / 38.67))"; Group = "Blutfette"; RefMin = -1; RefMax = 0.21; OptimalMin = -1; OptimalMax = 0.11; Description = "Atherogenic Index of Plasma = log10(Triglyceride / HDL), beide in mmol/l (Dobiášová 2001/2004). Korreliert mit kleinen, dichten LDL-Partikeln und dem Herz-Kreislauf-Risiko. < 0,11 niedriges, 0,11-0,21 mittleres, > 0,21 erhöhtes Risiko. Die mg/dl-Werte werden intern umgerechnet (TG / 88,57; HDL / 38,67)." },
            # v2.28.0 BUGFIX: Formel in EINFACHEN Anführungszeichen - in doppelten wurde $personal beim Laden
            # sofort (leer) eingesetzt -> "%0% / ([Math]::Pow(.Groesse / 100, 2))" -> Syntaxfehler, BMI nie berechnet.
            [PSCustomObject]@{ Name = "BMI"; Unit = "kg/m²"; RequiredMarkers = @("Gewicht"); Formula = '%0% / ([Math]::Pow($personal.Groesse / 100, 2))'; Group = "Stoffwechsel"; RefMin = 18.5; RefMax = 24.9; OptimalMin = 20; OptimalMax = 22; Description = "Body-Mass-Index. Verhältnis von Körpergewicht zu Körpergröße." },
            # v2.28.0: TyG = ln(TG x Glukose) / 2 - passend zum Grenzwert 4,68 (Simental-Mendía 2008 / Guerrero-Romero 2010).
            # Bisher ln(TG x Glukose / 2) (Werte ca. 8-9) bei Grenzwert 4,68 -> nie "optimal".
            [PSCustomObject]@{ Name = "TyG-Index"; Unit = ""; RequiredMarkers = @("Triglyceride (Trig)", "Glukose (Glu)"); Formula = "[Math]::Log((%0%) * (%1%)) / 2"; Group = "Stoffwechsel"; RefMin = 0; RefMax = 4.68; OptimalMin = 0; OptimalMax = 4.5; Description = "Triglyceride-Glucose Index = ln(Triglyceride [mg/dl] x Nüchternglukose [mg/dl]) / 2 (Simental-Mendía 2008, Guerrero-Romero 2010). Einfacher Surrogatmarker für Insulinresistenz, stark mit HOMA-IR korreliert. Werte > 4,68 sprechen für Insulinresistenz; < 4,5 optimal. Hinweis: In der Literatur ist auch die Variante ln(TG x Glukose / 2) mit Werten um 8-9 verbreitet (Grenzwert dort ca. 8,7) - nicht vermischen." },
            [PSCustomObject]@{ Name = "Triglyceride/HDL-Ratio"; Unit = ""; RequiredMarkers = @("Triglyceride (Trig)", "HDL-Cholesterin (HDL)"); Formula = "%0% / %1%"; Group = "Blutfette"; RefMin = 0; RefMax = 3.5; OptimalMin = 0; OptimalMax = 2.0; Description = "Metabolischer Gesundheitsmarker. Starker Prädiktor für Insulinresistenz, small-dense LDL und kardiovaskuläre Ereignisse (Salazar 2012). Optimal < 2." },
            [PSCustomObject]@{ Name = "Neutrophile/Lymphozyten-Ratio (NLR)"; Unit = ""; RequiredMarkers = @("Neutrophile Granulozyten (NEUT)", "Lymphozyten (LYMPH)"); Formula = "%0% / %1%"; Group = "Inflammation"; RefMin = 0.78; RefMax = 3.53; OptimalMin = 1.0; OptimalMax = 2.0; Description = "NLR (Zahorec 2021) - validierter Marker für systemische Inflammation. Erhöhte Werte (>3) assoziieren mit InflammAging, kardiovaskulärem Risiko und Gesamtmortalität." },
            # v2.28.0: PLR mit ABSOLUTER Lymphozytenzahl (Lymphozyten-% x Leukozyten / 100) - bisher Thrombozyten / Lymphozyten-% (ca. 8 statt ca. 140).
            [PSCustomObject]@{ Name = "Thrombozyten/Lymphozyten-Ratio (PLR)"; Unit = ""; RequiredMarkers = @("Thrombozyten (PLT)", "Lymphozyten (LYMPH)", "Leukozyten (WBC)"); Formula = "%0% / ((%1% * %2%) / 100)"; Group = "Inflammation"; RefMin = 36; RefMax = 180; OptimalMin = 70; OptimalMax = 130; Description = "PLR = Thrombozyten / absolute Lymphozyten (beide Tsd./µl); absolute Lymphozyten = Lymphozyten-% x Leukozyten / 100. Komplementärer Inflammationsmarker; erhöht bei chronischer Inflammation, onkologischer Belastung und kardiovaskulärem Risiko." },
            [PSCustomObject]@{ Name = "InflammAging-Score"; Unit = "Punkte"; RequiredMarkers = @("C-reaktives Protein (CRP)", "Neutrophile Granulozyten (NEUT)", "Lymphozyten (LYMPH)", "Thrombozyten (PLT)", "Leukozyten (WBC)"); Formula = "[INFLAMMAGING]"; Group = "Longevity"; RefMin = 0; RefMax = 100; OptimalMin = 70; OptimalMax = 100; Description = "Komposit-Score (Franceschi et al.) aus hsCRP, NLR und PLR. 100 = optimal (keine systemische Inflammation), 0 = ausgeprägtes InflammAging. Chronische niedriggradige Entzündung ist Haupttreiber des biologischen Alterns." },
            [PSCustomObject]@{ Name = "PhenoAge (biologisch)"; Unit = "Jahre"; RequiredMarkers = @("Albumin", "Kreatinin (Krea)", "Glukose (Glu)", "C-reaktives Protein (CRP)", "Lymphozyten (LYMPH)", "Mittleres Erythrozytenvolumen (MCV)", "Erythrozyten-Verteilungsbreite (RDW-CV)", "Alkalische Phosphatase (AP)", "Leukozyten (WBC)"); Formula = "[PHENOAGE]"; Group = "Longevity"; RefMin = 20; RefMax = 80; OptimalMin = 20; OptimalMax = 35; Description = "Biologisches Alter nach Levine et al. (2018, Aging). Berechnet aus 9 klinischen Laborwerten via Gompertz-Mortalitäts-Score. Validiert an NHANES-Kohorte. Prädiktor für All-Cause-Mortality, CVD und Krebs." },
            [PSCustomObject]@{ Name = "PhenoAge-Accel"; Unit = "Jahre"; RequiredMarkers = @("Albumin", "Kreatinin (Krea)", "Glukose (Glu)", "C-reaktives Protein (CRP)", "Lymphozyten (LYMPH)", "Mittleres Erythrozytenvolumen (MCV)", "Erythrozyten-Verteilungsbreite (RDW-CV)", "Alkalische Phosphatase (AP)", "Leukozyten (WBC)"); Formula = "[PHENOAGE_ACCEL]"; Group = "Longevity"; RefMin = -20; RefMax = 20; OptimalMin = -20; OptimalMax = -2; Description = "Biologisches minus chronologisches Alter. Negative Werte = jünger als das Kalenderalter (günstig). Werte >0 = beschleunigtes Altern. Direkter, sofort interpretierbarer Longevity-Marker." },
            [PSCustomObject]@{ Name = "Longevity-Score"; Unit = "Punkte"; RequiredMarkers = @("Langzeitzucker (HbA1c)", "Nüchterninsulin", "Vitamin D (25-OH)", "Homocystein", "C-reaktives Protein (CRP)", "NAD+", "Albumin"); Formula = "[LONGEVITY_SCORE]"; Group = "Longevity"; RefMin = 0; RefMax = 100; OptimalMin = 80; OptimalMax = 100; Description = "Mittelwert der Einzelbewertungen (0-100) von HbA1c, Nüchterninsulin, Vitamin D, Homocystein, CRP bzw. hs-CRP, NAD+ und Albumin am selben Testtag. Bewertet wird an den eigenen Referenz- und Optimalbereichen der Marker (wie die Ampel im Risiko-Cockpit): optimal = 100, im Referenzbereich 79 bis 50, außerhalb 49 bis 0. Berechnet ab 4 der 7 Werte; NAD+ ist optional." },
            # v2.27.0: echtes AHA-PREVENT-Basismodell (Khan et al., Circulation 2024). Pflicht nur noch
            # Gesamtcholesterin + HDL; Blutdruck, eGFR und Diabetes mit Ersatzwerten (siehe Calculate-PreventScore).
            # Grenzen nach PREVENT-ASCVD-Kategorien: < 3 % niedrig, 3-< 5 % grenzwertig, 5-< 10 % mittel, >= 10 % hoch.
            [PSCustomObject]@{ Name = "PREVENT-ASCVD-10Y"; Unit = "%"; RequiredMarkers = @("Cholesterin, gesamt (Chol)", "HDL-Cholesterin (HDL)"); Formula = "[PREVENT_SCORE]"; Group = "Risiko"; RefMin = 0; RefMax = 5; OptimalMin = 0; OptimalMax = 3; Description = "AHA-PREVENT (Khan et al., Circulation 2024): 10-Jahres-Risiko für atherosklerotische Herz-Kreislauf-Ereignisse (Herzinfarkt, Schlaganfall). Kategorien: < 3 % niedrig, 3 bis < 5 % grenzwertig, 5 bis < 10 % mittel, ab 10 % hoch. Eingaben: Alter, Geschlecht, Gesamt- und HDL-Cholesterin (Testtag), systolischer Blutdruck (Testtag, sonst Standardwert aus 'Persönliche Metriken'), eGFR (GFR-Marker, sonst aus Kreatinin nach CKD-EPI 2021), Diabetes (Checkbox oder HbA1c >= 6,5 %), Raucher, Blutdrucksenker, Lipidsenker. Gültig für 30-79 Jahre, Gesamtcholesterin 130-320 mg/dl, HDL 20-100 mg/dl, Blutdruck 90-180 mmHg, eGFR 15-140 - außerhalb wird kein Wert berechnet. Für Personen ohne bekannte Herz-Kreislauf-Erkrankung. Das Alter wird je Testtag aus dem Geburtsdatum berechnet ('Persönliche Metriken')." },
            # v2.27.0: PREVENT Herz-Kreislauf gesamt (ASCVD + Herzinsuffizienz). Kategorien: < 5 % niedrig,
            # 5-< 7,5 % grenzwertig, 7,5-< 20 % mittel, >= 20 % hoch.
            [PSCustomObject]@{ Name = "PREVENT-CVD-10Y"; Unit = "%"; RequiredMarkers = @("Cholesterin, gesamt (Chol)", "HDL-Cholesterin (HDL)"); Formula = "[PREVENT_CVD]"; Group = "Risiko"; RefMin = 0; RefMax = 7.5; OptimalMin = 0; OptimalMax = 5; Description = "AHA-PREVENT (Khan et al., Circulation 2024): 10-Jahres-Risiko für Herz-Kreislauf-Erkrankungen gesamt (Herzinfarkt, Schlaganfall UND Herzinsuffizienz). Kategorien: < 5 % niedrig, 5 bis < 7,5 % grenzwertig, 7,5 bis < 20 % mittel, ab 20 % hoch. Eingaben und Gültigkeitsbereich wie PREVENT-ASCVD-10Y. Welcher der beiden PREVENT-Werte im Risiko-Cockpit erscheint, legt die Auswahl 'PREVENT-Anzeige' unter 'Persönliche Metriken' fest." }
        )
        WarningThreshold = 10
        Personal = @{
            Name = ""
            Age = 37                         # Fallback ohne Geburtsdatum; beim Speichern = heutiges Alter
            Geburtsdatum = ""                # v2.29.0: JJJJ-MM-TT - Alter wird je Testtag berechnet
            Geschlecht = "männlich"
            Groesse = 0.0
            Gewicht = 64.5
            KFA = 0.0
            Grundumsatz = 1550
            WorkoutFreq = "5-6x/Woche"
            SystBlutdruck = 120
            Raucher = $false
            PersistentTooltips = $false      # v2.20.0: Hilfstexte dauerhaft anzeigen?
            Blutdrucksenker = $false         # v2.27.0: PREVENT - blutdrucksenkende Medikamente
            Lipidsenker = $false             # v2.27.0: PREVENT - Statin/Lipidsenker
            DiabetesDiagnose = $false        # v2.27.0: PREVENT - diagnostizierter Diabetes
            PreventAnzeige = "ASCVD"         # v2.27.0: Cockpit-Anzeige "ASCVD" (Standard) oder "CVD" (gesamt)
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
                    Markers = @("Homocystein", "Lp(a)", "NT-proBNP", "TnT (Troponin T)", "Systolischer Blutdruck", "ApoB", "PREVENT-ASCVD-10Y", "PREVENT-CVD-10Y", "Hochsensitives CRP (hs-CRP)")
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

            # v2.39.0: Eigene Ansichten separat pruefen; Fehler duerfen keine Ansichten loeschen.
            if ($loadedData.ContainsKey('DashboardViews')) {
                $configData['DashboardViews'] = $loadedData['DashboardViews']
                try { $configData['DashboardViews'] = @(ConvertTo-BtDashboardViews -Views $loadedData['DashboardViews']) }
                catch {
                    $dashboardViewsInvalid = $true
                    Write-Warning "Eigene Dashboard-Ansichten konnten nicht geladen werden: $($_.Exception.Message) Die Konfigurationsdatei bleibt unveraendert."
                }
            }

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
                # v2.27.0: Defaults für neue PREVENT-Felder (bestehende Configs kennen sie noch nicht)
                # v2.29.0: + Geburtsdatum (leer = wie bisher festes Alter)
                $preventDefaults = [ordered]@{ Blutdrucksenker = $false; Lipidsenker = $false; DiabetesDiagnose = $false; PreventAnzeige = "ASCVD"; Geburtsdatum = "" }
                foreach ($pdKey in $preventDefaults.Keys) {
                    if ($configData['Personal'] -is [hashtable]) {
                        if (-not $configData['Personal'].ContainsKey($pdKey)) { $configData['Personal'][$pdKey] = $preventDefaults[$pdKey] }
                    } elseif ($configData['Personal'].PSObject -and ($configData['Personal'].PSObject.Properties.Name -notcontains $pdKey)) {
                        $configData['Personal'] | Add-Member -NotePropertyName $pdKey -NotePropertyValue $preventDefaults[$pdKey] -Force
                    }
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
        DashboardViews = $configData.DashboardViews   # v2.39.0: auch ungueltigen Rohblock erhalten.
        Markers = [System.Collections.ArrayList]@($configData.Markers | ForEach-Object { [PSCustomObject]$_ })
        CalculatedMarkers = [System.Collections.ArrayList]@($configData.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ })
        WarningThreshold = $configData.WarningThreshold
        Personal = $configData.Personal
        AutoBackup = $configData.AutoBackup
        GeneticPredispositions = $configData.GeneticPredispositions
        # v2.25.0: erreichter Marker-Set-Stand (verhindert erneutes Nachziehen)
        MarkerSetVersion = $script:MarkerSetVersion
    }

    # v2.39.0: Load-Config speichert normalerweise sofort; ungueltige Ansichten nicht ueberschreiben.
    if (-not $dashboardViewsInvalid) { Save-Data -data @{ Config = $finalConfig } -type "Config" }
    return $finalConfig
}

function Assert-BtBloodStoredItem {
    # Nur bekannte skalare Item-Felder zulassen; unbekannte Formate nicht ueberschreiben.
    param($Item, [string]$DateString, [switch]$Prepared)
    if ($Item -is [System.Collections.IDictionary]) { $keys = @($Item.Keys) }
    elseif ($Item -is [PSCustomObject]) { $keys = @($Item.PSObject.Properties.Name) }
    else { throw 'Ungültiger Blutwert-Eintrag; es wurde nichts gespeichert.' }
    $required = @('Date', 'Name', 'Value', 'Unit')
    if ($Prepared) { $required += 'Note' }
    if (@($required | Where-Object { $keys -notcontains $_ }).Count -gt 0 -or @($keys | Where-Object { $_ -notin @('Date','Name','Value','Unit','Note') }).Count -gt 0) { throw 'Unbekanntes oder unvollständiges Blutwert-Format; es wurde nichts gespeichert.' }
    if ($Item.Date -isnot [string] -or $Item.Date -cne $DateString -or $Item.Name -isnot [string] -or [string]::IsNullOrWhiteSpace($Item.Name)) { throw 'Datum oder Markername eines Blutwerts ist ungültig.' }
    foreach ($field in @('Unit', 'Note')) { if ($null -ne $Item.$field -and $Item.$field -isnot [string]) { throw 'Einheit oder Notiz eines Blutwerts hat ein unbekanntes Format.' } }
    if ($null -eq $Item.Value -or [int][Type]::GetTypeCode($Item.Value.GetType()) -notin (5..15)) { throw 'Ein vorbereiteter Blutwert muss eine Zahl enthalten.' }
    $number = [double]$Item.Value
    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { throw 'Ein Blutwert muss eine endliche Zahl enthalten.' }
}

function Save-BtBloodValueItems {
    <#
    .SYNOPSIS
        v2.47.0: Vorbereitete Blutwerte eines Datums gemeinsam speichern; Anzahl zurueckgeben.
    .DESCRIPTION
        New-BtBloodValueItem bereitet die Eintraege vor. Hier erneut gegen aktuelle
        Konfiguration, bereits gespeicherte Werte und die gesamte Eingabeliste pruefen.
        Tagesdatei unter Schreibsperre frisch lesen, genau einen verschluesselten Stand
        schreiben. Strikter atomarer Modus: kein Kopier-Fallback und keine Teilspeicherung.
        Altbestand mit einem einzelnen Items-Objekt oder ohne Note bleibt lesbar.
        Unlesbare/mehrdeutige/unbekannte Bestandsformate werden niemals ueberschrieben.
        Keine Aenderung am Verlauf im Speicher; UI erst nach erfolgreicher Rueckkehr laden.
    #>
    param([Parameter(Mandatory)][string]$DateString, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Items)
    $ErrorActionPreference = 'Stop'
    if ($Items.Count -eq 0) { throw 'Keine Blutwerte zum Speichern übergeben.' }
    $parsedDate = [datetime]::MinValue
    if (-not [datetime]::TryParseExact($DateString, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedDate)) { throw 'Ungültiges Datum für die gemeinsame Speicherung.' }
    foreach ($item in $Items) { Assert-BtBloodStoredItem -Item $item -DateString $DateString -Prepared }
    $path, $directory = Get-DailyDataFilePath -dateString $DateString
    $null = [IO.Directory]::CreateDirectory($directory)
    $lock = $null
    try {
        try { $lock = [IO.File]::Open(($path + '.lock'), [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None) }
        catch { throw 'Die Tagesdatei wird gerade verwendet oder kann nicht gesperrt werden. Bitte erneut versuchen; es wurde nichts gespeichert.' }
        $existed = Test-Path -LiteralPath $path
        $originalHash = ''; $existing = @()
        if ($existed) {
            try {
                $originalHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
                $json = Read-ProtectedJsonFile -Path $path 3>$null
                if ([string]::IsNullOrWhiteSpace($json)) { throw 'Leer' }
                $store = $json | ConvertFrom-Json -ErrorAction Stop
                if ($store -isnot [PSCustomObject] -or @($store.PSObject.Properties.Name).Count -ne 1 -or @($store.PSObject.Properties.Name) -notcontains 'Items') { throw 'Format' }
                if ($store.Items -isnot [System.Collections.IList] -and $store.Items -isnot [PSCustomObject]) { throw 'Items-Format' }
                $existing = @($store.Items)
                $seen = @{}
                foreach ($savedItem in $existing) {
                    Assert-BtBloodStoredItem -Item $savedItem -DateString $DateString
                    if ($seen.ContainsKey($savedItem.Name)) { throw 'Bestehendes Duplikat' }
                    $seen[$savedItem.Name] = $true
                }
            } catch { throw 'Die vorhandene Tagesdatei ist nicht lesbar oder enthält ein unbekanntes/ungültiges Format beziehungsweise doppelte Marker. Sie bleibt unverändert.' }
        }
        $fresh = @()
        foreach ($item in $Items) {
            $number = [double]$item.Value
            $arguments = @{
                DateString = $DateString; MarkerQuery = $item.Name
                ValueText = $number.ToString('R', [cultureinfo]::InvariantCulture)
                ExistingItems = $existing; PendingItems = $fresh
            }
            if ($item.Name -eq 'HIV (Anti-HIV-1/2)') {
                if ($number -notin @(0, 1)) { throw 'HIV benötigt einen gültigen vorbereiteten Ergebniscode.' }
                $arguments.HivResult = if ($number -eq 1) { 'reaktiv' } else { 'nicht reaktiv' }
            } elseif ($item.Name -eq $script:ApoeMarkerName) {
                if ($number -ne [Math]::Truncate($number) -or $number -lt 1 -or $number -gt $script:ApoeGenotypeOptions.Count) { throw 'APOE benötigt einen gültigen vorbereiteten Genotypcode.' }
                $arguments.ApoeGenotype = $script:ApoeGenotypeOptions[[int]$number - 1]
            } elseif ($item.Name -eq 'Testosteron, gesamt') { $arguments.TestosteroneUnit = 'ng/dl' }
            $checked = New-BtBloodValueItem @arguments
            if ($checked.Name -cne $item.Name -or [string]$checked.Unit -cne [string]$item.Unit -or [double]$checked.Value -ne $number) { throw 'Marker, Einheit oder Wert hat sich seit der Eingabe geändert. Bitte die Eingaben erneut prüfen.' }
            # Automatische Notizen wurden schon bei der Eingabe erzeugt, nicht doppelt anhaengen.
            $fresh += [PSCustomObject]@{ Date = $DateString; Name = $checked.Name; Value = $number; Unit = $item.Unit; Note = $item.Note }
        }
        $payload = @{ Items = @($existing) + @($fresh) } | ConvertTo-Json -Depth 5
        # Auch Aenderungen durch aeltere, nicht sperrende Programmversionen erkennen.
        if ($existed) {
            if (-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -cne $originalHash) { throw 'Die Tagesdatei wurde zwischenzeitlich geändert. Bitte erneut laden; es wurde nichts gespeichert.' }
        } elseif (Test-Path -LiteralPath $path) { throw 'Die Tagesdatei wurde zwischenzeitlich angelegt. Bitte erneut laden; es wurde nichts gespeichert.' }
        Write-ProtectedJsonFile -Path $path -JsonString $payload -RequireAtomic
        return $fresh.Count
    } finally { if ($null -ne $lock) { $lock.Dispose() } }
}


function Save-Data {
    param ($data, $type, $dateString, [switch]$ThrowOnError)
    try {
        if (-not (Test-Path $dataDir)) { New-Item -Path $dataDir -ItemType Directory -Force | Out-Null }
        if ($type -eq "Config") {
            # v2.14.2 Bugfix: AutoBackup MUSS mit geschrieben werden, sonst gehen alle
            # Backup-Einstellungen nach Neustart verloren.
            $configToSave = @{
                DashboardViews = @(ConvertTo-BtDashboardViews -Views $data.Config.DashboardViews)
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
            # v2.47.0: Neuer gemeinsamer Pfad; bisherige NewItem-Aufrufer bleiben kompatibel.
            $hasNewItems = if ($data -is [System.Collections.IDictionary]) { $data.Contains('NewItems') } else { $null -ne $data.PSObject.Properties['NewItems'] }
            if ($hasNewItems) {
                $null = Save-BtBloodValueItems -DateString $dateString -Items $data.NewItems
                return
            }
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
    } catch {
        if ($ThrowOnError) { throw }
        [System.Windows.Forms.MessageBox]::Show("Fehler beim Speichern der $type-Daten: $($_.Exception.Message)")
    }
}

function Get-BtPathKey {
    <#
    .SYNOPSIS
        v2.26.1: Normalisierter Vergleichsschlüssel für Dateipfade
        (voller Pfad, ohne abschließendes Trennzeichen, Kleinschreibung).
    #>
    param([string]$Path)
    $full = $Path
    try { $full = [System.IO.Path]::GetFullPath($Path) } catch { $full = $Path }
    return $full.TrimEnd('\', '/').ToLowerInvariant()
}

function Load-AllHistoricalItems {
    $allItems = New-Object System.Collections.ArrayList
    # v2.26.1: Tagesdateien, die nicht gelesen werden konnten. Save-AllHistoricalData
    # löscht diese Dateien NIE (bisher: eine defekte Datei brach das Laden ab, alle
    # danach folgenden Tage fehlten im Speicher und wurden beim nächsten Speichern gelöscht).
    $script:HistoricalLoadFailedFiles = New-Object System.Collections.ArrayList
    try {
        # v2.33.0: Messungen (DEXA/Schlaf) im Unterordner "_messungen" sind keine Blutwerte
        $files = Get-ChildItem -Path $script:dailyDataDirBase -Filter "*.json" -Recurse -ErrorAction SilentlyContinue | Where-Object { -not (Test-IsMeasurementPath -Path $_.FullName) }
        foreach ($file in $files) {
            # v2.26.1: Fehler je Datei abfangen - eine defekte Datei stoppt das Laden nicht mehr
            try {
                # v2.19.0: Verschlüsselt lesen (auto-detect)
                $json = Read-ProtectedJsonFile -Path $file.FullName
                if ($json) {
                    $data = $json | ConvertFrom-Json
                    if ($data.Items) {
                        [void]$allItems.AddRange(@($data.Items | ForEach-Object { [PSCustomObject]$_ }))
                    }
                } elseif (-not [string]::IsNullOrWhiteSpace((Get-Content -LiteralPath $file.FullName -Raw -ErrorAction SilentlyContinue))) {
                    # Inhalt vorhanden, aber weder entschlüsselbar noch gültiges JSON
                    [void]$script:HistoricalLoadFailedFiles.Add($file.FullName)
                }
            } catch {
                [void]$script:HistoricalLoadFailedFiles.Add($file.FullName)
                Write-Warning "Tagesdatei nicht lesbar: $($file.FullName) - $($_.Exception.Message)"
            }
        }
    } catch { Write-Warning "Fehler beim Laden der Verlaufsdaten." }

    # v2.28.0: Hämatokrit einheitlich in %. Werte < 1,5 können physiologisch nur l/l sein
    # (z. B. 0,44 l/l = 44 %) und werden im Speicher umgerechnet; beim nächsten Speichern
    # wird der %-Wert geschrieben.
    foreach ($histItem in $allItems) {
        if ($histItem.Name -eq 'Hämatokrit (HCT/Hk)' -and $null -ne $histItem.Value) {
            $hctValue = 0.0
            if ([double]::TryParse([string]$histItem.Value, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$hctValue) -and $hctValue -gt 0 -and $hctValue -lt 1.5) {
                $histItem.Value = [Math]::Round($hctValue * 100, 1)
            }
        }
    }

    # v2.26.1: Nutzer einmalig je Fehlerbild informieren (nicht bei jedem Neuladen erneut)
    if ($script:HistoricalLoadFailedFiles.Count -gt 0) {
        $warnKey = ($script:HistoricalLoadFailedFiles | Sort-Object) -join '|'
        if ($script:HistoricalLoadWarnedKey -ne $warnKey) {
            $script:HistoricalLoadWarnedKey = $warnKey
            $shownFiles = @($script:HistoricalLoadFailedFiles | Select-Object -First 5 | ForEach-Object { "- " + (Split-Path -Path $_ -Leaf) })
            $moreText = if ($script:HistoricalLoadFailedFiles.Count -gt 5) { "`n- ... und $($script:HistoricalLoadFailedFiles.Count - 5) weitere" } else { "" }
            [System.Windows.Forms.MessageBox]::Show(
                "$($script:HistoricalLoadFailedFiles.Count) Tagesdatei(en) konnten nicht gelesen werden und werden nicht angezeigt:`n`n$($shownFiles -join "`n")$moreText`n`nDie Dateien bleiben unverändert erhalten und werden vom Programm NICHT gelöscht.`n`nMögliche Ursachen: Datei beschädigt oder mit einem anderen Windows-Benutzerprofil/PC verschlüsselt (DPAPI).",
                "Tagesdaten nicht lesbar", "OK", "Warning") | Out-Null
        }
    }
    return $allItems
}

function Save-AllHistoricalData {
    param ($AllItems)
    try {
        # v2.26.1 BUGFIX (Datenverlust): Reihenfolge umgestellt.
        # Bisher wurden ZUERST alle Tagesdateien gelöscht und danach neu geschrieben -
        # ein Fehler mitten im Schreiben (ungültiger Wert, Datei gesperrt, Platte voll)
        # vernichtete alle noch nicht geschriebenen Tage. Neu:
        #   Phase 1: alle Tagesdateien nur im Speicher vorbereiten (keine Dateiänderung)
        #   Phase 2: jede Datei atomar schreiben (Write-ProtectedJsonFile)
        #   Phase 3: ERST DANACH verwaiste Tagesdateien (vom Nutzer gelöschte Tage) entfernen
        #            (BUGFIX v2.7.3 bleibt erhalten); unlesbare Dateien werden nie gelöscht.

        # Phase 1: Vorbereiten
        $plannedFiles = @{}
        $groupedData = $AllItems | Group-Object -Property Date
        foreach ($group in $groupedData) {
            $dateString = $group.Name
            $itemsForDay = $group.Group
            $dailyDataFile, $dailyDataDir = Get-DailyDataFilePath -dateString $dateString
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
            # v2.19.0: DPAPI-verschlüsselt schreiben (JSON-Format unverändert)
            $jsonString = @{ Items = $savableItems } | ConvertTo-Json -Depth 5
            $plannedFiles[(Get-BtPathKey -Path $dailyDataFile)] = @{ File = $dailyDataFile; Dir = $dailyDataDir; Json = $jsonString }
        }

        # Schutzliste: beim Laden unlesbare Tagesdateien (siehe Load-AllHistoricalItems)
        $protectedKeys = @{}
        if ($script:HistoricalLoadFailedFiles) {
            foreach ($failedFile in $script:HistoricalLoadFailedFiles) { $protectedKeys[(Get-BtPathKey -Path $failedFile)] = $true }
        }

        # Phase 2: Schreiben (atomar je Datei - bei Fehler bleibt die Vorversion erhalten)
        foreach ($planKey in $plannedFiles.Keys) {
            $plan = $plannedFiles[$planKey]
            if (-not (Test-Path -LiteralPath $plan.Dir)) { New-Item -Path $plan.Dir -ItemType Directory -Force | Out-Null }
            if ($protectedKeys.ContainsKey($planKey) -and (Test-Path -LiteralPath $plan.File)) {
                # Unlesbare Datei würde überschrieben -> Original vorher als Kopie sichern
                $rescuePath = "$($plan.File).unlesbar_$(Get-Date -Format 'yyyyMMdd_HHmmss').bak"
                Copy-Item -LiteralPath $plan.File -Destination $rescuePath -Force -ErrorAction Stop
            }
            Write-ProtectedJsonFile -Path $plan.File -JsonString $plan.Json
        }

        # Phase 3: Erst nach vollständig erfolgreichem Schreiben verwaiste Tagesdateien entfernen
        if (Test-Path $script:dailyDataDirBase) {
            # v2.33.0: Messungen (DEXA/Schlaf) werden hier NIE gelöscht
            $existingFiles = Get-ChildItem -Path $script:dailyDataDirBase -Filter "*.json" -Recurse -ErrorAction SilentlyContinue | Where-Object { -not (Test-IsMeasurementPath -Path $_.FullName) }
            foreach ($file in $existingFiles) {
                $fileKey = Get-BtPathKey -Path $file.FullName
                if ($plannedFiles.ContainsKey($fileKey)) { continue }   # gerade geschrieben
                if ($protectedKeys.ContainsKey($fileKey)) { continue }  # unlesbar -> nie löschen
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction SilentlyContinue
            }
            # Leere Unterverzeichnisse (Monat/Jahr) aufräumen
            Get-ChildItem -Path $script:dailyDataDirBase -Directory -Recurse -ErrorAction SilentlyContinue |
                Sort-Object { $_.FullName.Length } -Descending |
                Where-Object { (Get-ChildItem -Path $_.FullName -ErrorAction SilentlyContinue).Count -eq 0 } |
                ForEach-Object { Remove-Item -Path $_.FullName -Force -ErrorAction SilentlyContinue }
        }

        # Überschriebene (vorher unlesbare) Dateien sind jetzt wieder gültig
        if ($script:HistoricalLoadFailedFiles -and $script:HistoricalLoadFailedFiles.Count -gt 0) {
            $stillFailed = @($script:HistoricalLoadFailedFiles | Where-Object { -not $plannedFiles.ContainsKey((Get-BtPathKey -Path $_)) })
            $script:HistoricalLoadFailedFiles = New-Object System.Collections.ArrayList
            foreach ($sf in $stillFailed) { [void]$script:HistoricalLoadFailedFiles.Add($sf) }
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ein schwerwiegender Fehler ist beim Speichern der Daten aufgetreten: $($_.Exception.Message)", "Speicherfehler", "OK", "Error")
    }
}

# =====================================================================
# v2.33.0: VORBEREITUNG WEITERER MESSWERTE - DEXA-Scan und Schlaf
# ---------------------------------------------------------------------
# Nur Grundarchitektur (Katalog, Datenmodell, Speichern/Laden) - noch KEINE
# Oberfläche. Die Werte erscheinen bewusst nirgends in der GUI.
#
# Ablage: <Datenordner>\data\_messungen\<Kategorie>\<JJJJ>\<JJJJ-MM-TT>.json
#   - DPAPI-verschlüsselt und atomar wie die Tagesdateien (Write-ProtectedJsonFile)
#   - liegt im "data"-Ordner, damit ALLE vorhandenen Sicherungswege sie
#     automatisch erfassen: .btbackup (portabel, entschlüsselt im Archiv),
#     Klartext-Export, Auto-Backup, Wiederherstellung, "Datenpfad ändern"
#   - Load-AllHistoricalItems und Save-AllHistoricalData überspringen den
#     Ordner (Test-IsMeasurementPath) - Blutwerte und Messungen bleiben getrennt
# Datum: DEXA = Scan-Tag; Schlaf = Tag des AUFWACHENS (eine Nacht je Datum).
# =====================================================================
$script:MeasurementFolderName    = '_messungen'
$script:MeasurementSchemaVersion = 1
$script:MeasurementCategories    = @('DEXA', 'Schlaf', 'Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet') # v2.41.0: + Tageswerte

function New-MeasurementDefinition {
    <# v2.33.0: Ein Eintrag des Messwert-Katalogs (Plaus* = erlaubter Eingabebereich). #>
    param(
        [Parameter(Mandatory)][string]$Key, [Parameter(Mandatory)][string]$Name, [string]$Unit = '',
        [Parameter(Mandatory)][string]$Group, [double]$PlausMin, [double]$PlausMax,
        $RefMin = $null, $RefMax = $null, $OptimalMin = $null, $OptimalMax = $null, [string]$Description = ''
    )
    # Bereiche immer als Zahl (Argumente wie -1.0 kommen im Befehlsmodus sonst als Text an)
    $toNumber = { param($x) if ($null -eq $x -or "$x" -eq '') { $null } else { [double]$x } }
    return [PSCustomObject]@{
        Key = $Key; Name = $Name; Unit = $Unit; Group = $Group; PlausMin = $PlausMin; PlausMax = $PlausMax
        RefMin = (& $toNumber $RefMin); RefMax = (& $toNumber $RefMax); OptimalMin = (& $toNumber $OptimalMin); OptimalMax = (& $toNumber $OptimalMax); Description = $Description
    }
}

function Get-MeasurementCatalog {
    <#
    .SYNOPSIS
        v2.33.0: Katalog aller vorbereiteten Messwerte einer Kategorie.
    .DESCRIPTION
        DEXA: alle üblichen Ergebnisse eines DXA-Scans (Körperzusammensetzung,
        Fettverteilung, Viszeralfett, Regionen, Knochendichte inkl. T-/Z-Score, TBS).
        Bewertungsbereiche nur dort, wo sie geschlechts- und altersunabhängig
        festgelegt sind (T-/Z-Score, Schlafdauer); sonst leer + Hinweis.
        Schlaf: Schlafdauer (h), Schlafscore (optional), Schlafphasen in Minuten.
    #>
    param([Parameter(Mandatory)][ValidateSet('DEXA', 'Schlaf', 'Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category)
    $catalog = New-Object System.Collections.ArrayList
    if ($Category -eq 'DEXA') {
        $gBody = 'DEXA Körperzusammensetzung'; $gDist = 'DEXA Fettverteilung'; $gVat = 'DEXA Viszeralfett'; $gReg = 'DEXA Regional'; $gBone = 'DEXA Knochendichte'
        $bodyDefs = @(
            @('Gesamtmasse', 'Gesamtmasse (DEXA)', 'kg', 20, 350, 'Summe aus Fett-, Mager- und Knochenmasse laut Scan (kann leicht vom Waagengewicht abweichen).'),
            @('Koerperfett_Prozent', 'Körperfettanteil gesamt', '%', 2, 75, 'Fettmasse / Gesamtmasse. Bezugswerte alters- und geschlechtsabhängig; DXA-Werte liegen meist über BIA-Werten.'),
            @('Fettmasse', 'Fettmasse gesamt', 'kg', 0.5, 250, 'Gesamte Fettmasse.'),
            @('Magermasse', 'Magermasse gesamt (Lean, ohne Knochen)', 'kg', 5, 150, 'Fettfreies Weichgewebe (Muskeln, Organe, Wasser) ohne Knochenmineral.'),
            @('FettfreieMasse', 'Fettfreie Masse (Lean + Knochenmineral)', 'kg', 5, 160, 'Magermasse plus Knochenmineralgehalt.'),
            @('BMC_Gesamt', 'Knochenmineralgehalt gesamt (BMC)', 'kg', 0.3, 6, 'Mineralgehalt des gesamten Skeletts.'),
            @('ALM', 'Appendikuläre Magermasse (ALM, Arme + Beine)', 'kg', 3, 60, 'Magermasse beider Arme und Beine - Näherung für die Skelettmuskelmasse.'),
            @('ALMI', 'ALM-Index (ALM / Größe²)', 'kg/m²', 2, 20, 'Sarkopenie: niedrige Muskelmasse bei < 7,0 kg/m² (Männer) bzw. < 5,5 kg/m² (Frauen) nach EWGSOP2 (Cruz-Jentoft 2019). Wird bei Bedarf aus ALM und Größe berechnet.'),
            @('LMI', 'Magermasse-Index (Lean / Größe²)', 'kg/m²', 5, 35, 'Magermasse bezogen auf die Körpergröße.'),
            @('FFMI', 'Fettfreie-Masse-Index (FFMI)', 'kg/m²', 8, 40, 'Fettfreie Masse bezogen auf die Körpergröße.'),
            @('FMI', 'Fettmasse-Index (Fett / Größe²)', 'kg/m²', 0.5, 70, 'Normalbereich nach Kelly 2009 (NHANES) etwa 3-6 kg/m² (Männer) bzw. 5-9 kg/m² (Frauen).')
        )
        foreach ($d in $bodyDefs) { [void]$catalog.Add((New-MeasurementDefinition -Key $d[0] -Name $d[1] -Unit $d[2] -Group $gBody -PlausMin $d[3] -PlausMax $d[4] -Description $d[5])) }
        $distDefs = @(
            @('Android_Fett_Prozent', 'Android-Fett (Bauchregion)', '%', 1, 80, 'Fettanteil der Bauchregion.'),
            @('Gynoid_Fett_Prozent', 'Gynoid-Fett (Hüft-/Oberschenkelregion)', '%', 1, 80, 'Fettanteil der Hüft-/Oberschenkelregion.'),
            @('Android_Fettmasse', 'Android-Fettmasse', 'kg', 0, 40, 'Fettmasse der Bauchregion.'),
            @('Gynoid_Fettmasse', 'Gynoid-Fettmasse', 'kg', 0, 50, 'Fettmasse der Hüft-/Oberschenkelregion.'),
            @('AG_Ratio', 'Android/Gynoid-Verhältnis (A/G)', '', 0.1, 3, 'Höher = mehr Bauchfett; Bezugswerte geschlechtsabhängig. Wird bei Bedarf aus Android-/Gynoid-Fett-% berechnet.'),
            @('Rumpf_Bein_Fett_Ratio', 'Rumpf-/Beinfett-Verhältnis', '', 0.1, 6, 'Zentrale gegenüber peripherer Fettverteilung.'),
            @('Rumpf_Gliedmassen_Fett_Ratio', 'Rumpf-/Gliedmaßenfett-Verhältnis', '', 0.1, 6, 'Rumpffett / Fett aller Gliedmaßen.')
        )
        foreach ($d in $distDefs) { [void]$catalog.Add((New-MeasurementDefinition -Key $d[0] -Name $d[1] -Unit $d[2] -Group $gDist -PlausMin $d[3] -PlausMax $d[4] -Description $d[5])) }
        $vatDefs = @(
            @('VAT_Masse', 'Viszerales Fett (VAT) - Masse', 'g', 0, 10000, 'Fett um die inneren Organe. Bezugswerte alters-/geschlechtsabhängig (z. B. LEAD-Kohorte, Ofenheimer 2020: Männer 40-50 J. im Mittel ca. 1,2 kg, 50-60 J. ca. 1,6 kg).'),
            @('VAT_Volumen', 'Viszerales Fett (VAT) - Volumen', 'cm³', 0, 11000, 'Volumen des viszeralen Fetts.'),
            @('VAT_Flaeche', 'Viszerales Fett (VAT) - Fläche', 'cm²', 0, 700, 'Querschnittsfläche des viszeralen Fetts.')
        )
        foreach ($d in $vatDefs) { [void]$catalog.Add((New-MeasurementDefinition -Key $d[0] -Name $d[1] -Unit $d[2] -Group $gVat -PlausMin $d[3] -PlausMax $d[4] -Description $d[5])) }
        # Regionen: Fett, Mager, Fett-%, Knochenmineral je Arm/Bein links/rechts, Rumpf, Kopf
        $regions = @(
            @('Arm_Links', 'Arm links', 25, 15, 1.0), @('Arm_Rechts', 'Arm rechts', 25, 15, 1.0),
            @('Bein_Links', 'Bein links', 50, 35, 2.0), @('Bein_Rechts', 'Bein rechts', 50, 35, 2.0),
            @('Rumpf', 'Rumpf', 100, 70, 2.5), @('Kopf', 'Kopf', 5, 8, 1.5)
        )
        foreach ($rg in $regions) {
            [void]$catalog.Add((New-MeasurementDefinition -Key "$($rg[0])_Fettmasse" -Name "$($rg[1]) - Fettmasse" -Unit 'kg' -Group $gReg -PlausMin 0 -PlausMax $rg[2] -Description "Fettmasse der Region $($rg[1])."))
            [void]$catalog.Add((New-MeasurementDefinition -Key "$($rg[0])_Magermasse" -Name "$($rg[1]) - Magermasse" -Unit 'kg' -Group $gReg -PlausMin 0 -PlausMax $rg[3] -Description "Magermasse (ohne Knochen) der Region $($rg[1])."))
            [void]$catalog.Add((New-MeasurementDefinition -Key "$($rg[0])_Fett_Prozent" -Name "$($rg[1]) - Fettanteil" -Unit '%' -Group $gReg -PlausMin 1 -PlausMax 80 -Description "Fettanteil der Region $($rg[1])."))
            [void]$catalog.Add((New-MeasurementDefinition -Key "$($rg[0])_BMC" -Name "$($rg[1]) - Knochenmineralgehalt" -Unit 'kg' -Group $gReg -PlausMin 0 -PlausMax $rg[4] -Description "Knochenmineralgehalt der Region $($rg[1])."))
        }
        # Knochendichte je Messort: BMD, T-Score, Z-Score
        $sites = @(@('LWS', 'Lendenwirbelsäule (L1-L4)'), @('Schenkelhals', 'Schenkelhals (Femurhals)'), @('Gesamthuefte', 'Gesamthüfte'), @('Radius33', 'Unterarm (Radius 33 %)'), @('Ganzkoerper', 'Ganzkörper'))
        foreach ($st in $sites) {
            [void]$catalog.Add((New-MeasurementDefinition -Key "BMD_$($st[0])" -Name "Knochendichte (BMD) $($st[1])" -Unit 'g/cm²' -Group $gBone -PlausMin 0.2 -PlausMax 2.5 -Description "Flächenbezogene Knochenmineraldichte $($st[1]); Werte sind gerätespezifisch (Hersteller-Referenzdatenbank)."))
            [void]$catalog.Add((New-MeasurementDefinition -Key "TScore_$($st[0])" -Name "T-Score $($st[1])" -Unit 'SD' -Group $gBone -PlausMin -8 -PlausMax 8 -RefMin -2.4 -RefMax 6 -OptimalMin -1.0 -OptimalMax 6 -Description "Vergleich mit jungen Erwachsenen. WHO: ab -1,0 normal, -1,0 bis -2,4 Osteopenie, -2,5 und tiefer Osteoporose (gilt für Männer ab 50 J. und Frauen nach der Menopause)."))
            [void]$catalog.Add((New-MeasurementDefinition -Key "ZScore_$($st[0])" -Name "Z-Score $($st[1])" -Unit 'SD' -Group $gBone -PlausMin -8 -PlausMax 8 -RefMin -1.9 -RefMax 6 -OptimalMin -1.9 -OptimalMax 6 -Description "Vergleich mit Gleichaltrigen. ISCD: -2,0 und tiefer = unter dem altersentsprechenden Bereich (maßgeblich für Männer unter 50 J. und Frauen vor der Menopause)."))
        }
        [void]$catalog.Add((New-MeasurementDefinition -Key 'TBS' -Name 'Trabecular Bone Score (TBS, LWS)' -Unit '' -Group $gBone -PlausMin 0.5 -PlausMax 2.0 -Description 'Knochen-Mikroarchitektur der Lendenwirbelsäule (Zusatzsoftware). Grenzwerte in der Literatur uneinheitlich (z. B. 1,350 / 1,200 bzw. 1,310 / 1,230) - daher noch ohne Bewertung.'))
    } elseif ($Category -eq 'Schlaf') {
        $gSleep = 'Schlaf'
        [void]$catalog.Add((New-MeasurementDefinition -Key 'Schlafdauer' -Name 'Schlafdauer' -Unit 'h' -Group $gSleep -PlausMin 0.5 -PlausMax 24 -RefMin 6 -RefMax 10 -OptimalMin 7 -OptimalMax 9 -Description 'Gesamtschlafzeit der Nacht (ohne Wachphasen). Erwachsene 26-64 J.: 7-9 h empfohlen, 6 bzw. 10 h noch vertretbar (National Sleep Foundation 2015); AASM/SRS 2015: mindestens 7 h. Datum = Tag des Aufwachens.'))
        [void]$catalog.Add((New-MeasurementDefinition -Key 'Schlafscore' -Name 'Schlafscore (optional)' -Unit 'Punkte' -Group $gSleep -PlausMin 0 -PlausMax 100 -Description 'Wert des Trackers bzw. der Uhr (0-100). Berechnung herstellerabhängig - daher ohne eigene Bewertung.'))
        [void]$catalog.Add((New-MeasurementDefinition -Key 'Phase_Wach' -Name 'Wachphasen' -Unit 'min' -Group $gSleep -PlausMin 0 -PlausMax 1440 -Description 'Wachzeit nach dem Einschlafen (optional).'))
        [void]$catalog.Add((New-MeasurementDefinition -Key 'Phase_Leicht' -Name 'Leichtschlaf (N1/N2)' -Unit 'min' -Group $gSleep -PlausMin 0 -PlausMax 1440 -Description 'Optional.'))
        [void]$catalog.Add((New-MeasurementDefinition -Key 'Phase_Tief' -Name 'Tiefschlaf (N3)' -Unit 'min' -Group $gSleep -PlausMin 0 -PlausMax 1440 -Description 'Optional.'))
        [void]$catalog.Add((New-MeasurementDefinition -Key 'Phase_REM' -Name 'REM-Schlaf' -Unit 'min' -Group $gSleep -PlausMin 0 -PlausMax 1440 -Description 'Optional.'))
    } else {
        # v2.35.0: Vitalwerte, Koerper, EKG, Fitness
        foreach ($extendedDefinition in (Get-ExtendedMeasurementDefinitions -Category $Category)) { [void]$catalog.Add($extendedDefinition) }
    }
    return , $catalog.ToArray()
}

function Resolve-MeasurementDefinition {
    <# v2.33.0: Katalogeintrag zu Schlüssel ODER Anzeigename (Groß/Klein egal), sonst $null. #>
    param([Parameter(Mandatory)][string]$Category, [Parameter(Mandatory)][string]$KeyOrName)
    $lookup = $KeyOrName.Trim()
    foreach ($definition in (Get-MeasurementCatalog -Category $Category)) {
        if ($definition.Key -ieq $lookup -or $definition.Name -ieq $lookup) { return $definition }
    }
    return $null
}

function ConvertTo-BtMeasurementDate {
    <#
    .SYNOPSIS
        v2.33.0: Datum -> 'yyyy-MM-dd'. Erlaubt [datetime], JJJJ-MM-TT, TT.MM.JJJJ, T.M.JJJJ.
        Kein Datum in der Zukunft.
    #>
    param([Parameter(Mandatory)]$Date)
    $parsedDate = $null
    if ($Date -is [datetime]) { $parsedDate = $Date.Date }
    else {
        $tmpDate = [datetime]::MinValue
        $formats = [string[]]@('yyyy-MM-dd', 'dd.MM.yyyy', 'd.M.yyyy')
        if ([datetime]::TryParseExact(([string]$Date).Trim(), $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$tmpDate)) { $parsedDate = $tmpDate.Date }
    }
    if ($null -eq $parsedDate) { throw "Ungültiges Datum: '$Date' (erwartet TT.MM.JJJJ oder JJJJ-MM-TT)." }
    if ($parsedDate -gt (Get-Date).Date) { throw "Das Datum $($parsedDate.ToString('dd.MM.yyyy')) liegt in der Zukunft." }
    return $parsedDate.ToString('yyyy-MM-dd')
}

function ConvertTo-BtDurationMinutes {
    <#
    .SYNOPSIS
        v2.33.0: Dauer -> Minuten. Stunden ODER Minuten, z. B. 95 | 1,5 | "1:35" | "1h 35min" |
        "1 Std 35 Min" | "1,5 h" | "95 min". Reine Zahlen gelten in -DefaultUnit.
    #>
    param([Parameter(Mandatory)]$Value, [ValidateSet('min', 'h')][string]$DefaultUnit = 'min')
    if ($null -eq $Value) { throw 'Dauer fehlt.' }
    $minutes = $null
    if ($Value -is [int] -or $Value -is [long] -or $Value -is [double] -or $Value -is [decimal] -or $Value -is [single]) {
        $minutes = if ($DefaultUnit -eq 'h') { [double]$Value * 60 } else { [double]$Value }
    } else {
        $text = ([string]$Value).Trim().ToLowerInvariant() -replace ',', '.'
        $numberPattern = '\d+(?:\.\d+)?'
        if ($text -match "^(\d+):([0-5]\d)$") {
            $minutes = [double]$Matches[1] * 60 + [double]$Matches[2]
        } elseif ($text -match "^($numberPattern)\s*(?:h|std|stunde|stunden)\.?\s*(?:($numberPattern)\s*(?:m|min|minute|minuten)?\.?)?$") {
            $minutes = [double]::Parse($Matches[1], [System.Globalization.CultureInfo]::InvariantCulture) * 60
            if ($Matches[2]) { $minutes += [double]::Parse($Matches[2], [System.Globalization.CultureInfo]::InvariantCulture) }
        } elseif ($text -match "^($numberPattern)\s*(?:m|min|minute|minuten)\.?$") {
            $minutes = [double]::Parse($Matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
        } elseif ($text -match "^($numberPattern)$") {
            $plainNumber = [double]::Parse($Matches[1], [System.Globalization.CultureInfo]::InvariantCulture)
            $minutes = if ($DefaultUnit -eq 'h') { $plainNumber * 60 } else { $plainNumber }
        }
    }
    if ($null -eq $minutes -or [double]::IsNaN($minutes) -or $minutes -lt 0) { throw "Ungültige Dauer: '$Value' (z. B. 95, 1:35, 1h 35min, 1,5 h)." }
    return [Math]::Round($minutes, 1)
}

function Get-MeasurementRoot {
    <# v2.33.0: Ablageordner der Messungen (im "data"-Ordner, siehe Kopfkommentar). #>
    return (Join-Path -Path $script:dailyDataDirBase -ChildPath $script:MeasurementFolderName)
}

function Test-IsMeasurementPath {
    <# v2.33.0: $true, wenn die Datei im Messungs-Ordner liegt (für Blutwert-Laden/-Speichern tabu). #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not $script:dailyDataDirBase) { return $false }
    try {
        $rootFull = [System.IO.Path]::GetFullPath((Get-MeasurementRoot)).TrimEnd('\', '/')
        $pathFull = [System.IO.Path]::GetFullPath($Path)
    } catch { return $false }
    return ($pathFull.StartsWith($rootFull + '\', [System.StringComparison]::OrdinalIgnoreCase) -or $pathFull.StartsWith($rootFull + '/', [System.StringComparison]::OrdinalIgnoreCase))
}

function Get-MeasurementFilePath {
    <# v2.33.0: <data>\_messungen\<Kategorie>\<JJJJ>\<JJJJ-MM-TT>.json #>
    param([Parameter(Mandatory)][ValidateSet('DEXA', 'Schlaf', 'Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category, [Parameter(Mandatory)]$Date)
    $isoDate = ConvertTo-BtMeasurementDate -Date $Date
    return (Join-Path (Join-Path (Join-Path (Get-MeasurementRoot) $Category) $isoDate.Substring(0, 4)) "$isoDate.json")
}

function New-DexaRecord {
    <#
    .SYNOPSIS
        v2.33.0: Prüft DEXA-Werte und erstellt einen Datensatz (speichert NICHT).
    .PARAMETER Values
        Hashtable Schlüssel ODER Anzeigename -> Wert (Zahl oder Text mit Komma), z. B.
        @{ Koerperfett_Prozent = 18,4; 'Fettmasse gesamt' = '14,7'; VAT_Masse = 620 }
    #>
    param([Parameter(Mandatory)]$Date, [Parameter(Mandatory)][hashtable]$Values, [string]$Geraet = '', [string]$Notiz = '')
    if ($Values.Count -eq 0) { throw 'Keine DEXA-Werte angegeben.' }
    $checkedValues = [ordered]@{}
    foreach ($inputKey in @($Values.Keys)) {
        $definition = Resolve-MeasurementDefinition -Category 'DEXA' -KeyOrName ([string]$inputKey)
        if (-not $definition) { throw "Unbekannter DEXA-Wert: '$inputKey'." }
        if ($checkedValues.Contains($definition.Key)) { throw "DEXA-Wert '$($definition.Name)' ist doppelt angegeben." }
        $rawValue = $Values[$inputKey]
        $numericValue = $null
        try { $numericValue = if ($rawValue -is [string]) { Parse-Number $rawValue } else { [double]$rawValue } } catch { $numericValue = $null }
        if ($null -eq $numericValue -or [double]::IsNaN($numericValue)) { throw "Ungültiger Zahlenwert für '$($definition.Name)': '$rawValue'." }
        if ($numericValue -lt $definition.PlausMin -or $numericValue -gt $definition.PlausMax) {
            throw "Unplausibler Wert für '$($definition.Name)': $numericValue $($definition.Unit) (erlaubt $($definition.PlausMin) bis $($definition.PlausMax))."
        }
        $checkedValues[$definition.Key] = [double]$numericValue
    }
    return [ordered]@{
        SchemaVersion = $script:MeasurementSchemaVersion; Kategorie = 'DEXA'; Datum = (ConvertTo-BtMeasurementDate -Date $Date)
        Werte = $checkedValues; Geraet = $Geraet; Notiz = $Notiz; Erfasst = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    }
}

function ConvertTo-BtSleepPhaseKey {
    <# v2.33.0: Schlafphasen-Bezeichnung -> Wach | Leicht | Tief | REM (sonst $null). #>
    param([string]$Name)
    switch -Regex ($Name.Trim().ToLowerInvariant()) {
        '^(wach|wachphase|wachphasen|wachzeit|awake|wake)$' { return 'Wach' }
        '^(leicht|leichtschlaf|light|n1|n2|n1n2|n1/n2|kernschlaf|core)$' { return 'Leicht' }
        '^(tief|tiefschlaf|deep|n3)$' { return 'Tief' }
        '^(rem|rem-schlaf|remschlaf|traumschlaf)$' { return 'REM' }
    }
    return $null
}

function New-SleepRecord {
    <#
    .SYNOPSIS
        v2.33.0: Prüft eine Schlaf-Erfassung und erstellt einen Datensatz (speichert NICHT).
    .PARAMETER Schlafdauer
        Pflicht. Stunden (7,5) oder Text ("7:30", "7h 30min", "450 min").
    .PARAMETER Schlafscore
        Optional, 0-100.
    .PARAMETER Phasen
        Optional: @{ Wach = 20; Leicht = 230; Tief = 95; REM = 105 } - Einheit siehe -PhasenEinheit,
        einzelne Werte auch als Text ("1:35", "1,5 h", "95 min").
    .PARAMETER Zyklen
        Optional: Liste der einzelnen Schlafzyklen - je Eintrag eine Dauer (wie Phasen) oder
        @{ Dauer = 95; Phasen = @{ Leicht = 50; Tief = 25; REM = 20 } }.
    .OUTPUTS
        Datensatz (geordnete Hashtable). Unstimmigkeiten (z. B. Summe der Phasen weicht von
        der Schlafdauer ab) stehen in 'Hinweise' - sie verhindern das Speichern nicht.
    #>
    param(
        [Parameter(Mandatory)]$Date, [Parameter(Mandatory)]$Schlafdauer, $Schlafscore = $null,
        [hashtable]$Phasen = $null, [object[]]$Zyklen = $null, [ValidateSet('min', 'h')][string]$PhasenEinheit = 'min',
        [string]$Quelle = '', [string]$Notiz = ''
    )
    $hints = New-Object System.Collections.ArrayList
    $durationMinutes = ConvertTo-BtDurationMinutes -Value $Schlafdauer -DefaultUnit 'h'
    $durationHours = [Math]::Round($durationMinutes / 60, 2)
    if ($durationHours -lt 0.5 -or $durationHours -gt 24) { throw "Unplausible Schlafdauer: $durationHours h (erlaubt 0,5 bis 24 h)." }
    $sleepValues = [ordered]@{ Schlafdauer = $durationHours }
    if ($null -ne $Schlafscore -and "$Schlafscore" -ne '') {
        $scoreValue = $null
        try { $scoreValue = if ($Schlafscore -is [string]) { Parse-Number $Schlafscore } else { [double]$Schlafscore } } catch { $scoreValue = $null }
        if ($null -eq $scoreValue -or $scoreValue -lt 0 -or $scoreValue -gt 100) { throw "Ungültiger Schlafscore: '$Schlafscore' (erlaubt 0 bis 100)." }
        $sleepValues['Schlafscore'] = [double]$scoreValue
    }
    $convertPhases = {
        param([hashtable]$PhaseTable, [string]$Context)
        $phaseResult = [ordered]@{}
        foreach ($phaseInput in @($PhaseTable.Keys)) {
            $phaseKey = ConvertTo-BtSleepPhaseKey -Name ([string]$phaseInput)
            if (-not $phaseKey) { throw "Unbekannte Schlafphase '$phaseInput'$Context (erlaubt: Wach, Leicht, Tief, REM)." }
            if ($phaseResult.Contains($phaseKey)) { throw "Schlafphase '$phaseKey'$Context ist doppelt angegeben." }
            $phaseMinutes = ConvertTo-BtDurationMinutes -Value $PhaseTable[$phaseInput] -DefaultUnit $PhasenEinheit
            if ($phaseMinutes -gt 1440) { throw "Unplausible Dauer für '$phaseKey'$Context`: $phaseMinutes min." }
            $phaseResult[$phaseKey] = $phaseMinutes
        }
        return $phaseResult
    }
    $phaseMinutesTable = $null
    if ($Phasen -and $Phasen.Count -gt 0) {
        $phaseMinutesTable = & $convertPhases $Phasen ''
        $sleepPhaseSum = 0.0
        foreach ($sleepPhase in @('Leicht', 'Tief', 'REM')) { if ($phaseMinutesTable.Contains($sleepPhase)) { $sleepPhaseSum += $phaseMinutesTable[$sleepPhase] } }
        if ($sleepPhaseSum -gt 0) {
            $allowedDiff = [Math]::Max(15, 0.1 * $durationMinutes)
            if ([Math]::Abs($sleepPhaseSum - $durationMinutes) -gt $allowedDiff) {
                [void]$hints.Add("Summe Leicht + Tief + REM ($([Math]::Round($sleepPhaseSum / 60, 2)) h) weicht von der Schlafdauer ($durationHours h) ab.")
            }
        }
    }
    $cycleList = New-Object System.Collections.ArrayList
    if ($Zyklen) {
        $cycleNumber = 0
        foreach ($cycleInput in $Zyklen) {
            $cycleNumber++
            $cycleDurationInput = $cycleInput; $cyclePhaseInput = $null
            if ($cycleInput -is [hashtable] -or $cycleInput -is [System.Collections.Specialized.OrderedDictionary]) {
                $cycleDurationInput = $cycleInput['Dauer']; $cyclePhaseInput = $cycleInput['Phasen']
            }
            $cycleMinutes = ConvertTo-BtDurationMinutes -Value $cycleDurationInput -DefaultUnit $PhasenEinheit
            if ($cycleMinutes -lt 5 -or $cycleMinutes -gt 300) { throw "Unplausible Dauer für Schlafzyklus $cycleNumber`: $cycleMinutes min (erlaubt 5 bis 300 min)." }
            $cycleEntry = [ordered]@{ Nr = $cycleNumber; DauerMinuten = $cycleMinutes }
            if ($cyclePhaseInput -is [hashtable] -and $cyclePhaseInput.Count -gt 0) { $cycleEntry['PhasenMinuten'] = & $convertPhases $cyclePhaseInput " (Zyklus $cycleNumber)" }
            [void]$cycleList.Add($cycleEntry)
        }
        $cycleSum = ($cycleList | ForEach-Object { $_.DauerMinuten } | Measure-Object -Sum).Sum
        if ([Math]::Abs($cycleSum - $durationMinutes) -gt [Math]::Max(30, 0.2 * $durationMinutes)) {
            [void]$hints.Add("Summe der Schlafzyklen ($([Math]::Round($cycleSum / 60, 2)) h) weicht deutlich von der Schlafdauer ($durationHours h) ab.")
        }
    }
    foreach ($hint in $hints) { Write-Warning "Schlaf $Date`: $hint" }
    return [ordered]@{
        SchemaVersion = $script:MeasurementSchemaVersion; Kategorie = 'Schlaf'; Datum = (ConvertTo-BtMeasurementDate -Date $Date)
        Werte = $sleepValues; PhasenMinuten = $phaseMinutesTable; Zyklen = @($cycleList.ToArray()); Quelle = $Quelle; Notiz = $Notiz
        Hinweise = @($hints.ToArray()); Erfasst = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    }
}

function Save-MeasurementRecord {
    <#
    .SYNOPSIS
        v2.33.0: Speichert einen Datensatz von New-DexaRecord / New-SleepRecord
        (verschlüsselt, atomar). Ein Datensatz je Kategorie und Datum - ein vorhandener
        Datensatz desselben Tages wird ersetzt.
    .OUTPUTS
        Pfad der gespeicherten Datei.
    #>
    param([Parameter(Mandatory)]$Record)
    $category = [string]$Record['Kategorie']
    if ($category -eq 'Aktivitaet') { Assert-BtActivityDayRecord -Record $Record }
    if ($script:MeasurementCategories -notcontains $category) { throw "Unbekannte Kategorie '$category'." }
    if ($script:MeasurementMultiCategories -contains $category) {
        # v2.35.0: Tagesdatensatz mit Liste "Messungen" (Add-MeasurementEntry / Remove-MeasurementEntry)
        if (@($Record['Messungen'] | Where-Object { $null -ne $_ }).Count -eq 0) { throw 'Der Tagesdatensatz enthält keine Messungen.' }
    } elseif (-not $Record['Werte'] -or @($Record['Werte'].Keys).Count -eq 0) { throw 'Der Datensatz enthält keine Werte.' }
    $filePath = Get-MeasurementFilePath -Category $category -Date $Record['Datum']
    $fileDir = Split-Path -Path $filePath -Parent
    if (-not (Test-Path -LiteralPath $fileDir)) { New-Item -ItemType Directory -Path $fileDir -Force | Out-Null }
    Write-ProtectedJsonFile -Path $filePath -JsonString ($Record | ConvertTo-Json -Depth 8)
    return $filePath
}

function Get-MeasurementRecords {
    <#
    .SYNOPSIS
        v2.33.0: Lädt alle Datensätze einer Kategorie (optional Zeitraum), nach Datum sortiert.
        Unlesbare Dateien werden übersprungen und in $script:MeasurementLoadFailedFiles gemeldet.
    #>
    param([Parameter(Mandatory)][ValidateSet('DEXA', 'Schlaf', 'Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category, $From = $null, $To = $null)
    $script:MeasurementLoadFailedFiles = New-Object System.Collections.ArrayList
    $records = New-Object System.Collections.ArrayList
    $categoryDir = Join-Path (Get-MeasurementRoot) $Category
    if (-not (Test-Path -LiteralPath $categoryDir)) { return , @() }
    $fromIso = if ($null -ne $From) { ConvertTo-BtMeasurementDate -Date $From } else { $null }
    $toIso = if ($null -ne $To) { ConvertTo-BtMeasurementDate -Date $To } else { $null }
    foreach ($file in @(Get-ChildItem -LiteralPath $categoryDir -Filter '*.json' -Recurse -File -ErrorAction SilentlyContinue)) {
        try {
            $json = Read-ProtectedJsonFile -Path $file.FullName
            if (-not $json) { throw 'leer oder nicht lesbar' }
            $loadedRecord = $json | ConvertFrom-Json
            if ([string]$loadedRecord.Kategorie -ne $Category -or -not $loadedRecord.Datum) { throw 'kein gültiger Datensatz' }
            if ($Category -eq 'Aktivitaet') { Assert-BtActivityDayRecord -Record $loadedRecord }
            if ($fromIso -and [string]$loadedRecord.Datum -lt $fromIso) { continue }
            if ($toIso -and [string]$loadedRecord.Datum -gt $toIso) { continue }
            [void]$records.Add($loadedRecord)
        } catch {
            [void]$script:MeasurementLoadFailedFiles.Add($file.FullName)
            Write-Warning "Messung nicht lesbar: $($file.FullName) - $($_.Exception.Message)"
        }
    }
    return , @($records | Sort-Object { [string]$_.Datum })
}

function Remove-MeasurementRecord {
    <# v2.33.0: Löscht den Datensatz einer Kategorie an einem Datum. $true = gelöscht. #>
    param([Parameter(Mandatory)][ValidateSet('DEXA', 'Schlaf', 'Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category, [Parameter(Mandatory)]$Date)
    $filePath = Get-MeasurementFilePath -Category $Category -Date $Date
    if (-not (Test-Path -LiteralPath $filePath)) { return $false }
    Remove-Item -LiteralPath $filePath -Force -ErrorAction Stop
    return $true
}

function Get-DexaDerivedValues {
    <#
    .SYNOPSIS
        v2.33.0: Berechnet fehlende DEXA-Kennzahlen aus vorhandenen Werten (werden nicht
        gespeichert): ALMI, LMI, FFMI, FMI (je / Größe²), A/G-Verhältnis, Körperfettanteil,
        fettfreie Masse. Bereits gemessene Werte werden nie überschrieben.
    #>
    param([Parameter(Mandatory)]$Record, [double]$HeightCm = 0)
    $measured = @{}
    $valueSource = $Record.Werte
    if ($valueSource -is [System.Collections.IDictionary]) { foreach ($k in $valueSource.Keys) { $measured[[string]$k] = [double]$valueSource[$k] } }
    else { foreach ($p in $valueSource.PSObject.Properties) { $measured[$p.Name] = [double]$p.Value } }
    $derived = [ordered]@{}
    $heightSquared = if ($HeightCm -gt 0) { [Math]::Pow($HeightCm / 100, 2) } else { 0 }
    if ($heightSquared -gt 0) {
        if ($measured.ContainsKey('ALM') -and -not $measured.ContainsKey('ALMI')) { $derived['ALMI'] = [Math]::Round($measured['ALM'] / $heightSquared, 2) }
        if ($measured.ContainsKey('Magermasse') -and -not $measured.ContainsKey('LMI')) { $derived['LMI'] = [Math]::Round($measured['Magermasse'] / $heightSquared, 2) }
        if ($measured.ContainsKey('Fettmasse') -and -not $measured.ContainsKey('FMI')) { $derived['FMI'] = [Math]::Round($measured['Fettmasse'] / $heightSquared, 2) }
    }
    $fatFree = $null
    if ($measured.ContainsKey('FettfreieMasse')) { $fatFree = $measured['FettfreieMasse'] }
    elseif ($measured.ContainsKey('Magermasse') -and $measured.ContainsKey('BMC_Gesamt')) { $fatFree = $measured['Magermasse'] + $measured['BMC_Gesamt']; $derived['FettfreieMasse'] = [Math]::Round($fatFree, 2) }
    if ($heightSquared -gt 0 -and $null -ne $fatFree -and -not $measured.ContainsKey('FFMI')) { $derived['FFMI'] = [Math]::Round($fatFree / $heightSquared, 2) }
    if ($measured.ContainsKey('Android_Fett_Prozent') -and $measured.ContainsKey('Gynoid_Fett_Prozent') -and $measured['Gynoid_Fett_Prozent'] -gt 0 -and -not $measured.ContainsKey('AG_Ratio')) {
        $derived['AG_Ratio'] = [Math]::Round($measured['Android_Fett_Prozent'] / $measured['Gynoid_Fett_Prozent'], 2)
    }
    if ($measured.ContainsKey('Fettmasse') -and $measured.ContainsKey('Gesamtmasse') -and $measured['Gesamtmasse'] -gt 0 -and -not $measured.ContainsKey('Koerperfett_Prozent')) {
        $derived['Koerperfett_Prozent'] = [Math]::Round(100 * $measured['Fettmasse'] / $measured['Gesamtmasse'], 1)
    }
    return $derived
}

function Get-SleepPhaseShares {
    <# v2.33.0: Anteile von Leicht-, Tief- und REM-Schlaf an der Summe der Schlafphasen (%). #>
    param([Parameter(Mandatory)]$Record)
    $phaseSource = $Record.PhasenMinuten
    if ($null -eq $phaseSource) { return [ordered]@{} }
    $phaseTable = @{}
    if ($phaseSource -is [System.Collections.IDictionary]) { foreach ($k in $phaseSource.Keys) { $phaseTable[[string]$k] = [double]$phaseSource[$k] } }
    else { foreach ($p in $phaseSource.PSObject.Properties) { $phaseTable[$p.Name] = [double]$p.Value } }
    $sleepSum = 0.0
    foreach ($phaseName in @('Leicht', 'Tief', 'REM')) { if ($phaseTable.ContainsKey($phaseName)) { $sleepSum += $phaseTable[$phaseName] } }
    $shares = [ordered]@{}
    if ($sleepSum -le 0) { return $shares }
    foreach ($phaseName in @('Leicht', 'Tief', 'REM')) { if ($phaseTable.ContainsKey($phaseName)) { $shares[$phaseName] = [Math]::Round(100 * $phaseTable[$phaseName] / $sleepSum, 1) } }
    return $shares
}

# =====================================================================
# v2.35.0: WEITERE MESSWERTE - Vitalwerte, Körper, EKG, Fitness (VO2max)
# ---------------------------------------------------------------------
# Nur Grundarchitektur (wie v2.33.0) - noch KEINE Oberfläche.
# Unterschied zu DEXA/Schlaf: MEHRERE Messungen je Tag, jede mit Uhrzeit.
#   Ablage: <data>\_messungen\<Kategorie>\<JJJJ>\<JJJJ-MM-TT>.json (verschlüsselt, atomar)
#   Inhalt: @{ SchemaVersion = 2; Kategorie; Datum; Messungen = @( Eintrag, ... ) }
#   Eintrag: @{ Kategorie; Datum; Zeit = 'HH:mm' ('' = unbekannt); Werte = @{ Schlüssel = Zahl };
#             Kontext; Details = @{ ... }; Quelle; Notiz; Hinweise; Erfasst }
#   Ablauf: New-VitalEntry / New-BodyEntry / New-EkgEntry / New-FitnessEntry (prüfen)
#           -> Add-MeasurementEntry (speichern) / Get-MeasurementEntries / Remove-MeasurementEntry
# Bewertungsbereiche nur, wo eine Leitlinie sie geschlechts- und altersunabhängig festlegt.
# Die Blutmarker "Gewicht" und "Systolischer Blutdruck" (Blutwert-Ablage, BMI-Marker, PREVENT)
# bleiben unverändert - eine Verbindung kommt erst mit der Oberfläche.
# Bewusst NICHT enthalten: Stress-Werte (Herstellerformeln, nicht validiert - Entscheidung 08.10.2026).
# =====================================================================
$script:MeasurementMultiCategories   = @('Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')
$script:MeasurementMultiSchemaVersion = 2
$script:VitalContexts          = @('Ruhe', 'Morgens nach dem Aufstehen', 'Abends', 'Nach Belastung', 'Während Belastung', 'Nachts/Schlaf', 'Sonstiges')
$script:TemperatureSites       = @('Oral', 'Rektal', 'Ohr', 'Stirn', 'Achsel')
$script:BloodPressureArms      = @('Links', 'Rechts')
$script:BodyPositions          = @('Sitzend', 'Liegend', 'Stehend')
$script:BodyCompositionMethods = @('Waage (BIA)', 'DEXA', 'Caliper', 'Sonstige')
$script:EkgRhythms             = @('Sinusrhythmus', 'Vorhofflimmern', 'Nicht eindeutig', 'Niedrige Herzfrequenz', 'Hohe Herzfrequenz', 'Schlechte Aufzeichnung', 'Sonstiges')
$script:EkgTypes               = @('Uhr (1 Kanal)', '6 Kanäle', '12 Kanäle', 'Langzeit-EKG', 'Belastungs-EKG')
$script:Vo2maxMethods          = @('Spiroergometrie (Labor)', 'Feldtest', 'Schätzung Uhr/App')

function Get-ExtendedMeasurementDefinitions {
    <#
    .SYNOPSIS
        v2.35.0: Katalog der Kategorien Vitalwerte, Koerper, EKG, Fitness (Aufruf über Get-MeasurementCatalog).
    .DESCRIPTION
        Quellen der Bewertungsbereiche: Blutdruck ESC 2024 (Praxismessung), BMI WHO,
        Herzfrequenz < 50 bpm = Sinusbradykardie sowie PQ > 200 ms und QRS 110-119 / ab 120 ms
        (ACC/AHA/HRS-Bradykardie-Leitlinie 2018), QTc-Grenzen AHA/ACCF/HRS 2009 (geschlechtsabhängig,
        daher nur Beschreibung). Ohne Leitlinienbereich: nur Plausibilitätsprüfung, keine Bewertung.
    #>
    param([Parameter(Mandatory)][ValidateSet('Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category)
    $defs = New-Object System.Collections.ArrayList
    switch ($Category) {
        'Vitalwerte' {
            [void]$defs.Add((New-MeasurementDefinition -Key 'Herzfrequenz' -Name 'Herzfrequenz' -Unit 'bpm' -Group 'Vitalwerte' -PlausMin 20 -PlausMax 250 -RefMin 50 -RefMax 100 -Description 'Bewertung nur für Messungen in Ruhe. Unter 50 bpm: Sinusbradykardie (ACC/AHA/HRS 2018) - bei Ausdauertrainierten oft ohne Krankheitswert; über 100 bpm: Tachykardie.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'Blutdruck_Sys' -Name 'Blutdruck systolisch' -Unit 'mmHg' -Group 'Vitalwerte' -PlausMin 50 -PlausMax 300 -RefMax 139 -OptimalMax 119 -Description 'ESC 2024 (Praxismessung): nicht erhöht < 120/70, erhöht 120-139 bzw. 70-89, Hypertonie ab 140/90 mmHg. Für Heim- und Langzeitmessungen nennt die Leitlinie niedrigere Grenzwerte. Immer zusammen mit dem diastolischen Wert.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'Blutdruck_Dia' -Name 'Blutdruck diastolisch' -Unit 'mmHg' -Group 'Vitalwerte' -PlausMin 20 -PlausMax 200 -RefMax 89 -OptimalMax 69 -Description 'ESC 2024 (Praxismessung): nicht erhöht < 70, erhöht 70-89, Hypertonie ab 90 mmHg. Immer zusammen mit dem systolischen Wert.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'Koerpertemperatur' -Name 'Körpertemperatur' -Unit '°C' -Group 'Vitalwerte' -PlausMin 30 -PlausMax 44 -Description 'Der Normalwert hängt vom Messort ab (rektal am höchsten, Achsel am niedrigsten) - Messort angeben. Ohne Bewertung.'))
        }
        'Koerper' {
            [void]$defs.Add((New-MeasurementDefinition -Key 'Gewicht' -Name 'Gewicht' -Unit 'kg' -Group 'Körper' -PlausMin 20 -PlausMax 350 -Description 'Waagengewicht. Unabhängig vom gleichnamigen Blutmarker (Blutwert-Ablage).'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'BMI' -Name 'BMI' -Unit 'kg/m²' -Group 'Körper' -PlausMin 10 -PlausMax 80 -RefMin 18.5 -RefMax 24.9 -Description 'WHO: 18,5-24,9 Normalgewicht. Bei viel Muskelmasse wenig aussagekräftig (Körperfettanteil bzw. FFMI beachten). Wird bei Bedarf aus Gewicht und Größe berechnet.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'Koerperfett_Prozent' -Name 'Körperfettanteil' -Unit '%' -Group 'Körper' -PlausMin 2 -PlausMax 75 -Description 'Messmethode angeben - BIA-Waagen weichen deutlich von DEXA ab. Ohne Bewertung (alters- und geschlechtsabhängig).'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'FettfreieMasse' -Name 'Fettfreie Masse' -Unit 'kg' -Group 'Körper' -PlausMin 5 -PlausMax 200 -Description 'Gewicht minus Fettmasse. Wird bei Bedarf aus Gewicht und Körperfettanteil berechnet.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'FFMI' -Name 'Fettfreie-Masse-Index (FFMI)' -Unit 'kg/m²' -Group 'Körper' -PlausMin 8 -PlausMax 40 -Description 'Fettfreie Masse / Größe². Wird bei Bedarf berechnet. Ohne Bewertung.'))
        }
        'EKG' {
            [void]$defs.Add((New-MeasurementDefinition -Key 'Herzfrequenz' -Name 'Herzfrequenz (EKG)' -Unit 'bpm' -Group 'EKG' -PlausMin 20 -PlausMax 300 -RefMin 50 -RefMax 100 -Description 'Unter 50 bpm: Sinusbradykardie (ACC/AHA/HRS 2018), über 100 bpm: Tachykardie.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'PQ' -Name 'PQ-Zeit (PR-Intervall)' -Unit 'ms' -Group 'EKG' -PlausMin 40 -PlausMax 600 -RefMin 120 -RefMax 200 -Description 'Üblicher Normbereich 120-200 ms. Über 200 ms: AV-Block I. Grades (ACC/AHA/HRS 2018); unter 120 ms: verkürzte Überleitung.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'QRS' -Name 'QRS-Dauer' -Unit 'ms' -Group 'EKG' -PlausMin 40 -PlausMax 300 -RefMax 109 -Description '110-119 ms: inkompletter, ab 120 ms: kompletter Schenkelblock (ACC/AHA/HRS 2018).'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'QT' -Name 'QT-Zeit' -Unit 'ms' -Group 'EKG' -PlausMin 150 -PlausMax 800 -Description 'Frequenzabhängig - bewertet wird die frequenzkorrigierte QTc.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'QTc' -Name 'QTc (Gerät/Befund)' -Unit 'ms' -Group 'EKG' -PlausMin 200 -PlausMax 800 -Description 'Verlängert ab 450 ms (Männer) bzw. 460 ms (Frauen), deutlich ab 500 ms (AHA/ACCF/HRS 2009). Geschlechtsabhängig - daher ohne festen Bereich.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'QTc_Bazett' -Name 'QTc nach Bazett' -Unit 'ms' -Group 'EKG' -PlausMin 200 -PlausMax 800 -Description 'QT / Wurzel(RR), RR in Sekunden. Wird aus QT und Herzfrequenz berechnet; überschätzt die QTc bei hoher Herzfrequenz.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'QTc_Fridericia' -Name 'QTc nach Fridericia' -Unit 'ms' -Group 'EKG' -PlausMin 200 -PlausMax 800 -Description 'QT / Kubikwurzel(RR), RR in Sekunden. Wird aus QT und Herzfrequenz berechnet; bei hoher Herzfrequenz genauer als Bazett.'))
        }
        'Aktivitaet' {
            # Nur technische Eingabegrenzen; 2^53-1 verhindert Genauigkeitsverlust bei Schrittzahlen.
            [void]$defs.Add((New-MeasurementDefinition -Key 'Aktivitaetsenergie' -Name 'Aktivitätsenergie' -Unit 'kcal' -Group 'Aktivität' -PlausMin 0 -PlausMax ([double]::MaxValue) -Description 'Konsolidierte aktive Energie eines Tages, ohne Ruheenergie. Keine Addition mehrerer Gerätesummen. Quellenbeschreibung: Apple HealthKit activeEnergyBurned, https://developer.apple.com/documentation/healthkit/hkquantitytypeidentifier/activeenergyburned . Keine medizinische Bewertung.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'Schritte' -Name 'Schritte' -Unit 'Schritte' -Group 'Aktivität' -PlausMin 0 -PlausMax 9007199254740991 -Description 'Konsolidierte ganzzahlige Schrittzahl eines Tages. Fehlende Tage sind keine Nullwerte. Quellenbeschreibung: Apple HealthKit stepCount, https://developer.apple.com/documentation/healthkit/hkquantitytypeidentifier/stepcount . Keine medizinische Bewertung.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'CardioErholung' -Name 'Cardio-Erholung (Tagesmittel)' -Unit 'bpm' -Group 'Aktivität' -PlausMin (-[double]::MaxValue) -PlausMax ([double]::MaxValue) -Description 'Übernommenes Tagesmittel der Herzfrequenz-Erholung nach Belastung. Erholungsintervall, Verfahren und bekannte Messungsanzahl separat erhalten. Apple beschreibt den Ein-Minuten-Typ in WWDC22: https://developer.apple.com/videos/play/wwdc2022/10005/ . Kein Intervall automatisch ergänzen; keine medizinische Bewertung.'))
        }
        'Fitness' {
            [void]$defs.Add((New-MeasurementDefinition -Key 'CardioErholung' -Name 'Cardio-Erholung' -Unit 'bpm' -Group 'Fitness' -PlausMin (-[double]::MaxValue) -PlausMax ([double]::MaxValue) -Description 'Quellenwert der Herzfrequenz-Erholung je Training. Intervall und Verfahren separat speichern; unbekannte oder unterschiedliche Intervalle nicht zusammenfassen. Apple beschreibt den Ein-Minuten-Typ in WWDC22: https://developer.apple.com/videos/play/wwdc2022/10005/ . Negative Quellwerte bleiben mit Hinweis erhalten. Keine medizinische Bewertung.'))
            [void]$defs.Add((New-MeasurementDefinition -Key 'VO2max' -Name 'VO2max' -Unit 'ml/kg/min' -Group 'Fitness' -PlausMin 10 -PlausMax 100 -Description 'Maximale Sauerstoffaufnahme. Referenzstandard ist die Spiroergometrie; Schätzungen von Uhren weichen ab - Methode angeben. Normwerte alters- und geschlechtsabhängig (ACSM), daher ohne festen Bereich. Höhere kardiorespiratorische Fitness geht mit geringerer Sterblichkeit einher (Mandsager 2018, JAMA Netw Open).'))
        }
    }
    return $defs.ToArray()
}

function ConvertTo-BtMeasurementTime {
    <#
    .SYNOPSIS
        v2.35.0: Uhrzeit -> 'HH:mm' ('' = unbekannt). Erlaubt [datetime], "7:05", "07:05", "7.05", "07:05:30".
    #>
    param($Time)
    if ($null -eq $Time) { return '' }
    if ($Time -is [datetime]) { return $Time.ToString('HH:mm') }
    $text = ([string]$Time).Trim()
    if ($text -eq '') { return '' }
    if ($text -match '^(\d{1,2})[:.](\d{2})(?::\d{2})?$') {
        $hours = [int]$Matches[1]; $minutes = [int]$Matches[2]
        if ($hours -le 23 -and $minutes -le 59) { return ('{0:00}:{1:00}' -f $hours, $minutes) }
    }
    throw "Ungültige Uhrzeit: '$Time' (erwartet HH:MM, z. B. 07:30)."
}

function Resolve-BtMeasurementOption {
    <# v2.35.0: Auswahlwert prüfen (Groß/Klein egal) -> Schreibweise der Liste; leer -> ''. #>
    param([string]$Value, [string[]]$Allowed, [string]$Label)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    foreach ($option in $Allowed) { if ($option -ieq $Value.Trim()) { return $option } }
    throw "Ungültige Angabe für $Label`: '$Value' (erlaubt: $($Allowed -join ', '))."
}

function ConvertTo-CheckedMeasurementValues {
    <#
    .SYNOPSIS
        v2.35.0: Prüft Werte gegen den Katalog einer Kategorie: Schlüssel oder Anzeigename -> Zahl
        im Plausibilitätsbereich (Text mit Komma erlaubt). Leere Felder werden ignoriert.
    .OUTPUTS
        Geordnete Hashtable Schlüssel -> [double].
    #>
    param([Parameter(Mandatory)][string]$Category, [hashtable]$Values)
    $checked = [ordered]@{}
    if (-not $Values) { return $checked }
    foreach ($inputKey in @($Values.Keys)) {
        $rawValue = $Values[$inputKey]
        if ($null -eq $rawValue -or ([string]$rawValue).Trim() -eq '') { continue }
        $definition = Resolve-MeasurementDefinition -Category $Category -KeyOrName ([string]$inputKey)
        if (-not $definition) { throw "Unbekannter Messwert '$inputKey' in Kategorie '$Category'." }
        if ($checked.Contains($definition.Key)) { throw "Messwert '$($definition.Name)' ist doppelt angegeben." }
        $numericValue = $null
        try { $numericValue = if ($rawValue -is [string]) { Parse-Number $rawValue } else { [double]$rawValue } } catch { $numericValue = $null }
        if ($null -eq $numericValue -or [double]::IsNaN($numericValue) -or [double]::IsInfinity($numericValue)) { throw "Ungültiger Zahlenwert für '$($definition.Name)': '$rawValue'." }
        if ($numericValue -lt $definition.PlausMin -or $numericValue -gt $definition.PlausMax) {
            throw "Unplausibler Wert für '$($definition.Name)': $numericValue $($definition.Unit) (erlaubt $($definition.PlausMin) bis $($definition.PlausMax))."
        }
        $checked[$definition.Key] = [double]$numericValue
    }
    return $checked
}

function New-MeasurementEntryObject {
    <# v2.35.0 (intern): baut einen geprüften Messungs-Eintrag; Datum/Uhrzeit nicht in der Zukunft. #>
    param([string]$Category, $Date, $Zeit, $Werte, [string]$Kontext, $Details, [string]$Quelle, [string]$Notiz, $Hinweise)
    $isoDate = ConvertTo-BtMeasurementDate -Date $Date
    $timeText = ConvertTo-BtMeasurementTime -Time $Zeit
    if ($timeText -and $isoDate -eq (Get-Date).ToString('yyyy-MM-dd')) {
        $entryMoment = [datetime]::ParseExact("$isoDate $timeText", 'yyyy-MM-dd HH:mm', [System.Globalization.CultureInfo]::InvariantCulture)
        if ($entryMoment -gt (Get-Date).AddMinutes(5)) { throw "Die Uhrzeit $timeText liegt in der Zukunft." }
    }
    $hintList = @(@($Hinweise) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    foreach ($hint in $hintList) { Write-Warning "$Category $isoDate $timeText`: $hint" }
    if ($null -eq $Details) { $Details = [ordered]@{} }
    return [ordered]@{
        Kategorie = $Category; Datum = $isoDate; Zeit = $timeText; Werte = $Werte; Kontext = $Kontext; Details = $Details
        Quelle = $Quelle; Notiz = $Notiz; Hinweise = $hintList; Erfasst = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')
    }
}

function New-VitalEntry {
    <#
    .SYNOPSIS
        v2.35.0: Prüft eine Vitalwert-Messung und erstellt einen Eintrag (speichert NICHT -> Add-MeasurementEntry).
    .PARAMETER Werte
        @{ Herzfrequenz = 58; Blutdruck_Sys = 124; Blutdruck_Dia = 79; Koerpertemperatur = '36,6' }
        (Schlüssel oder Anzeigename). Blutdruck nur paarweise, systolisch > diastolisch.
    .PARAMETER Kontext
        $script:VitalContexts (z. B. 'Ruhe', 'Morgens nach dem Aufstehen').
    .PARAMETER Arm / Koerperlage
        Nur mit Blutdruck: $script:BloodPressureArms / $script:BodyPositions.
    .PARAMETER Messort
        Nur mit Körpertemperatur: $script:TemperatureSites.
    #>
    param(
        [Parameter(Mandatory)]$Date, $Zeit = '', [Parameter(Mandatory)][hashtable]$Werte, [string]$Kontext = '',
        [string]$Arm = '', [string]$Koerperlage = '', [string]$Messort = '', [string]$Quelle = '', [string]$Notiz = ''
    )
    $checked = ConvertTo-CheckedMeasurementValues -Category 'Vitalwerte' -Values $Werte
    if ($checked.Count -eq 0) { throw 'Keine Vitalwerte angegeben.' }
    $hasSys = $checked.Contains('Blutdruck_Sys'); $hasDia = $checked.Contains('Blutdruck_Dia')
    if ($hasSys -ne $hasDia) { throw 'Blutdruck bitte immer mit systolischem UND diastolischem Wert angeben.' }
    if ($hasSys -and $checked['Blutdruck_Sys'] -le $checked['Blutdruck_Dia']) { throw "Der systolische Wert ($($checked['Blutdruck_Sys'])) muss über dem diastolischen ($($checked['Blutdruck_Dia'])) liegen." }
    $armValue = Resolve-BtMeasurementOption -Value $Arm -Allowed $script:BloodPressureArms -Label 'Arm'
    $positionValue = Resolve-BtMeasurementOption -Value $Koerperlage -Allowed $script:BodyPositions -Label 'Körperlage'
    $siteValue = Resolve-BtMeasurementOption -Value $Messort -Allowed $script:TemperatureSites -Label 'Messort'
    if (($armValue -or $positionValue) -and -not $hasSys) { throw 'Arm und Körperlage nur zusammen mit einem Blutdruckwert angeben.' }
    if ($siteValue -and -not $checked.Contains('Koerpertemperatur')) { throw 'Messort nur zusammen mit der Körpertemperatur angeben.' }
    $details = [ordered]@{}
    if ($armValue) { $details['Arm'] = $armValue }
    if ($positionValue) { $details['Koerperlage'] = $positionValue }
    if ($siteValue) { $details['Messort'] = $siteValue }
    $hints = @()
    if ($checked.Contains('Koerpertemperatur') -and -not $siteValue) { $hints += 'Körpertemperatur ohne Messort - der Normalwert hängt vom Messort ab.' }
    $contextValue = Resolve-BtMeasurementOption -Value $Kontext -Allowed $script:VitalContexts -Label 'Kontext'
    return (New-MeasurementEntryObject -Category 'Vitalwerte' -Date $Date -Zeit $Zeit -Werte $checked -Kontext $contextValue -Details $details -Quelle $Quelle -Notiz $Notiz -Hinweise $hints)
}

function New-BodyEntry {
    <#
    .SYNOPSIS
        v2.35.0: Prüft eine Körpermessung (Waage o. Ä.) und erstellt einen Eintrag (speichert NICHT).
    .PARAMETER Werte
        @{ Gewicht = 82.4; Koerperfett_Prozent = 14.8; FettfreieMasse = 70.2; BMI = 26.0 }
    .PARAMETER Methode
        Für Körperfett/fettfreie Masse: $script:BodyCompositionMethods.
    #>
    param(
        [Parameter(Mandatory)]$Date, $Zeit = '', [Parameter(Mandatory)][hashtable]$Werte, [string]$Methode = '',
        [string]$Kontext = '', [string]$Quelle = '', [string]$Notiz = ''
    )
    $checked = ConvertTo-CheckedMeasurementValues -Category 'Koerper' -Values $Werte
    if ($checked.Count -eq 0) { throw 'Keine Körperwerte angegeben.' }
    if ($checked.Contains('Gewicht') -and $checked.Contains('FettfreieMasse') -and $checked['FettfreieMasse'] -ge $checked['Gewicht']) {
        throw "Die fettfreie Masse ($($checked['FettfreieMasse']) kg) muss kleiner als das Gewicht ($($checked['Gewicht']) kg) sein."
    }
    $methodValue = Resolve-BtMeasurementOption -Value $Methode -Allowed $script:BodyCompositionMethods -Label 'Methode'
    $hints = @()
    $hasComposition = $checked.Contains('Koerperfett_Prozent') -or $checked.Contains('FettfreieMasse')
    if ($hasComposition -and -not $methodValue) { $hints += 'Messmethode fehlt - BIA-Waage und DEXA weichen deutlich voneinander ab.' }
    if ($checked.Contains('Gewicht') -and $checked.Contains('Koerperfett_Prozent') -and $checked.Contains('FettfreieMasse')) {
        $expectedFatFree = $checked['Gewicht'] * (1 - $checked['Koerperfett_Prozent'] / 100)
        if ([Math]::Abs($expectedFatFree - $checked['FettfreieMasse']) -gt 2) { $hints += "Fettfreie Masse passt nicht zu Gewicht und Körperfettanteil (erwartet ca. $([Math]::Round($expectedFatFree, 1)) kg)." }
    }
    $details = [ordered]@{}
    if ($methodValue) { $details['Methode'] = $methodValue }
    $contextValue = Resolve-BtMeasurementOption -Value $Kontext -Allowed $script:VitalContexts -Label 'Kontext'
    return (New-MeasurementEntryObject -Category 'Koerper' -Date $Date -Zeit $Zeit -Werte $checked -Kontext $contextValue -Details $details -Quelle $Quelle -Notiz $Notiz -Hinweise $hints)
}

function New-EkgEntry {
    <#
    .SYNOPSIS
        v2.35.0: Prüft ein EKG und erstellt einen Eintrag (speichert NICHT). Mindestens Rhythmus,
        ein Messwert, Befund oder Dokument. Rohdaten der Kurve werden nicht gespeichert.
    .PARAMETER Werte
        Optional: @{ Herzfrequenz = 62; PQ = 160; QRS = 96; QT = 392 } (QTc wird bei Bedarf berechnet,
        Get-EkgDerivedValues).
    .PARAMETER Rhythmus
        $script:EkgRhythms (Einstufungen wie bei Uhr-EKGs, z. B. 'Sinusrhythmus', 'Vorhofflimmern').
    .PARAMETER Typ
        $script:EkgTypes (z. B. 'Uhr (1 Kanal)', '12 Kanäle').
    .PARAMETER Dokument
        Relativer Verweis auf eine Datei im Datenordner (z. B. PDF des EKGs) - kein absoluter Pfad.
    #>
    param(
        [Parameter(Mandatory)]$Date, $Zeit = '', [hashtable]$Werte = @{}, [string]$Rhythmus = '', [string]$Typ = '',
        [string]$Befund = '', [string]$Dokument = '', [string]$Kontext = '', [string]$Quelle = '', [string]$Notiz = ''
    )
    $checked = ConvertTo-CheckedMeasurementValues -Category 'EKG' -Values $Werte
    $rhythmValue = Resolve-BtMeasurementOption -Value $Rhythmus -Allowed $script:EkgRhythms -Label 'Rhythmus'
    $typeValue = Resolve-BtMeasurementOption -Value $Typ -Allowed $script:EkgTypes -Label 'EKG-Typ'
    $documentValue = ([string]$Dokument).Trim()
    if ($documentValue) {
        if ([System.IO.Path]::IsPathRooted($documentValue) -or $documentValue -match '^([A-Za-z]:|[\\/])' -or $documentValue -match '(^|[\\/])\.\.([\\/]|$)') { throw "Dokument bitte als relativen Verweis im Datenordner angeben (kein absoluter Pfad, kein '..'): '$documentValue'." }
        if ($documentValue.IndexOfAny([System.IO.Path]::GetInvalidPathChars()) -ge 0) { throw "Ungültiger Dokumentverweis: '$documentValue'." }
    }
    if ($checked.Count -eq 0 -and -not $rhythmValue -and [string]::IsNullOrWhiteSpace($Befund) -and -not $documentValue) {
        throw 'EKG ohne Inhalt: mindestens Rhythmus, einen Messwert, einen Befund oder ein Dokument angeben.'
    }
    $details = [ordered]@{}
    if ($rhythmValue) { $details['Rhythmus'] = $rhythmValue }
    if ($typeValue) { $details['Typ'] = $typeValue }
    if (-not [string]::IsNullOrWhiteSpace($Befund)) { $details['Befund'] = $Befund.Trim() }
    if ($documentValue) { $details['Dokument'] = $documentValue }
    $contextValue = Resolve-BtMeasurementOption -Value $Kontext -Allowed $script:VitalContexts -Label 'Kontext'
    return (New-MeasurementEntryObject -Category 'EKG' -Date $Date -Zeit $Zeit -Werte $checked -Kontext $contextValue -Details $details -Quelle $Quelle -Notiz $Notiz -Hinweise @())
}

function Get-BtActivityHints {
    # Technische Hinweise, keine medizinische Einstufung. Quellwerte unveraendert lassen.
    param($Values)
    $hints = @()
    foreach ($key in @('Aktivitaetsenergie', 'Schritte', 'CardioErholung')) {
        if (-not $Values.Contains($key)) { continue }
        if ($key -eq 'CardioErholung' -and $Values[$key] -lt 0) { $hints += 'Negative Cardio-Erholung aus der Quelle übernommen. Vorzeichen, Messverfahren und Erholungsintervall prüfen; keine automatische Korrektur.' }
        $limit = $script:ActivityWarningUpperBounds[$key]
        if ($null -eq $limit) { continue }
        $limitNumber = $null
        try { $limitNumber = if ($limit -is [string]) { Parse-Number $limit } else { [double]$limit } } catch { $limitNumber = $null }
        if ($null -eq $limitNumber -or [double]::IsNaN($limitNumber) -or [double]::IsInfinity($limitNumber) -or $limitNumber -lt 0) { throw 'Eine technische Aktivitäts-Warnschwelle in USER-KONFIGURATION ist ungültig.' }
        if ($Values[$key] -gt $limitNumber) { $hints += "Technische Warnschwelle für $key überschritten. Quellwert bleibt unverändert; dies ist keine medizinische Bewertung." }
    }
    return $hints
}

function New-BtCardioRecoveryDetails {
    # $null bleibt unbekannt. Insbesondere Quelle/Geraet niemals als Intervall-Nachweis verwenden.
    param([ValidateSet('Training', 'Tagesmittel')][string]$Datenart, $ErholungsintervallSekunden = $null, $AnzahlMessungen = $null, [string]$Methode = '')
    $interval = $null; $count = $null
    foreach ($field in @('ErholungsintervallSekunden', 'AnzahlMessungen')) {
        $raw = if ($field -eq 'ErholungsintervallSekunden') { $ErholungsintervallSekunden } else { $AnzahlMessungen }
        if ($null -eq $raw -or ($raw -is [string] -and [string]::IsNullOrWhiteSpace($raw))) { continue }
        $number = $null
        try { if ($raw -is [bool]) { throw 'bool' }; $number = if ($raw -is [string]) { Parse-Number $raw } else { [double]$raw } } catch { $number = $null }
        if ($null -eq $number -or [double]::IsNaN($number) -or [double]::IsInfinity($number) -or $number -le 0 -or $number -gt 9007199254740991) { throw "Ungültige Angabe für $field (positive endliche Zahl oder unbekannt erforderlich)." }
        if ($field -eq 'AnzahlMessungen') {
            if ($number -ne [Math]::Truncate($number)) { throw 'AnzahlMessungen muss ganzzahlig sein.' }
            $count = [long]$number
        } else { $interval = [double]$number }
    }
    if ($Datenart -eq 'Training') { $count = 1 }
    return [ordered]@{ Datenart = $Datenart; ErholungsintervallSekunden = $interval; AnzahlMessungen = $count; Methode = $Methode.Trim() }
}

function New-ActivityEntry {
    <#
    .SYNOPSIS
        v2.41.0: Konsolidierte Tageswerte; nur erzeugen, speichern mit Add-MeasurementEntry.
    .DESCRIPTION
        Werte: Aktivitaetsenergie (kcal, aktive Energie ohne Ruheenergie), Schritte (ganze Anzahl),
        optional CardioErholung (bpm, ausdruecklich uebernommenes Tagesmittel).
        Keine Aufteilung oder Addition von Quellenwerten. Ein gemeinsamer Eintrag je Tag,
        ohne Uhrzeit. -Replace beim Speichern ersetzt den GESAMTEN Tages-Eintrag.
        Quelle/Geraet sind optionale Herkunftsangaben, keine Bindung an einen Hersteller.
        Vergangene Tage gelten standardmaessig als abgeschlossen, heute als unvollstaendig;
        -TagesAbgeschlossen $false kennzeichnet auch unvollstaendige vergangene Tage.
        Fehlende Werte bleiben fehlend, explizite Nullwerte bleiben erhalten.
    #>
    param(
        [Parameter(Mandatory)]$Date, [Parameter(Mandatory)][hashtable]$Werte,
        $TagesAbgeschlossen = $null, [string]$Quelle = '', [string]$Geraet = '', [string]$Notiz = '',
        $ErholungsintervallSekunden = $null, $AnzahlCardioMessungen = $null, [string]$ErholungsMethode = ''
    )
    foreach ($raw in $Werte.Values) { if ($raw -is [bool]) { throw 'Aktivitätswerte müssen Zahlen sein, keine Wahrheitswerte.' } }
    $checked = ConvertTo-CheckedMeasurementValues -Category 'Aktivitaet' -Values $Werte
    if ($checked.Count -eq 0) { throw 'Keine Aktivitätswerte angegeben.' }
    if ($checked.Contains('Schritte') -and $checked['Schritte'] -ne [Math]::Truncate($checked['Schritte'])) { throw 'Schritte als Tagesgesamtwert müssen ganzzahlig sein.' }
    $isoDate = ConvertTo-BtMeasurementDate -Date $Date
    $complete = ($isoDate -lt (Get-Date).ToString('yyyy-MM-dd'))
    if ($null -ne $TagesAbgeschlossen) {
        if ($TagesAbgeschlossen -isnot [bool]) { throw 'TagesAbgeschlossen muss $true oder $false sein.' }
        $complete = $TagesAbgeschlossen
    }
    $details = [ordered]@{ Datenart = 'Tageswerte'; Herkunft = 'Quelle'; TagesAbgeschlossen = [bool]$complete; Geraet = $Geraet.Trim() }
    $hints = @(Get-BtActivityHints -Values $checked)
    if (-not $complete) { $hints += 'Unvollständiger Tag: für spätere Tagesdurchschnitte nicht als vollständigen Tageswert verwenden.' }
    if ($checked.Contains('CardioErholung')) {
        $cardio = New-BtCardioRecoveryDetails -Datenart 'Tagesmittel' -ErholungsintervallSekunden $ErholungsintervallSekunden -AnzahlMessungen $AnzahlCardioMessungen -Methode $ErholungsMethode
        $details['CardioErholung'] = $cardio
        if ($null -eq $cardio.ErholungsintervallSekunden) { $hints += 'Erholungsintervall unbekannt: Cardio-Erholung nicht automatisch mit anderen Messungen zusammenfassen.' }
    } elseif ($null -ne $ErholungsintervallSekunden -or $null -ne $AnzahlCardioMessungen -or $ErholungsMethode) { throw 'Angaben zur Cardio-Erholung benötigen einen Cardio-Erholungswert.' }
    return (New-MeasurementEntryObject -Category 'Aktivitaet' -Date $isoDate -Zeit '' -Werte $checked -Kontext 'Tag' -Details $details -Quelle $Quelle.Trim() -Notiz $Notiz -Hinweise $hints)
}

# =====================================================================
# v2.42.0: Uebernommene Zeitraum-Mittel (Engine, ohne GUI/Aggregation)
# Get-ActivityPeriodRecords liefert wie Get-MeasurementEntries per Komma-Return.
# Ablage getrennt von Tagesdateien: _messungen\_zeitraeume\Aktivitaet.json.
# =====================================================================
function New-ActivityPeriodRecord {
    <#
    .SYNOPSIS
        Ein uebernommenes Zeitraum-Mittel vorbereiten; speichert NICHT.
    .DESCRIPTION
        Je Eintrag genau eine Messgroesse. From/To inklusive, keine Zukunft.
        Energie/Schritte = Mittel pro Tag, Schritte duerfen hier Nachkommastellen haben.
        Cardio: Basis Trainingswerte, Tagesmittel oder Unbekannt; niemals automatisch mischen.
        Anzahl = bekannte beruecksichtigte Tage bzw. Trainingsmessungen; null bleibt unbekannt.
        Quelle/Geraet/Methode sind frei waehlbar. Kein Import und keine Berechnung.
        Id/Erfasst erlauben Bearbeitung vorhandener Eintraege ohne neue Identitaet.
    #>
    param(
        [Parameter(Mandatory)]$From, [Parameter(Mandatory)]$To,
        [Parameter(Mandatory)][string]$ValueKey, [Parameter(Mandatory)]$Value,
        [string]$Basis = '', $Anzahl = $null, $ErholungsintervallSekunden = $null,
        [string]$Quelle = '', [string]$Geraet = '', [string]$Methode = '', [string]$Notiz = '',
        [string]$Id = '', [string]$Erfasst = ''
    )
    $fromIso = ConvertTo-BtMeasurementDate -Date $From
    $toIso = ConvertTo-BtMeasurementDate -Date $To
    if ($fromIso -gt $toIso) { throw 'Der Beginn des Zeitraums muss vor oder auf seinem Ende liegen.' }
    $definition = Resolve-MeasurementDefinition -Category 'Aktivitaet' -KeyOrName $ValueKey
    if (-not $definition) { throw 'Unbekannte Messgröße für ein Aktivitäts-Zeitraummittel.' }
    if ($null -eq $Value -or $Value -is [bool] -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { throw 'Ein Zeitraum-Mittel benötigt einen Zahlenwert; fehlend ist nicht null.' }
    $values = ConvertTo-CheckedMeasurementValues -Category 'Aktivitaet' -Values @{ $definition.Key = $Value }
    if ($values.Count -ne 1) { throw 'Ein Zeitraum-Mittel benötigt genau einen Zahlenwert.' }
    $isCardio = ($definition.Key -eq 'CardioErholung')
    $basisValue = $Basis.Trim()
    if (-not $basisValue) { $basisValue = if ($isCardio) { 'Unbekannt' } else { 'Tageswerte' } }
    $allowed = if ($isCardio) { @('Trainingswerte', 'Tagesmittel', 'Unbekannt') } else { @('Tageswerte') }
    $basisValue = Resolve-BtMeasurementOption -Value $basisValue -Allowed $allowed -Label 'Berechnungsbasis'
    if (-not $isCardio -and $null -ne $ErholungsintervallSekunden) { throw 'Ein Erholungsintervall ist nur für Cardio-Erholung zulässig.' }
    # Wiederverwendete Zahlenpruefung; Datenart des Ergebnisses wird hier nicht uebernommen.
    $details = New-BtCardioRecoveryDetails -Datenart 'Tagesmittel' -ErholungsintervallSekunden $ErholungsintervallSekunden -AnzahlMessungen $Anzahl
    if ($basisValue -eq 'Unbekannt' -and $null -ne $details.AnzahlMessungen) { throw 'Bei bekannter Anzahl bitte auch die Berechnungsbasis angeben.' }
    $calendarDays = ([datetime]::ParseExact($toIso, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture) - [datetime]::ParseExact($fromIso, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)).Days + 1
    if ($basisValue -in @('Tageswerte', 'Tagesmittel') -and $null -ne $details.AnzahlMessungen -and $details.AnzahlMessungen -gt $calendarDays) { throw 'Die Anzahl berücksichtigter Tage übersteigt die Länge des Zeitraums.' }
    $parsedId = [guid]::Empty
    if (-not $Id) { $Id = [guid]::NewGuid().ToString('D') }
    if (-not [guid]::TryParseExact($Id, 'D', [ref]$parsedId) -or $parsedId -eq [guid]::Empty) { throw 'Ungültige Kennung des Zeitraum-Eintrags.' }
    $parsedTime = [datetime]::MinValue
    if (-not $Erfasst) { $Erfasst = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss') }
    if (-not [datetime]::TryParseExact($Erfasst, 'yyyy-MM-ddTHH:mm:ss', [cultureinfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsedTime)) { throw 'Ungültiger Erfassungszeitpunkt des Zeitraum-Eintrags.' }
    $hints = @(Get-BtActivityHints -Values $values)
    if ($null -eq $details.AnzahlMessungen) { $hints += 'Anzahl berücksichtigter Tage oder Messungen unbekannt; keine vollständige Datenabdeckung ableiten.' }
    if ($isCardio -and $basisValue -eq 'Unbekannt') { $hints += 'Berechnungsbasis unbekannt: nicht automatisch mit Trainingswerten oder Tagesmitteln vergleichen.' }
    if ($isCardio -and $null -eq $details.ErholungsintervallSekunden) { $hints += 'Erholungsintervall unbekannt: Cardio-Erholung nicht automatisch mit anderen Messungen zusammenfassen.' }
    return [ordered]@{
        Id = $parsedId.ToString('D'); Datenart = 'Zeitraummittel'; Herkunft = 'Quelle'
        Von = $fromIso; Bis = $toIso; Messgroesse = $definition.Key; Wert = $values[$definition.Key]; Einheit = $definition.Unit
        Basis = $basisValue; Anzahl = $details.AnzahlMessungen; ErholungsintervallSekunden = $details.ErholungsintervallSekunden
        Quelle = $Quelle.Trim(); Geraet = $Geraet.Trim(); Methode = $Methode.Trim(); Notiz = $Notiz
        Hinweise = @($hints); Erfasst = $Erfasst
    }
}

function ConvertTo-BtActivityPeriodRecord {
    # Validiert vollstaendig vor Schreiben. Unbekannte Felder niemals stillschweigend verlieren.
    param([Parameter(Mandatory)]$Record)
    $required = @('Id','Datenart','Herkunft','Von','Bis','Messgroesse','Wert','Einheit','Basis','Anzahl','ErholungsintervallSekunden','Quelle','Geraet','Methode','Notiz','Hinweise','Erfasst')
    if ($Record -is [System.Collections.IDictionary]) { $keys = @($Record.Keys) }
    elseif ($Record -is [System.Management.Automation.PSCustomObject]) { $keys = @($Record.PSObject.Properties.Name) }
    else { throw 'Ungültiger Zeitraum-Eintrag.' }
    if ($keys.Count -ne $required.Count -or @($required | Where-Object { $keys -notcontains $_ }).Count -gt 0) { throw 'Unbekanntes oder unvollständiges Zeitraum-Format; keine Änderung vorgenommen.' }
    foreach ($field in @('Id','Datenart','Herkunft','Von','Bis','Messgroesse','Einheit','Basis','Quelle','Geraet','Methode','Notiz','Erfasst')) {
        if ($Record.$field -isnot [string]) { throw 'Ungültiges Textfeld im Zeitraum-Eintrag.' }
    }
    if (-not $Record.Id -or -not $Record.Erfasst -or -not $Record.Basis -or $Record.Datenart -cne 'Zeitraummittel' -or $Record.Herkunft -cne 'Quelle') { throw 'Unpassende Datenart oder Herkunft des Zeitraum-Eintrags.' }
    if ($Record.Hinweise -isnot [System.Collections.IList] -or @($Record.Hinweise | Where-Object { $_ -isnot [string] }).Count -gt 0) { throw 'Ungültige Hinweisliste im Zeitraum-Eintrag.' }
    $normalized = New-ActivityPeriodRecord -From $Record.Von -To $Record.Bis -ValueKey $Record.Messgroesse -Value $Record.Wert -Basis $Record.Basis -Anzahl $Record.Anzahl -ErholungsintervallSekunden $Record.ErholungsintervallSekunden -Quelle $Record.Quelle -Geraet $Record.Geraet -Methode $Record.Methode -Notiz $Record.Notiz -Id $Record.Id -Erfasst $Record.Erfasst
    if ($normalized.Einheit -cne $Record.Einheit -or $normalized.Messgroesse -cne $Record.Messgroesse -or $normalized.Von -cne $Record.Von -or $normalized.Bis -cne $Record.Bis) { throw 'Unpassende Einheit, Messgröße oder Datumsform im Zeitraum-Eintrag.' }
    # Historische Hinweise erhalten; aktuelle Warnschwellen aendern keinen gespeicherten Inhalt.
    $normalized.Hinweise = @($Record.Hinweise)
    return $normalized
}

function Get-BtActivityPeriodPath {
    return (Join-Path (Join-Path (Get-MeasurementRoot) '_zeitraeume') 'Aktivitaet.json')
}

function Get-BtActivityPeriodIdentity {
    param($Record)
    # Feste Feldreihenfolge, JSON-Escaping statt mehrdeutiger Trennzeichen. Keine Wert-/Anzahlanteile.
    return (ConvertTo-Json -InputObject @($Record.Von, $Record.Bis, $Record.Messgroesse, $Record.Basis, $Record.ErholungsintervallSekunden, $Record.Quelle.ToLowerInvariant(), $Record.Geraet.ToLowerInvariant(), $Record.Methode.ToLowerInvariant()) -Compress)
}

function Read-BtActivityPeriodStore {
    # Fehler blockieren die gesamte Ablage statt still Eintraege beim naechsten Speichern zu verlieren.
    $path = Get-BtActivityPeriodPath
    if (-not (Test-Path -LiteralPath $path)) { return , @() }
    try {
        $json = Read-ProtectedJsonFile -Path $path 3>$null
        if ([string]::IsNullOrWhiteSpace($json)) { throw 'Leer' }
        $store = $json | ConvertFrom-Json -ErrorAction Stop
        $keys = @($store.PSObject.Properties.Name)
        if ($keys.Count -ne 3 -or @(@('SchemaVersion','Kategorie','Eintraege') | Where-Object { $keys -notcontains $_ }).Count -gt 0 -or $store.SchemaVersion -ne 1 -or $store.Kategorie -cne 'AktivitaetZeitraeume' -or $store.Eintraege -isnot [System.Collections.IList]) { throw 'Format' }
        $records = @(); $seenIds = @{}; $seenIdentities = @{}
        foreach ($entry in @($store.Eintraege)) {
            $record = ConvertTo-BtActivityPeriodRecord -Record $entry
            $identity = Get-BtActivityPeriodIdentity -Record $record
            if ($seenIds.ContainsKey($record.Id) -or $seenIdentities.ContainsKey($identity)) { throw 'Duplikat' }
            $seenIds[$record.Id] = $true; $seenIdentities[$identity] = $true
            $records += [PSCustomObject]$record
        }
        return , $records
    } catch { throw 'Die Zeitraum-Ablage ist nicht lesbar oder hat ein unbekanntes/ungültiges Format. Sie bleibt unverändert.' }
}

function Get-ActivityPeriodRecords {
    <#
    .SYNOPSIS
        Quellenmittel laden, optional nach Messgroesse und Zeitraum filtern; keine Berechnung.
    .DESCRIPTION
        Standardfilter: alle ueberlappenden Eintraege, immer mit ihrem GANZEN Originalzeitraum.
        -ExactPeriod benoetigt From UND To und liefert nur genau diesen Zeitraum.
        Array per Komma-Return; Aufrufer NICHT zusaetzlich in @() einschliessen.
    #>
    param($From = $null, $To = $null, [string]$ValueKey = '', [switch]$ExactPeriod)
    $fromIso = if ($null -ne $From) { ConvertTo-BtMeasurementDate -Date $From } else { '' }
    $toIso = if ($null -ne $To) { ConvertTo-BtMeasurementDate -Date $To } else { '' }
    if ($fromIso -and $toIso -and $fromIso -gt $toIso) { throw 'Ungültiger Suchzeitraum.' }
    if ($ExactPeriod -and (-not $fromIso -or -not $toIso)) { throw 'ExactPeriod benötigt Anfang und Ende des Zeitraums.' }
    $key = ''
    if ($ValueKey) {
        $definition = Resolve-MeasurementDefinition -Category 'Aktivitaet' -KeyOrName $ValueKey
        if (-not $definition) { throw 'Unbekannte Messgröße für die Zeitraum-Suche.' }
        $key = $definition.Key
    }
    $records = Read-BtActivityPeriodStore
    $selected = @($records | Where-Object {
        (-not $key -or $_.Messgroesse -eq $key) -and
        (-not $fromIso -or $_.Bis -ge $fromIso) -and (-not $toIso -or $_.Von -le $toIso) -and
        (-not $ExactPeriod -or ($_.Von -eq $fromIso -and $_.Bis -eq $toIso))
    } | Sort-Object Von, Bis, Messgroesse, Id)
    return , $selected
}

function Save-ActivityPeriodRecord {
    <# Neuer Eintrag; vorhandene Id nur mit -Replace. Keine Addition, keine automatische Ueberschreibung. #>
    param([Parameter(Mandatory)]$Record, [switch]$Replace)
    $ErrorActionPreference = 'Stop'
    $candidate = ConvertTo-BtActivityPeriodRecord -Record $Record
    $path = Get-BtActivityPeriodPath; $directory = Split-Path -Parent $path
    $null = [System.IO.Directory]::CreateDirectory($directory)
    $lock = $null
    try {
        $lock = [System.IO.File]::Open(($path + '.lock'), [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $records = Read-BtActivityPeriodStore
        $existing = @($records | Where-Object { $_.Id -eq $candidate.Id })
        if ($existing.Count -gt 0 -and -not $Replace) { throw 'Vorhandener Eintrag' }
        if ($Replace -and $existing.Count -ne 1) { throw 'Ersetzen ohne Ziel' }
        $identity = Get-BtActivityPeriodIdentity -Record $candidate
        if (@($records | Where-Object { $_.Id -ne $candidate.Id -and (Get-BtActivityPeriodIdentity -Record $_) -ceq $identity }).Count -gt 0) { throw 'Doppelte Quelle fuer diesen Zeitraum' }
        $next = @(); $replaced = $false
        foreach ($entry in $records) {
            if ($entry.Id -eq $candidate.Id) { $next += $candidate; $replaced = $true } else { $next += $entry }
        }
        if (-not $replaced) { $next += $candidate }
        $store = [ordered]@{ SchemaVersion = 1; Kategorie = 'AktivitaetZeitraeume'; Eintraege = @($next) }
        Write-ProtectedJsonFile -Path $path -JsonString ($store | ConvertTo-Json -Depth 8) 3>$null
    } catch { throw 'Zeitraum-Mittel nicht gespeichert. Ablage/Schreibzugriff prüfen; doppelte Quellenwerte sind gesperrt und Ersetzen benötigt eine vorhandene Kennung mit -Replace.' }
    finally { if ($lock) { $lock.Dispose() } }
    return $path
}

function Remove-ActivityPeriodRecord {
    # Genau eine Kennung entfernen. Eine leere Liste bleibt explizit als [] gespeichert.
    param([Parameter(Mandatory)][string]$Id)
    $parsedId = [guid]::Empty
    if (-not [guid]::TryParseExact($Id, 'D', [ref]$parsedId) -or $parsedId -eq [guid]::Empty) { throw 'Ungültige Kennung des Zeitraum-Eintrags.' }
    $path = Get-BtActivityPeriodPath
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    $lock = $null; $ErrorActionPreference = 'Stop'
    try {
        $lock = [System.IO.File]::Open(($path + '.lock'), [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $records = Read-BtActivityPeriodStore
        $kept = @($records | Where-Object { $_.Id -ne $parsedId.ToString('D') })
        if ($kept.Count -eq $records.Count) { return $false }
        $store = [ordered]@{ SchemaVersion = 1; Kategorie = 'AktivitaetZeitraeume'; Eintraege = @($kept) }
        Write-ProtectedJsonFile -Path $path -JsonString ($store | ConvertTo-Json -Depth 8) 3>$null
        return $true
    } catch { throw 'Zeitraum-Mittel nicht entfernt. Ablage und Schreibzugriff prüfen; unlesbare oder unbekannte Formate werden nicht überschrieben.' }
    finally { if ($lock) { $lock.Dispose() } }
}


# =====================================================================
# v2.45.0: YTD-/Sechsmonats-Auswertung (nur lesen/rechnen, keine GUI)
# =====================================================================
function Get-ActivitySummaryWindow {
    <# Kalendergrenzen inklusive. Sechs Monate sind kein festes 180-Tage-Fenster. #>
    param([ValidateSet('YTD', 'Last6Months')][string]$Period = 'YTD', $AsOf = (Get-Date))
    $iso = ConvertTo-BtMeasurementDate -Date $AsOf
    $end = [datetime]::ParseExact($iso, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)
    $start = if ($Period -eq 'YTD') { [datetime]::new($end.Year, 1, 1) } else { $end.AddMonths(-6).AddDays(1) }
    return [PSCustomObject]@{ Zeitraum = $Period; Von = $start.ToString('yyyy-MM-dd'); Bis = $iso; Kalendertage = ($end - $start).Days + 1 }
}

function Get-BtActivityAggregate {
    # Eingaben sind bereits gepruefte, vergleichbare Einzelwerte, niemals Zeitraum-Mittel.
    param([object[]]$Rows, [int]$Digits, [int]$CalendarDays, [string]$To)
    $rowsSorted = @($Rows | Sort-Object Datum, Zeit)
    if ($rowsSorted.Count -eq 0) { return $null }
    # Normierte Summation verhindert Ueberlauf auch bei grossen endlichen Quellwerten.
    $scale = 0.0
    foreach ($row in $rowsSorted) { $scale = [Math]::Max($scale, [Math]::Abs([double]$row.Wert)) }
    $sum = 0.0
    if ($scale -gt 0) { foreach ($row in $rowsSorted) { $sum += ([double]$row.Wert / $scale) } }
    $mean = if ($scale -eq 0) { 0.0 } else { $scale * [Math]::Max(-1.0, [Math]::Min(1.0, $sum / $rowsSorted.Count)) }
    $dates = @($rowsSorted | Select-Object -ExpandProperty Datum -Unique)
    $first = $rowsSorted[0]; $last = $rowsSorted[-1]
    $daysSince = ([datetime]::ParseExact($To, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture) - [datetime]::ParseExact($last.Datum, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture)).Days
    $hints = @()
    if ($first.Basis -eq 'Trainingswerte') {
        $hints += 'Mittel pro erfasster Trainingsmessung. Tage ohne Training sind keine Nullwerte; die Vollständigkeit aller Trainings ist unbekannt.'
    } else {
        $hints += 'Gleichgewichtetes Mittel der vorhandenen abgeschlossenen Tage. Fehlende Tage werden nicht ergänzt.'
        if ($first.Basis -eq 'Tagesmittel') { $hints += 'Mittel der Tagesmittel; kein nach Trainingsanzahl gewichteter Trainingsdurchschnitt.' }
    }
    if ($daysSince -gt 0) { $hints += "Letzter verwendeter Wert liegt $daysSince Tage vor dem Stichtag." }
    if (@($rowsSorted | Where-Object { $_.Wert -lt 0 }).Count -gt 0) { $hints += 'Negative Cardio-Erholungswerte wurden unverändert berücksichtigt; Vorzeichen und Quellenverfahren prüfen.' }
    return [PSCustomObject]@{
        Basis = $first.Basis; ErholungsintervallSekunden = $first.ErholungsintervallSekunden; Methode = $first.Methode
        Quelle = $(if (@($rowsSorted | Select-Object -ExpandProperty Quelle -Unique).Count -eq 1) { $first.Quelle } else { '' })
        Geraet = $(if (@($rowsSorted | Select-Object -ExpandProperty Geraet -Unique).Count -eq 1) { $first.Geraet } else { '' })
        Quellen = @($rowsSorted | Select-Object -ExpandProperty Quelle -Unique | Sort-Object)
        Geraete = @($rowsSorted | Select-Object -ExpandProperty Geraet -Unique | Sort-Object)
        Mittelwert = [double]$mean; AnzeigeMittelwert = [Math]::Round($mean, $Digits, [MidpointRounding]::AwayFromZero)
        Anzahl = $rowsSorted.Count; Nenner = $(if ($first.Basis -eq 'Trainingswerte') { 'Trainingsmessungen' } else { 'Tage' })
        TageMitWert = $dates.Count; Kalendertage = $CalendarDays
        AbdeckungProzent = $(if ($first.Basis -eq 'Trainingswerte') { $null } else { 100.0 * $dates.Count / $CalendarDays })
        ErsterWertAm = $first.Datum; LetzterWertAm = $last.Datum; TageSeitLetztemWert = $daysSince
        Hinweise = @($hints)
    }
}

function Measure-ActivitySummary {
    <#
    .SYNOPSIS
        Reine Berechnung aus uebergebenen Tages-/Trainingswerten und Quellenmitteln.
    .DESCRIPTION
        Kein Dateizugriff und keine Speicherung. Ergebnis = ein Objekt mit Fenster und
        Ergebnissen fuer Energie, Schritte, Cardio. Berechnet und Quellenmittel bleiben
        getrennte Arrays. BevorzugteDatenart gibt exakten Quellenmitteln Vorrang, waehlt
        aber bei mehreren Quellen NICHT willkuerlich einen Wert aus.
        Tageswerte: vorhandene abgeschlossene Tage gleich gewichten. Training und
        Cardio-Tagesmittel bleiben getrennt; gruppieren nach bekanntem Intervall,
        bekanntem Verfahren, Quelle und Geraet. Unbekanntes nicht als Gleichheit werten.
        Quellenmittel nur bei exakt passendem Fenster vergleichen. Andere ueberlappende
        Zeitraeume mit Originalgrenzen melden; nie aufteilen, mitteln oder hinzurechnen.
        Differenz = Quellenwert minus eigener Mittelwert, keine medizinische Bewertung.
        Gerundet wird nur AnzeigeMittelwert/AnzeigeDifferenz (AwayFromZero).
        Arrays duerfen auch die einzelnen Objekte der New-...-Funktionen enthalten.
    #>
    param(
        [ValidateSet('YTD', 'Last6Months')][string]$Period = 'YTD', $AsOf = (Get-Date),
        [object[]]$ActivityEntries = @(), [object[]]$FitnessEntries = @(), [object[]]$PeriodRecords = @()
    )
    $window = Get-ActivitySummaryWindow -Period $Period -AsOf $AsOf
    $digits = $script:ActivitySummaryDecimals
    if ($null -eq $digits -or $digits -is [bool] -or $digits -is [string] -or $digits -lt 0 -or $digits -gt 15 -or $digits -ne [Math]::Truncate($digits)) { throw 'ActivitySummaryDecimals muss eine ganze Zahl von 0 bis 15 sein.' }
    $keys = @('Aktivitaetsenergie', 'Schritte', 'CardioErholung')
    $rowsByKey = @{}; $excludedByKey = @{}
    foreach ($key in $keys) { $rowsByKey[$key] = New-Object System.Collections.ArrayList; $excludedByKey[$key] = New-Object System.Collections.ArrayList }
    $seenDays = @{}; $seenTrainings = @{}
    foreach ($category in @('Aktivitaet', 'Fitness')) {
        $entries = if ($category -eq 'Aktivitaet') { $ActivityEntries } else { $FitnessEntries }
        foreach ($entry in $entries) {
            if ($null -eq $entry) { continue }
            $date = ConvertTo-BtMeasurementDate -Date $entry.Datum
            if ($date -lt $window.Von -or $date -gt $window.Bis) { continue }
            $time = ConvertTo-BtMeasurementTime -Time $entry.Zeit
            if ($category -eq 'Aktivitaet') {
                if ($time -ne '' -or $entry.Details.Datenart -ne 'Tageswerte' -or $entry.Details.TagesAbgeschlossen -isnot [bool]) { throw 'Ungültige Aktivitäts-Tagesdaten: Auswertung abgebrochen, keine Daten geändert.' }
                if ($seenDays.ContainsKey($date)) { throw 'Doppelte Aktivitäts-Tageswerte: keine automatische Auswahl oder Addition.' }
                $seenDays[$date] = $true
            }
            foreach ($key in $keys) {
                if ($category -eq 'Fitness' -and $key -ne 'CardioErholung') { continue }
                $raw = $entry.Werte.$key
                # Fehlend bleibt fehlend, insbesondere kein [double]$null -> 0.
                if ($null -eq $raw -or ($raw -is [string] -and [string]::IsNullOrWhiteSpace($raw))) { continue }
                if ($raw -is [bool]) { throw 'Ungültiger Wahrheitswert in den Aktivitätsdaten.' }
                $checked = ConvertTo-CheckedMeasurementValues -Category 'Aktivitaet' -Values @{ $key = $raw }
                if ($checked.Count -ne 1) { throw 'Ungültiger Aktivitätswert: Auswertung abgebrochen.' }
                $value = $checked[$key]
                if ($key -eq 'Schritte' -and $value -ne [Math]::Truncate($value)) { throw 'Einzelne Schritt-Tageswerte müssen ganzzahlig sein.' }
                $reason = ''; $interval = $null; $method = ''; $measurementCount = $null
                $basis = 'Tageswerte'
                if ($category -eq 'Aktivitaet' -and -not $entry.Details.TagesAbgeschlossen) { $reason = 'UnvollstaendigerTag' }
                if ($key -eq 'CardioErholung') {
                    $expected = if ($category -eq 'Fitness') { 'Training' } else { 'Tagesmittel' }
                    $basis = if ($category -eq 'Fitness') { 'Trainingswerte' } else { 'Tagesmittel' }
                    $cardio = $entry.Details.CardioErholung
                    if ($null -eq $cardio -or $cardio.Datenart -ne $expected) { throw 'Unpassende Cardio-Datenart: Training und Tagesmittel dürfen nicht vermischt werden.' }
                    $details = New-BtCardioRecoveryDetails -Datenart $expected -ErholungsintervallSekunden $cardio.ErholungsintervallSekunden -AnzahlMessungen $cardio.AnzahlMessungen -Methode ([string]$cardio.Methode)
                    $interval = $details.ErholungsintervallSekunden; $method = $details.Methode; $measurementCount = $details.AnzahlMessungen
                    if (-not $reason) {
                        if ($null -eq $interval) { $reason = 'ErholungsintervallUnbekannt' }
                        elseif ([string]::IsNullOrWhiteSpace($method)) { $reason = 'ErholungsverfahrenUnbekannt' }
                    }
                    if ($category -eq 'Fitness') {
                        $identity = $date + 'T' + $time
                        if ($seenTrainings.ContainsKey($identity)) { throw 'Doppelte Cardio-Trainingsmessung: keine automatische Auswahl oder Addition.' }
                        $seenTrainings[$identity] = $true
                    }
                }
                $row = [PSCustomObject]@{
                    Datum = $date; Zeit = $time; Wert = $value; Basis = $basis
                    ErholungsintervallSekunden = $interval; Methode = $method; AnzahlMessungen = $measurementCount
                    Quelle = ([string]$entry.Quelle).Trim(); Geraet = ([string]$entry.Details.Geraet).Trim()
                }
                if ($reason) {
                    [void]$excludedByKey[$key].Add([PSCustomObject]@{ Grund = $reason; Messung = $row })
                } else { [void]$rowsByKey[$key].Add($row) }
            }
        }
    }
    $periods = @(); $seenIds = @{}; $seenIdentities = @{}
    foreach ($record in $PeriodRecords) {
        if ($null -eq $record) { continue }
        $normalized = ConvertTo-BtActivityPeriodRecord -Record $record
        $identity = Get-BtActivityPeriodIdentity -Record $normalized
        if ($seenIds.ContainsKey($normalized.Id) -or $seenIdentities.ContainsKey($identity)) { throw 'Doppelte Zeitraum-Mittel: keine automatische Auswahl oder Addition.' }
        $seenIds[$normalized.Id] = $true; $seenIdentities[$identity] = $true
        if ($normalized.Von -le $window.Bis -and $normalized.Bis -ge $window.Von) { $periods += [PSCustomObject]$normalized }
    }
    $results = @()
    foreach ($key in $keys) {
        $groups = @{}
        foreach ($row in $rowsByKey[$key]) {
            $groupKey = 'Tageswerte'
            if ($key -eq 'CardioErholung') {
                $groupKey = ConvertTo-Json -InputObject @($row.Basis, $row.ErholungsintervallSekunden, $row.Methode.ToLowerInvariant(), $row.Quelle.ToLowerInvariant(), $row.Geraet.ToLowerInvariant()) -Compress
            }
            if (-not $groups.ContainsKey($groupKey)) { $groups[$groupKey] = New-Object System.Collections.ArrayList }
            [void]$groups[$groupKey].Add($row)
        }
        $calculated = @()
        foreach ($groupKey in @($groups.Keys | Sort-Object)) {
            $aggregate = Get-BtActivityAggregate -Rows @($groups[$groupKey]) -Digits $digits -CalendarDays $window.Kalendertage -To $window.Bis
            $calculated += $aggregate
        }
        $sources = @($periods | Where-Object { $_.Messgroesse -eq $key -and $_.Von -eq $window.Von -and $_.Bis -eq $window.Bis } | Sort-Object Quelle, Geraet, Basis, Methode, Id)
        $otherPeriods = @($periods | Where-Object { $_.Messgroesse -eq $key -and ($_.Von -ne $window.Von -or $_.Bis -ne $window.Bis) } | Sort-Object Von, Bis, Quelle, Id)
        $comparisons = @(); $hints = @()
        foreach ($source in $sources) {
            $matches = @($calculated | Where-Object {
                $_.Basis -eq $source.Basis -and ($key -ne 'CardioErholung' -or (
                    $null -ne $source.ErholungsintervallSekunden -and $_.ErholungsintervallSekunden -eq $source.ErholungsintervallSekunden -and
                    $source.Methode -and $_.Methode -ieq $source.Methode -and $_.Quelle -ieq $source.Quelle -and $_.Geraet -ieq $source.Geraet
                ))
            })
            if ($matches.Count -eq 0) { $hints += 'Ein Quellenmittel hat keine passende Berechnungsgruppe; es bleibt ohne automatischen Vergleich erhalten.' }
            foreach ($aggregate in $matches) {
                $difference = [double]$source.Wert - $aggregate.Mittelwert
                $differenceFinite = -not ([double]::IsInfinity($difference) -or [double]::IsNaN($difference))
                $sameCoverage = $null
                if ($aggregate.Basis -ne 'Trainingswerte') { $sameCoverage = ($aggregate.Anzahl -eq $window.Kalendertage -and $source.Anzahl -eq $window.Kalendertage) }
                $comparisons += [PSCustomObject]@{
                    Quellenmittel = $source; Berechnet = $aggregate
                    Differenz = $(if ($differenceFinite) { $difference } else { $null })
                    AnzeigeDifferenz = $(if ($differenceFinite) { [Math]::Round($difference, $digits, [MidpointRounding]::AwayFromZero) } else { $null })
                    Abweichend = ([double]$source.Wert -ne $aggregate.Mittelwert)
                    GleicheTagesabdeckung = $sameCoverage
                    Hinweis = $(if (-not $differenceFinite) { 'Differenz außerhalb des darstellbaren Zahlenbereichs; beide Ausgangswerte bleiben erhalten.' } elseif ($sameCoverage -eq $true) { 'Beide Mittel betreffen alle Tage des Fensters; Quellen und Verfahren können dennoch abweichen.' } else { 'Unterschiedliche oder unbekannte Datenabdeckung: rechnerische Differenz, kein Nachweis eines Fehlers. Quellenmittel hat Vorrang.' })
                }
            }
        }
        if ($calculated.Count -eq 0) { $hints += 'Keine verwendbaren Einzelwerte für einen eigenen Durchschnitt vorhanden.' }
        if ($sources.Count -gt 0) { $hints += 'Exakte Quellenmittel bleiben maßgeblich und unverändert neben der eigenen Berechnung erhalten; mehrere Quellen werden nicht automatisch zusammengefasst.' }
        if ($otherPeriods.Count -gt 0) { $hints += 'Weitere Quellenmittel überlappen das Fenster, passen aber nicht exakt. Keine anteilige Umrechnung und kein Vergleich als gleiches Zeitfenster.' }
        if ($excludedByKey[$key].Count -gt 0) { $hints += 'Einzelwerte wurden nicht zusammengefasst: unvollständige Tage oder unbekanntes Cardio-Intervall/Verfahren. Gründe und Quellwerte stehen in NichtZusammengefasst.' }
        if (@($comparisons | Where-Object Abweichend).Count -gt 0) { $hints += 'Quellenmittel und eigener Durchschnitt weichen ab. Datenabdeckung und Herkunft vor einer Interpretation prüfen.' }
        $definition = Resolve-MeasurementDefinition -Category 'Aktivitaet' -KeyOrName $key
        $results += [PSCustomObject]@{
            Messgroesse = $key; Einheit = $definition.Unit
            BevorzugteDatenart = $(if ($sources.Count -gt 0) { 'Quellenmittel' } elseif ($calculated.Count -gt 0) { 'Berechnet' } else { 'Keine' })
            Berechnet = @($calculated); Quellenmittel = @($sources); Abweichungen = @($comparisons); AndereZeitraeume = @($otherPeriods)
            NichtZusammengefasst = @($excludedByKey[$key] | Sort-Object { $_.Messung.Datum }, { $_.Messung.Zeit }, Grund)
            Hinweise = @($hints | Select-Object -Unique)
        }
    }
    return [PSCustomObject]@{ Zeitraum = $window.Zeitraum; Von = $window.Von; Bis = $window.Bis; Kalendertage = $window.Kalendertage; Ergebnisse = @($results) }
}

function Get-ActivitySummary {
    <#
    .SYNOPSIS
        Gespeicherte Daten fuer YTD oder die letzten sechs Kalendermonate auswerten.
    .EXAMPLE
        Get-ActivitySummary -Period YTD -AsOf '2026-10-09'
    .EXAMPLE
        Get-ActivitySummary -Period Last6Months -AsOf '2026-10-09'
    .DESCRIPTION
        Liefert ein Ergebnisobjekt, keine Speicherung/GUI. Unlesbare Tagesdateien
        blockieren die Auswertung statt einen scheinbar vollstaendigen Mittelwert
        zu liefern. Die reine Rechenfunktion Measure-ActivitySummary ist ohne Ablage nutzbar.
    #>
    param([ValidateSet('YTD', 'Last6Months')][string]$Period = 'YTD', $AsOf = (Get-Date))
    $window = Get-ActivitySummaryWindow -Period $Period -AsOf $AsOf
    $days = Get-MeasurementEntries -Category Aktivitaet -From $window.Von -To $window.Bis 3>$null
    if ($script:MeasurementLoadFailedFiles.Count -gt 0) { throw 'Mindestens eine Aktivitätsdatei ist nicht lesbar. Auswertung abgebrochen; Dateien bleiben unverändert.' }
    $trainings = Get-MeasurementEntries -Category Fitness -From $window.Von -To $window.Bis 3>$null
    if ($script:MeasurementLoadFailedFiles.Count -gt 0) { throw 'Mindestens eine Fitnessdatei ist nicht lesbar. Auswertung abgebrochen; Dateien bleiben unverändert.' }
    $periods = Get-ActivityPeriodRecords -From $window.Von -To $window.Bis
    return (Measure-ActivitySummary -Period $Period -AsOf $window.Bis -ActivityEntries $days -FitnessEntries $trainings -PeriodRecords $periods)
}


function Assert-BtActivityDayRecord {
    # Schuetzt neue Tagesdateien vor stiller Normalisierung/Verlust bei unpassendem Format.
    param($Record)
    if ([string]$Record.Kategorie -ne 'Aktivitaet' -or $Record.SchemaVersion -ne 2) { throw 'Unpassendes Aktivitäts-Tagesformat; Datei unverändert lassen.' }
    $entries = @(@($Record.Messungen) | Where-Object { $null -ne $_ })
    if ($entries.Count -ne 1) { throw 'Eine Aktivitäts-Tagesdatei muss genau einen konsolidierten Tages-Eintrag enthalten.' }
    $entry = $entries[0]
    if ([string]$entry.Kategorie -ne 'Aktivitaet' -or [string]$entry.Datum -ne [string]$Record.Datum -or [string]$entry.Zeit -ne '' -or $entry.Details.Datenart -ne 'Tageswerte' -or $entry.Details.TagesAbgeschlossen -isnot [bool]) { throw 'Unpassender Aktivitäts-Tages-Eintrag; Datei unverändert lassen.' }
}


function New-FitnessEntry {
    <#
    .SYNOPSIS
        v2.41.0: Fitness-Messung (VO2max und/oder CardioErholung je Training), speichert NICHT.
        Cardio-Tagesmittel dagegen mit New-ActivityEntry erzeugen. Quelle/Geraet frei waehlbar.
        -ErholungsintervallSekunden bleibt ohne Angabe unbekannt (z. B. 60 oder 120 explizit).
        -ErholungsMethode dokumentiert das Quellenverfahren, -Methode gilt weiterhin fuer VO2max.
    .PARAMETER Methode
        $script:Vo2maxMethods - Labor und Uhr-Schätzung sind nicht gleichwertig.
    #>
    param([Parameter(Mandatory)]$Date, $Zeit = '', [Parameter(Mandatory)][hashtable]$Werte, [string]$Methode = '', [string]$Quelle = '', [string]$Notiz = '',
          [string]$Geraet = '', $ErholungsintervallSekunden = $null, [string]$ErholungsMethode = '')
    $checked = ConvertTo-CheckedMeasurementValues -Category 'Fitness' -Values $Werte
    if ($checked.Count -eq 0) { throw 'Keine Fitnesswerte angegeben.' }
    $methodValue = Resolve-BtMeasurementOption -Value $Methode -Allowed $script:Vo2maxMethods -Label 'Methode'
    $hints = @()
    if ($checked.Contains('VO2max') -and -not $methodValue) { $hints += 'Messmethode fehlt - Laborwert und Schätzung der Uhr sind nicht gleichwertig.' }
    $details = [ordered]@{}
    if ($methodValue) { $details['Methode'] = $methodValue }
    if ($Geraet.Trim()) { $details['Geraet'] = $Geraet.Trim() }
    if ($checked.Contains('CardioErholung')) {
        foreach ($key in $Werte.Keys) {
            $definition = Resolve-MeasurementDefinition -Category 'Fitness' -KeyOrName $key
            if ($definition -and $definition.Key -eq 'CardioErholung' -and $Werte[$key] -is [bool]) { throw 'Cardio-Erholung muss eine Zahl sein, kein Wahrheitswert.' }
        }
        $cardio = New-BtCardioRecoveryDetails -Datenart 'Training' -ErholungsintervallSekunden $ErholungsintervallSekunden -Methode $ErholungsMethode
        $details['CardioErholung'] = $cardio; $details['Herkunft'] = 'Quelle'
        $hints += @(Get-BtActivityHints -Values $checked)
        if ($null -eq $cardio.ErholungsintervallSekunden) { $hints += 'Erholungsintervall unbekannt: Cardio-Erholung nicht automatisch mit anderen Messungen zusammenfassen.' }
    } elseif ($null -ne $ErholungsintervallSekunden -or $ErholungsMethode) { throw 'Angaben zur Cardio-Erholung benötigen einen Cardio-Erholungswert.' }
    return (New-MeasurementEntryObject -Category 'Fitness' -Date $Date -Zeit $Zeit -Werte $checked -Kontext '' -Details $details -Quelle $Quelle -Notiz $Notiz -Hinweise $hints)
}

function Read-MeasurementDayRecord {
    <#
    .SYNOPSIS
        v2.35.0 (intern): Tagesdatei einer Kategorie mit Mehrfachmessungen laden.
        Keine Datei -> $null. Unlesbar oder unpassend -> Fehler; die Datei wird NIE überschrieben.
    #>
    param([Parameter(Mandatory)][string]$Category, [Parameter(Mandatory)]$Date)
    $filePath = Get-MeasurementFilePath -Category $Category -Date $Date
    if (-not (Test-Path -LiteralPath $filePath)) { return $null }
    $record = $null
    try {
        $json = Read-ProtectedJsonFile -Path $filePath
        if ($json) { $record = $json | ConvertFrom-Json }
    } catch { $record = $null }
    if ($null -eq $record -or [string]$record.Kategorie -ne $Category -or [string]$record.Datum -ne (ConvertTo-BtMeasurementDate -Date $Date)) {
        throw "Messungsdatei nicht lesbar oder unpassend: $filePath - es wurde nichts geändert."
    }
    if ($Category -eq 'Aktivitaet') { Assert-BtActivityDayRecord -Record $record }
    return $record
}

function Add-MeasurementEntry {
    <#
    .SYNOPSIS
        v2.35.0: Speichert einen Eintrag (New-VitalEntry / New-BodyEntry / New-EkgEntry / New-FitnessEntry)
        in der Tagesdatei - verschlüsselt, atomar, Einträge nach Uhrzeit sortiert.
    .PARAMETER Replace
        Ersetzt einen vorhandenen Eintrag mit GLEICHER Uhrzeit (ohne -Replace: Fehler, nichts geändert).
    .OUTPUTS
        Pfad der Tagesdatei.
    #>
    param([Parameter(Mandatory)]$Entry, [switch]$Replace)
    $category = [string]$Entry['Kategorie']
    if ($script:MeasurementMultiCategories -notcontains $category) { throw "Kategorie '$category' unterstützt keine Einzelmessungen (DEXA/Schlaf: Save-MeasurementRecord)." }
    if ($category -eq 'Aktivitaet') {
        Assert-BtActivityDayRecord -Record @{ SchemaVersion = $script:MeasurementMultiSchemaVersion; Kategorie = $category; Datum = $Entry['Datum']; Messungen = @($Entry) }
    }
    $isoDate = ConvertTo-BtMeasurementDate -Date $Entry['Datum']
    $timeText = ConvertTo-BtMeasurementTime -Time $Entry['Zeit']
    $existing = Read-MeasurementDayRecord -Category $category -Date $isoDate
    $entries = @()
    if ($existing) { $entries = @(@($existing.Messungen) | Where-Object { $null -ne $_ }) }
    $sameTime = @($entries | Where-Object { [string]$_.Zeit -eq $timeText })
    if ($sameTime.Count -gt 0 -and -not $Replace) {
        $timeLabel = if ($timeText) { "um $timeText" } else { 'ohne Uhrzeit' }
        throw "Für $category am $isoDate gibt es bereits eine Messung $timeLabel - zum Ersetzen -Replace verwenden."
    }
    $kept = @($entries | Where-Object { [string]$_.Zeit -ne $timeText })
    $allEntries = @($kept) + @(, $Entry)
    $sorted = @($allEntries | Sort-Object -Property @{ Expression = { [string]$_.Zeit } })
    $dayRecord = [ordered]@{ SchemaVersion = $script:MeasurementMultiSchemaVersion; Kategorie = $category; Datum = $isoDate; Messungen = $sorted }
    return (Save-MeasurementRecord -Record $dayRecord)
}

function Remove-MeasurementEntry {
    <#
    .SYNOPSIS
        v2.35.0: Löscht EINEN Eintrag (Kategorie, Datum, Uhrzeit; '' = ohne Uhrzeit). Letzter Eintrag
        des Tages -> Tagesdatei wird gelöscht. $true = gelöscht, $false = nicht vorhanden.
    #>
    param([Parameter(Mandatory)][ValidateSet('Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category, [Parameter(Mandatory)]$Date, $Zeit = '')
    $existing = Read-MeasurementDayRecord -Category $Category -Date $Date
    if (-not $existing) { return $false }
    $timeText = ConvertTo-BtMeasurementTime -Time $Zeit
    $all = @(@($existing.Messungen) | Where-Object { $null -ne $_ })
    $kept = @($all | Where-Object { [string]$_.Zeit -ne $timeText })
    if ($kept.Count -eq $all.Count) { return $false }
    if ($kept.Count -eq 0) { return (Remove-MeasurementRecord -Category $Category -Date $Date) }
    $dayRecord = [ordered]@{ SchemaVersion = $script:MeasurementMultiSchemaVersion; Kategorie = $Category; Datum = [string]$existing.Datum; Messungen = $kept }
    $null = Save-MeasurementRecord -Record $dayRecord
    return $true
}

function Get-MeasurementEntries {
    <#
    .SYNOPSIS
        v2.35.0: Alle Einzelmessungen einer Kategorie (optional Zeitraum und nur mit bestimmtem Wert),
        chronologisch nach Datum und Uhrzeit. Unlesbare Tagesdateien: siehe Get-MeasurementRecords.
    .PARAMETER ValueKey
        Optional: nur Einträge mit diesem Wert (Schlüssel oder Anzeigename, z. B. 'Blutdruck_Sys').
    .OUTPUTS
        Array von PSCustomObject { Datum; Zeit; Werte; Kontext; Details; Quelle; Notiz; Hinweise; Erfasst }.
        Wie Get-MeasurementRecords: Aufrufer OHNE zusätzliches @() übernehmen.
    #>
    param([Parameter(Mandatory)][ValidateSet('Vitalwerte', 'Koerper', 'EKG', 'Fitness', 'Aktivitaet')][string]$Category, $From = $null, $To = $null, [string]$ValueKey = '')
    $filterKey = ''
    if ($ValueKey) {
        $definition = Resolve-MeasurementDefinition -Category $Category -KeyOrName $ValueKey
        if (-not $definition) { throw "Unbekannter Messwert '$ValueKey' in Kategorie '$Category'." }
        $filterKey = $definition.Key
    }
    $result = New-Object System.Collections.ArrayList
    foreach ($day in (Get-MeasurementRecords -Category $Category -From $From -To $To)) {
        foreach ($entry in @(@($day.Messungen) | Where-Object { $null -ne $_ })) {
            if ($filterKey -and -not (@($entry.Werte.PSObject.Properties.Name) -contains $filterKey)) { continue }
            [void]$result.Add([PSCustomObject]@{
                Datum = [string]$day.Datum; Zeit = [string]$entry.Zeit; Werte = $entry.Werte; Kontext = [string]$entry.Kontext
                Details = $entry.Details; Quelle = [string]$entry.Quelle; Notiz = [string]$entry.Notiz; Hinweise = @($entry.Hinweise); Erfasst = [string]$entry.Erfasst
            })
        }
    }
    return , @($result | Sort-Object -Property @{ Expression = { [string]$_.Datum } }, @{ Expression = { [string]$_.Zeit } })
}

function ConvertTo-BtMeasurementValueTable {
    <# v2.35.0 (intern): Werte (Hashtable/[ordered] ODER PSCustomObject aus JSON) -> Hashtable Schlüssel -> [double]. #>
    param($Values)
    $table = @{}
    if ($null -eq $Values) { return $table }
    if ($Values -is [System.Collections.IDictionary]) { foreach ($k in $Values.Keys) { $table[[string]$k] = [double]$Values[$k] } }
    else { foreach ($p in $Values.PSObject.Properties) { $table[$p.Name] = [double]$p.Value } }
    return $table
}

function Get-BodyDerivedValues {
    <#
    .SYNOPSIS
        v2.35.0: Fehlende Körper-Kennzahlen aus vorhandenen Werten (werden nicht gespeichert):
        BMI (Gewicht / Größe²), fettfreie Masse (Gewicht x (1 - KF %)), Körperfettanteil
        (aus Gewicht und fettfreier Masse), FFMI (fettfreie Masse / Größe²).
        Gemessene Werte werden nie überschrieben.
    #>
    param([Parameter(Mandatory)]$Entry, [double]$HeightCm = 0)
    $measured = ConvertTo-BtMeasurementValueTable -Values $Entry.Werte
    $derived = [ordered]@{}
    $heightSquared = if ($HeightCm -gt 0) { [Math]::Pow($HeightCm / 100, 2) } else { 0 }
    $weight = if ($measured.ContainsKey('Gewicht')) { $measured['Gewicht'] } else { $null }
    if ($null -ne $weight -and $heightSquared -gt 0 -and -not $measured.ContainsKey('BMI')) { $derived['BMI'] = [Math]::Round($weight / $heightSquared, 1) }
    $fatFree = $null
    if ($measured.ContainsKey('FettfreieMasse')) { $fatFree = $measured['FettfreieMasse'] }
    elseif ($null -ne $weight -and $measured.ContainsKey('Koerperfett_Prozent')) {
        $fatFree = $weight * (1 - $measured['Koerperfett_Prozent'] / 100)
        $derived['FettfreieMasse'] = [Math]::Round($fatFree, 1)
    }
    if ($null -ne $weight -and $weight -gt 0 -and $measured.ContainsKey('FettfreieMasse') -and -not $measured.ContainsKey('Koerperfett_Prozent')) {
        $derived['Koerperfett_Prozent'] = [Math]::Round(100 * ($weight - $measured['FettfreieMasse']) / $weight, 1)
    }
    if ($null -ne $fatFree -and $heightSquared -gt 0 -and -not $measured.ContainsKey('FFMI')) { $derived['FFMI'] = [Math]::Round($fatFree / $heightSquared, 1) }
    return $derived
}

function Get-EkgDerivedValues {
    <#
    .SYNOPSIS
        v2.35.0: QTc aus QT und Herzfrequenz (werden nicht gespeichert), RR = 60 / Herzfrequenz (s):
        Bazett QT / Wurzel(RR), Fridericia QT / Kubikwurzel(RR), ganze ms. Gemessene Werte bleiben.
    #>
    param([Parameter(Mandatory)]$Entry)
    $measured = ConvertTo-BtMeasurementValueTable -Values $Entry.Werte
    $derived = [ordered]@{}
    if ($measured.ContainsKey('QT') -and $measured.ContainsKey('Herzfrequenz') -and $measured['Herzfrequenz'] -gt 0) {
        $rrSeconds = 60 / $measured['Herzfrequenz']
        if (-not $measured.ContainsKey('QTc_Bazett')) { $derived['QTc_Bazett'] = [Math]::Round($measured['QT'] / [Math]::Sqrt($rrSeconds), 0, [MidpointRounding]::AwayFromZero) }
        if (-not $measured.ContainsKey('QTc_Fridericia')) { $derived['QTc_Fridericia'] = [Math]::Round($measured['QT'] / [Math]::Pow($rrSeconds, 1 / 3), 0, [MidpointRounding]::AwayFromZero) }
    }
    return $derived
}

function ConvertTo-BtBirthDate {
    <#
    .SYNOPSIS
        v2.29.0: Wandelt ein Geburtsdatum (Text) in [datetime] um.
    .DESCRIPTION
        Erlaubt TT.MM.JJJJ / T.M.JJJJ (Eingabe) und JJJJ-MM-TT (Speicherformat).
    .OUTPUTS
        [datetime] (nur Datum) oder $null bei leerer/ungültiger Eingabe.
    #>
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $formats = [string[]]@('dd.MM.yyyy', 'd.M.yyyy', 'yyyy-MM-dd')
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($Text.Trim(), $formats, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) {
        return $parsed.Date
    }
    return $null
}

function Get-AgeAtDate {
    <#
    .SYNOPSIS
        v2.29.0: Alter in vollendeten Lebensjahren an einem (Test-)Datum.
    .DESCRIPTION
        Quelle ist Personal.Geburtsdatum (JJJJ-MM-TT). Fehlt es (Configs vor v2.29.0),
        wird wie bisher das feste Personal.Age verwendet. Ohne Datum gilt heute.
    .OUTPUTS
        [double] Alter oder $null (keine Angabe, ungültiges Datum, Datum vor der Geburt).
    #>
    param($Personal, $Date = $null)
    if ($null -eq $Personal) { return $null }
    $birth = $null
    $rawBirth = $Personal.Geburtsdatum
    if ($rawBirth -is [datetime]) { $birth = $rawBirth.Date }
    elseif ($rawBirth) { $birth = ConvertTo-BtBirthDate -Text ([string]$rawBirth) }
    if ($null -eq $birth) {
        $fixedAge = 0.0
        if ($null -ne $Personal.Age -and [double]::TryParse([string]$Personal.Age, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$fixedAge) -and $fixedAge -gt 0) {
            return $fixedAge
        }
        return $null
    }
    $refDate = (Get-Date).Date
    if ($null -ne $Date -and [string]$Date -ne '') {
        try { $refDate = ([datetime]$Date).Date } catch { return $null }
    }
    if ($refDate -lt $birth) { return $null }
    $years = $refDate.Year - $birth.Year
    if ($refDate.Month -lt $birth.Month -or ($refDate.Month -eq $birth.Month -and $refDate.Day -lt $birth.Day)) { $years-- }
    return [double]$years
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
        Array [CRP mg/L, Neutrophile %, Lymphozyten %, Thrombozyten Tsd./µL, Leukozyten Tsd./µL]
        v2.28.0: Leukozyten (5. Wert) für die ABSOLUTE Lymphozytenzahl. Fehlt er (alte
        Config ohne Korrektur), wird der Score ohne PLR-Anteil berechnet (CRP 62,5 %,
        NLR 37,5 %) statt mit einem falschen PLR.
    #>
    param([double[]]$values)
    try {
        $crp    = [double]$values[0]
        $neut   = [double]$values[1]
        $lymph  = [double]$values[2]
        $plt    = [double]$values[3]
        $wbc    = if ($values.Count -ge 5) { [double]$values[4] } else { 0 }

        if ($lymph -le 0) { return 0 }

        # NLR: Prozent / Prozent = Verhältnis der absoluten Zahlen (Leukozyten kürzen sich)
        $nlr = $neut / $lymph

        # Score-Komponenten (0-100, höher = besser)
        # v2.28.0: 0.0/100.0 statt 0/100 - sonst wählt PowerShell Math.Min(int,int) und rundet die Teilscores auf ganze Zahlen
        $crpScore  = [Math]::Max(0.0, [Math]::Min(100.0, (1 - $crp / 3.0) * 100))         # <1 optimal, >3 schlecht
        $nlrScore  = [Math]::Max(0.0, [Math]::Min(100.0, (1 - ($nlr - 1.0) / 2.5) * 100)) # 1-3 Range

        if ($wbc -gt 0) {
            # v2.28.0 BUGFIX: PLR = Thrombozyten / ABSOLUTE Lymphozyten (Tsd./µl).
            # Bisher Thrombozyten / Lymphozyten-% (ca. 8 statt ca. 140) -> PLR-Anteil immer 100.
            $lymphAbs = $lymph * $wbc / 100
            $plr = $plt / $lymphAbs
            $plrScore  = [Math]::Max(0.0, [Math]::Min(100.0, (1 - ($plr - 70) / 130) * 100))  # 70-200 Range
            # Gewichteter Mittelwert: CRP dominant (Evidenzstärke)
            $score = ($crpScore * 0.5) + ($nlrScore * 0.3) + ($plrScore * 0.2)
        } else {
            # Ohne Leukozyten kein korrekter PLR -> nur CRP + NLR (Gewichte 0,5 : 0,3 normiert)
            $score = (($crpScore * 0.5) + ($nlrScore * 0.3)) / 0.8
        }
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
        [double]$value,
        $date = $null          # v2.29.0: Testtag des Wertes (für das Alter am Testtag)
    )
    $score = 50
    switch ($markerName) {
        "PhenoAge (biologisch)" {
            # Score basiert auf Differenz zu chronol. Alter - wird über PhenoAge-Accel abgedeckt
            # v2.29.0: Alter am Testtag des Wertes (Geburtsdatum), sonst festes Alter
            $chronoAge = Get-AgeAtDate -Personal $script:data.Config.Personal -Date $date
            if ($null -eq $chronoAge) { return "-" }
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
    # v2.31.0: [double] - bei ganzzahligem $score wählte PowerShell Math.Round(decimal) ->
    # Ergebnis [decimal], die Ampelfarbe im Tab "Longevity" fehlte (Prüfung auf double/int).
    return [Math]::Round([double]$score, 0)
}

function Get-MarkerRangeScore {
    <#
    .SYNOPSIS
        v2.30.0: Bewertet einen Messwert mit 0-100 Punkten an den Bereichen seines Markers.
    .DESCRIPTION
        Gleiche Einteilung wie die Ampel im Risiko-Cockpit:
          im Optimalbereich                    -> 100            (grün)
          im Referenzbereich, nicht optimal    -> 79 bis 50      (gelb, linear bis zur Referenzgrenze)
          außerhalb des Referenzbereichs       -> 49 bis 0       (rot, 0 ab einem weiteren Abstand
                                                                  in Breite der gelben Zone)
        Ohne Optimalbereich: im Referenzbereich 79 (Cockpit: "Akzeptabel").
    .OUTPUTS
        [double] Punkte (1 Nachkommastelle, kaufmännisch gerundet) oder $null ohne Bereiche.
    #>
    param([double]$Value, $RefMin, $RefMax, $OptimalMin, $OptimalMax)
    $toNum = { param($x) if ($null -eq $x -or [string]$x -eq '') { $null } else { try { [double]$x } catch { $null } } }
    $rMin = & $toNum $RefMin; $rMax = & $toNum $RefMax; $oMin = & $toNum $OptimalMin; $oMax = & $toNum $OptimalMax
    $hasOptimal = ($null -ne $oMin -and $null -ne $oMax)
    if (-not $hasOptimal) {
        if ($null -eq $rMin -or $null -eq $rMax) { return $null }
        if ($Value -ge $rMin -and $Value -le $rMax) { return 79.0 }
        $oMin = $rMin; $oMax = $rMax
    }
    if ($null -eq $rMin) { $rMin = $oMin }
    if ($null -eq $rMax) { $rMax = $oMax }
    $rMin = [Math]::Min($rMin, $oMin); $rMax = [Math]::Max($rMax, $oMax)
    if ($Value -ge $oMin -and $Value -le $oMax) { return 100.0 }
    if ($Value -gt $oMax) { $zone = $rMax - $oMax; $dist = $Value - $oMax } else { $zone = $oMin - $rMin; $dist = $oMin - $Value }
    if ($hasOptimal -and $zone -gt 0 -and $dist -le $zone) { return [Math]::Round(79.0 - 29.0 * ($dist / $zone), 1, [MidpointRounding]::AwayFromZero) }
    # Außerhalb des Referenzbereichs: Skala = Breite der gelben Zone (sonst Breite des Optimal-/Referenzbereichs)
    $scale = if ($hasOptimal -and $zone -gt 0) { $zone } else { $oMax - $oMin }
    if ($scale -le 0) { $scale = [Math]::Max([Math]::Abs($oMax), 1.0) }
    $over = if ($hasOptimal -and $zone -gt 0) { $dist - $zone } else { $dist }
    return [Math]::Round([Math]::Max(0.0, 49.0 * (1.0 - $over / $scale)), 1, [MidpointRounding]::AwayFromZero)
}

function Get-LongevityMarkerScore {
    <#
    .SYNOPSIS
        v2.30.0: 0-100-Bewertung eines Laborwerts an den Bereichen aus den Marker-Einstellungen.
    #>
    param([string]$MarkerName, [double]$Value)
    $markerCfg = @($script:data.Config.Markers) | Where-Object { $_.Name -eq $MarkerName } | Select-Object -First 1
    if (-not $markerCfg) { return $null }
    return Get-MarkerRangeScore -Value $Value -RefMin $markerCfg.RefMin -RefMax $markerCfg.RefMax -OptimalMin $markerCfg.OptimalMin -OptimalMax $markerCfg.OptimalMax
}

function Get-OptimalRangeText {
    <#
    .SYNOPSIS
        v2.30.0: Optimalbereich eines Markers als Text (Spalte "Ideal" im Tab Longevity).
    #>
    param([string]$MarkerName, [string]$Fallback = '')
    $markerCfg = @($script:data.Config.Markers) | Where-Object { $_.Name -eq $MarkerName } | Select-Object -First 1
    if (-not $markerCfg) { return $Fallback }
    $oMin = $null; $oMax = $null
    try { if ($null -ne $markerCfg.OptimalMin -and [string]$markerCfg.OptimalMin -ne '') { $oMin = [double]$markerCfg.OptimalMin } } catch { $oMin = $null }
    try { if ($null -ne $markerCfg.OptimalMax -and [string]$markerCfg.OptimalMax -ne '') { $oMax = [double]$markerCfg.OptimalMax } } catch { $oMax = $null }
    if ($null -eq $oMin -or $null -eq $oMax) { return $Fallback }
    $ci = [System.Globalization.CultureInfo]::InvariantCulture
    $unitText = if ($markerCfg.Unit) { " $($markerCfg.Unit)" } else { '' }
    if ($oMin -le 0) { return "≤ " + $oMax.ToString($ci) + $unitText }
    return $oMin.ToString($ci) + " - " + $oMax.ToString($ci) + $unitText
}

function Calculate-LongevityScore {
    <#
    .SYNOPSIS
        v2.30.0: Longevity-Score 0-100 = Mittelwert der Einzelbewertungen am Testtag.
    .DESCRIPTION
        Bis v2.29.0: feste Formeln, die den eigenen Optimalbereichen widersprachen
        (HbA1c 5,2 % = 0 Punkte; gesunder Datensatz ca. 33/100), und Berechnung nur,
        wenn alle 7 Werte inkl. NAD+ vorlagen.
        Neu: jeder vorhandene Wert wird mit Get-LongevityMarkerScore an den Bereichen
        seines Markers bewertet. Mindestens $script:LongevityScoreMinComponents Werte.
    .PARAMETER values
        Werte in der Reihenfolge der RequiredMarkers ($null = am Testtag nicht gemessen).
    .PARAMETER names
        Tatsächlich verwendete Markernamen (z. B. hs-CRP statt CRP) - deren Bereiche zählen.
    .OUTPUTS
        [double] Score (1 Nachkommastelle) oder $null (zu wenige Werte).
    #>
    param($values, [string[]]$names)
    $scores = New-Object System.Collections.ArrayList
    $valueList = @($values)
    for ($i = 0; $i -lt $valueList.Count; $i++) {
        if ($null -eq $valueList[$i] -or $null -eq $names -or $i -ge $names.Count -or [string]::IsNullOrEmpty($names[$i])) { continue }
        $componentScore = Get-LongevityMarkerScore -MarkerName $names[$i] -Value ([double]$valueList[$i])
        if ($null -ne $componentScore) { [void]$scores.Add([double]$componentScore) }
    }
    $minComponents = if ($script:LongevityScoreMinComponents) { [int]$script:LongevityScoreMinComponents } else { 4 }
    if ($scores.Count -lt $minComponents) { return $null }
    return [Math]::Round(($scores | Measure-Object -Average).Average, 1, [MidpointRounding]::AwayFromZero)   # kaufmännisch gerundet
}
function Get-EgfrCkdEpi2021 {
    <#
    .SYNOPSIS
        v2.27.0: eGFR in ml/min/1,73 m² nach CKD-EPI 2021 (Kreatinin, ohne Ethnie-Faktor).
    .DESCRIPTION
        Inker et al., NEJM 2021. eGFR = 142 x min(Scr/k,1)^a x max(Scr/k,1)^-1,200
        x 0,9938^Alter x 1,012 [Frau]; k = 0,7 (Frau) / 0,9 (Mann),
        a = -0,241 (Frau) / -0,302 (Mann). Kreatinin > 20 wird als µmol/l
        interpretiert und durch 88,4 geteilt (deutsche Labore nutzen beide Einheiten).
    .OUTPUTS
        [double] eGFR (1 Nachkommastelle) oder $null bei ungültiger Eingabe.
    #>
    param([double]$Creatinine, [double]$Age, [bool]$Female)
    if ($Creatinine -le 0 -or $Age -le 0) { return $null }
    $scr = if ($Creatinine -gt 20) { $Creatinine / 88.4 } else { $Creatinine }
    if ($Female) { $kappa = 0.7; $alpha = -0.241; $sexFactor = 1.012 } else { $kappa = 0.9; $alpha = -0.302; $sexFactor = 1.0 }
    $ratio = $scr / $kappa
    $egfr = 142 * [Math]::Pow([Math]::Min($ratio, 1), $alpha) * [Math]::Pow([Math]::Max($ratio, 1), -1.200) * [Math]::Pow(0.9938, $Age) * $sexFactor
    return [Math]::Round($egfr, 1)
}

function Calculate-PreventRisk {
    <#
    .SYNOPSIS
        v2.27.0: AHA-PREVENT-Basismodell (Khan et al., Circulation 2024;149:430-449).
    .DESCRIPTION
        10-Jahres-Risiko in % für:
          ASCVD = atherosklerotische Herz-Kreislauf-Erkrankung (Herzinfarkt, Schlaganfall)
          CVD   = Herz-Kreislauf gesamt (ASCVD + Herzinsuffizienz)
        Logistisches Modell mit zentrierten Prädiktoren; Cholesterin intern in mmol/l
        (mg/dl x 0,02586). Gültigkeitsbereich wie im offiziellen Rechner: Alter 30-79,
        Gesamtcholesterin 130-320 mg/dl, HDL 20-100 mg/dl, SBP 90-180 mmHg,
        eGFR 15-140 ml/min/1,73 m². Außerhalb -> $null (kein Wert).
        Validiert gegen preventr::estimate_risk (Frau 50 J., SBP 160 behandelt,
        TC 200, HDL 45, Diabetes, eGFR 90): CVD 14,7 %, ASCVD 9,2 %.
    .OUTPUTS
        [double] Risiko in % (2 Nachkommastellen) oder $null.
    #>
    param(
        [ValidateSet('ASCVD', 'CVD')][string]$Outcome = 'ASCVD',
        [double]$TotalCholesterol, [double]$Hdl, [double]$Sbp, [double]$Egfr, [double]$Age,
        [bool]$Female, [bool]$Diabetes, [bool]$Smoker, [bool]$BpTreatment, [bool]$Statin
    )
    if ($Age -lt 30 -or $Age -gt 79) { return $null }
    if ($TotalCholesterol -lt 130 -or $TotalCholesterol -gt 320) { return $null }
    if ($Hdl -lt 20 -or $Hdl -gt 100) { return $null }
    if ($Sbp -lt 90 -or $Sbp -gt 180) { return $null }
    if ($Egfr -lt 15 -or $Egfr -gt 140) { return $null }

    # Reihenfolge: Konstante, Alter, NonHDL, HDL, SBP<110, SBP>=110, Diabetes, Raucher,
    # eGFR<60, eGFR>=60, Blutdrucksenker, Statin, BPtx*SBP, Statin*NonHDL, Alter*NonHDL,
    # Alter*HDL, Alter*SBP, Alter*Diabetes, Alter*Raucher, Alter*eGFR<60
    $coefficients = @{
        'CVD|F'   = @(-3.307728, 0.7939329, 0.0305239, -0.1606857, -0.2394003, 0.3600781, 0.8667604, 0.5360739, 0.6045917, 0.0433769, 0.3151672, -0.1477655, -0.0663612, 0.1197879, -0.0819715, 0.0306769, -0.0946348, -0.27057, -0.078715, -0.1637806)
        'CVD|M'   = @(-3.031168, 0.7688528, 0.0736174, -0.0954431, -0.4347345, 0.3362658, 0.7692857, 0.4386871, 0.5378979, 0.0164827, 0.288879, -0.1337349, -0.0475924, 0.150273, -0.0517874, 0.0191169, -0.1049477, -0.2251948, -0.0895067, -0.1543702)
        'ASCVD|F' = @(-3.819975, 0.719883, 0.1176967, -0.151185, -0.0835358, 0.3592852, 0.8348585, 0.4831078, 0.4864619, 0.0397779, 0.2265309, -0.0592374, -0.0395762, 0.0844423, -0.0567839, 0.0325692, -0.1035985, -0.2417542, -0.0791142, -0.1671492)
        'ASCVD|M' = @(-3.500655, 0.7099847, 0.1658663, -0.1144285, -0.2837212, 0.3239977, 0.7189597, 0.3956973, 0.3690075, 0.0203619, 0.2036522, -0.0865581, -0.0322916, 0.114563, -0.0300005, 0.0232747, -0.0927024, -0.2018525, -0.0970527, -0.1217081)
    }
    $b = $coefficients["$Outcome|$(if ($Female) { 'F' } else { 'M' })"]

    $tcMmol  = $TotalCholesterol * 0.02586
    $hdlMmol = $Hdl * 0.02586
    $cAge   = ($Age - 55) / 10
    $cNonHdl = $tcMmol - $hdlMmol - 3.5
    $cHdl   = ($hdlMmol - 1.3) / 0.3
    $cSbp1  = ([Math]::Min($Sbp, 110) - 110) / 20
    $cSbp2  = ([Math]::Max($Sbp, 110) - 130) / 20
    $cEgfr1 = ([Math]::Min($Egfr, 60) - 60) / -15
    $cEgfr2 = ([Math]::Max($Egfr, 60) - 90) / -15
    $dm  = if ($Diabetes) { 1 } else { 0 }
    $smk = if ($Smoker) { 1 } else { 0 }
    $bpt = if ($BpTreatment) { 1 } else { 0 }
    $sta = if ($Statin) { 1 } else { 0 }

    $x = $b[0] + $b[1] * $cAge + $b[2] * $cNonHdl + $b[3] * $cHdl + $b[4] * $cSbp1 + $b[5] * $cSbp2 +
         $b[6] * $dm + $b[7] * $smk + $b[8] * $cEgfr1 + $b[9] * $cEgfr2 + $b[10] * $bpt + $b[11] * $sta +
         $b[12] * $bpt * $cSbp2 + $b[13] * $sta * $cNonHdl + $b[14] * $cAge * $cNonHdl + $b[15] * $cAge * $cHdl +
         $b[16] * $cAge * $cSbp2 + $b[17] * $cAge * $dm + $b[18] * $cAge * $smk + $b[19] * $cAge * $cEgfr1

    return [Math]::Round(100 * [Math]::Exp($x) / (1 + [Math]::Exp($x)), 2)
}

function Calculate-PreventScore {
    <#
    .SYNOPSIS
        v2.27.0: PREVENT-Risiko für EIN Testdatum (ersetzt die bisherige Näherungsformel,
        die z. B. für einen gesunden 38-Jährigen 35 % statt < 1 % lieferte).
    .DESCRIPTION
        Eingaben je Testdatum:
          Gesamtcholesterin + HDL : Pflicht (Messwerte am Testtag)
          Systolischer Blutdruck  : Messwert am Testtag, sonst Personal.SystBlutdruck
          eGFR                    : GFR-Marker am Testtag, sonst aus Kreatinin (CKD-EPI 2021)
          Diabetes                : Personal.DiabetesDiagnose ODER HbA1c >= 6,5 % (>= 48 mmol/mol)
          Raucher / Blutdrucksenker / Lipidsenker : Persönliche Metriken
          Alter                   : am Testtag aus dem Geburtsdatum (v2.29.0), sonst festes Alter
          Geschlecht              : Persönliche Metriken
    .OUTPUTS
        [double] Risiko in % oder $null (fehlende/ungültige Eingaben).
    #>
    param(
        $ValuesOnDate,
        $Personal,
        [ValidateSet('ASCVD', 'CVD')][string]$Outcome = 'ASCVD',
        $TestDate = $null      # v2.29.0: Testtag (Alter am Testtag); leer = heute
    )
    # Messwerte des Testtags nach Markername (erster Treffer zählt)
    $valueByName = @{}
    foreach ($dayItem in @($ValuesOnDate)) {
        if ($null -eq $dayItem -or $null -eq $dayItem.Value) { continue }
        $itemName = [string]$dayItem.Name
        if (-not $valueByName.ContainsKey($itemName)) { $valueByName[$itemName] = [double]$dayItem.Value }
    }
    $tc  = $valueByName['Cholesterin, gesamt (Chol)']
    $hdl = $valueByName['HDL-Cholesterin (HDL)']
    if ($null -eq $tc -or $null -eq $hdl) { return $null }

    # v2.29.0: Alter am Testtag (Geburtsdatum); ohne Geburtsdatum wie bisher festes Alter
    $age = Get-AgeAtDate -Personal $Personal -Date $TestDate
    if ($null -eq $age) { return $null }
    $female = ([string]$Personal.Geschlecht -eq 'weiblich')

    $sbp = $valueByName['Systolischer Blutdruck']
    if ($null -eq $sbp -and $Personal.SystBlutdruck) { $sbp = [double]$Personal.SystBlutdruck }
    if ($null -eq $sbp) { return $null }

    $egfr = $valueByName['Glomeruläre Filtrationsrate (GFR)']
    if ($null -eq $egfr) {
        $crea = $valueByName['Kreatinin (Krea)']
        if ($null -ne $crea) { $egfr = Get-EgfrCkdEpi2021 -Creatinine $crea -Age $age -Female $female }
    }
    if ($null -eq $egfr) { return $null }

    $hba1c = $valueByName['Langzeitzucker (HbA1c)']
    $diabetes = [bool]$Personal.DiabetesDiagnose
    if ($null -ne $hba1c) {
        if (($hba1c -lt 20 -and $hba1c -ge 6.5) -or $hba1c -ge 48) { $diabetes = $true }
    }

    return Calculate-PreventRisk -Outcome $Outcome -TotalCholesterol $tc -Hdl $hdl -Sbp $sbp -Egfr $egfr -Age $age `
        -Female $female -Diabetes $diabetes -Smoker ([bool]$Personal.Raucher) `
        -BpTreatment ([bool]$Personal.Blutdrucksenker) -Statin ([bool]$Personal.Lipidsenker)
}

function Get-CalculatedValuesForMarker {
    param($markerName)
    $calculatedValues = New-Object System.Collections.ArrayList
    $markerConfig = [System.Collections.ArrayList]@($script:data.Config.CalculatedMarkers | ForEach-Object { [PSCustomObject]$_ }) | Where-Object { $_.Name -eq $markerName }
    if (-not $markerConfig) { return $calculatedValues }

    $specialFormulas = @("[LONGEVITY_SCORE]", "[PREVENT_SCORE]", "[PREVENT_CVD]", "[PHENOAGE]", "[PHENOAGE_ACCEL]", "[INFLAMMAGING]")   # v2.27.0: + [PREVENT_CVD]
    $isSpecialFormula = $specialFormulas -contains $markerConfig.Formula
    $groupedByDate = $script:allHistoricalItems | Group-Object -Property Date
    
    foreach ($dateGroup in $groupedByDate) {
        $date = $dateGroup.Name
        $valuesOnDate = $dateGroup.Group
        
        $formula = $markerConfig.Formula
        $formula = $formula -replace '\$personal\.(\w+)', ('$script:data.Config.Personal.{0}' -f '$1')

        $allMarkersFound = $true
        $markerValues = @()
        $markerNamesUsed = @()   # v2.30.0: tatsächlich verwendete Marker (Bereiche für den Longevity-Score)
        for ($i = 0; $i -lt $markerConfig.RequiredMarkers.Count; $i++) {
            $requiredMarkerName = $markerConfig.RequiredMarkers[$i]
            $foundMarker = $null
            # v2.31.0: Score-Formeln nehmen den Vorrang-Marker zuerst (hs-CRP vor CRP)
            if ($isSpecialFormula -and $script:MarkerPreferred -and $script:MarkerPreferred.ContainsKey($requiredMarkerName)) {
                foreach ($preferredName in $script:MarkerPreferred[$requiredMarkerName]) {
                    $foundMarker = $valuesOnDate | Where-Object { $_.Name -eq $preferredName } | Select-Object -First 1
                    if ($foundMarker) { break }
                }
            }
            if (-not $foundMarker) { $foundMarker = $valuesOnDate | Where-Object { $_.Name -eq $requiredMarkerName } | Select-Object -First 1 }
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
                    $markerNamesUsed += [string]$foundMarker.Name
                } else {
                    $formula = $formula -replace "%$i%", $foundMarker.Value.ToString([System.Globalization.CultureInfo]::InvariantCulture)
                }
            } elseif ($markerConfig.Formula -eq "[LONGEVITY_SCORE]") {
                # v2.30.0: Longevity-Score braucht nicht alle Werte (Mindestanzahl prüft Calculate-LongevityScore)
                $markerValues += $null
                $markerNamesUsed += ''
            } else {
                $allMarkersFound = $false; break
            }
        }

        if ($allMarkersFound) {
            try {
                $resultValue = 0
                if ($markerConfig.Formula -eq "[LONGEVITY_SCORE]") {
                    $resultValue = Calculate-LongevityScore -values $markerValues -names $markerNamesUsed
                    if ($null -eq $resultValue) { continue }
                } elseif ($markerConfig.Formula -eq "[PREVENT_SCORE]" -or $markerConfig.Formula -eq "[PREVENT_CVD]") {
                    # v2.27.0: echtes PREVENT-Basismodell; [PREVENT_SCORE] = 10 J. ASCVD, [PREVENT_CVD] = 10 J. Herz-Kreislauf gesamt
                    $preventOutcome = if ($markerConfig.Formula -eq "[PREVENT_CVD]") { 'CVD' } else { 'ASCVD' }
                    $resultValue = Calculate-PreventScore -ValuesOnDate $valuesOnDate -Personal $script:data.Config.Personal -Outcome $preventOutcome -TestDate $date
                    if ($null -eq $resultValue) { continue }
                } elseif ($markerConfig.Formula -eq "[PHENOAGE]") {
                    # v2.29.0: Alter am Testtag (Geburtsdatum) statt festem Alter
                    $chronoAge = Get-AgeAtDate -Personal $script:data.Config.Personal -Date $date
                    if ($null -eq $chronoAge) { continue }
                    $resultValue = Calculate-PhenoAge -values $markerValues -chronoAge $chronoAge
                    if ($null -eq $resultValue) { continue }
                } elseif ($markerConfig.Formula -eq "[PHENOAGE_ACCEL]") {
                    $chronoAge = Get-AgeAtDate -Personal $script:data.Config.Personal -Date $date
                    if ($null -eq $chronoAge) { continue }
                    $bioAge = Calculate-PhenoAge -values $markerValues -chronoAge $chronoAge
                    if ($null -eq $bioAge) { continue }
                    $resultValue = $bioAge - $chronoAge
                } elseif ($markerConfig.Formula -eq "[INFLAMMAGING]") {
                    $resultValue = Calculate-InflammAgingScore -values $markerValues
                    if ($null -eq $resultValue) { continue }
                } else {
                    $resultValue = Invoke-Expression $formula
                }
                # v2.28.0: Keine Anzeige von ungültigen Ergebnissen (z. B. BMI ohne Körpergröße -> Division durch 0)
                if ($null -eq $resultValue -or [double]::IsNaN([double]$resultValue) -or [double]::IsInfinity([double]$resultValue)) { continue }
                
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

    # v2.27.1: Alle drei Buttons einheitlich in EINER Reihe unter der Tabelle, gleiche Größe,
    # Gesamtbreite = Tabellenbreite (460 px = 3 x 146 px + 2 x 11 px Abstand).
    $editBtnY = 350; $editBtnW = 146; $editBtnH = 34; $editBtnGap = 11
    $deleteButton = New-Object System.Windows.Forms.Button; $deleteButton.Text = "Auswahl löschen"; $deleteButton.Location = New-Object System.Drawing.Point(10, $editBtnY); $deleteButton.Size = New-Object System.Drawing.Size($editBtnW, $editBtnH)
    $saveButton = New-Object System.Windows.Forms.Button; $saveButton.Text = "Änderungen übernehmen"; $saveButton.Location = New-Object System.Drawing.Point((10 + $editBtnW + $editBtnGap), $editBtnY); $saveButton.Size = New-Object System.Drawing.Size($editBtnW, $editBtnH)
    $closeButton = New-Object System.Windows.Forms.Button; $closeButton.Text = "Abbrechen"; $closeButton.Location = New-Object System.Drawing.Point((10 + 2 * ($editBtnW + $editBtnGap)), $editBtnY); $closeButton.Size = New-Object System.Drawing.Size($editBtnW, $editBtnH)
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

# ---------- v2.26.2: Sichere Wiederherstellung (Einstellungen > "Sicherung wiederherstellen (ZIP)") ----------

function Get-UnreadableJsonFiles {
    <#
    .SYNOPSIS
        v2.26.2: Prüft alle JSON-Dateien eines Wiederherstellungs-Ordners auf Lesbarkeit.
    .DESCRIPTION
        Auto-Backup- und Deinstallations-ZIPs enthalten die Dateien so, wie sie auf
        der Platte liegen - also DPAPI-verschlüsselt für den Windows-Benutzer des
        Quell-PCs. Auf einem anderen PC/Benutzerprofil sind sie nicht lesbar.
        Leere Dateien und der Ordner "lib" werden ignoriert.
    .OUTPUTS
        Array relativer Pfade der NICHT lesbaren Dateien (leer = alles lesbar).
    #>
    param([Parameter(Mandatory)][string]$RootDir)
    $unreadable = New-Object System.Collections.ArrayList
    $rootFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($RootDir).TrimEnd('\', '/')
    $jsonFiles = @(Get-ChildItem -LiteralPath $rootFull -Recurse -File -Filter "*.json" -Force -ErrorAction SilentlyContinue)
    foreach ($jsonFile in $jsonFiles) {
        $relPath = $jsonFile.FullName.Substring($rootFull.Length).TrimStart('\', '/')
        if ((($relPath -split '[\\/]')[0]) -eq 'lib') { continue }
        $isReadable = $false
        try {
            $json = Read-ProtectedJsonFile -Path $jsonFile.FullName
            if ($json) {
                $null = $json | ConvertFrom-Json -ErrorAction Stop
                $isReadable = $true
            } elseif ([string]::IsNullOrWhiteSpace((Get-Content -LiteralPath $jsonFile.FullName -Raw -ErrorAction SilentlyContinue))) {
                $isReadable = $true   # leere Datei - kein Inhalt, der verloren gehen kann
            }
        } catch { $isReadable = $false }
        if (-not $isReadable) { [void]$unreadable.Add($relPath) }
    }
    return $unreadable.ToArray()
}

function Invoke-SafeDataRestore {
    <#
    .SYNOPSIS
        v2.26.2: Stellt den Datenordner transaktional aus einem entpackten Backup wieder her.
    .DESCRIPTION
        Bis v2.26.1 löschte der Modus "Alles löschen" ZUERST den gesamten Datenordner
        und kopierte DANACH das Backup. Ein Fehler beim Kopieren (Datei gesperrt,
        Platte voll, defektes ZIP) hinterließ einen leeren oder halben Datenbestand.
        Ablauf jetzt:
          1. Zielzustand vollständig in einem Staging-Ordner NEBEN dem Datenordner
             aufbauen (merge: aktueller Stand + Backup darüber; clean/overwrite:
             nur Backup). Bis hierhin ist nichts verändert.
          2. Austausch per Umbenennen: bisherige Einträge -> Sicherungsordner,
             Staging-Einträge -> Datenordner.
          3. Bei jedem Fehler: Rückweg (alles zurück an den alten Platz).
          4. Erst nach Erfolg werden Sicherungs- und Staging-Ordner entfernt.
        Der Ordner "lib" (heruntergeladene PDF-Bibliothek, keine Nutzerdaten, ggf.
        geladen/gesperrt) bleibt unangetastet; fehlt er, wird er aus dem Backup
        übernommen. Das laufende Script wird nie verschoben.
    .PARAMETER Mode
        clean     = Datenordner exakt durch das Backup ersetzen
        merge     = Backup ergänzt den Bestand, gleichnamige Dateien werden überschrieben
        overwrite = wie clean (bisher: kein Bestand vorhanden)
    .OUTPUTS
        $true bei Erfolg, sonst Fehler (bestehende Daten unverändert).
    #>
    param(
        [Parameter(Mandatory)][string]$SourceDir,
        [Parameter(Mandatory)][string]$DataDir,
        [ValidateSet('clean', 'merge', 'overwrite')][string]$Mode = 'clean',
        [string]$ProtectedFile = $null
    )
    $excludeNames = @('lib')
    $timestamp  = Get-Date -Format 'yyyyMMdd_HHmmss'
    $sourceFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceDir).TrimEnd('\', '/')
    $dataFull   = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DataDir).TrimEnd('\', '/')
    $stageDir   = "$dataFull.restore_$timestamp"
    $oldDir     = "$dataFull.vorrestore_$timestamp"
    $protectedKey = if ($ProtectedFile) { Get-BtPathKey -Path $ProtectedFile } else { $null }
    $movedOut = New-Object System.Collections.ArrayList   # Namen: Datenordner -> Sicherungsordner
    $movedIn  = New-Object System.Collections.ArrayList   # Namen: Staging -> Datenordner

    # Datei/Ordner umbenennen, 3 Versuche (kurzzeitige Sperren durch Virenscanner/Sync)
    $moveEntry = {
        param([string]$Source, [string]$Destination)
        for ($moveAttempt = 1; $moveAttempt -le 3; $moveAttempt++) {
            try {
                if ([System.IO.Directory]::Exists($Source)) { [System.IO.Directory]::Move($Source, $Destination) }
                else { [System.IO.File]::Move($Source, $Destination) }
                return
            } catch { if ($moveAttempt -eq 3) { throw }; Start-Sleep -Milliseconds 200 }
        }
    }
    # Dateibaum kopieren (deterministisch, ohne die Verschachtelungs-Eigenheiten von Copy-Item -Recurse)
    $copyTree = {
        param([string]$From, [string]$To, [string[]]$SkipTopNames)
        $fromFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($From).TrimEnd('\', '/')
        foreach ($srcFile in @(Get-ChildItem -LiteralPath $fromFull -Recurse -File -Force -ErrorAction Stop)) {
            $relPath = $srcFile.FullName.Substring($fromFull.Length).TrimStart('\', '/')
            if ($SkipTopNames -and ($SkipTopNames -contains (($relPath -split '[\\/]')[0]))) { continue }
            if ($protectedKey -and ((Get-BtPathKey -Path $srcFile.FullName) -eq $protectedKey)) { continue }
            $targetFile = Join-Path $To $relPath
            $targetDir  = Split-Path -Path $targetFile -Parent
            if (-not (Test-Path -LiteralPath $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
            Copy-Item -LiteralPath $srcFile.FullName -Destination $targetFile -Force -ErrorAction Stop
        }
    }

    # Schutz: Datenordner direkt im Laufwerksstamm (z. B. "D:\") wird nicht unterstützt -
    # Staging/Sicherung müssen NEBEN dem Datenordner liegen können.
    if ([string]::IsNullOrEmpty((Split-Path -Path $dataFull -Parent))) {
        throw "Der Datenordner '$DataDir' liegt im Laufwerksstamm - Wiederherstellung nicht möglich. Bestehende Daten wurden NICHT verändert."
    }

    try {
        # 1. Zielzustand im Staging-Ordner aufbauen (bestehende Daten unberührt)
        New-Item -ItemType Directory -Path $stageDir -Force -ErrorAction Stop | Out-Null
        if ($Mode -eq 'merge' -and (Test-Path -LiteralPath $dataFull)) {
            & $copyTree $dataFull $stageDir $excludeNames
        }
        & $copyTree $sourceFull $stageDir @()

        # 2a. Bisherige Einträge in den Sicherungsordner verschieben
        if (-not (Test-Path -LiteralPath $dataFull)) { New-Item -ItemType Directory -Path $dataFull -Force -ErrorAction Stop | Out-Null }
        New-Item -ItemType Directory -Path $oldDir -Force -ErrorAction Stop | Out-Null
        foreach ($entry in @(Get-ChildItem -LiteralPath $dataFull -Force -ErrorAction Stop)) {
            if ($excludeNames -contains $entry.Name) { continue }
            if ($protectedKey -and ((Get-BtPathKey -Path $entry.FullName) -eq $protectedKey)) { continue }
            try {
                & $moveEntry $entry.FullName (Join-Path $oldDir $entry.Name)
            } catch {
                throw "'$($entry.Name)' im Datenordner ist gesperrt (z. B. geöffnetes Dokument, Explorer-Vorschau oder Synchronisierung). Wiederherstellung abgebrochen - bestehende Daten wurden NICHT verändert.`n`nDetails: $($_.Exception.Message)"
            }
            [void]$movedOut.Add($entry.Name)
        }

        # 2b. Staging-Einträge in den Datenordner übernehmen
        foreach ($entry in @(Get-ChildItem -LiteralPath $stageDir -Force -ErrorAction Stop)) {
            $destination = Join-Path $dataFull $entry.Name
            if (Test-Path -LiteralPath $destination) { continue }   # z. B. "lib" bleibt bestehen
            & $moveEntry $entry.FullName $destination
            [void]$movedIn.Add($entry.Name)
        }

        # 3. Erfolg: Sicherungs- und Staging-Ordner entfernen
        Remove-Item -LiteralPath $oldDir -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $stageDir -Recurse -Force -ErrorAction SilentlyContinue
        return $true
    } catch {
        $originalError = $_
        # Rückweg: neue Einträge zurück ins Staging, alte Einträge zurück in den Datenordner
        $rollbackFailed = $false
        for ($i = $movedIn.Count - 1; $i -ge 0; $i--) {
            try { & $moveEntry (Join-Path $dataFull $movedIn[$i]) (Join-Path $stageDir $movedIn[$i]) } catch { $rollbackFailed = $true }
        }
        for ($i = $movedOut.Count - 1; $i -ge 0; $i--) {
            try { & $moveEntry (Join-Path $oldDir $movedOut[$i]) (Join-Path $dataFull $movedOut[$i]) } catch { $rollbackFailed = $true }
        }
        if (Test-Path -LiteralPath $stageDir) { Remove-Item -LiteralPath $stageDir -Recurse -Force -ErrorAction SilentlyContinue }
        if ($rollbackFailed) {
            throw "Wiederherstellung fehlgeschlagen UND der alte Stand konnte nicht vollständig zurückgeholt werden. Bitte NICHTS löschen - der vorherige Stand liegt unter:`n$oldDir`n`nUrsache: $($originalError.Exception.Message)"
        }
        # Sicherungsordner nur entfernen, wenn er nach dem Rückweg wirklich leer ist
        if ((Test-Path -LiteralPath $oldDir) -and (@(Get-ChildItem -LiteralPath $oldDir -Force -ErrorAction SilentlyContinue).Count -eq 0)) {
            Remove-Item -LiteralPath $oldDir -Force -ErrorAction SilentlyContinue
        }
        throw $originalError
    }
}

function Move-DataDirectorySafe {
    <#
    .SYNOPSIS
        v2.27.2: Verschiebt den Datenordner sicher an einen neuen Ort (Einstellungen > "Datenpfad ändern...").
    .DESCRIPTION
        Bis v2.27.1: Move-Item "<Daten>\*" -> Ziel, danach Remove-Item des alten Ordners -
        ohne Prüfung, ob das Verschieben vollständig war. Move-Item verschiebt Ordner nur
        innerhalb DESSELBEN Laufwerks (Microsoft Learn, Move-Item > Notes); bei einem
        anderen Laufwerk blieb der Ordner "data" zurück und wurde anschließend gelöscht.
        Ablauf jetzt:
          1. Alle Dateien ins Ziel KOPIEREN (laufwerksübergreifend möglich)
          2. Jede Kopie prüfen (vorhanden, gleiche Größe, gleicher SHA256-Hash)
          3. $SwitchAction ausführen (Registry-Eintrag DataPath umstellen)
          4. ERST DANACH den alten Ordner löschen (gesperrte Reste bleiben stehen
             und werden gemeldet - die Daten liegen dann bereits geprüft im Ziel)
        Schlägt 1-3 fehl, werden die angelegten Kopien entfernt; der alte Datenordner
        bleibt unverändert und aktiv.
    .OUTPUTS
        PSCustomObject @{ FileCount; OldDirRemoved }
    #>
    param(
        [Parameter(Mandatory)][string]$SourceDir,
        [Parameter(Mandatory)][string]$TargetDir,
        [Parameter(Mandatory)][scriptblock]$SwitchAction
    )
    $sourceFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($SourceDir).TrimEnd('\', '/')
    $targetFull = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($TargetDir).TrimEnd('\', '/')
    $sourceKey = Get-BtPathKey -Path $sourceFull
    $targetKey = Get-BtPathKey -Path $targetFull
    $sep = [string][System.IO.Path]::DirectorySeparatorChar
    if ($targetKey -eq $sourceKey -or $targetKey.StartsWith($sourceKey + $sep)) {
        throw "Der Zielordner darf nicht im aktuellen Datenordner liegen. Es wurde nichts verändert."
    }

    $createdFiles = New-Object System.Collections.ArrayList
    $createdDirs  = New-Object System.Collections.ArrayList
    try {
        # 1. Kopieren
        if (-not (Test-Path -LiteralPath $targetFull)) {
            New-Item -ItemType Directory -Path $targetFull -Force -ErrorAction Stop | Out-Null
            [void]$createdDirs.Add($targetFull)
        }
        $sourceFiles = @(Get-ChildItem -LiteralPath $sourceFull -Recurse -File -Force -ErrorAction Stop)
        foreach ($srcFile in $sourceFiles) {
            $relPath = $srcFile.FullName.Substring($sourceFull.Length).TrimStart('\', '/')
            $targetFile = Join-Path $targetFull $relPath
            $targetDir = Split-Path -Path $targetFile -Parent
            if (-not (Test-Path -LiteralPath $targetDir)) {
                # Alle fehlenden Ebenen merken (oben -> unten), damit das Aufräumen sie wieder entfernt
                $missingDirs = New-Object System.Collections.ArrayList
                $probeDir = $targetDir
                while ($probeDir -and -not (Test-Path -LiteralPath $probeDir)) { $missingDirs.Insert(0, $probeDir); $probeDir = Split-Path -Path $probeDir -Parent }
                New-Item -ItemType Directory -Path $targetDir -Force -ErrorAction Stop | Out-Null
                foreach ($missingDir in $missingDirs) { [void]$createdDirs.Add($missingDir) }
            }
            Copy-Item -LiteralPath $srcFile.FullName -Destination $targetFile -Force -ErrorAction Stop
            [void]$createdFiles.Add($targetFile)
        }

        # 2. Prüfen (Anzahl, Größe, SHA256)
        foreach ($srcFile in $sourceFiles) {
            $relPath = $srcFile.FullName.Substring($sourceFull.Length).TrimStart('\', '/')
            $targetFile = Join-Path $targetFull $relPath
            if (-not (Test-Path -LiteralPath $targetFile)) { throw "Kopie fehlt: $relPath" }
            if ((Get-Item -LiteralPath $targetFile -Force).Length -ne $srcFile.Length) { throw "Kopie unvollständig (Größe): $relPath" }
            $srcHash = (Get-FileHash -LiteralPath $srcFile.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
            $dstHash = (Get-FileHash -LiteralPath $targetFile -Algorithm SHA256 -ErrorAction Stop).Hash
            if ($srcHash -ne $dstHash) { throw "Kopie fehlerhaft (Prüfsumme): $relPath" }
        }

        # 3. Umschalten (z. B. Registry DataPath)
        & $SwitchAction
    } catch {
        $originalError = $_
        # Angelegte Kopien entfernen - der alte Datenordner bleibt unverändert aktiv
        for ($i = $createdFiles.Count - 1; $i -ge 0; $i--) { Remove-Item -LiteralPath $createdFiles[$i] -Force -ErrorAction SilentlyContinue }
        for ($i = $createdDirs.Count - 1; $i -ge 0; $i--) {
            if ((Test-Path -LiteralPath $createdDirs[$i]) -and @(Get-ChildItem -LiteralPath $createdDirs[$i] -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                Remove-Item -LiteralPath $createdDirs[$i] -Force -ErrorAction SilentlyContinue
            }
        }
        throw $originalError
    }

    # 4. Alten Ordner entfernen (best effort - Daten liegen geprüft im Ziel)
    Remove-Item -LiteralPath $sourceFull -Recurse -Force -ErrorAction SilentlyContinue
    return [PSCustomObject]@{ FileCount = $sourceFiles.Count; OldDirRemoved = (-not (Test-Path -LiteralPath $sourceFull)) }
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
    # v2.27.2 BUGFIX (Datenverlust): Umzug über Move-DataDirectorySafe (kopieren -> prüfen ->
    # Registry umstellen -> erst dann alten Ordner löschen). Bisher: Move-Item + Remove-Item ohne
    # Prüfung - bei einem anderen Laufwerk wurde der Ordner "data" nicht verschoben, aber gelöscht.
    $moveDataButton.Add_Click({
        $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $folderBrowser.Description = "Wählen Sie ein neues, leeres Verzeichnis für die Anwendungsdaten."
        if ($folderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $newPath = $folderBrowser.SelectedPath
            if ($newPath -eq $script:dataDir) { [System.Windows.Forms.MessageBox]::Show("Der ausgewählte Pfad ist bereits der aktuelle Speicherort.", "Information", "OK", "Information"); return }
            if (Get-ChildItem -Path $newPath) { [System.Windows.Forms.MessageBox]::Show("Das ausgewählte Verzeichnis ist nicht leer. Bitte wählen Sie einen leeren Ordner aus, um Datenkonflikte zu vermeiden.", "Fehler", "OK", "Error"); return }
            try {
                $popupForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
                $moveResult = Move-DataDirectorySafe -SourceDir $script:dataDir -TargetDir $newPath -SwitchAction {
                    Set-ItemProperty -Path $script:regPath -Name "DataPath" -Value $newPath -ErrorAction Stop
                }
                $popupForm.Cursor = [System.Windows.Forms.Cursors]::Default
                $moveMsg = "Der Datenpfad wurde erfolgreich nach `"$newPath`" geändert ($($moveResult.FileCount) Dateien kopiert und geprüft)."
                if (-not $moveResult.OldDirRemoved) {
                    $moveMsg += "`n`nHinweis: Der alte Ordner konnte nicht vollständig entfernt werden (Datei in Benutzung):`n$($script:dataDir)`nAlle Daten liegen geprüft am neuen Ort - den alten Ordner bitte nach dem Neustart manuell löschen."
                }
                $moveMsg += "`n`nDas Programm wird nun neu gestartet."
                [System.Windows.Forms.MessageBox]::Show($moveMsg, "Erfolg", "OK", "Information")
                # v2.27.2: Script-Pfad wie bei der Wiederherstellung ermitteln ($MyInvocation.MyCommand.Path ist im Event-Handler leer)
                $scriptPath = $PSCommandPath
                if (-not $scriptPath) { $scriptPath = $MyInvocation.ScriptName }
                if ($scriptPath -and (Test-Path $scriptPath)) {
                    Start-Process powershell -ArgumentList "-File `"$scriptPath`""
                } else {
                    [System.Windows.Forms.MessageBox]::Show("Bitte starten Sie den Blood-Tracker manuell neu.", "Neustart", "OK", "Information")
                }
                $popupForm.Close(); $mainForm.Close()
            } catch {
                $popupForm.Cursor = [System.Windows.Forms.Cursors]::Default
                [System.Windows.Forms.MessageBox]::Show("Fehler beim Verschieben der Daten: $($_.Exception.Message)`n`nEs wurde nichts verändert - die Daten liegen weiterhin unter:`n$($script:dataDir)", "Fehler", "OK", "Error")
            }
        }
    })
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

            # Prüfe ob der UserData-Ordner einen Unterordner mit App-Name enthält (Backup-Struktur: UserData\PSC.Blood-Tracker\...)
            # (v2.26.2: aus Schritt 4c vorgezogen, da die Lesbarkeitsprüfung die Quelle braucht)
            $innerAppDir = Join-Path -Path $userDataSource -ChildPath "PSC.Blood-Tracker"
            $restoreSource = if (Test-Path $innerAppDir) { $innerAppDir } else { $userDataSource }

            # Schritt 2b (v2.26.2): Sind die Backup-Dateien auf DIESEM PC/Benutzerprofil lesbar?
            # ZIP-Backups enthalten DPAPI-verschlüsselte Dateien - von einem anderen PC/Profil
            # wären sie unlesbar und würden lesbare Daten ersetzen.
            $unreadableFiles = @(Get-UnreadableJsonFiles -RootDir $restoreSource)
            if ($unreadableFiles.Count -gt 0) {
                $shownList = ($unreadableFiles | Select-Object -First 5 | ForEach-Object { "- $_" }) -join "`n"
                $moreText = if ($unreadableFiles.Count -gt 5) { "`n- ... und $($unreadableFiles.Count - 5) weitere" } else { "" }
                $continueChoice = [System.Windows.Forms.MessageBox]::Show(
                    "Das Backup enthält $($unreadableFiles.Count) Datei(en), die auf diesem PC bzw. mit diesem Windows-Benutzer NICHT lesbar sind:`n`n$shownList$moreText`n`nUrsache meist: Das Backup stammt von einem anderen PC oder Benutzerprofil (Windows-Verschlüsselung DPAPI). Für einen PC-Wechsel bitte den portablen Export (.btbackup) verwenden.`n`nTrotzdem wiederherstellen? (nicht empfohlen)`n`n[Ja] = Trotzdem wiederherstellen`n[Nein] = Abbrechen, nichts verändern",
                    "Backup auf diesem PC nicht lesbar",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Warning,
                    [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
                if ($continueChoice -ne "Yes") {
                    Remove-Item -Path $tempDir -Recurse -Force -ErrorAction SilentlyContinue
                    return
                }
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
            # v2.26.2 BUGFIX (Datenverlust): Bisher wurde im Modus "clean" ZUERST alles
            # gelöscht und DANACH kopiert. Jetzt transaktional (Staging + Austausch +
            # Rückweg bei Fehler) - Details siehe Invoke-SafeDataRestore.
            # 4a-4c: clean / merge / overwrite (Verschachtelte und flache Struktur über $restoreSource)
            $runningScript = $PSCommandPath
            if (-not $runningScript) { $runningScript = $MyInvocation.ScriptName }
            $null = Invoke-SafeDataRestore -SourceDir $restoreSource -DataDir $script:dataDir -Mode $restoreMode -ProtectedFile $runningScript

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
            # v2.31.3: @() - ein Einzelergebnis hat in Windows PowerShell 5.1 kein .Count
            $calcData = @(Get-CalculatedValuesForMarker -markerName $marker.Name)
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

# ╔══════════════════════════════════════════════════════════════════════════════════════╗
# ║ v2.34.0: RELEASE-NOTES - Logik und Reiter (Integration A; v2.34.1: Modul v0.2.0)     ║
# ║ Quelle: Modul BloodTracker_ReleaseNotes v0.2.0 - isoliert getestet, unveraendert     ║
# ║ uebernommen. Keine Abhaengigkeit von Funktionen in Show-DataEntryForm.               ║
# ║ Ab v2.34.0 ist DIESES Script massgeblich; das Modul dient nur dem isolierten Test.   ║
# ╚══════════════════════════════════════════════════════════════════════════════════════╝
# ---------------------------------------------------------------------------------
# Konstanten
# ---------------------------------------------------------------------------------
$script:ReleaseNoteTypes = @('Neu', 'Bugfix', 'Anpassung', 'Sicherheit', 'Hinweis')

# Gemeinsames UI-Theme für Release-Notes und Dashboard (Spec Abschnitt 9).
# Im Prototyp aus Dashboard_Theme.json überschreibbar, im Haupt-Script fest.
$script:HealthUiTheme = @{
    Page          = '#F4F4F2'
    Card          = '#FFFFFF'
    Border        = '#E2E1DC'
    TextPrimary   = '#0B0B0B'
    TextSecondary = '#52514E'
    TextMuted     = '#6B6A65'
    Link          = '#1C5CAB'
    Accent        = '#2A78D6'
    SelectionBack = '#E6EEF9'
    Highlight     = '#FDEDC6'
    StatusGood    = '#0CA30C'
    StatusWarning = '#FAB219'
    StatusCritical = '#D03B3B'
    StatusNone    = '#9A9993'
    FontFamily    = 'Segoe UI'
}

# ---------------------------------------------------------------------------------
# Hilfsfunktionen (ohne Oberfläche, testbar)
# ---------------------------------------------------------------------------------
function Get-RnField {
    <# Liest ein Feld aus Hashtable/[ordered] ODER PSCustomObject (PS-5.1-sicher). #>
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        return $null
    }
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-ReleaseVersion {
    <#
    .SYNOPSIS
        Text → [version]. Nimmt nur den führenden Zahlenteil ("2.34.0 BETA-TEST" → 2.34.0).
    .OUTPUTS
        [version] oder $null.
    #>
    param($Text)
    if ($null -eq $Text) { return $null }
    $match = [regex]::Match(([string]$Text).Trim(), '^\d+(\.\d+){1,3}')
    if (-not $match.Success) { return $null }
    $parsed = $null
    if ([version]::TryParse($match.Value, [ref]$parsed)) { return $parsed }
    return $null
}

function ConvertTo-RnDate {
    <# 'yyyy-MM-dd' → [datetime] oder $null (leer/ungültig). #>
    param($Text)
    if ([string]::IsNullOrWhiteSpace([string]$Text)) { return $null }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact(([string]$Text).Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    return $null
}

function Format-ReleaseDate {
    <# Anzeige 'TT.MM.JJJJ'; ohne (gültiges) Datum leer – dann zählt nur die Version (v0.2.0). #>
    param($Text)
    $parsed = ConvertTo-RnDate -Text $Text
    if ($null -eq $parsed) { return '' }
    return $parsed.ToString('dd.MM.yyyy')
}

function Test-RnPlaceholder {
    <#
    .SYNOPSIS
        $true, wenn ein Text leer oder ein Platzhalter ist.
    .DESCRIPTION
        Platzhalter: leer, nur Satzzeichen ("…", "..."), Spitzklammer-Platzhalter ‹…›,
        oder die Wörter TODO / TBD / FIXME / PLATZHALTER. "..." INNERHALB eines echten
        Textes (z. B. Schaltfläche "Datenpfad ändern...") ist KEIN Platzhalter.
    #>
    param($Text)
    $value = [string]$Text
    if ([string]::IsNullOrWhiteSpace($value)) { return $true }
    $trimmed = $value.Trim()
    if ($trimmed -match '^[\W_]+$') { return $true }
    if ($trimmed.IndexOf([char]0x2039) -ge 0 -or $trimmed.IndexOf([char]0x203A) -ge 0) { return $true }
    if ($trimmed -match '(?i)\b(TODO|TBD|FIXME|PLATZHALTER)\b') { return $true }
    return $false
}

function Get-ReleaseNotesSorted {
    <#
    .SYNOPSIS
        Sortiert numerisch nach Version, neueste zuerst (2.10.0 vor 2.9.0).
        Ungültige Versionen stehen am Ende (die Konsistenzprüfung meldet sie).
    .OUTPUTS
        Array (Aufrufer IMMER mit @() übernehmen).
    #>
    param($Notes)
    $valid = @(); $invalid = @()
    foreach ($note in @($Notes)) {
        if ($null -eq $note) { continue }
        if ($null -ne (ConvertTo-ReleaseVersion -Text (Get-RnField $note 'Version'))) { $valid += , $note } else { $invalid += , $note }
    }
    $sorted = @($valid | Sort-Object -Property @{ Expression = { ConvertTo-ReleaseVersion -Text (Get-RnField $_ 'Version') } } -Descending)
    return @($sorted + $invalid)
}

function Test-RnTextMatch {
    <# Groß-/Kleinschreibung egal, Teilstring. Leerer Suchtext = Treffer. #>
    param([string]$Text, [string]$Search)
    if ([string]::IsNullOrWhiteSpace($Search)) { return $true }
    if ([string]::IsNullOrEmpty($Text)) { return $false }
    return ($Text.IndexOf($Search.Trim(), [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
}

function Get-ReleaseNotesView {
    <#
    .SYNOPSIS
        Wendet Suche und Typ-Filter an (Spec 8, Regeln 1–4).
    .DESCRIPTION
        1. Suche und Typ-Filter wirken gemeinsam.
        2. Treffer in Version oder Titel qualifiziert die Version; der Typ-Filter begrenzt
           weiterhin die sichtbaren Einträge.
        3. Treffer nur im Eintragstext: nur die passenden Einträge des gewählten Typs.
        4. Nur Versionen mit mindestens einem verbleibenden Eintrag.
    .PARAMETER TypeFilter
        'Alle' (oder leer) bzw. einer der Typen aus $script:ReleaseNoteTypes.
    .OUTPUTS
        Array von PSCustomObject { Version; VersionObj; Datum; DatumText; Titel; Eintraege; IsCurrent }
        in Anzeige-Reihenfolge. Aufrufer IMMER mit @() übernehmen.
    #>
    param($Notes, [string]$SearchText = '', [string]$TypeFilter = 'Alle', $CurrentVersion = $null)
    $currentVer = ConvertTo-ReleaseVersion -Text $CurrentVersion
    $filterAll = ([string]::IsNullOrWhiteSpace($TypeFilter) -or $TypeFilter -eq 'Alle')
    $search = if ($null -eq $SearchText) { '' } else { $SearchText.Trim() }
    $result = @()
    foreach ($note in (Get-ReleaseNotesSorted -Notes $Notes)) {
        $noteVersion = [string](Get-RnField $note 'Version')
        $title   = [string](Get-RnField $note 'Titel')
        $entries = @(Get-RnField $note 'Eintraege' | Where-Object { $null -ne $_ })
        # Typ-Filter
        $typed = @($entries | Where-Object { $filterAll -or ([string](Get-RnField $_ 'Typ')) -eq $TypeFilter })
        # Suche
        $headerHit = (Test-RnTextMatch -Text $noteVersion -Search $search) -or (Test-RnTextMatch -Text $title -Search $search)
        if ($headerHit) { $visible = $typed }
        else { $visible = @($typed | Where-Object { Test-RnTextMatch -Text ([string](Get-RnField $_ 'Text')) -Search $search }) }
        if (@($visible).Count -eq 0) { continue }
        $versionObj = ConvertTo-ReleaseVersion -Text $noteVersion
        $result += , ([PSCustomObject]@{
            Version    = $noteVersion
            VersionObj = $versionObj
            Datum      = [string](Get-RnField $note 'Datum')
            DatumText  = Format-ReleaseDate -Text (Get-RnField $note 'Datum')
            Titel      = $title
            Eintraege  = @($visible)
            IsCurrent  = ($null -ne $currentVer -and $null -ne $versionObj -and $versionObj -eq $currentVer)
        })
    }
    return $result
}

function Resolve-ReleaseNotesSelection {
    <#
    .SYNOPSIS
        Spec 8, Regel 5: bisherige Auswahl bleibt, wenn noch enthalten; sonst erster Treffer.
    .PARAMETER PreviousKey
        v0.2.0: Schlüssel des bisher gewählten Listeneintrags (Get-ReleaseNotesGroups .Key) –
        hat Vorrang vor -PreviousVersion.
    .OUTPUTS
        Index in -ViewItems oder -1 (keine Treffer).
    #>
    param($ViewItems, [string]$PreviousVersion, [string]$PreviousKey = '')
    $items = @($ViewItems)
    if ($items.Count -eq 0) { return -1 }
    if (-not [string]::IsNullOrEmpty($PreviousKey)) {
        for ($i = 0; $i -lt $items.Count; $i++) { if ([string]$items[$i].Key -eq $PreviousKey) { return $i } }
    }
    if (-not [string]::IsNullOrEmpty($PreviousVersion)) {
        for ($i = 0; $i -lt $items.Count; $i++) { if ($items[$i].Version -eq $PreviousVersion) { return $i } }
    }
    return 0
}

function Get-ReleaseNotesGroups {
    <#
    .SYNOPSIS
        v0.2.0: Fasst die Anzeige-Einträge (Get-ReleaseNotesView) für die Liste zusammen –
        alle Versionen mit GLEICHEM Datum bilden einen Listeneintrag. Versionen ohne (gültiges)
        Datum bleiben einzeln. Nur Anzeige; das Datenmodell bleibt ein Eintrag je Version.
    .DESCRIPTION
        Reihenfolge = Position der neuesten Version je Gruppe (die Eingabe ist numerisch absteigend
        sortiert, die erste Version einer Gruppe ist ihre neueste). Versions enthält nur die nach
        Suche/Typ-Filter sichtbaren Versionen, Count deren Anzahl.
    .OUTPUTS
        Array von PSCustomObject { Key; Version; Datum; DatumText; Titel; Versions[]; Count;
        IsCurrent; HasDate }. Key: "D:JJJJ-MM-TT" bzw. "V:<Version>". Aufrufer IMMER mit @() übernehmen.
    #>
    param($ViewItems)
    $order = New-Object System.Collections.ArrayList
    $members = @{}
    foreach ($item in @($ViewItems | Where-Object { $null -ne $_ })) {
        $date = ConvertTo-RnDate -Text $item.Datum
        if ($null -ne $date) { $key = 'D:' + $date.ToString('yyyy-MM-dd') } else { $key = 'V:' + [string]$item.Version }
        if (-not $members.ContainsKey($key)) {
            $members[$key] = New-Object System.Collections.ArrayList
            [void]$order.Add($key)
        }
        [void]$members[$key].Add($item)
    }
    $groups = @()
    foreach ($key in $order) {
        $versions = @($members[$key])
        $lead = $versions[0]
        $groups += , ([PSCustomObject]@{
            Key       = $key
            Version   = $lead.Version
            Datum     = $lead.Datum
            DatumText = $lead.DatumText
            Titel     = $lead.Titel
            Versions  = $versions
            Count     = $versions.Count
            IsCurrent = (@($versions | Where-Object { $_.IsCurrent }).Count -gt 0)
            HasDate   = $key.StartsWith('D:')
        })
    }
    return $groups
}

function Get-RnGroupItemHeight {
    <# v0.2.0: Höhe eines Listeneintrags – mit Datum 3 Zeilen (58 px), ohne Datum 2 Zeilen (42 px). #>
    param($Group)
    if ($null -ne $Group -and $Group.HasDate) { return 58 }
    return 42
}

function Get-ReleaseEntriesGrouped {
    <#
    .SYNOPSIS
        Einträge einer Version nach Typ gruppiert, feste Typ-Reihenfolge, innerhalb des
        Typs Original-Reihenfolge. Unbekannte Typen am Ende unter "Sonstiges".
    .OUTPUTS
        Array von PSCustomObject { Typ; Texte[] }.
    #>
    param($Entries)
    $groups = @()
    $all = @($Entries | Where-Object { $null -ne $_ })
    foreach ($type in $script:ReleaseNoteTypes) {
        $texts = @($all | Where-Object { ([string](Get-RnField $_ 'Typ')) -eq $type } | ForEach-Object { [string](Get-RnField $_ 'Text') })
        if ($texts.Count -gt 0) { $groups += , ([PSCustomObject]@{ Typ = $type; Texte = $texts }) }
    }
    $other = @($all | Where-Object { $script:ReleaseNoteTypes -notcontains ([string](Get-RnField $_ 'Typ')) } | ForEach-Object { [string](Get-RnField $_ 'Text') })
    if ($other.Count -gt 0) { $groups += , ([PSCustomObject]@{ Typ = 'Sonstiges'; Texte = $other }) }
    return $groups
}

function Get-ScriptHeaderVersion {
    <# Liest "# Version x.y.z" aus den ersten 10 Zeilen einer Script-Datei. $null, wenn nicht gefunden. #>
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return $null }
    foreach ($line in @(Get-Content -LiteralPath $Path -TotalCount 10 -ErrorAction SilentlyContinue)) {
        if ($line -match '^\s*#\s*Version\s+(\d+(\.\d+){1,3})') { return $Matches[1] }
    }
    return $null
}

function Get-RnEntrySignature {
    <# Vergleichsschlüssel eines Versionseintrags (Titel + Einträge) für die Release-Prüfung. #>
    param($Note)
    $parts = @([string](Get-RnField $Note 'Titel'))
    foreach ($entry in @(Get-RnField $Note 'Eintraege' | Where-Object { $null -ne $_ })) {
        $parts += ('{0}|{1}' -f [string](Get-RnField $entry 'Typ'), [string](Get-RnField $entry 'Text'))
    }
    return ($parts -join "`n")
}

function Test-ReleaseNotesConsistency {
    <#
    .SYNOPSIS
        Prüft die Release-Notes (Spec 12.9). Ändert nichts.
    .PARAMETER CurrentVersion
        Laufende Version ($Version), Zusätze wie " BETA-TEST" werden ignoriert.
    .PARAMETER HeaderVersion
        Version aus der Kopfzeile "# Version x.y.z" (optional, Get-ScriptHeaderVersion).
    .PARAMETER PreviousNotes
        Nur Release-Check: Release-Notes des vorherigen Dateistands (veröffentlichte Einträge
        dürfen sich nicht ändern; ein vorher LEERES Datum darf ergänzt werden → Warnung).
    .PARAMETER Today
        Stichtag für die Zukunftsprüfung (Tests); Standard heute.
    .OUTPUTS
        PSCustomObject { IsValid; Errors[]; Warnings[] }
    #>
    param($Notes, $CurrentVersion, $HeaderVersion = $null, $PreviousNotes = $null, [datetime]$Today = (Get-Date).Date)
    $errors = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    $list = @($Notes | Where-Object { $null -ne $_ })
    if ($list.Count -eq 0) {
        [void]$errors.Add('Keine Release-Notes vorhanden.')
        return [PSCustomObject]@{ IsValid = $false; Errors = @($errors); Warnings = @($warnings) }
    }

    # 1. Versionen gültig, eindeutig, numerisch absteigend (in der gespeicherten Reihenfolge)
    $seen = @{}
    $previousVer = $null
    for ($i = 0; $i -lt $list.Count; $i++) {
        $versionText = [string](Get-RnField $list[$i] 'Version')
        $ver = ConvertTo-ReleaseVersion -Text $versionText
        if ($null -eq $ver -or $versionText.Trim() -ne $ver.ToString()) {
            [void]$errors.Add("Eintrag $($i + 1): ungültige Version '$versionText' (erwartet z. B. 2.34.0).")
            continue
        }
        $key = $ver.ToString()
        if ($seen.ContainsKey($key)) { [void]$errors.Add("Version $key ist doppelt vorhanden.") } else { $seen[$key] = $true }
        if ($null -ne $previousVer -and $ver -ge $previousVer) {
            [void]$errors.Add("Reihenfolge: $key steht unter $previousVer – Versionen müssen numerisch absteigend sortiert sein.")
        }
        $previousVer = $ver
    }

    # 2. Laufende Version, Kopfzeile und oberster Eintrag stimmen überein
    $top = $list[0]
    $topVer = ConvertTo-ReleaseVersion -Text (Get-RnField $top 'Version')
    $curVer = ConvertTo-ReleaseVersion -Text $CurrentVersion
    if ($null -eq $curVer) { [void]$errors.Add("Laufende Version '$CurrentVersion' ist ungültig.") }
    elseif ($null -ne $topVer -and $curVer -ne $topVer) { [void]$errors.Add("Laufende Version $curVer ≠ oberster Release-Notes-Eintrag $topVer.") }
    if ($null -ne $HeaderVersion) {
        $headVer = ConvertTo-ReleaseVersion -Text $HeaderVersion
        if ($null -eq $headVer) { [void]$errors.Add("Kopfzeilen-Version '$HeaderVersion' ist ungültig.") }
        elseif ($null -ne $curVer -and $headVer -ne $curVer) { [void]$errors.Add("Kopfzeile '# Version $headVer' ≠ laufende Version $curVer.") }
    }

    # 3. Oberster Eintrag vollständig
    $topDate = ConvertTo-RnDate -Text (Get-RnField $top 'Datum')
    if ($null -eq $topDate) { [void]$errors.Add("Oberster Eintrag: Datum fehlt oder ist ungültig (Format JJJJ-MM-TT).") }
    elseif ($topDate -gt $Today) { [void]$warnings.Add("Oberster Eintrag: Datum $($topDate.ToString('dd.MM.yyyy')) liegt in der Zukunft – Datum = Auslieferungsdatum.") }
    if (Test-RnPlaceholder -Text (Get-RnField $top 'Titel')) { [void]$errors.Add("Oberster Eintrag: Titel fehlt oder ist ein Platzhalter.") }
    $topEntries = @(Get-RnField $top 'Eintraege' | Where-Object { $null -ne $_ -and -not (Test-RnPlaceholder -Text (Get-RnField $_ 'Text')) })
    if ($topEntries.Count -eq 0) { [void]$errors.Add("Oberster Eintrag: mindestens ein ausgefüllter Eintrag erforderlich.") }

    # 4.–6. Typen, Platzhalter, Datumsformat aller Einträge
    for ($i = 0; $i -lt $list.Count; $i++) {
        $label = "Version $([string](Get-RnField $list[$i] 'Version'))"
        $dateText = [string](Get-RnField $list[$i] 'Datum')
        if (-not [string]::IsNullOrWhiteSpace($dateText) -and $null -eq (ConvertTo-RnDate -Text $dateText)) {
            [void]$errors.Add("${label}: Datum '$dateText' ungültig (Format JJJJ-MM-TT oder leer).")
        }
        if ($i -gt 0 -and (Test-RnPlaceholder -Text (Get-RnField $list[$i] 'Titel'))) { [void]$errors.Add("${label}: Titel fehlt oder ist ein Platzhalter.") }
        $entries = @(Get-RnField $list[$i] 'Eintraege' | Where-Object { $null -ne $_ })
        if ($entries.Count -eq 0) { [void]$errors.Add("${label}: keine Einträge.") }
        $n = 0
        foreach ($entry in $entries) {
            $n++
            $type = [string](Get-RnField $entry 'Typ')
            if ($script:ReleaseNoteTypes -notcontains $type) { [void]$errors.Add("${label}, Eintrag ${n}: unbekannter Typ '$type' (erlaubt: $($script:ReleaseNoteTypes -join ', ')).") }
            if (Test-RnPlaceholder -Text (Get-RnField $entry 'Text')) { [void]$errors.Add("${label}, Eintrag ${n}: Text fehlt oder ist ein Platzhalter.") }
        }
    }

    # 7. Nur Release-Check: veröffentlichte Einträge unverändert
    if ($null -ne $PreviousNotes) {
        $current = @{}
        foreach ($note in $list) {
            $v = ConvertTo-ReleaseVersion -Text (Get-RnField $note 'Version')
            if ($null -ne $v) { $current[$v.ToString()] = $note }
        }
        foreach ($old in @($PreviousNotes | Where-Object { $null -ne $_ })) {
            $oldVer = ConvertTo-ReleaseVersion -Text (Get-RnField $old 'Version')
            if ($null -eq $oldVer) { continue }
            $key = $oldVer.ToString()
            if (-not $current.ContainsKey($key)) { [void]$errors.Add("Veröffentlichte Version $key wurde entfernt."); continue }
            $new = $current[$key]
            if ((Get-RnEntrySignature -Note $old) -ne (Get-RnEntrySignature -Note $new)) {
                [void]$errors.Add("Veröffentlichte Version ${key}: Titel oder Einträge wurden geändert.")
            }
            $oldDate = [string](Get-RnField $old 'Datum'); $newDate = [string](Get-RnField $new 'Datum')
            if ($oldDate -ne $newDate) {
                if ([string]::IsNullOrWhiteSpace($oldDate)) { [void]$warnings.Add("Version ${key}: Datum ergänzt ($newDate).") }
                else { [void]$errors.Add("Veröffentlichte Version ${key}: Datum geändert ($oldDate → $newDate).") }
            }
        }
    }
    return [PSCustomObject]@{ IsValid = ($errors.Count -eq 0); Errors = @($errors); Warnings = @($warnings) }
}

function ConvertTo-ReleaseNotesMarkdown {
    <#
    .SYNOPSIS
        Erzeugt den Inhalt einer CHANGELOG.md (nur Text, schreibt nichts).
    #>
    param($Notes, [string]$AppName = 'Blood-Tracker')
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("# Changelog – $AppName")
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Automatisch erzeugt aus `$script:ReleaseNotes` – nicht von Hand bearbeiten.')
    foreach ($note in (Get-ReleaseNotesSorted -Notes $Notes)) {
        $dateText = [string](Get-RnField $note 'Datum')
        $heading = '## [{0}]' -f [string](Get-RnField $note 'Version')
        if (-not [string]::IsNullOrWhiteSpace($dateText)) { $heading += ' – ' + $dateText.Trim() }   # v0.2.0: ohne Datum nur die Version
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine($heading)
        $title = [string](Get-RnField $note 'Titel')
        if (-not [string]::IsNullOrWhiteSpace($title)) { [void]$sb.AppendLine(''); [void]$sb.AppendLine("**$title**") }
        foreach ($group in (Get-ReleaseEntriesGrouped -Entries (Get-RnField $note 'Eintraege'))) {
            [void]$sb.AppendLine('')
            [void]$sb.AppendLine("### $($group.Typ)")
            foreach ($text in $group.Texte) { [void]$sb.AppendLine("- $text") }
        }
    }
    return $sb.ToString()
}

function Export-ReleaseNotesMarkdown {
    <#
    .SYNOPSIS
        Schreibt CHANGELOG.md – NUR als ausdrücklicher Entwicklungsschritt (Spec 12.9).
        Verweigert Ziele innerhalb eines Ordners, der als Datenordner übergeben wird.
    #>
    param($Notes, [Parameter(Mandatory)][string]$Path, [string]$AppName = 'Blood-Tracker', [string]$ForbiddenRoot = $null)
    $full = [System.IO.Path]::GetFullPath($Path)
    if ($ForbiddenRoot) {
        $root = [System.IO.Path]::GetFullPath($ForbiddenRoot).TrimEnd('\', '/')
        if ($full.StartsWith($root + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Export in den Datenordner ist nicht erlaubt: $full"
        }
    }
    $content = ConvertTo-ReleaseNotesMarkdown -Notes $Notes -AppName $AppName
    [System.IO.File]::WriteAllText($full, $content, (New-Object System.Text.UTF8Encoding($false)))
    return $full
}

# ---------------------------------------------------------------------------------
# Oberfläche (WinForms) – Spec Abschnitt 8
# Zustand im $script:-Scope (Event-Handler laufen in eigenem Scope, PS-5.1-Falle 9).
# ---------------------------------------------------------------------------------
function Get-RnColor {
    param([string]$Token)
    $hex = $script:HealthUiTheme[$Token]
    if ([string]::IsNullOrWhiteSpace($hex)) { $hex = '#000000' }
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

function Initialize-ReleaseNotesTab {
    <#
    .SYNOPSIS
        Baut den Reiter "Release-Notes" in eine vorhandene TabPage.
    .PARAMETER Notes
        $script:ReleaseNotes
    .PARAMETER CurrentVersion
        Laufende Version ($Version), Zusätze wie " BETA-TEST" erlaubt.
    .PARAMETER ToolTip
        Optional: vorhandenes ToolTip-Objekt des Hauptfensters.
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TabPage]$TabPage,
        $Notes,
        $CurrentVersion,
        $ToolTip = $null
    )
    $script:rnNotes = @($Notes)
    $script:rnCurrentVersion = $CurrentVersion
    $script:rnView = @()
    $script:rnGroups = @()          # v0.2.0: Listeneinträge (Versionen gleichen Datums zusammengefasst)
    $script:rnBusy = $false
    $script:rnSelectedKey = $null   # v0.2.0: Auswahl = Listeneintrag (Tag bzw. Version ohne Datum)
    $family = $script:HealthUiTheme['FontFamily']
    $script:rnFonts = @{
        Title   = New-Object System.Drawing.Font($family, 16, [System.Drawing.FontStyle]::Bold)
        Normal  = New-Object System.Drawing.Font($family, 9)
        Bold    = New-Object System.Drawing.Font($family, 9, [System.Drawing.FontStyle]::Bold)
        ListVer = New-Object System.Drawing.Font($family, 10, [System.Drawing.FontStyle]::Bold)
        DetHead = New-Object System.Drawing.Font($family, 14, [System.Drawing.FontStyle]::Bold)
        DetSub  = New-Object System.Drawing.Font($family, 11)
        DetType = New-Object System.Drawing.Font($family, 9, [System.Drawing.FontStyle]::Bold)
        DetText = New-Object System.Drawing.Font($family, 10)
    }

    $TabPage.BackColor = Get-RnColor 'Page'
    # Feste Inhaltsfläche 1152 × 760 in einem Scroll-Panel (Spec 3 / E5)
    $contentPanel = New-Object System.Windows.Forms.Panel
    $contentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $contentPanel.AutoScroll = $true
    $contentPanel.AutoScrollMinSize = New-Object System.Drawing.Size(1152, 760)
    $contentPanel.BackColor = Get-RnColor 'Page'
    $TabPage.Controls.Add($contentPanel)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Release-Notes'
    $title.Font = $script:rnFonts.Title
    $title.ForeColor = Get-RnColor 'TextPrimary'
    $title.Location = New-Object System.Drawing.Point(16, 14)
    $title.AutoSize = $true
    $contentPanel.Controls.Add($title)

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = 'Suche:'; $searchLabel.Font = $script:rnFonts.Normal; $searchLabel.ForeColor = Get-RnColor 'TextSecondary'
    $searchLabel.Location = New-Object System.Drawing.Point(332, 22); $searchLabel.AutoSize = $true
    $contentPanel.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Font = $script:rnFonts.Normal
    $searchBox.Location = New-Object System.Drawing.Point(384, 19); $searchBox.Size = New-Object System.Drawing.Size(230, 23)
    $searchBox.TabIndex = 0
    $contentPanel.Controls.Add($searchBox)

    $typeLabel = New-Object System.Windows.Forms.Label
    $typeLabel.Text = 'Typ:'; $typeLabel.Font = $script:rnFonts.Normal; $typeLabel.ForeColor = Get-RnColor 'TextSecondary'
    $typeLabel.Location = New-Object System.Drawing.Point(634, 22); $typeLabel.AutoSize = $true
    $contentPanel.Controls.Add($typeLabel)

    $typeCombo = New-Object System.Windows.Forms.ComboBox
    $typeCombo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $typeCombo.Font = $script:rnFonts.Normal
    $typeCombo.Location = New-Object System.Drawing.Point(670, 19); $typeCombo.Size = New-Object System.Drawing.Size(120, 23)
    [void]$typeCombo.Items.Add('Alle')
    foreach ($type in $script:ReleaseNoteTypes) { [void]$typeCombo.Items.Add($type) }
    $typeCombo.SelectedIndex = 0
    $typeCombo.TabIndex = 1
    $contentPanel.Controls.Add($typeCombo)

    $resetButton = New-Object System.Windows.Forms.Button
    $resetButton.Text = 'Filter zurücksetzen'
    $resetButton.Font = $script:rnFonts.Normal
    $resetButton.Location = New-Object System.Drawing.Point(806, 18); $resetButton.Size = New-Object System.Drawing.Size(140, 26)
    $resetButton.Enabled = $false
    $resetButton.TabIndex = 2
    $contentPanel.Controls.Add($resetButton)

    $currentLabel = New-Object System.Windows.Forms.Label
    $currentLabel.Font = $script:rnFonts.Normal; $currentLabel.ForeColor = Get-RnColor 'TextMuted'
    $currentLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $currentLabel.Location = New-Object System.Drawing.Point(962, 20); $currentLabel.Size = New-Object System.Drawing.Size(174, 22)
    $currentVersionObj = ConvertTo-ReleaseVersion -Text $CurrentVersion
    $currentLabel.Text = if ($null -ne $currentVersionObj) { "aktuelle Version: $currentVersionObj" } else { '' }
    $contentPanel.Controls.Add($currentLabel)

    $list = New-Object System.Windows.Forms.ListBox
    $list.DrawMode = [System.Windows.Forms.DrawMode]::OwnerDrawVariable   # v0.2.0: Einträge ohne Datum niedriger
    $list.ItemHeight = 58
    $list.IntegralHeight = $false
    $list.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $list.BackColor = Get-RnColor 'Card'
    $list.Location = New-Object System.Drawing.Point(16, 64); $list.Size = New-Object System.Drawing.Size(300, 680)
    $list.TabIndex = 3
    $contentPanel.Controls.Add($list)

    $detailCard = New-Object System.Windows.Forms.Panel
    $detailCard.BackColor = Get-RnColor 'Card'
    $detailCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $detailCard.Location = New-Object System.Drawing.Point(332, 64); $detailCard.Size = New-Object System.Drawing.Size(804, 680)
    $contentPanel.Controls.Add($detailCard)

    $details = New-Object System.Windows.Forms.RichTextBox
    $details.ReadOnly = $true
    $details.BackColor = Get-RnColor 'Card'     # NACH ReadOnly setzen, sonst grau
    $details.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $details.DetectUrls = $false
    $details.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $details.Location = New-Object System.Drawing.Point(16, 12); $details.Size = New-Object System.Drawing.Size(770, 654)
    $details.TabIndex = 4
    $detailCard.Controls.Add($details)

    $emptyLabel = New-Object System.Windows.Forms.Label
    $emptyLabel.Text = 'Keine passenden Release-Notes'
    $emptyLabel.Font = $script:rnFonts.DetSub
    $emptyLabel.ForeColor = Get-RnColor 'TextSecondary'
    $emptyLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $emptyLabel.Location = New-Object System.Drawing.Point(0, 280); $emptyLabel.Size = New-Object System.Drawing.Size(802, 28)
    $emptyLabel.Visible = $false
    $detailCard.Controls.Add($emptyLabel)

    $emptyReset = New-Object System.Windows.Forms.Button
    $emptyReset.Text = 'Filter zurücksetzen'
    $emptyReset.Font = $script:rnFonts.Normal
    $emptyReset.Location = New-Object System.Drawing.Point(331, 316); $emptyReset.Size = New-Object System.Drawing.Size(140, 28)
    $emptyReset.Visible = $false
    $detailCard.Controls.Add($emptyReset)

    if ($ToolTip) {
        $ToolTip.SetToolTip($searchBox, 'Sucht in Version, Titel und Einträgen (Groß-/Kleinschreibung egal). Esc leert die Suche.')
        $ToolTip.SetToolTip($typeCombo, 'Zeigt nur Einträge dieses Typs.')
    }

    $script:rnControls = @{
        Search = $searchBox; Type = $typeCombo; Reset = $resetButton; List = $list
        Details = $details; Empty = $emptyLabel; EmptyReset = $emptyReset
    }

    # --- Ereignisse ---
    $searchBox.Add_TextChanged({ Update-ReleaseNotesTab })
    $searchBox.Add_KeyDown({
        $keyArgs = $args[1]
        if ($keyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $keyArgs.Handled = $true; $keyArgs.SuppressKeyPress = $true
            $script:rnControls.Search.Text = ''
        }
    })
    $typeCombo.Add_SelectedIndexChanged({ Update-ReleaseNotesTab })
    $resetButton.Add_Click({ Reset-ReleaseNotesFilter })
    $emptyReset.Add_Click({ Reset-ReleaseNotesFilter })
    $list.Add_SelectedIndexChanged({
        if ($script:rnBusy) { return }
        $index = $script:rnControls.List.SelectedIndex
        if ($index -ge 0 -and $index -lt @($script:rnGroups).Count) {
            $group = @($script:rnGroups)[$index]
            $script:rnSelectedKey = $group.Key
            Show-ReleaseNoteDetails -Items $group.Versions
        }
    })
    $list.Add_MeasureItem({
        $measureArgs = $args[1]
        $groups = @($script:rnGroups)
        if ($measureArgs.Index -ge 0 -and $measureArgs.Index -lt $groups.Count) { $measureArgs.ItemHeight = Get-RnGroupItemHeight -Group $groups[$measureArgs.Index] }
        else { $measureArgs.ItemHeight = 58 }
    })
    $list.Add_DrawItem({ Invoke-ReleaseNotesDrawItem -DrawArgs $args[1] })

    Update-ReleaseNotesTab
}

function Reset-ReleaseNotesFilter {
    $script:rnBusy = $true
    try {
        $script:rnControls.Search.Text = ''
        $script:rnControls.Type.SelectedIndex = 0
    } finally { $script:rnBusy = $false }
    Update-ReleaseNotesTab
}

function Update-ReleaseNotesTab {
    <# Filter anwenden, Liste neu füllen, Auswahl nach Regel 5/6 setzen. #>
    if ($script:rnBusy -or -not $script:rnControls) { return }
    $controls = $script:rnControls
    $search = [string]$controls.Search.Text
    $type = [string]$controls.Type.SelectedItem
    $script:rnView = @(Get-ReleaseNotesView -Notes $script:rnNotes -SearchText $search -TypeFilter $type -CurrentVersion $script:rnCurrentVersion)
    $script:rnGroups = @(Get-ReleaseNotesGroups -ViewItems $script:rnView)   # v0.2.0
    $filterActive = (-not [string]::IsNullOrWhiteSpace($search)) -or ($type -ne 'Alle')
    $controls.Reset.Enabled = $filterActive

    $script:rnBusy = $true
    try {
        $controls.List.BeginUpdate()
        $controls.List.Items.Clear()
        foreach ($group in $script:rnGroups) { [void]$controls.List.Items.Add($group.Version) }
        $controls.List.EndUpdate()
        $index = Resolve-ReleaseNotesSelection -ViewItems $script:rnGroups -PreviousKey $script:rnSelectedKey
        $controls.List.SelectedIndex = $index
    } finally { $script:rnBusy = $false }

    if ($index -lt 0) {
        # Regel 6: keine Treffer → Details leeren, Hinweis + Zurücksetzen
        $controls.Details.Clear()
        $controls.Details.Visible = $false
        $controls.Empty.Visible = $true
        $controls.EmptyReset.Visible = $true
    } else {
        $controls.Details.Visible = $true
        $controls.Empty.Visible = $false
        $controls.EmptyReset.Visible = $false
        $script:rnSelectedKey = @($script:rnGroups)[$index].Key
        Show-ReleaseNoteDetails -Items (@($script:rnGroups)[$index].Versions)
    }
}

function Invoke-ReleaseNotesDrawItem {
    <#
        Zeichnet einen Listeneintrag (v0.2.0: Gruppe): neueste Version (+ "aktuell"),
        mit Datum: "TT.MM.JJJJ" bzw. "TT.MM.JJJJ · n Versionen" und Titel (3 Zeilen);
        ohne Datum: nur Version und Titel (2 Zeilen).
    #>
    param($DrawArgs)
    $e = $DrawArgs
    if ($null -eq $e -or $e.Index -lt 0 -or $e.Index -ge @($script:rnGroups).Count) { return }
    $item = @($script:rnGroups)[$e.Index]
    $bounds = $e.Bounds
    $selected = (($e.State -band [System.Windows.Forms.DrawItemState]::Selected) -eq [System.Windows.Forms.DrawItemState]::Selected)
    $backColor = if ($selected) { Get-RnColor 'SelectionBack' } else { Get-RnColor 'Card' }
    $backBrush = New-Object System.Drawing.SolidBrush($backColor)
    $linePen = New-Object System.Drawing.Pen((Get-RnColor 'Border'), 1)
    try {
        $e.Graphics.FillRectangle($backBrush, $bounds)
        if ($selected) {
            $accentBrush = New-Object System.Drawing.SolidBrush((Get-RnColor 'Link'))
            try { $e.Graphics.FillRectangle($accentBrush, $bounds.X, $bounds.Y, 3, $bounds.Height) } finally { $accentBrush.Dispose() }
        }
        $e.Graphics.DrawLine($linePen, $bounds.X, $bounds.Bottom - 1, $bounds.Right, $bounds.Bottom - 1)
        $flags = [System.Windows.Forms.TextFormatFlags]::NoPrefix -bor [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis
        $x = $bounds.X + 12
        $width = $bounds.Width - 24
        $versionRect = New-Object System.Drawing.Rectangle($x, ($bounds.Y + 6), $width, 18)
        [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $item.Version, $script:rnFonts.ListVer, $versionRect, (Get-RnColor 'TextPrimary'), $flags)
        if ($item.IsCurrent) {
            $versionWidth = [System.Windows.Forms.TextRenderer]::MeasureText($e.Graphics, $item.Version, $script:rnFonts.ListVer, $versionRect.Size, $flags).Width
            $badgeRect = New-Object System.Drawing.Rectangle(($x + $versionWidth + 6), ($bounds.Y + 6), 56, 18)
            $badgePen = New-Object System.Drawing.Pen((Get-RnColor 'Link'), 1)
            try { $e.Graphics.DrawRectangle($badgePen, $badgeRect) } finally { $badgePen.Dispose() }
            [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, 'aktuell', $script:rnFonts.Normal, $badgeRect, (Get-RnColor 'Link'),
                ($flags -bor [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter))
        }
        $titleY = $bounds.Y + 24
        if ($item.HasDate) {
            $dateText = $item.DatumText
            if ($item.Count -gt 1) { $dateText += " · $($item.Count) Versionen" }
            $dateRect = New-Object System.Drawing.Rectangle($x, ($bounds.Y + 24), $width, 16)
            [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $dateText, $script:rnFonts.Normal, $dateRect, (Get-RnColor 'TextMuted'), $flags)
            $titleY = $bounds.Y + 40
        }
        $titleRect = New-Object System.Drawing.Rectangle($x, $titleY, $width, 16)
        [System.Windows.Forms.TextRenderer]::DrawText($e.Graphics, $item.Titel, $script:rnFonts.Normal, $titleRect, (Get-RnColor 'TextSecondary'), $flags)
        if (($e.State -band [System.Windows.Forms.DrawItemState]::Focus) -eq [System.Windows.Forms.DrawItemState]::Focus) { $e.DrawFocusRectangle() }
    } finally {
        $backBrush.Dispose(); $linePen.Dispose()
    }
}

function Show-ReleaseNoteDetails {
    <#
        Füllt die Detailansicht; Suchtreffer werden hinterlegt.
        v0.2.0: -Items = alle (sichtbaren) Versionen eines Listeneintrags, neueste zuerst,
        je "Version x · Datum" (ohne Datum nur "Version x"), getrennt durch eine feine Linie.
    #>
    param($Items)
    $versions = @($Items | Where-Object { $null -ne $_ })
    $rtb = $script:rnControls.Details
    $rtb.SuspendLayout()
    try {
        $rtb.Clear()
        $append = {
            param([string]$Text, $Font, [string]$ColorToken, [bool]$Bullet = $false)
            $rtb.SelectionStart = $rtb.TextLength
            $rtb.SelectionLength = 0
            $rtb.SelectionFont = $Font
            $rtb.SelectionColor = Get-RnColor $ColorToken
            $rtb.SelectionBullet = $Bullet
            # AppendText statt SelectedText: bewährtes Muster für schreibgeschützte RichTextBoxen
            $rtb.AppendText($Text)
        }
        $separator = ([string][char]0x2014) * 32
        $first = $true
        foreach ($ViewItem in $versions) {
            if (-not $first) {
                & $append "`n" $script:rnFonts.DetText 'TextPrimary'
                & $append ($separator + "`n`n") $script:rnFonts.DetText 'Border'
            }
            $first = $false
            $headline = "Version $($ViewItem.Version)"
            if (-not [string]::IsNullOrEmpty($ViewItem.DatumText)) { $headline += " · $($ViewItem.DatumText)" }
            if ($ViewItem.IsCurrent) { $headline += '  (aktuell)' }
            & $append ($headline + "`n") $script:rnFonts.DetHead 'TextPrimary'
            if (-not [string]::IsNullOrWhiteSpace($ViewItem.Titel)) { & $append ($ViewItem.Titel + "`n") $script:rnFonts.DetSub 'TextSecondary' }
            foreach ($group in (Get-ReleaseEntriesGrouped -Entries $ViewItem.Eintraege)) {
                & $append "`n" $script:rnFonts.DetText 'TextPrimary'
                & $append ($group.Typ.ToUpperInvariant() + "`n") $script:rnFonts.DetType 'TextSecondary'
                foreach ($text in $group.Texte) { & $append ($text + "`n") $script:rnFonts.DetText 'TextPrimary' $true }
            }
        }
        $rtb.SelectionBullet = $false
        # Suchtreffer hinterlegen (Highlight-Farbe, Text bleibt in Textfarbe)
        $search = ([string]$script:rnControls.Search.Text).Trim()
        if ($search.Length -gt 0) {
            $start = 0
            while ($start -lt $rtb.TextLength) {
                $hit = $rtb.Find($search, $start, [System.Windows.Forms.RichTextBoxFinds]::None)
                if ($hit -lt 0) { break }
                $rtb.SelectionBackColor = Get-RnColor 'Highlight'
                $start = $hit + $search.Length
            }
        }
        $rtb.SelectionStart = 0
        $rtb.SelectionLength = 0
        $rtb.ScrollToCaret()
    } finally { $rtb.ResumeLayout() }
}

# ---------------------------------------------------------------------------------
# Einbindung in das Hauptfenster (v2.34.0, nicht Teil des Moduls)
# ---------------------------------------------------------------------------------
function Invoke-BtDrawTabHeader {
    # v2.43.0: Nur den Inhalt des nativen Reiterkopfs zeichnen; Rahmen bleibt bei Windows.
    param([System.Windows.Forms.TabControl]$TabControl, $DrawArgs)
    if ($null -eq $TabControl -or $null -eq $DrawArgs -or $DrawArgs.Index -lt 0 -or $DrawArgs.Index -ge $TabControl.TabPages.Count) { return }
    $bounds = $DrawArgs.Bounds
    if ($bounds.Width -lt 8 -or $bounds.Height -lt 4) { return }
    $page = $TabControl.TabPages[$DrawArgs.Index]
    $selected = ($DrawArgs.Index -eq $TabControl.SelectedIndex)
    $backColor = [System.Drawing.SystemColors]::Control
    $foreColor = if ($TabControl.Enabled -and $page.Enabled) { $TabControl.ForeColor } else { [System.Drawing.SystemColors]::GrayText }
    $background = New-Object System.Drawing.SolidBrush($backColor)
    try {
        $DrawArgs.Graphics.FillRectangle($background, $bounds)
        if ($selected) {
            $accent = New-Object System.Drawing.SolidBrush($TabControl.Tag.HeaderStyle.Color)
            try { $DrawArgs.Graphics.FillRectangle($accent, ($bounds.X + 4), ($bounds.Y + 1), ($bounds.Width - 8), 2) }
            finally { $accent.Dispose() }
        }
        $flags = [System.Windows.Forms.TextFormatFlags]::HorizontalCenter -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter -bor [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor [System.Windows.Forms.TextFormatFlags]::NoPrefix -bor [System.Windows.Forms.TextFormatFlags]::PreserveGraphicsClipping -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
        $headerFont = $TabControl.Font
        if ($script:ActiveTabBoldEnabled -and -not $selected) { $headerFont = $TabControl.Tag.HeaderStyle.NormalFont }
        [System.Windows.Forms.TextRenderer]::DrawText($DrawArgs.Graphics, $page.Text, $headerFont, $bounds, $foreColor, $flags)
        if ($selected -and $TabControl.Focused) {
            $focusBounds = New-Object System.Drawing.Rectangle(($bounds.X + 3), ($bounds.Y + 4), ($bounds.Width - 6), ([Math]::Max(1, $bounds.Height - 6)))
            [System.Windows.Forms.ControlPaint]::DrawFocusRectangle($DrawArgs.Graphics, $focusBounds, $foreColor, $backColor)
        }
    } finally { $background.Dispose() }
}

function Select-BtNavigationPage {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl,
        [Parameter(Mandatory)][System.Windows.Forms.TabPage]$TabPage
    )
    if ($TabControl.IsDisposed -or $TabPage.IsDisposed) { throw 'Die Navigation wurde bereits geschlossen.' }
    $state = if ($TabControl.Tag) { $TabControl.Tag.Navigation } else { $null }
    if ($null -ne $state) {
        $group = @($state.Groups | Where-Object { $_.Pages -contains $TabPage })
        if ($group.Count -ne 1) { throw 'Der Zielreiter gehört nicht zu dieser Navigation.' }
        Set-BtNavigationCategory -State $state -Group $group[0] -Page $TabPage
    } else {
        if (-not $TabControl.TabPages.Contains($TabPage)) { throw 'Der Zielreiter gehört nicht zu dieser Navigation.' }
        if ($TabControl.SelectedTab -ne $TabPage) { $TabControl.SelectedTab = $TabPage }
    }
}

function Invoke-BtNavigationRefresh {
    param($State)
    if ($State.Disposed -or $State.Switching -or $null -eq $State.Tabs.SelectedTab) { return }
    $State.LastPages[$State.Current.Name] = $State.Tabs.SelectedTab
    if ($null -ne $State.Refresh) { & $State.Refresh }
}

function Set-BtNavigationCategory {
    param($State, $Group, [System.Windows.Forms.TabPage]$Page)
    if ($State.Disposed -or $State.Switching) { return }
    if ($null -eq $Page) {
        $Page = $State.LastPages[$Group.Name]
        if ($null -eq $Page) { $Page = $Group.Pages[0] }
    }
    if ($Group.Pages -notcontains $Page) { throw 'Der Reiter gehört nicht zur gewählten Kategorie.' }
    if ($State.Current -eq $Group) {
        if ($State.Tabs.SelectedTab -ne $Page) { $State.Tabs.SelectedTab = $Page }
        return
    }
    $State.Switching = $true
    $State.Tabs.SuspendLayout()
    try {
        if ($null -ne $State.Current -and $null -ne $State.Tabs.SelectedTab) {
            $State.LastPages[$State.Current.Name] = $State.Tabs.SelectedTab
        }
        # Nur umhaengen: Controls, Eingaben und Filter bleiben dieselben Instanzen.
        $State.Tabs.TabPages.Clear()
        foreach ($item in $Group.Pages) { $State.Tabs.TabPages.Add($item) }
        $State.Current = $Group
        $State.Categories.SelectedTab = $Group.CategoryPage
        $State.Tabs.SelectedTab = $Page
        $State.LastPages[$Group.Name] = $Page
    } finally {
        $State.Tabs.ResumeLayout($true)
        $State.Switching = $false
    }
    Invoke-BtNavigationRefresh -State $State
}

function Initialize-BtCategoryNavigation {
    param(
        [System.Windows.Forms.Form]$Form,
        [System.Windows.Forms.TabControl]$Tabs,
        [System.Windows.Forms.Button]$SettingsButton,
        [object[]]$Groups
    )
    $pages = @($Groups | ForEach-Object { $_.Pages })
    if ($pages.Count -ne $Tabs.TabCount -or @($pages | Select-Object -Unique).Count -ne $pages.Count) { throw 'Ungültige Zuordnung der Navigationsreiter.' }
    foreach ($page in $pages) { if (-not $Tabs.TabPages.Contains($page)) { throw 'Ein Navigationsreiter fehlt.' } }
    $categories = New-Object Windows.Forms.TabControl
    $categories.Location = New-Object Drawing.Point(10, 10)
    $categories.Size = New-Object Drawing.Size(($SettingsButton.Left - 18), 28)
    $categories.Anchor = 'Top,Left,Right'; $categories.TabIndex = 0
    $Form.Controls.Add($categories)
    Initialize-BtActiveTabIndicator -TabControl $categories
    $state = [PSCustomObject]@{
        Tabs = $Tabs; Categories = $categories; Groups = $Groups; Current = $null
        LastPages = @{}; Switching = $true; Disposed = $false; Refresh = $null
    }
    $Tabs.Tag.Navigation = $state; $categories.Tag.Navigation = $state
    foreach ($group in $Groups) {
        $page = New-Object Windows.Forms.TabPage; $page.Text = $group.Name
        $group | Add-Member -MemberType NoteProperty -Name CategoryPage -Value $page
        $categories.TabPages.Add($page)
    }
    $categories.Add_SelectedIndexChanged({
        $state = $args[0].Tag.Navigation
        if ($state.Switching -or $state.Disposed -or $null -eq $args[0].SelectedTab) { return }
        $group = @($state.Groups | Where-Object { $_.CategoryPage -eq $state.Categories.SelectedTab })[0]
        Set-BtNavigationCategory -State $state -Group $group
    })
    $Tabs.Add_SelectedIndexChanged({ Invoke-BtNavigationRefresh -State $args[0].Tag.Navigation })
    $categories.Add_Disposed({
        $state = $args[0].Tag.Navigation; $state.Disposed = $true; $state.Refresh = $null
        # Nicht sichtbare Seiten haben keinen Parent und muessen ebenfalls freigegeben werden.
        foreach ($group in $state.Groups) {
            foreach ($page in $group.Pages) { if (-not $page.IsDisposed) { $page.Dispose() } }
        }
    })
    $state.Switching = $false
    Set-BtNavigationCategory -State $state -Group $Groups[0] -Page $Groups[0].Pages[0]
    return $state
}

function Initialize-BtActiveTabIndicator {
    param([Parameter(Mandatory)][System.Windows.Forms.TabControl]$TabControl)
    if ($null -eq $TabControl.Tag) { $TabControl.Tag = [PSCustomObject]@{ HeaderStyle = $null; Navigation = $null } }
    if (-not $script:ActiveTabIndicatorEnabled -or $null -ne $TabControl.Tag.HeaderStyle) { return }
    $colorText = [string]$script:ActiveTabIndicatorColor
    if ($colorText -notmatch '^#[0-9a-fA-F]{6}$') {
        Write-Warning 'Ungültige Reiter-Markierungsfarbe in USER-KONFIGURATION; dezentes Rot wird verwendet.'
        $colorText = '#B84040'
    }
    $style = [PSCustomObject]@{
        NormalFont = $TabControl.Font; BoldFont = $null
        Color = [Drawing.ColorTranslator]::FromHtml($colorText)
    }
    $TabControl.Tag.HeaderStyle = $style
    if ($script:ActiveTabBoldEnabled) {
        foreach ($page in $TabControl.TabPages) { $page.Font = $style.NormalFont }
        $TabControl.Add_ControlAdded({
            if ($args[1].Control -is [Windows.Forms.TabPage]) { $args[1].Control.Font = $args[0].Tag.HeaderStyle.NormalFont }
        })
        $style.BoldFont = New-Object Drawing.Font($style.NormalFont, ($style.NormalFont.Style -bor [Drawing.FontStyle]::Bold))
        $TabControl.Font = $style.BoldFont
    }
    $TabControl.Add_Disposed({
        $style = $args[0].Tag.HeaderStyle
        if ($null -ne $style.BoldFont) { $style.BoldFont.Dispose() }
    })
    $TabControl.DrawMode = [Windows.Forms.TabDrawMode]::OwnerDrawFixed
    $TabControl.Add_DrawItem({ Invoke-BtDrawTabHeader -TabControl $args[0] -DrawArgs $args[1] })
    $TabControl.Add_SelectedIndexChanged({ $args[0].Invalidate() })
    $TabControl.Add_GotFocus({ $args[0].Invalidate() })
    $TabControl.Add_LostFocus({ $args[0].Invalidate() })
}

function Resolve-TabHeaderOverlap {
    <#
    .SYNOPSIS
        v2.34.0: Reiterkoepfe duerfen die Schaltflaechen rechts daneben ("globale
        Einstellungen", "Beenden") nicht ueberdecken.
    .DESCRIPTION
        Schaetzt die Breite aller Reiterkoepfe wie das Windows-Reitersteuerelement
        (Textbreite + 2 x Innenabstand + 1) und verringert bei Bedarf den waagerechten
        Innenabstand in 2-px-Schritten bis minimal 2 px (bei Fettschrift 4 px). VOR dem ersten Anzeigen aufrufen
        (dann ohne Neuaufbau des Fensters). Nur Darstellung, keine Daten.
    .OUTPUTS
        [int] geschaetzter rechter Rand der Reiterkoepfe (Fensterkoordinaten); $null ohne Pruefung.
    #>
    param(
        [System.Windows.Forms.TabControl]$TabControl,
        [System.Windows.Forms.Control]$Obstacle,
        [int]$Margin = 8
    )
    if ($null -eq $TabControl -or $null -eq $Obstacle -or $TabControl.TabPages.Count -eq 0) { return $null }
    $minimumPadX = if ($script:ActiveTabIndicatorEnabled -and $script:ActiveTabBoldEnabled) { 4 } else { 2 }
    $padX = [Math]::Max($minimumPadX, [int]$TabControl.Padding.X)
    $textWidth = 0
    foreach ($page in $TabControl.TabPages) {
        $textWidth += [System.Windows.Forms.TextRenderer]::MeasureText([string]$page.Text, $TabControl.Font).Width
    }
    $count = $TabControl.TabPages.Count
    $right = $TabControl.Left + $textWidth + $count * (2 * $padX + 1)
    while (($right + $Margin) -gt $Obstacle.Left -and $padX -gt $minimumPadX) {
        $padX = [Math]::Max($minimumPadX, $padX - 2)
        $right = $TabControl.Left + $textWidth + $count * (2 * $padX + 1)
    }
    if ($padX -ne [int]$TabControl.Padding.X) {
        $TabControl.Padding = New-Object System.Drawing.Point($padX, $TabControl.Padding.Y)
    }
    return [int]$right
}

function Invoke-ReleaseNotesStartCheck {
    <#
    .SYNOPSIS
        v2.34.0 (Spec 12.9, Integrationshinweis R4): Konsistenzpruefung der Release-Notes beim
        Start - NUR im BETA-Testmodus ($Version enthaelt "BETA", gesetzt vom BETA-Starter).
    .DESCRIPTION
        Prueft Kopfzeile "# Version", $Version und den obersten Eintrag von $script:ReleaseNotes
        (Test-ReleaseNotesConsistency). Abweichungen -> EINE Warnmeldung; das Programm laeuft
        immer weiter. Im normalen Start passiert nichts.
    .OUTPUTS
        Meldungszeilen (string[]); leer, wenn alles passt oder kein BETA-Start.
        Aufrufer IMMER mit @() uebernehmen.
    #>
    param(
        [string]$CurrentVersion = [string]$script:Version,
        [string]$ScriptPath = [string]$script:BtScriptPath,
        $Notes = $script:ReleaseNotes
    )
    if ($CurrentVersion -notmatch 'BETA') { return @() }
    $lines = @()
    try {
        $check = Test-ReleaseNotesConsistency -Notes $Notes -CurrentVersion $CurrentVersion -HeaderVersion (Get-ScriptHeaderVersion -Path $ScriptPath)
        $lines = @(@($check.Errors) | ForEach-Object { "Fehler: $_" }) + @(@($check.Warnings) | ForEach-Object { "Hinweis: $_" })
    } catch {
        $lines = @("Fehler: Prüfung nicht möglich – $($_.Exception.Message)")
    }
    if ($lines.Count -gt 0) {
        $shown = @($lines | Select-Object -First 12)
        $text = "Release-Notes: Abweichungen gefunden (nur BETA-Hinweis – das Programm läuft normal weiter).`n`n" + ($shown -join "`n")
        if ($lines.Count -gt $shown.Count) { $text += "`n… und $($lines.Count - $shown.Count) weitere" }
        $text += "`n`nRegeln: Regelblock oben im Script."
        [System.Windows.Forms.MessageBox]::Show($text, 'Blood-Tracker BETA – Release-Notes-Prüfung', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
    return $lines
}
# ════ Ende RELEASE-NOTES (v2.34.0) ════════════════════════════════════════════════════

# ╔══════════════════════════════════════════════════════════════════════════════════════╗
# ║ v2.35.0: SORTIEREN IN TABELLEN (Risiko-Cockpit, Longevity-Indizes)                   ║
# ║ Klick auf eine Spaltenüberschrift: 1. aufsteigend, 2. absteigend, 3. Standard.       ║
# ║ Zahlen nach Zahlenwert (auch "5.2 %", "612 (21.24 nmol/l)"), Datum chronologisch,    ║
# ║ Rang-Spalten (Bewertung) über eine Hilfsspalte, Text alphabetisch (de-DE).           ║
# ║ Leere Werte und "-" stehen IMMER unten. Gleichstand: Folgespalten, dann Einfüge-     ║
# ║ reihenfolge (Row.Tag). Zustand je Tabelle in $script:BtGridSort - bleibt beim        ║
# ║ Neuaufbau (Filter, Aktualisierung) erhalten.                                         ║
# ╚══════════════════════════════════════════════════════════════════════════════════════╝
$script:BtGridSort = @{}
$script:BtGridTextCulture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')

function Get-BtGridSortKey {
    <#
    .SYNOPSIS
        v2.35.0: Sortierschlüssel eines Zellwerts.
    .PARAMETER Mode
        Auto/Number: Zahl (auch führende Zahl in Text, z. B. "5.2 %", "< 5.2"), sonst Text.
        Date: TT.MM.JJJJ bzw. JJJJ-MM-TT. Text: immer alphabetisch.
    .OUTPUTS
        PSCustomObject { Empty; Kind (0 = Zahl, 1 = Text); Number; Text }
    #>
    param($Value, [ValidateSet('Auto', 'Number', 'Date', 'Text')][string]$Mode = 'Auto')
    if ($null -eq $Value -or $Value -is [System.DBNull]) { return [PSCustomObject]@{ Empty = $true; Kind = 2; Number = 0.0; Text = '' } }
    $isNumericType = ($Value -is [double] -or $Value -is [int] -or $Value -is [long] -or $Value -is [decimal] -or $Value -is [single] -or $Value -is [int16] -or $Value -is [byte])
    if ($Mode -ne 'Text' -and $isNumericType) {
        $numeric = [double]$Value
        if ([double]::IsNaN($numeric)) { return [PSCustomObject]@{ Empty = $true; Kind = 2; Number = 0.0; Text = '' } }
        return [PSCustomObject]@{ Empty = $false; Kind = 0; Number = $numeric; Text = '' }
    }
    if ($Value -is [datetime]) { return [PSCustomObject]@{ Empty = $false; Kind = 0; Number = [double]$Value.Ticks; Text = '' } }
    $text = ([string]$Value).Trim()
    if ($text -eq '' -or $text -match '^[-–—]+$') { return [PSCustomObject]@{ Empty = $true; Kind = 2; Number = 0.0; Text = '' } }
    if ($Mode -eq 'Date') {
        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParseExact($text, [string[]]@('dd.MM.yyyy', 'd.M.yyyy', 'yyyy-MM-dd'), [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsedDate)) {
            return [PSCustomObject]@{ Empty = $false; Kind = 0; Number = [double]$parsedDate.Ticks; Text = '' }
        }
    } elseif ($Mode -ne 'Text') {
        $numberMatch = [regex]::Match($text, '^[<>≤≥~]?\s*([-+]?\d+(?:[.,]\d+)?)')
        if ($numberMatch.Success) {
            $number = [double]::Parse($numberMatch.Groups[1].Value.Replace(',', '.'), [System.Globalization.CultureInfo]::InvariantCulture)
            return [PSCustomObject]@{ Empty = $false; Kind = 0; Number = $number; Text = '' }
        }
    }
    # Warnsymbol vor Markernamen ("⚠ Homocystein") beim Sortieren ignorieren
    $clean = $text -replace '^[⚠\s]+', ''
    return [PSCustomObject]@{ Empty = $false; Kind = 1; Number = 0.0; Text = $clean }
}

function Compare-BtGridSortKey {
    <# v2.35.0: Vergleich zweier Sortierschlüssel; leere Werte stehen IMMER unten (unabhängig von -Descending). #>
    param($A, $B, [switch]$Descending)
    if ($A.Empty -and $B.Empty) { return 0 }
    if ($A.Empty) { return 1 }
    if ($B.Empty) { return -1 }
    $result = 0
    if ($A.Kind -ne $B.Kind) { $result = [Math]::Sign($A.Kind - $B.Kind) }
    elseif ($A.Kind -eq 0) { $result = ([double]$A.Number).CompareTo([double]$B.Number) }
    else { $result = [string]::Compare([string]$A.Text, [string]$B.Text, $script:BtGridTextCulture, [System.Globalization.CompareOptions]::IgnoreCase) }
    if ($Descending) { $result = -$result }
    return [int][Math]::Sign($result)
}

function Get-BtGridCellSortValue {
    <# v2.35.0: Vergleichswert einer Zeile: Einfügereihenfolge (Row.Tag), Rang-Hilfsspalte, echter Gruppenname (Cell.Tag) oder Zellwert. #>
    param($Row, [string]$ColumnName, $Config)
    if ($ColumnName -eq '__Order') { return $Row.Tag }
    if ($Config.RankColumns.ContainsKey($ColumnName)) { return $Row.Cells[[string]$Config.RankColumns[$ColumnName]].Value }
    $cell = $Row.Cells[$ColumnName]
    if ($ColumnName -eq $Config.GroupColumn -and $null -ne $cell.Tag) { return $cell.Tag }
    return $cell.Value
}

function Compare-BtGridRows {
    <#
    .SYNOPSIS
        v2.35.0: Vergleicht zwei Tabellenzeilen nach der gewählten Spalte ($Config.ActiveColumn,
        $Config.ActiveDescending) bzw. ohne Auswahl nach der Standardspalte (aufsteigend), danach
        Folgespalten ($Config.TieBreak, aufsteigend) und zuletzt nach der Einfügereihenfolge.
    #>
    param($Row1, $Row2, $Config)
    $column = [string]$Config.ActiveColumn
    $descending = [bool]$Config.ActiveDescending
    if ([string]::IsNullOrEmpty($column)) { $column = [string]$Config.DefaultColumn; $descending = $false }
    $steps = @(@{ Name = $column; Desc = $descending })
    foreach ($tieColumn in @($Config.TieBreak)) { if ($tieColumn -and $tieColumn -ne $column) { $steps += , @{ Name = $tieColumn; Desc = $false } } }
    if ($column -ne '__Order') { $steps += , @{ Name = '__Order'; Desc = $false } }
    foreach ($step in $steps) {
        $mode = 'Auto'
        if ($step.Name -eq '__Order' -or $Config.RankColumns.ContainsKey($step.Name)) { $mode = 'Number' }
        elseif ($Config.Modes.ContainsKey($step.Name)) { $mode = [string]$Config.Modes[$step.Name] }
        $key1 = Get-BtGridSortKey -Value (Get-BtGridCellSortValue -Row $Row1 -ColumnName $step.Name -Config $Config) -Mode $mode
        $key2 = Get-BtGridSortKey -Value (Get-BtGridCellSortValue -Row $Row2 -ColumnName $step.Name -Config $Config) -Mode $mode
        $result = Compare-BtGridSortKey -A $key1 -B $key2 -Descending:([bool]$step.Desc)
        if ($result -ne 0) { return $result }
    }
    return 0
}

function Get-BtGridNextSortState {
    <# v2.35.0: Klickfolge je Spalte: aufsteigend -> absteigend -> Standard; andere Spalte: aufsteigend. #>
    param([string]$CurrentColumn, [bool]$CurrentDescending, [string]$ClickedColumn)
    if ($CurrentColumn -ne $ClickedColumn) { return [PSCustomObject]@{ Column = $ClickedColumn; Descending = $false } }
    if (-not $CurrentDescending) { return [PSCustomObject]@{ Column = $ClickedColumn; Descending = $true } }
    return [PSCustomObject]@{ Column = ''; Descending = $false }
}

function Update-BtGridGroupDisplay {
    <#
    .SYNOPSIS
        v2.35.0: Gruppenspalte - mit -Collapse nur in der ersten Zeile einer Gruppe, sonst in jeder
        Zeile. Der echte Name steht in Cell.Tag (ersetzt die Ausblend-Schleife bis v2.34.1, die mit
        der bereits geleerten Vorzeile verglich).
    #>
    param($Rows, [Parameter(Mandatory)][string]$GroupColumn, [bool]$Collapse)
    $previous = $null
    $isFirst = $true
    foreach ($row in $Rows) {
        if ($row.IsNewRow) { continue }
        $cell = $row.Cells[$GroupColumn]
        if ($null -eq $cell.Tag) { $cell.Tag = [string]$cell.Value }
        $realName = [string]$cell.Tag
        if ($Collapse -and -not $isFirst -and $realName -eq $previous) { $cell.Value = '' } else { $cell.Value = $realName }
        $previous = $realName
        $isFirst = $false
    }
}

function Invoke-BtGridApplySort {
    <# v2.35.0: Sortiert eine registrierte Tabelle nach dem gespeicherten Zustand (oder Standard) und setzt die Sortierpfeile. #>
    param([Parameter(Mandatory)]$Grid)
    $config = $script:BtGridSort[[string]$Grid.Name]
    if (-not $config) { return }
    if ($Grid.Rows.Count -gt 1) {
        # Die Vergleichslogik steckt im SortCompare-Ereignis (Invoke-BtGridSortCompare) - die
        # übergebene Spalte und Richtung dienen nur als Auslöser.
        $Grid.Sort($Grid.Columns[[string]$config.SortHostColumn], [System.ComponentModel.ListSortDirection]::Ascending)
    }
    foreach ($column in $Grid.Columns) { $column.HeaderCell.SortGlyphDirection = [System.Windows.Forms.SortOrder]::None }
    if ($config.ActiveColumn) {
        $glyph = [System.Windows.Forms.SortOrder]::Ascending
        if ($config.ActiveDescending) { $glyph = [System.Windows.Forms.SortOrder]::Descending }
        $Grid.Columns[[string]$config.ActiveColumn].HeaderCell.SortGlyphDirection = $glyph
    }
    if ($config.GroupColumn) {
        $collapse = (-not $config.ActiveColumn) -or ($config.ActiveColumn -eq $config.GroupColumn)
        Update-BtGridGroupDisplay -Rows $Grid.Rows -GroupColumn $config.GroupColumn -Collapse $collapse
    }
}

function Invoke-BtGridHeaderClick {
    <# v2.35.0: Ereignis ColumnHeaderMouseClick (nur linke Maustaste). #>
    param($Grid, $ClickArgs)
    if ($null -eq $ClickArgs -or $ClickArgs.ColumnIndex -lt 0 -or $ClickArgs.Button -ne [System.Windows.Forms.MouseButtons]::Left) { return }
    $config = $script:BtGridSort[[string]$Grid.Name]
    if (-not $config) { return }
    $next = Get-BtGridNextSortState -CurrentColumn ([string]$config.ActiveColumn) -CurrentDescending ([bool]$config.ActiveDescending) -ClickedColumn ([string]$Grid.Columns[$ClickArgs.ColumnIndex].Name)
    $config.ActiveColumn = $next.Column
    $config.ActiveDescending = $next.Descending
    try { Invoke-BtGridApplySort -Grid $Grid } catch { Write-Warning "Sortieren: $($_.Exception.Message)" }
}

function Invoke-BtGridSortCompare {
    <# v2.35.0: Ereignis SortCompare - vergleicht ganze Zeilen (Compare-BtGridRows); bricht nie mit Fehler ab. #>
    param($Grid, $CompareArgs)
    $config = $script:BtGridSort[[string]$Grid.Name]
    if (-not $config) { return }
    $result = 0
    try { $result = Compare-BtGridRows -Row1 $Grid.Rows[$CompareArgs.RowIndex1] -Row2 $Grid.Rows[$CompareArgs.RowIndex2] -Config $config } catch { $result = 0 }
    $CompareArgs.SortResult = [int]$result
    $CompareArgs.Handled = $true
}

function Register-BtGridSorting {
    <#
    .SYNOPSIS
        v2.35.0: Macht alle Spalten einer Tabelle sortierbar (Klickfolge auf/ab/Standard).
    .PARAMETER Key
        Name der Tabelle (wird als Grid.Name gesetzt, Schlüssel in $script:BtGridSort).
    .PARAMETER Modes
        @{ Spalte = 'Auto' | 'Number' | 'Date' | 'Text' } - fehlende Spalten: Auto.
    .PARAMETER RankColumns
        @{ Spalte = 'Hilfsspalte' } - sortiert über den Zahlenwert der (unsichtbaren) Hilfsspalte.
    .PARAMETER DefaultColumn
        Standardreihenfolge: Spaltenname oder '__Order' (Einfügereihenfolge = Row.Tag).
    .PARAMETER TieBreak
        Folgespalten bei Gleichstand (immer aufsteigend).
    .PARAMETER GroupColumn
        Spalte mit Gruppennamen (echter Name in Cell.Tag) - siehe Update-BtGridGroupDisplay.
    #>
    param(
        [Parameter(Mandatory)]$Grid, [Parameter(Mandatory)][string]$Key, [hashtable]$Modes = @{}, [hashtable]$RankColumns = @{},
        [string]$DefaultColumn = '__Order', [string[]]$TieBreak = @(), [string]$GroupColumn = ''
    )
    $Grid.Name = $Key
    $hostColumn = ''
    foreach ($column in $Grid.Columns) {
        $column.SortMode = [System.Windows.Forms.DataGridViewColumnSortMode]::Programmatic
        if (-not $hostColumn -and $column.Visible) { $hostColumn = [string]$column.Name }
    }
    $script:BtGridSort[$Key] = @{
        Modes = $Modes; RankColumns = $RankColumns; DefaultColumn = $DefaultColumn; TieBreak = @($TieBreak); GroupColumn = $GroupColumn
        SortHostColumn = $hostColumn; ActiveColumn = ''; ActiveDescending = $false
    }
    $Grid.Add_ColumnHeaderMouseClick({ Invoke-BtGridHeaderClick -Grid $args[0] -ClickArgs $args[1] })
    $Grid.Add_SortCompare({ Invoke-BtGridSortCompare -Grid $args[0] -CompareArgs $args[1] })
}
# ════ Ende SORTIEREN (v2.35.0) ═══════════════════════════════════════════════════════════

# ╔══════════════════════════════════════════════════════════════════════════════════════╗
# ║ v2.36.0: DASHBOARD - Logik und Reiter (Integration B)                                ║
# ║ Quelle: Modul BloodTracker_Dashboard v0.1.1 - isoliert getestet, unveraendert        ║
# ║ uebernommen. Keine Abhaengigkeit von Funktionen in Show-DataEntryForm.               ║
# ║ Ab v2.36.0 ist DIESES Script massgeblich; das Modul dient nur dem isolierten Test.   ║
# ╚══════════════════════════════════════════════════════════════════════════════════════╝
# ---------------------------------------------------------------------------------
# Konstanten (Spec v0.3.1, Teil B)
# ---------------------------------------------------------------------------------
# Kacheln (Spec 12.6, Entscheidung 08.10.2026). Bewertung immer mit den Bereichen der eigenen Konfiguration.
$script:HealthDashKeyMarkers = @(
    @{ Key = 'ApoB';        Marker = 'ApoB';                        Short = 'ApoB' }
    @{ Key = 'HbA1c';       Marker = 'Langzeitzucker (HbA1c)';      Short = 'HbA1c' }
    @{ Key = 'Insulin';     Marker = 'Nüchterninsulin';             Short = 'Nüchterninsulin' }
    @{ Key = 'CRP';         Marker = 'Hochsensitives CRP (hs-CRP)'; Short = 'hs-CRP'; Fallback = 'C-reaktives Protein (CRP)'; FallbackShort = 'CRP' }
    @{ Key = 'Homocystein'; Marker = 'Homocystein';                 Short = 'Homocystein' }
    @{ Key = 'PhenoAge';    Marker = 'PhenoAge-Accel';              Short = 'PhenoAge − Alter'
       Tooltip = 'Klinisches Phenotypic Age (Levine 2018, aus 9 Blutwerten) minus Lebensalter am Testtag. Kein PhenoAgeAccel im Sinne der Literatur und kein DNA-Methylierungs-Alter.' }
)
$script:HealthDashScoreExcludeGroups = @('Genetik', 'Infektionskrankheiten')   # Spec 12.3 Nr. 1
$script:HealthDashScoreExclude       = @()                                     # zusätzliche Marker (Name)
$script:HealthDashAbsoluteTrend      = @('PhenoAge-Accel')                     # Spec 12.5 (F9)
$script:HealthDashTrendRelative      = 0.02                                    # ±2 % (Anzeigeregel)
$script:HealthDashTrendAbsolute      = 0.5                                     # ±0,5 Jahre (F9)
$script:HealthDashWindowDays         = 365                                     # "365 Tage" (inklusive)
$script:HealthDashMinMarkers         = 5                                       # Produktregel 12.3 Nr. 6
$script:HealthDashMinGroups          = 2
$script:HealthDashSparklinePoints    = 12
$script:HealthDashTestosterone = @{
    Total = 'Testosteron, gesamt'; Free = 'Freies Testosteron'; Shbg = 'Sexualhormon-bindendes Globulin (SHBG)'; Albumin = 'Albumin'
    NgdlToNmol = 0.0347; AlbuminDefaultGl = 43.0; AlbuminMinGdl = 1.0; AlbuminMaxGdl = 7.0
}
$script:HealthDashMeaningText = 'Der Score fasst die hinterlegten Markerbewertungen zusammen. 100 Punkte bedeuten die höchste Bewertung nach diesen Regeln, nicht vollständige Gesundheit.'
$script:HealthDashWeightingText = 'Jeder Funktionsbereich zählt gleich viel. Innerhalb eines Bereichs zählen alle Marker gleich.'

# ---------------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------------
function Get-HdField {
    <# Feld aus Hashtable/[ordered] ODER PSCustomObject lesen (PS-5.1-sicher). #>
    param($Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { if ($Object.Contains($Name)) { return $Object[$Name] }; return $null }
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }
    return $null
}

function ConvertTo-HdNumber {
    <# Bereichs- oder Messwert -> [double]; leer, nicht konvertierbar oder nicht endlich -> $null. #>
    param($Value)
    if ($null -eq $Value -or [string]$Value -eq '') { return $null }
    $number = $null
    try { $number = [double]$Value } catch { return $null }
    if ([double]::IsNaN($number) -or [double]::IsInfinity($number)) { return $null }
    return $number
}

function ConvertTo-HdDate {
    <# 'yyyy-MM-dd' oder [datetime] -> Datum (ohne Uhrzeit); sonst $null. #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return $Value.Date }
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact(([string]$Value).Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
    return $null
}

function Format-HdDate {
    param($Date)
    if ($null -eq $Date) { return '' }
    return ([datetime]$Date).ToString('dd.MM.yyyy')
}

function Format-HdNumber {
    <#
        Zahl für die Anzeige (de-DE): bis zu -MaxDecimals Nachkommastellen ohne Nullen am Ende,
        negatives Vorzeichen als "−", optional immer mit Vorzeichen (-Signed, z. B. Änderungen).
    #>
    param($Value, [int]$MaxDecimals = 2, [switch]$Signed)
    if ($null -eq $Value) { return '–' }
    $number = [double]$Value
    $rounded = [Math]::Round($number, $MaxDecimals, [MidpointRounding]::AwayFromZero)
    $culture = [System.Globalization.CultureInfo]::GetCultureInfo('de-DE')
    $text = [Math]::Abs($rounded).ToString('0.' + ('#' * [Math]::Max($MaxDecimals, 0)), $culture).TrimEnd(',')
    if ($MaxDecimals -le 0) { $text = [Math]::Abs($rounded).ToString('0', $culture) }
    if ($rounded -lt 0) { return ([string][char]0x2212) + $text }
    if ($Signed -and $rounded -gt 0) { return '+' + $text }
    if ($Signed -and $rounded -eq 0) { return '±0' }
    return $text
}

function Get-HdMarkerConfig {
    <# Marker-Konfiguration (gemessen oder berechnet) zu einem Namen; $null, wenn unbekannt. #>
    param($Config, [string]$Name)
    # Heißer Pfad: $marker.Name direkt (Hashtable und PSCustomObject, kein StrictMode)
    foreach ($marker in @(Get-HdField $Config 'Markers')) { if ($null -ne $marker -and [string]$marker.Name -eq $Name) { return [PSCustomObject]@{ Config = $marker; IsCalculated = $false } } }
    foreach ($marker in @(Get-HdField $Config 'CalculatedMarkers')) { if ($null -ne $marker -and [string]$marker.Name -eq $Name) { return [PSCustomObject]@{ Config = $marker; IsCalculated = $true } } }
    return $null
}

function Get-HdRangeInfo {
    <#
    .SYNOPSIS
        Bereiche eines Markers (Spec 12.1). Rateable = Get-MarkerRangeScore kann bewerten
        (Optimalbereich vollständig ODER Referenzbereich vollständig). Ungültig: Min > Max.
    #>
    param($MarkerConfig)
    $refMin = ConvertTo-HdNumber (Get-HdField $MarkerConfig 'RefMin'); $refMax = ConvertTo-HdNumber (Get-HdField $MarkerConfig 'RefMax')
    $optMin = ConvertTo-HdNumber (Get-HdField $MarkerConfig 'OptimalMin'); $optMax = ConvertTo-HdNumber (Get-HdField $MarkerConfig 'OptimalMax')
    $invalid = (($null -ne $refMin -and $null -ne $refMax -and $refMin -gt $refMax) -or ($null -ne $optMin -and $null -ne $optMax -and $optMin -gt $optMax))
    $hasOptimal = ($null -ne $optMin -and $null -ne $optMax)
    $hasReference = ($null -ne $refMin -and $null -ne $refMax)
    $reason = ''
    if ($invalid) { $reason = 'keine Bewertung (Bereichskonfiguration ungültig)' } elseif (-not ($hasOptimal -or $hasReference)) { $reason = 'keine Bewertung' }
    return [PSCustomObject]@{
        RefMin = $refMin; RefMax = $refMax; OptMin = $optMin; OptMax = $optMax; HasOptimal = $hasOptimal; HasReference = $hasReference
        Invalid = $invalid; Rateable = ((-not $invalid) -and ($hasOptimal -or $hasReference)); Reason = $reason
    }
}

function Get-HdValueRating {
    <# Bewertung eines Einzelwerts (Spec 12.1) über Get-MarkerRangeScore; Statuswörter wie Risiko-Cockpit. #>
    param([double]$Value, $Range)
    if (-not $Range.Rateable) { return [PSCustomObject]@{ Score = $null; Status = $Range.Reason; Level = 'None' } }
    $score = Get-MarkerRangeScore -Value $Value -RefMin $Range.RefMin -RefMax $Range.RefMax -OptimalMin $Range.OptMin -OptimalMax $Range.OptMax
    if ($null -eq $score) { return [PSCustomObject]@{ Score = $null; Status = 'keine Bewertung'; Level = 'None' } }
    $score = [double]$score
    if ($score -ge 100) { return [PSCustomObject]@{ Score = $score; Status = 'Optimal'; Level = 'Good' } }
    if ($score -ge 50) { return [PSCustomObject]@{ Score = $score; Status = 'Akzeptabel'; Level = 'Warning' } }
    return [PSCustomObject]@{ Score = $score; Status = 'Außerhalb der Norm'; Level = 'Critical' }
}

function Get-HdScoreStatus {
    <# Status eines zusammengefassten Scores aus dem GERUNDETEN Wert (Spec 12.3 Nr. 7). #>
    param($RoundedScore)
    if ($null -eq $RoundedScore) { return [PSCustomObject]@{ Status = ''; Level = 'None' } }
    if ($RoundedScore -ge 80) { return [PSCustomObject]@{ Status = 'Optimal'; Level = 'Good' } }
    if ($RoundedScore -ge 50) { return [PSCustomObject]@{ Status = 'Akzeptabel'; Level = 'Warning' } }
    return [PSCustomObject]@{ Status = 'Optimierungsbedarf'; Level = 'Critical' }
}

function Test-HdEntryValid {
    <# Ein Eintrag ist gültig, wenn der Wert endlich ist und die Einheit leer oder gleich der Konfiguration ist. #>
    param($Entry, [string]$Unit)
    if ($null -eq (ConvertTo-HdNumber (Get-HdField $Entry 'Value'))) { return $false }
    $entryUnit = [string](Get-HdField $Entry 'Unit')
    if ([string]::IsNullOrWhiteSpace($entryUnit) -or [string]::IsNullOrWhiteSpace($Unit)) { return $true }
    return ($entryUnit.Trim() -eq $Unit.Trim())
}

function Get-HdDayState {
    <#
        Bewertet die Einträge EINES Tages: ein ungültiger Eintrag -> 'Invalid'; verschiedene Werte -> 'Conflict';
        keine Einträge -> 'None'; sonst 'Ok' + Value (Spec 12.2, ohne Rückgriff). Schleife statt Pipeline (Tempo).
    #>
    param($DayEntries, [string]$Unit)
    $distinct = @{}
    foreach ($entry in @($DayEntries)) {
        if ($null -eq $entry) { continue }
        if (-not (Test-HdEntryValid -Entry $entry -Unit $Unit)) { return [PSCustomObject]@{ State = 'Invalid'; Value = $null } }
        $distinct[[double](Get-HdField $entry 'Value')] = $true
    }
    if ($distinct.Count -eq 0) { return [PSCustomObject]@{ State = 'None'; Value = $null } }
    if ($distinct.Count -gt 1) { return [PSCustomObject]@{ State = 'Conflict'; Value = $null } }
    return [PSCustomObject]@{ State = 'Ok'; Value = [double]@($distinct.Keys)[0] }
}

function Get-HdDayValue {
    <#
        Wert eines Markers an EINEM Tag: kein Eintrag -> State 'None'; ein ungültiger Eintrag -> 'Invalid';
        verschiedene Werte -> 'Conflict'; sonst 'Ok' + Value (Spec 12.2, ohne Rückgriff).
    #>
    param($Entries, [datetime]$Date, [string]$Unit)
    $target = $Date.Date
    $onDay = New-Object System.Collections.ArrayList
    foreach ($entry in @($Entries)) { if ($null -ne $entry -and (ConvertTo-HdDate (Get-HdField $entry 'Date')) -eq $target) { [void]$onDay.Add($entry) } }
    return (Get-HdDayState -DayEntries $onDay -Unit $Unit)
}

function Get-HdEntryDays {
    <#
        Einträge je Tag (Datum <= Stichtag), aufsteigend; jedes Datum wird nur EINMAL umgewandelt. Ausgabe: { Date; Entries }.
        Heißer Pfad: direkter Feldzugriff ($entry.Date gilt für Hashtable und PSCustomObject, kein StrictMode) statt Get-HdField.
    #>
    param($Entries, [datetime]$AsOf)
    $limit = $AsOf.Date
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $byDay = @{}
    $order = New-Object 'System.Collections.Generic.List[datetime]'
    foreach ($entry in @($Entries)) {
        if ($null -eq $entry) { continue }
        $raw = $entry.Date
        $date = $null
        if ($raw -is [datetime]) { $date = $raw.Date }
        elseif ($null -ne $raw) {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParseExact(([string]$raw).Trim(), 'yyyy-MM-dd', $culture, [System.Globalization.DateTimeStyles]::None, [ref]$parsed)) { $date = $parsed.Date }
        }
        if ($null -eq $date -or $date -gt $limit) { continue }
        if (-not $byDay.ContainsKey($date)) { $byDay[$date] = New-Object System.Collections.ArrayList; $order.Add($date) }
        [void]$byDay[$date].Add($entry)
    }
    $order.Sort()
    foreach ($date in $order) { [PSCustomObject]@{ Date = $date; Entries = $byDay[$date] } }
}

function Get-HdCurrentValue {
    <#
    .SYNOPSIS
        "Aktueller Wert" eines Markers zum Stichtag (Spec 12.2): spätester Eintrag mit Datum <= AsOf.
    .OUTPUTS
        PSCustomObject { State; Value; Date; Unit; Score; Status; Level; Counts; Stale; Text; Range }
        State: NoData | Ok | Stale | Conflict | Invalid | Unrated | RangeInvalid.
        Counts = $true nur für zulässige, bewertbare Werte im 365-Tage-Fenster.
    #>
    param($Entries, $MarkerConfig, [datetime]$AsOf, [int]$WindowDays = $script:HealthDashWindowDays)
    $unit = [string](Get-HdField $MarkerConfig 'Unit')
    $range = Get-HdRangeInfo $MarkerConfig
    $days = @(Get-HdEntryDays -Entries $Entries -AsOf $AsOf)
    $result = [PSCustomObject]@{ State = 'NoData'; Value = $null; Date = $null; Unit = $unit; Score = $null; Status = 'keine Daten'; Level = 'None'; Counts = $false; Stale = $false; Text = 'keine Daten'; Range = $range }
    if ($days.Count -eq 0) { return $result }
    $latest = $days[-1].Date
    $result.Date = $latest
    $day = Get-HdDayState -DayEntries $days[-1].Entries -Unit $unit
    if ($day.State -eq 'Invalid') { $result.State = 'Invalid'; $result.Status = "Ungültiger Wert am $(Format-HdDate $latest)"; $result.Text = $result.Status; return $result }
    if ($day.State -eq 'Conflict') { $result.State = 'Conflict'; $result.Status = "Konflikt: mehrere Werte am $(Format-HdDate $latest)"; $result.Text = $result.Status; return $result }
    $result.Value = $day.Value
    $result.Stale = ($latest -lt $AsOf.Date.AddDays(-$WindowDays))
    if ($range.Invalid) { $result.State = 'RangeInvalid'; $result.Status = $range.Reason; $result.Text = $range.Reason; return $result }
    if (-not $range.Rateable) { $result.State = 'Unrated'; $result.Status = 'keine Bewertung'; $result.Text = 'keine Bewertung'; return $result }
    $rating = Get-HdValueRating -Value $day.Value -Range $range
    $result.Score = $rating.Score; $result.Status = $rating.Status; $result.Level = $rating.Level
    if ($result.Stale) { $result.State = 'Stale'; $result.Text = "$(Format-HdDate $latest) · veraltet" }
    else { $result.State = 'Ok'; $result.Counts = ($null -ne $rating.Score); $result.Text = Format-HdDate $latest }
    return $result
}

function Get-HdHistory {
    <# Gültige Tageswerte eines Markers (ohne Konflikt-/Ungültig-Tage), Datum <= AsOf, aufsteigend. #>
    param($Entries, $MarkerConfig, [datetime]$AsOf)
    $unit = [string](Get-HdField $MarkerConfig 'Unit')
    $points = @()
    foreach ($day in @(Get-HdEntryDays -Entries $Entries -AsOf $AsOf)) {
        $state = Get-HdDayState -DayEntries $day.Entries -Unit $unit
        if ($state.State -eq 'Ok') { $points += , ([PSCustomObject]@{ Date = $day.Date; Value = [double]$state.Value }) }
    }
    return $points
}

function Get-HdTrend {
    <#
    .SYNOPSIS
        Trend aus aktuellem und vorherigem Wert DESSELBEN Markers (Spec 12.5).
        Relativ (>= +2 % / <= -2 %) nur für positive Vorwerte; absolute Änderung (±0,5) für Marker in
        $script:HealthDashAbsoluteTrend. Sonst kein Pfeil.
    .OUTPUTS
        PSCustomObject { Direction ('Up'|'Down'|'Flat'|''); Change; Text }
    #>
    param([double]$Previous, [double]$Current, [string]$MarkerName)
    $none = [PSCustomObject]@{ Direction = ''; Change = $null; Text = '' }
    if ($script:HealthDashAbsoluteTrend -contains $MarkerName) {
        $diff = [Math]::Round($Current - $Previous, 6)
        $direction = 'Flat'
        if ($diff -ge $script:HealthDashTrendAbsolute) { $direction = 'Up' } elseif ($diff -le -$script:HealthDashTrendAbsolute) { $direction = 'Down' }
        return [PSCustomObject]@{ Direction = $direction; Change = $diff; Text = (Format-HdNumber $diff -MaxDecimals 1 -Signed) }
    }
    if ($Previous -le 0) { return $none }
    $relative = [Math]::Round(($Current - $Previous) / $Previous, 6)
    $direction = 'Flat'
    if ($relative -ge $script:HealthDashTrendRelative) { $direction = 'Up' } elseif ($relative -le -$script:HealthDashTrendRelative) { $direction = 'Down' }
    return [PSCustomObject]@{ Direction = $direction; Change = $relative; Text = ((Format-HdNumber ($relative * 100) -MaxDecimals 1 -Signed) + ' %') }
}

function New-HdItemIndex {
    <# Messwerte nach Markername gruppieren (Hashtable Name -> ArrayList). #>
    param($Items)
    $index = @{}
    foreach ($item in @($Items)) {
        if ($null -eq $item) { continue }
        $name = [string]$item.Name
        if (-not $index.ContainsKey($name)) { $index[$name] = New-Object System.Collections.ArrayList }
        [void]$index[$name].Add($item)
    }
    return $index
}

# ---------------------------------------------------------------------------------
# Gesundheitsscore, Vergleich (Spec 12.3, 12.4)
# ---------------------------------------------------------------------------------
function Get-HealthDashScore {
    <#
    .SYNOPSIS
        Gesundheitsscore zum Stichtag (Spec 12.3): gemessene Marker mit bewertbaren Bereichen,
        Bereiche gleich gewichtet (oder einfacher Mittelwert), Mindestmenge 5 Marker aus 2 Bereichen,
        kaufmännisch gerundeter Endwert, Status aus dem gerundeten Wert.
    .PARAMETER Experimental
        E6: "100 im Referenzbereich" - nur im Prototyp, nur Gesamtscore.
    .OUTPUTS
        PSCustomObject { Score (int oder $null); Raw; Status; Level; MarkerCount; GroupCount; Groups[]; Included[]; NotRated[]; Reason; Mode; Experimental }
    #>
    param($Config, [hashtable]$Index, [datetime]$AsOf, [ValidateSet('Weighted', 'Simple')][string]$Mode = 'Weighted', [switch]$Experimental)
    $groups = [ordered]@{}
    $included = New-Object System.Collections.ArrayList
    $notRated = New-Object System.Collections.ArrayList
    foreach ($marker in @(Get-HdField $Config 'Markers')) {
        if ($null -eq $marker) { continue }
        $name = [string](Get-HdField $marker 'Name')
        $group = [string](Get-HdField $marker 'Group')
        if ($script:HealthDashScoreExcludeGroups -contains $group -or $script:HealthDashScoreExclude -contains $name) { continue }
        if (-not $Index.ContainsKey($name)) { continue }
        $current = Get-HdCurrentValue -Entries $Index[$name] -MarkerConfig $marker -AsOf $AsOf
        if ($current.State -eq 'NoData') { continue }
        if (-not $current.Counts) {
            $reason = $current.Text
            if ($current.State -eq 'Stale') { $reason = "älter als $($script:HealthDashWindowDays) Tage ($(Format-HdDate $current.Date))" }
            [void]$notRated.Add([PSCustomObject]@{ Name = $name; Reason = $reason })
            continue
        }
        $score = [double]$current.Score
        if ($Experimental) {
            $low = if ($null -ne $current.Range.RefMin) { $current.Range.RefMin } else { $current.Range.OptMin }
            $high = if ($null -ne $current.Range.RefMax) { $current.Range.RefMax } else { $current.Range.OptMax }
            if ($null -ne $low -and $null -ne $high -and $current.Value -ge $low -and $current.Value -le $high) { $score = 100.0 }
        }
        if ([string]::IsNullOrWhiteSpace($group)) { $group = 'Sonstige' }
        if (-not $groups.Contains($group)) { $groups[$group] = New-Object System.Collections.ArrayList }
        [void]$groups[$group].Add([PSCustomObject]@{ Name = $name; Score = $score; Value = $current.Value; Date = $current.Date })
        [void]$included.Add($name)
    }
    $groupList = @()
    foreach ($groupName in $groups.Keys) {
        $scores = @($groups[$groupName] | ForEach-Object { $_.Score })
        $groupList += , ([PSCustomObject]@{ Name = $groupName; Score = (($scores | Measure-Object -Average).Average); Count = $scores.Count; Markers = @($groups[$groupName]) })
    }
    $markerCount = $included.Count
    $result = [PSCustomObject]@{
        Score = $null; Raw = $null; Status = ''; Level = 'None'; MarkerCount = $markerCount; GroupCount = $groupList.Count
        Groups = @($groupList | Sort-Object -Property Score, Name); Included = @($included); NotRated = @($notRated); Reason = ''; Mode = $Mode; Experimental = [bool]$Experimental
    }
    if ($Index.Count -eq 0) { $result.Reason = 'Noch keine Bluttests erfasst'; return $result }
    if ($markerCount -lt $script:HealthDashMinMarkers -or $groupList.Count -lt $script:HealthDashMinGroups) {
        $result.Reason = "Zu wenige Werte der letzten $($script:HealthDashWindowDays) Tage ($markerCount)"
        return $result
    }
    if ($Mode -eq 'Weighted') { $raw = (@($groupList | ForEach-Object { $_.Score }) | Measure-Object -Average).Average }
    else { $raw = (@($groupList | ForEach-Object { $_.Markers } | ForEach-Object { $_.Score }) | Measure-Object -Average).Average }
    $rounded = [int][Math]::Round([double]$raw, 0, [MidpointRounding]::AwayFromZero)
    $status = Get-HdScoreStatus -RoundedScore $rounded
    $result.Raw = [double]$raw; $result.Score = $rounded; $result.Status = $status.Status; $result.Level = $status.Level
    return $result
}

function Get-HdTestDates {
    <# Alle Testtage (Datum <= Stichtag) der gemessenen Werte, absteigend. #>
    param([hashtable]$Index, [datetime]$AsOf)
    $dates = @{}
    foreach ($name in $Index.Keys) { foreach ($day in @(Get-HdEntryDays -Entries $Index[$name] -AsOf $AsOf)) { $dates[$day.Date] = $day.Date } }
    return @($dates.Values | Sort-Object -Descending)
}

function Get-HealthDashComparison {
    <#
    .SYNOPSIS
        Vergleich (Spec 12.4, E8): Hauptscore (Stichtag heute) gegen Score am vorherigen Testtag D_prev
        (nur Werte <= D_prev, 365-Tage-Fenster relativ zu D_prev). Delta aus GERUNDETEN Werten.
    .OUTPUTS
        PSCustomObject { Available; Delta; Direction; PrevDate; PrevScore; PrevMarkerCount; DataChanged; Text }
    #>
    param($Config, [hashtable]$Index, $Main, [datetime]$Today, [string]$Mode = 'Weighted', [switch]$Experimental)
    $none = [PSCustomObject]@{ Available = $false; Delta = $null; Direction = ''; PrevDate = $null; PrevScore = $null; PrevMarkerCount = 0; DataChanged = $false; Text = 'Kein vergleichbarer vorheriger Score'; Previous = $null }
    $dates = @(Get-HdTestDates -Index $Index -AsOf $Today)
    if ($dates.Count -lt 2 -or $null -eq $Main -or $null -eq $Main.Score) { return $none }
    $prevDate = [datetime]$dates[1]
    $previous = Get-HealthDashScore -Config $Config -Index $Index -AsOf $prevDate -Mode $Mode -Experimental:$Experimental
    $none.PrevDate = $prevDate; $none.Previous = $previous; $none.PrevMarkerCount = $previous.MarkerCount
    if ($null -eq $previous.Score) { return $none }
    $delta = [int]$Main.Score - [int]$previous.Score
    $direction = 'Flat'; if ($delta -gt 0) { $direction = 'Up' } elseif ($delta -lt 0) { $direction = 'Down' }
    $changed = ((@($Main.Included | Sort-Object) -join '|') -ne (@($previous.Included | Sort-Object) -join '|'))
    $text = "$(Format-HdNumber $delta -MaxDecimals 0 -Signed) seit $(Format-HdDate $prevDate)"
    if ($changed) { $text += ' · Datengrundlage verändert' }
    return [PSCustomObject]@{ Available = $true; Delta = $delta; Direction = $direction; PrevDate = $prevDate; PrevScore = $previous.Score; PrevMarkerCount = $previous.MarkerCount; DataChanged = $changed; Text = $text; Previous = $previous }
}

# ---------------------------------------------------------------------------------
# Kacheln, Longevity, Vorbelastungen, Testosteron (Spec 12.5 - 12.8)
# ---------------------------------------------------------------------------------
function Get-HdMarkerEntries {
    <# Einträge eines Markers: gemessen aus dem Index, berechnet über den Provider (Get-CalculatedValuesForMarker). #>
    param([string]$Name, [hashtable]$Index, [scriptblock]$CalcProvider, [bool]$IsCalculated)
    if ($IsCalculated) {
        if ($null -eq $CalcProvider) { return @() }
        return @(& $CalcProvider $Name)
    }
    if ($Index.ContainsKey($Name)) { return @($Index[$Name]) }
    return @()
}

function Get-HealthDashTile {
    <# Daten einer Kachel (ein Marker): aktueller Wert, Status, Verlauf, Trend. #>
    param($Config, [hashtable]$Index, [datetime]$AsOf, [scriptblock]$CalcProvider, [string]$MarkerName, [string]$Title, [string]$Key, [string]$Tooltip = '')
    $markerInfo = Get-HdMarkerConfig -Config $Config -Name $MarkerName
    $tile = [PSCustomObject]@{
        Key = $Key; Title = $Title; MarkerName = $MarkerName; LongName = $MarkerName; Tooltip = $Tooltip; Unit = ''
        State = 'NoConfig'; Value = $null; Date = $null; Status = 'Marker nicht konfiguriert'; Level = 'None'; Stale = $false; DateText = ''
        History = @(); Trend = [PSCustomObject]@{ Direction = ''; Change = $null; Text = '' }
    }
    if (-not $markerInfo) { return $tile }
    $entries = Get-HdMarkerEntries -Name $MarkerName -Index $Index -CalcProvider $CalcProvider -IsCalculated $markerInfo.IsCalculated
    $current = Get-HdCurrentValue -Entries $entries -MarkerConfig $markerInfo.Config -AsOf $AsOf
    $history = @(Get-HdHistory -Entries $entries -MarkerConfig $markerInfo.Config -AsOf $AsOf)
    $tile.Unit = [string](Get-HdField $markerInfo.Config 'Unit')
    $tile.State = $current.State; $tile.Value = $current.Value; $tile.Date = $current.Date; $tile.Status = $current.Status; $tile.Level = $current.Level; $tile.Stale = $current.Stale
    if ($current.State -eq 'NoData') { $tile.Status = 'keine Daten'; $tile.Level = 'None' }
    elseif ($current.State -eq 'Stale') { $tile.DateText = "$(Format-HdDate $current.Date) · veraltet" }
    else { $tile.DateText = Format-HdDate $current.Date }
    $tile.History = @($history | Select-Object -Last $script:HealthDashSparklinePoints)
    if ($null -ne $current.Value -and $current.State -notin @('Invalid', 'Conflict')) {
        $before = @($history | Where-Object { $_.Date -lt $current.Date })
        if ($before.Count -gt 0) { $tile.Trend = Get-HdTrend -Previous $before[-1].Value -Current $current.Value -MarkerName $MarkerName }
    }
    return $tile
}

function Get-HealthDashTiles {
    <# Alle Kacheln aus $script:HealthDashKeyMarkers; hs-CRP/CRP nach Ersatzregel (Spec 12.6). #>
    param($Config, [hashtable]$Index, [datetime]$AsOf, [scriptblock]$CalcProvider)
    $tiles = @()
    foreach ($definition in $script:HealthDashKeyMarkers) {
        $markerName = [string]$definition.Marker; $title = [string]$definition.Short
        if ($definition.ContainsKey('Fallback') -and $definition.Fallback) {
            $primaryCfg = Get-HdMarkerConfig -Config $Config -Name $definition.Marker
            $fallbackCfg = Get-HdMarkerConfig -Config $Config -Name $definition.Fallback
            $primary = $null; $fallback = $null
            if ($primaryCfg) { $primary = Get-HdCurrentValue -Entries (Get-HdMarkerEntries -Name $definition.Marker -Index $Index -CalcProvider $CalcProvider -IsCalculated $primaryCfg.IsCalculated) -MarkerConfig $primaryCfg.Config -AsOf $AsOf }
            if ($fallbackCfg) { $fallback = Get-HdCurrentValue -Entries (Get-HdMarkerEntries -Name $definition.Fallback -Index $Index -CalcProvider $CalcProvider -IsCalculated $fallbackCfg.IsCalculated) -MarkerConfig $fallbackCfg.Config -AsOf $AsOf }
            $useFallback = $false
            if ($primary -and $primary.State -eq 'Ok') { $useFallback = $false }
            elseif ($fallback -and $fallback.State -eq 'Ok') { $useFallback = $true }
            elseif ($fallback -and $fallback.State -ne 'NoData' -and (-not $primary -or $primary.State -eq 'NoData' -or ($null -ne $fallback.Date -and $null -ne $primary.Date -and $fallback.Date -gt $primary.Date))) { $useFallback = $true }
            if ($useFallback) { $markerName = [string]$definition.Fallback; $title = [string]$definition.FallbackShort }
        }
        $tooltip = if ($definition.ContainsKey('Tooltip')) { [string]$definition.Tooltip } else { '' }
        $tiles += , (Get-HealthDashTile -Config $Config -Index $Index -AsOf $AsOf -CalcProvider $CalcProvider -MarkerName $markerName -Title $title -Key ([string]$definition.Key) -Tooltip $tooltip)
    }
    return $tiles
}

function Get-HealthDashLongevity {
    <# Longevity-Score als Zweitwert (Spec 4.2): letzter berechneter Wert, kaufmännisch gerundet, Status wie Longevity-Reiter. #>
    param($Entries, [datetime]$AsOf)
    $points = @(@($Entries) | Where-Object { $null -ne $_ } | ForEach-Object {
        $d = ConvertTo-HdDate (Get-HdField $_ 'Date'); $v = ConvertTo-HdNumber (Get-HdField $_ 'Value')
        if ($null -ne $d -and $null -ne $v -and $d -le $AsOf.Date) { [PSCustomObject]@{ Date = $d; Value = $v } }
    } | Sort-Object -Property Date)
    if ($points.Count -eq 0) { return [PSCustomObject]@{ Available = $false; Score = $null; Status = ''; Level = 'None'; Date = $null; Stale = $false; Text = 'noch nicht berechenbar' } }
    $last = $points[-1]
    $rounded = [int][Math]::Round([double]$last.Value, 0, [MidpointRounding]::AwayFromZero)
    $status = Get-HdScoreStatus -RoundedScore $rounded
    $stale = ($last.Date -lt $AsOf.Date.AddDays(-$script:HealthDashWindowDays))
    $text = if ($stale) { "$(Format-HdDate $last.Date) · veraltet" } else { "Stand $(Format-HdDate $last.Date)" }
    return [PSCustomObject]@{ Available = $true; Score = $rounded; Status = $status.Status; Level = $status.Level; Date = $last.Date; Stale = $stale; Text = $text }
}

function Get-HealthDashPredispositions {
    <#
    .SYNOPSIS
        Vorbelastungs-Scores (Spec 12.7): je aktiver Vorbelastung m = eindeutige Marker, k = davon mit
        zulässigem, bewertbarem aktuellem Wert; Score = gerundeter Mittelwert (ab k >= 1); Fokus = 3 schlechteste.
    .OUTPUTS
        PSCustomObject { State ('Disabled'|'NoneActive'|'Active'); Entries[] }
    #>
    param($Config, [hashtable]$Index, [datetime]$AsOf, [scriptblock]$CalcProvider)
    $gp = Get-HdField $Config 'GeneticPredispositions'
    if (-not $gp -or -not [bool](Get-HdField $gp 'Enabled')) { return [PSCustomObject]@{ State = 'Disabled'; Entries = @() } }
    $all = @(@(Get-HdField $gp 'BuiltIn') + @(Get-HdField $gp 'Custom') | Where-Object { $null -ne $_ -and [bool](Get-HdField $_ 'Active') })
    if ($all.Count -eq 0) { return [PSCustomObject]@{ State = 'NoneActive'; Entries = @() } }
    $entries = @()
    foreach ($predisposition in $all) {
        $names = New-Object System.Collections.ArrayList
        foreach ($markerName in @(Get-HdField $predisposition 'Markers')) { $n = [string]$markerName; if ($n -and -not $names.Contains($n)) { [void]$names.Add($n) } }
        $rated = @(); $notRated = @()
        foreach ($name in $names) {
            $info = Get-HdMarkerConfig -Config $Config -Name $name
            if (-not $info) { $notRated += , ([PSCustomObject]@{ Name = $name; Reason = 'unbekannter Marker' }); continue }
            $current = Get-HdCurrentValue -Entries (Get-HdMarkerEntries -Name $name -Index $Index -CalcProvider $CalcProvider -IsCalculated $info.IsCalculated) -MarkerConfig $info.Config -AsOf $AsOf
            if ($current.Counts) {
                $rated += , ([PSCustomObject]@{ Name = $name; Value = $current.Value; Unit = [string](Get-HdField $info.Config 'Unit'); Score = [double]$current.Score; Status = $current.Status; Level = $current.Level; Date = $current.Date; IsCalculated = $info.IsCalculated })
            } else {
                $reason = switch ($current.State) { 'NoData' { 'kein Wert' } 'Stale' { "kein aktueller Wert (letzter am $(Format-HdDate $current.Date))" } default { $current.Text } }
                $notRated += , ([PSCustomObject]@{ Name = $name; Reason = $reason })
            }
        }
        $k = $rated.Count; $m = $names.Count
        $score = $null; $status = [PSCustomObject]@{ Status = 'Keine aktuellen bewertbaren Werte'; Level = 'None' }
        if ($k -ge 1) {
            $score = [int][Math]::Round([double](($rated | ForEach-Object { $_.Score } | Measure-Object -Average).Average), 0, [MidpointRounding]::AwayFromZero)
            $status = Get-HdScoreStatus -RoundedScore $score
        }
        $entries += , ([PSCustomObject]@{
            Name = [string](Get-HdField $predisposition 'Name'); Score = $score; Status = $status.Status; Level = $status.Level; K = $k; M = $m
            Focus = @($rated | Sort-Object -Property Score, Name | Select-Object -First 3); Rated = $rated; NotRated = $notRated
            CalculatedNames = @($rated | Where-Object { $_.IsCalculated } | ForEach-Object { $_.Name })
        })
    }
    return [PSCustomObject]@{ State = 'Active'; Entries = $entries }
}

function Get-HealthDashPredispositionTiles {
    <# v2.37.0: Alle Marker je aktiver Vorbelastung; Reihenfolge erhalten, je Gruppe eindeutig. #>
    param($Config, [hashtable]$Index, [datetime]$AsOf, [scriptblock]$CalcProvider)
    $gp = Get-HdField $Config 'GeneticPredispositions'
    if (-not $gp -or -not [bool](Get-HdField $gp 'Enabled')) {
        return [PSCustomObject]@{ State = 'Disabled'; Groups = @() }
    }
    $active = @(@(Get-HdField $gp 'BuiltIn') + @(Get-HdField $gp 'Custom') | Where-Object { $null -ne $_ -and [bool](Get-HdField $_ 'Active') })
    $groups = @(); $tileCache = @{}
    foreach ($predisposition in $active) {
        $seen = @{}; $tiles = @()
        foreach ($rawName in @(Get-HdField $predisposition 'Markers')) {
            $name = ([string]$rawName).Trim()
            if (-not $name -or $seen.ContainsKey($name)) { continue }
            $seen[$name] = $true
            if (-not $tileCache.ContainsKey($name)) {
                $tileCache[$name] = Get-HealthDashTile -Config $Config -Index $Index -AsOf $AsOf -CalcProvider $CalcProvider -MarkerName $name -Title $name -Key $name
            }
            $tiles += , $tileCache[$name]
        }
        $groups += , ([PSCustomObject]@{ Name = [string](Get-HdField $predisposition 'Name'); Tiles = @($tiles) })
    }
    $state = if ($groups.Count -gt 0) { 'Active' } else { 'NoneActive' }
    return [PSCustomObject]@{ State = $state; Groups = @($groups) }
}

function Get-HealthDashAbnormalities {
    <# v2.38.0: Eindeutige aktuelle Auffaelligkeiten; vorhandene Bewertung, keine eigene Skala. #>
    param($Config, [hashtable]$Index, [datetime]$AsOf, [scriptblock]$CalcProvider)
    $seen = @{}; $ranked = @()
    foreach ($definition in @(@(Get-HdField $Config 'Markers') + @(Get-HdField $Config 'CalculatedMarkers'))) {
        if ($null -eq $definition) { continue }
        $name = [string](Get-HdField $definition 'Name')
        if ([string]::IsNullOrWhiteSpace($name) -or $seen.ContainsKey($name)) { continue }
        $seen[$name] = $true
        # Gleiche Konfigurationsaufloesung wie die Kacheln (gemessen vor berechnet).
        $info = Get-HdMarkerConfig -Config $Config -Name $name
        if (-not $info -or ([string](Get-HdField $info.Config 'Group')).Trim() -eq 'Genetik') { continue }
        $range = Get-HdRangeInfo -MarkerConfig $info.Config
        if (-not $range.HasReference -or $range.Invalid) { continue }
        $tile = Get-HealthDashTile -Config $Config -Index $Index -AsOf $AsOf -CalcProvider $CalcProvider -MarkerName $name -Title $name -Key $name
        if ($tile.State -ne 'Ok' -or $tile.Level -ne 'Critical' -or $tile.Stale) { continue }
        if ($tile.Value -ge $range.RefMin -and $tile.Value -le $range.RefMax) { continue }
        $rating = Get-HdValueRating -Value $tile.Value -Range $range
        if ($null -eq $rating.Score) { continue }
        $ranked += , ([PSCustomObject]@{ Score = [double]$rating.Score; Name = $name; Tile = $tile })
    }
    $tiles = @($ranked | Sort-Object -Property Score, Name | ForEach-Object { $_.Tile })
    return [PSCustomObject]@{ Count = $tiles.Count; Tiles = $tiles }
}

function Get-HdVermeulenFreeTestosterone {
    <#
    .SYNOPSIS
        Freies Testosteron nach Vermeulen 1999 (Spec 12.8), numerisch stabile Lösung.
        K(SHBG) = 1,0e9 l/mol, K(Alb) = 3,6e4 l/mol, Albumin 69 000 g/mol.
    .OUTPUTS
        PSCustomObject { FreePgMl; FreePercent; FreeMolL }
    #>
    param([double]$TotalNmol, [double]$ShbgNmol, [double]$AlbuminGl)
    $kShbg = 1.0e9; $kAlb = 3.6e4
    $tt = $TotalNmol * 1e-9; $shbg = $ShbgNmol * 1e-9; $alb = $AlbuminGl / 69000.0
    $n = 1 + $kAlb * $alb
    $a = $n * $kShbg
    $b = $n + $kShbg * ($shbg - $tt)
    $root = [Math]::Sqrt($b * $b + 4 * $a * $tt)
    if ($b -ge 0) { $ft = 2 * $tt / ($b + $root) } else { $ft = (-$b + $root) / (2 * $a) }
    return [PSCustomObject]@{ FreePgMl = $ft * 288.42e9; FreePercent = 100 * $ft / $tt; FreeMolL = $ft }
}

function Get-HealthDashTestosterone {
    <#
    .SYNOPSIS
        Karte Testosteron (Spec 4.5, 12.8): drei Messwerte mit Bewertung + berechneter Abschnitt
        (FAI, freies Testosteron nach Vermeulen) am letzten GEMEINSAMEN Tag von Gesamt-Testosteron und SHBG.
    #>
    param($Config, [hashtable]$Index, [datetime]$AsOf)
    $t = $script:HealthDashTestosterone
    $rows = @()
    foreach ($row in @(@{ Key = 'Total'; Name = $t.Total; Label = 'Testosteron, gesamt' }, @{ Key = 'Free'; Name = $t.Free; Label = 'Freies Testosteron (gemessen)' }, @{ Key = 'Shbg'; Name = $t.Shbg; Label = 'SHBG' })) {
        $info = Get-HdMarkerConfig -Config $Config -Name $row.Name
        $entry = [PSCustomObject]@{ Key = $row.Key; Label = $row.Label; MarkerName = $row.Name; Configured = [bool]$info; State = 'NoData'; Value = $null; Unit = ''; Date = $null; DateText = ''; Status = 'keine Daten'; Level = 'None'; Range = $null; ExtraText = '' }
        if ($info) {
            $current = Get-HdCurrentValue -Entries (Get-HdMarkerEntries -Name $row.Name -Index $Index -IsCalculated $false) -MarkerConfig $info.Config -AsOf $AsOf
            $entry.State = $current.State; $entry.Value = $current.Value; $entry.Unit = $current.Unit; $entry.Date = $current.Date; $entry.Status = $current.Status; $entry.Level = $current.Level; $entry.Range = $current.Range
            if ($current.State -eq 'NoData') { $entry.Status = 'keine Daten' }
            if ($null -ne $current.Date) { $entry.DateText = if ($current.Stale) { "$(Format-HdDate $current.Date) · veraltet" } else { Format-HdDate $current.Date } }
            if ($row.Key -eq 'Total' -and $null -ne $current.Value) { $entry.ExtraText = "($(Format-HdNumber ($current.Value * $t.NgdlToNmol) -MaxDecimals 2) nmol/l)" }
        } else { $entry.Status = 'Marker nicht konfiguriert' }
        $rows += , $entry
    }
    $computed = [PSCustomObject]@{ Available = $false; Reason = ''; Date = $null; FreePgMl = $null; FreePercent = $null; Fai = $null; AlbuminEstimated = $false; AlbuminGdl = $null }
    $ttInfo = Get-HdMarkerConfig -Config $Config -Name $t.Total; $shbgInfo = Get-HdMarkerConfig -Config $Config -Name $t.Shbg; $albInfo = Get-HdMarkerConfig -Config $Config -Name $t.Albumin
    if (-not $ttInfo -or -not $shbgInfo) { $computed.Reason = 'Gesamt-Testosteron oder SHBG ist nicht konfiguriert'; return [PSCustomObject]@{ Rows = $rows; Computed = $computed } }
    $ttEntries = Get-HdMarkerEntries -Name $t.Total -Index $Index -IsCalculated $false
    $shbgEntries = Get-HdMarkerEntries -Name $t.Shbg -Index $Index -IsCalculated $false
    $ttDates = @{}; foreach ($e in $ttEntries) { $d = ConvertTo-HdDate (Get-HdField $e 'Date'); if ($null -ne $d -and $d -le $AsOf.Date) { $ttDates[$d.ToString('yyyy-MM-dd')] = $d } }
    $common = @()
    foreach ($e in $shbgEntries) { $d = ConvertTo-HdDate (Get-HdField $e 'Date'); if ($null -ne $d -and $ttDates.ContainsKey($d.ToString('yyyy-MM-dd'))) { $common += $d } }
    $common = @($common | Sort-Object -Unique)
    if ($common.Count -eq 0) { $computed.Reason = 'kein gemeinsamer Testtag von Gesamt-Testosteron und SHBG'; return [PSCustomObject]@{ Rows = $rows; Computed = $computed } }
    $day = [datetime]$common[-1]
    $computed.Date = $day
    $ttDay = Get-HdDayValue -Entries $ttEntries -Date $day -Unit ([string](Get-HdField $ttInfo.Config 'Unit'))
    $shbgDay = Get-HdDayValue -Entries $shbgEntries -Date $day -Unit ([string](Get-HdField $shbgInfo.Config 'Unit'))
    if ($ttDay.State -ne 'Ok' -or $ttDay.Value -le 0) { $computed.Reason = 'Gesamt-Testosteron ungültig'; return [PSCustomObject]@{ Rows = $rows; Computed = $computed } }
    if ($shbgDay.State -ne 'Ok' -or $shbgDay.Value -le 0) { $computed.Reason = 'SHBG ungültig'; return [PSCustomObject]@{ Rows = $rows; Computed = $computed } }
    $albuminGl = $t.AlbuminDefaultGl; $estimated = $true
    if ($albInfo) {
        $albDay = Get-HdDayValue -Entries (Get-HdMarkerEntries -Name $t.Albumin -Index $Index -IsCalculated $false) -Date $day -Unit ([string](Get-HdField $albInfo.Config 'Unit'))
        if ($albDay.State -ne 'None') {
            if ($albDay.State -ne 'Ok' -or $albDay.Value -lt $t.AlbuminMinGdl -or $albDay.Value -gt $t.AlbuminMaxGdl) {
                $computed.Reason = 'Albumin ungültig – keine Berechnung'; return [PSCustomObject]@{ Rows = $rows; Computed = $computed }
            }
            $albuminGl = $albDay.Value * 10; $estimated = $false
        }
    }
    $totalNmol = $ttDay.Value * $t.NgdlToNmol
    $free = Get-HdVermeulenFreeTestosterone -TotalNmol $totalNmol -ShbgNmol $shbgDay.Value -AlbuminGl $albuminGl
    $computed.Available = $true; $computed.FreePgMl = $free.FreePgMl; $computed.FreePercent = $free.FreePercent
    $computed.Fai = $totalNmol / $shbgDay.Value * 100; $computed.AlbuminEstimated = $estimated; $computed.AlbuminGdl = $albuminGl / 10
    return [PSCustomObject]@{ Rows = $rows; Computed = $computed }
}

# ---------------------------------------------------------------------------------
# Gesamtmodell und Erklärung
# ---------------------------------------------------------------------------------
function Get-HdCustomViews {
    # v2.40.0: Ungueltige Ansichten blockieren nicht die eingebauten Ansichten.
    param($Config, [hashtable]$Index, [datetime]$AsOf, [scriptblock]$CalcProvider)
    try { $definitions = @(ConvertTo-BtDashboardViews -Views $Config.DashboardViews) }
    catch { return [PSCustomObject]@{ Views = @(); Error = 'Eigene Ansichten sind fehlerhaft gespeichert. Die Konfiguration bleibt unverändert; Erstellen und Bearbeiten sind gesperrt.' } }
    $views = @(); $cache = @{}
    foreach ($definition in $definitions) {
        $tiles = @()
        foreach ($markerName in @($definition.Markers)) {
            if (-not $cache.ContainsKey($markerName)) {
                $tile = Get-HealthDashTile -Config $Config -Index $Index -AsOf $AsOf -CalcProvider $CalcProvider -MarkerName $markerName -Title $markerName -Key $markerName
                if ($tile.State -eq 'NoConfig') { $tile.Status = 'Marker nicht mehr verfügbar' }
                $cache[$markerName] = $tile
            }
            $tiles += $cache[$markerName]
        }
        $views += [PSCustomObject]@{ Id = $definition.Id; Name = $definition.Name; Tiles = @($tiles) }
    }
    return [PSCustomObject]@{ Views = @($views); Error = '' }
}

function Save-HdDashboardViews {
    # Erst persistieren, dann den aktuellen Arbeitsspeicher aktualisieren.
    param($Views)
    $ErrorActionPreference = 'Stop'
    if (-not $script:data -or -not $script:data.Config) { throw 'Die Konfiguration ist nicht geladen.' }
    $null = @(ConvertTo-BtDashboardViews -Views $script:data.Config.DashboardViews)
    $normalized = @(ConvertTo-BtDashboardViews -Views $Views)
    $candidateConfig = @{}
    foreach ($key in $script:data.Config.Keys) { $candidateConfig[$key] = $script:data.Config[$key] }
    $candidateConfig.DashboardViews = $normalized
    Save-Data -data @{ Config = $candidateConfig } -type 'Config' -ThrowOnError
    $script:data.Config.DashboardViews = $normalized
}

function Remove-HdDashboardView {
    param([string]$Id)
    $views = @(ConvertTo-BtDashboardViews -Views $script:data.Config.DashboardViews)
    if (@($views | Where-Object { $_.Id -eq $Id }).Count -ne 1) { throw 'Die eigene Ansicht ist nicht mehr vorhanden.' }
    Save-HdDashboardViews -Views @($views | Where-Object { $_.Id -ne $Id })
}

function Resolve-HdViewRow {
    param($Row)
    $query = $Row.Combo.Text.Trim()
    # Ein geloeschter Marker darf unveraendert erhalten bleiben, aber nicht neu angelegt werden.
    if ($Row.Original -and $query -ieq $Row.Original) { return $Row.Original }
    return Resolve-ListSelection -Query $query -Items $Row.Editor.Items
}

function Update-HdViewRowSearch {
    param($Combo, [switch]$All)
    $row = $Combo.Tag
    if ($row.Busy) { return }
    $typed = $Combo.Text; $caret = $Combo.SelectionStart
    $items = if ($All -or [string]::IsNullOrWhiteSpace($typed)) { @($row.Editor.Items) } else { @(Get-ListMatches -Query $typed -Items $row.Editor.Items) }
    $row.Busy = $true
    $Combo.BeginUpdate()
    try {
        $Combo.Items.Clear()
        if ($items.Count -gt 0) { $Combo.Items.AddRange([object[]]$items) }
        $Combo.SelectedIndex = -1; $Combo.Text = $typed
        $Combo.SelectionStart = [Math]::Min($caret, $typed.Length); $Combo.SelectionLength = 0
    } finally { $Combo.EndUpdate(); $row.Busy = $false }
    if (-not $All) {
        $Combo.ForeColor = if ($items.Count -eq 0 -and $typed) { [System.Drawing.Color]::Firebrick } else { [System.Drawing.SystemColors]::WindowText }
        if ($Combo.Focused) {
            # Automatisches Aufklappen darf nicht ueber DropDown die Trefferliste zuruecksetzen.
            $row.Busy = $true
            try { $Combo.DroppedDown = ($items.Count -gt 0); $Combo.SelectionStart = [Math]::Min($caret, $typed.Length); $Combo.SelectionLength = 0 }
            finally { $row.Busy = $false }
        }
    }
}

function Add-HdViewEditorRow {
    param($Editor, [string]$MarkerName = '')
    $panel = New-Object System.Windows.Forms.Panel
    $panel.Size = New-Object System.Drawing.Size(690, 56); $panel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 4)
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $combo.Location = New-Object System.Drawing.Point(0, 0); $combo.Size = New-Object System.Drawing.Size(638, 25)
    $combo.DropDownWidth = 638; $combo.IntegralHeight = $false; $combo.MaxDropDownItems = 12
    $combo.AccessibleName = 'Marker suchen oder auswählen'; $combo.TabIndex = 0
    if ($Editor.Items.Count -gt 0) { $combo.Items.AddRange([object[]]$Editor.Items) }
    $combo.Text = $MarkerName
    $hint = New-Object System.Windows.Forms.Label
    $hint.Location = New-Object System.Drawing.Point(0, 29); $hint.Size = New-Object System.Drawing.Size(680, 23)
    $hint.ForeColor = [System.Drawing.Color]::Firebrick; $hint.UseMnemonic = $false
    if ($MarkerName -and $Editor.Items -notcontains $MarkerName) { $hint.Text = 'Marker nicht mehr verfügbar – bleibt erhalten, bis Sie diese Zeile entfernen oder ändern.' }
    $remove = New-Object System.Windows.Forms.Button
    $remove.Text = '×'; $remove.Location = New-Object System.Drawing.Point(648, 0); $remove.Size = New-Object System.Drawing.Size(36, 26)
    $remove.TabIndex = 1; $remove.AccessibleName = 'Diese Markerzeile entfernen'
    $row = @{ Editor = $Editor; Panel = $panel; Combo = $combo; Hint = $hint; Original = $MarkerName; Busy = $false }
    $combo.Tag = $row; $remove.Tag = $row
    $combo.Add_TextUpdate({ Update-HdViewRowSearch -Combo $args[0] })
    $combo.Add_DropDown({ Update-HdViewRowSearch -Combo $args[0] -All })
    $combo.Add_SelectionChangeCommitted({
        $rowState = $args[0].Tag
        $rowState.Hint.Text = ''
        $args[0].ForeColor = [System.Drawing.SystemColors]::WindowText
    })
    $combo.Add_KeyDown({
        if ($args[1].KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $resolved = Resolve-HdViewRow -Row $args[0].Tag
            if ($resolved) { $args[0].Text = $resolved; $args[0].DroppedDown = $false; $args[0].Tag.Hint.Text = '' }
            else { $args[0].Tag.Hint.Text = 'Bitte einen eindeutigen Marker auswählen.' }
            $args[1].Handled = $true; $args[1].SuppressKeyPress = $true
        }
    })
    $remove.Add_Click({
        $rowState = $args[0].Tag; $editorState = $rowState.Editor
        $editorState.Rows.Remove($rowState)
        $editorState.RowPanel.Controls.Remove($rowState.Panel)
        $rowState.Panel.Dispose()
        $editorState.Error.Text = ''; $editorState.Add.Focus() | Out-Null
    })
    $panel.Controls.AddRange([System.Windows.Forms.Control[]]@($combo, $remove, $hint))
    $panel.TabIndex = $Editor.Rows.Count
    [void]$Editor.Rows.Add($row); $Editor.RowPanel.Controls.Add($panel)
    $Editor.RowPanel.ScrollControlIntoView($panel)
    return $row
}

function Get-HdViewEditorCandidate {
    param($Editor)
    if ([string]::IsNullOrWhiteSpace($Editor.Name.Text)) { throw 'Bitte einen Namen für die Ansicht eingeben.' }
    $name = $Editor.Name.Text.Trim()
    if ($name -match '^(Standard|Vorbelastungen|Aktuelle Auffälligkeiten(?:\s*\(\d+\))?)$') { throw 'Dieser Name ist für eine eingebaute Ansicht reserviert. Bitte einen anderen Namen wählen.' }
    if (@($Editor.Views | Where-Object { $_.Id -ne $Editor.Id -and $_.Name -ieq $name }).Count -gt 0) { throw 'Eine eigene Ansicht mit diesem Namen ist bereits vorhanden. Bitte einen anderen Namen wählen.' }
    $markers = @(); $seen = @{}; $position = 0
    foreach ($row in $Editor.Rows) {
        $position++
        $marker = Resolve-HdViewRow -Row $row
        if (-not $marker) { throw "Zeile ${position}: Bitte einen eindeutigen Marker auswählen oder die Zeile entfernen." }
        if ($seen.ContainsKey($marker)) { throw "Zeile ${position}: Dieser Marker ist bereits ausgewählt." }
        $seen[$marker] = $true; $markers += $marker
    }
    if ($markers.Count -eq 0) { throw 'Bitte mindestens einen Marker hinzufügen.' }
    $entry = [PSCustomObject]@{ Id = $Editor.Id; Name = $Editor.Name.Text.Trim(); Markers = @($markers) }
    $views = @(); $replaced = $false
    foreach ($view in $Editor.Views) {
        if ($view.Id -eq $Editor.Id) { $views += $entry; $replaced = $true } else { $views += $view }
    }
    if (-not $replaced) { $views += $entry }
    return @(ConvertTo-BtDashboardViews -Views $views)
}

function New-HdViewEditor {
    # Dialog wird getrennt aufgebaut, damit dieselben Bedienelemente ohne Start der App testbar sind.
    param($Config, [string]$Id = '', [scriptblock]$SaveViews = { param($Views) Save-HdDashboardViews -Views $Views })
    $views = @(ConvertTo-BtDashboardViews -Views $Config.DashboardViews)
    $existing = @($views | Where-Object { $_.Id -eq $Id })
    if ($Id -and $existing.Count -ne 1) { throw 'Die eigene Ansicht ist nicht mehr vorhanden.' }
    $items = @(@($Config.Markers) + @($Config.CalculatedMarkers) | Where-Object { $_ -and $_.Name } | ForEach-Object { [string]$_.Name } | Sort-Object -Unique)
    $form = New-Object System.Windows.Forms.Form
    $form.Text = if ($Id) { 'Eigene Ansicht bearbeiten' } else { 'Eigene Ansicht erstellen' }
    $form.ClientSize = New-Object System.Drawing.Size(752, 580)
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.ShowInTaskbar = $false
    $form.Font = $script:hdFonts.Normal
    $label = New-Object System.Windows.Forms.Label
    $label.Text = 'Name der Ansicht'; $label.AutoSize = $true; $label.Location = New-Object System.Drawing.Point(20, 18)
    $nameBox = New-Object System.Windows.Forms.TextBox
    $nameBox.Location = New-Object System.Drawing.Point(20, 42); $nameBox.Size = New-Object System.Drawing.Size(712, 25); $nameBox.TabIndex = 0
    $help = New-Object System.Windows.Forms.Label
    $help.Text = 'Marker auswählen oder suchen (auch Kürzel/Aliase). Die Reihenfolge bestimmt die Kacheln.'
    $help.Location = New-Object System.Drawing.Point(20, 82); $help.Size = New-Object System.Drawing.Size(712, 40)
    $rowsPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $rowsPanel.Location = New-Object System.Drawing.Point(20, 168); $rowsPanel.Size = New-Object System.Drawing.Size(712, 300)
    $rowsPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::TopDown
    $rowsPanel.WrapContents = $false; $rowsPanel.AutoScroll = $true; $rowsPanel.TabIndex = 2
    $add = New-Object System.Windows.Forms.Button
    $add.Text = '+ Marker hinzufügen'; $add.Location = New-Object System.Drawing.Point(20, 126); $add.Size = New-Object System.Drawing.Size(170, 30); $add.TabIndex = 1
    $errorLabel = New-Object System.Windows.Forms.Label
    $errorLabel.Location = New-Object System.Drawing.Point(20, 476); $errorLabel.Size = New-Object System.Drawing.Size(712, 48)
    $errorLabel.ForeColor = [System.Drawing.Color]::Firebrick; $errorLabel.UseMnemonic = $false
    $save = New-Object System.Windows.Forms.Button
    $save.Text = 'Speichern'; $save.Location = New-Object System.Drawing.Point(512, 536); $save.Size = New-Object System.Drawing.Size(106, 30); $save.TabIndex = 3
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Abbrechen'; $cancel.Location = New-Object System.Drawing.Point(626, 536); $cancel.Size = New-Object System.Drawing.Size(106, 30); $cancel.TabIndex = 4
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $form.CancelButton = $cancel
    $state = @{ Form = $form; Name = $nameBox; Rows = (New-Object System.Collections.ArrayList); RowPanel = $rowsPanel; Items = @($items); Views = @($views); Id = $Id; Error = $errorLabel; Add = $add; Save = $save; Cancel = $cancel; SaveViews = $SaveViews }
    if (-not $Id) { $state.Id = [guid]::NewGuid().ToString('D') } else { $nameBox.Text = $existing[0].Name }
    $form.Tag = $state; $add.Tag = $state; $save.Tag = $state
    $form.Controls.AddRange([System.Windows.Forms.Control[]]@($label, $nameBox, $help, $rowsPanel, $add, $errorLabel, $save, $cancel))
    if ($Id -and $existing[0].Markers.Count -gt 0) { foreach ($marker in $existing[0].Markers) { $null = Add-HdViewEditorRow -Editor $state -MarkerName $marker } }
    else { $null = Add-HdViewEditorRow -Editor $state }
    $rowsPanel.AutoScrollPosition = New-Object System.Drawing.Point(0, 0)
    $add.Add_Click({ $newRow = Add-HdViewEditorRow -Editor $args[0].Tag; $newRow.Combo.Focus() | Out-Null })
    $save.Add_Click({
        $editorState = $args[0].Tag
        $editorState.Error.Text = ''
        try { $candidate = @(Get-HdViewEditorCandidate -Editor $editorState) }
        catch { $editorState.Error.Text = $_.Exception.Message; return }
        try { & $editorState.SaveViews $candidate }
        catch { $editorState.Error.Text = 'Speichern fehlgeschlagen. Die Ansicht wurde nicht übernommen. Bitte Speicherort und Schreibzugriff prüfen und erneut versuchen.'; return }
        $editorState.Form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $editorState.Form.Close()
    })
    $form.ActiveControl = $nameBox
    return $form
}

function Show-HdViewEditor {
    param([string]$Id = '')
    $dialog = $null
    try {
        $dialog = New-HdViewEditor -Config $script:data.Config -Id $Id
        if ($dialog.ShowDialog($script:hdControls.Content.FindForm()) -eq [System.Windows.Forms.DialogResult]::OK) {
            $selectedId = $dialog.Tag.Id
            Update-HealthDashboardTab -Force
            Set-HealthDashboardView -Name ('Custom:' + $selectedId)
        }
    } catch { [System.Windows.Forms.MessageBox]::Show('Die eigene Ansicht konnte nicht geöffnet werden. Bitte die gespeicherten Ansichtsdefinitionen prüfen.', 'Eigene Ansicht', 'OK', 'Warning') | Out-Null }
    finally { if ($dialog) { $dialog.Dispose() } }
}

function Show-HdDeleteViewConfirmation {
    if ($script:hdSelectedView -notlike 'Custom:*') { return }
    $id = $script:hdSelectedView.Substring(7)
    $view = @($script:hdModel.CustomViews.Views | Where-Object { $_.Id -eq $id })
    if ($view.Count -ne 1) { return }
    $answer = [System.Windows.Forms.MessageBox]::Show($script:hdControls.Content.FindForm(), ('Eigene Ansicht „' + $view[0].Name + '“ löschen? Die Messwerte bleiben erhalten.'), 'Ansicht löschen', [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question, [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    try { Remove-HdDashboardView -Id $id }
    catch { [System.Windows.Forms.MessageBox]::Show('Löschen fehlgeschlagen. Die Ansicht wurde nicht entfernt. Bitte Speicherort und Schreibzugriff prüfen.', 'Eigene Ansicht', 'OK', 'Warning') | Out-Null; return }
    Update-HealthDashboardTab -Force
    Set-HealthDashboardView -Name 'Standard'
}


function Get-HealthDashboardModel {
    <#
    .SYNOPSIS
        Alle Daten des Reiters "Dashboard" (ohne Oberfläche, testbar).
    .PARAMETER Items
        Messwerte { Date 'yyyy-MM-dd'; Name; Value; Unit } (wie $script:allHistoricalItems).
    .PARAMETER CalcProvider
        { param($Name) Get-CalculatedValuesForMarker -markerName $Name } - berechnete Marker.
    .PARAMETER ScoreMode / Experimental
        Prototyp-Schalter (F1-Vergleich, E6). Haupt-Script: Weighted, ohne Experimental.
    #>
    param($Config, $Items, [datetime]$Today = (Get-Date).Date, [scriptblock]$CalcProvider = $null,
          [ValidateSet('Weighted', 'Simple')][string]$ScoreMode = 'Weighted', [switch]$Experimental)
    $index = New-HdItemIndex -Items $Items
    $main = Get-HealthDashScore -Config $Config -Index $index -AsOf $Today -Mode $ScoreMode -Experimental:$Experimental
    $comparison = Get-HealthDashComparison -Config $Config -Index $index -Main $main -Today $Today -Mode $ScoreMode -Experimental:$Experimental
    $longevityEntries = @()
    if ($CalcProvider -and (Get-HdMarkerConfig -Config $Config -Name 'Longevity-Score')) { $longevityEntries = @(& $CalcProvider 'Longevity-Score') }
    $dates = @(Get-HdTestDates -Index $index -AsOf $Today)
    return [PSCustomObject]@{
        Today = $Today.Date; LastTestDate = $(if ($dates.Count -gt 0) { [datetime]$dates[0] } else { $null }); HasData = ($index.Count -gt 0)
        Score = $main; Comparison = $comparison; Longevity = (Get-HealthDashLongevity -Entries $longevityEntries -AsOf $Today)
        Tiles = @(Get-HealthDashTiles -Config $Config -Index $index -AsOf $Today -CalcProvider $CalcProvider)
        Predispositions = (Get-HealthDashPredispositions -Config $Config -Index $index -AsOf $Today -CalcProvider $CalcProvider)
        PredispositionTiles = (Get-HealthDashPredispositionTiles -Config $Config -Index $index -AsOf $Today -CalcProvider $CalcProvider)
        Abnormalities = (Get-HealthDashAbnormalities -Config $Config -Index $index -AsOf $Today -CalcProvider $CalcProvider)
        CustomViews = (Get-HdCustomViews -Config $Config -Index $index -AsOf $Today -CalcProvider $CalcProvider)
        Testosterone = (Get-HealthDashTestosterone -Config $Config -Index $index -AsOf $Today)
        ScoreMode = $ScoreMode; Experimental = [bool]$Experimental
    }
}

function ConvertTo-HdCanonicalText {
    <#
        v0.1.1: Eindeutige Textform eines Werts für die Prüfsumme - Schlüssel/Eigenschaften SORTIERT
        (die Reihenfolge einer Hashtable ändert sich beim Zuweisen, ConvertTo-Json wäre nicht stabil).
    #>
    param($Value, [int]$Depth = 0)
    if ($null -eq $Value) { return '~' }
    if ($Depth -gt 8) { return '...' }
    if ($Value -is [string]) { return '"' + $Value + '"' }
    if ($Value -is [System.Collections.IDictionary]) {
        $parts = foreach ($key in @($Value.Keys | ForEach-Object { [string]$_ } | Sort-Object)) { $key + '=' + (ConvertTo-HdCanonicalText -Value $Value[$key] -Depth ($Depth + 1)) }
        return '{' + (@($parts) -join ';') + '}'
    }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $parts = foreach ($property in @($Value.PSObject.Properties | Sort-Object -Property Name)) { $property.Name + '=' + (ConvertTo-HdCanonicalText -Value $property.Value -Depth ($Depth + 1)) }
        return '{' + (@($parts) -join ';') + '}'
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $parts = foreach ($element in $Value) { ConvertTo-HdCanonicalText -Value $element -Depth ($Depth + 1) }
        return '[' + (@($parts) -join ',') + ']'
    }
    return [string]$Value
}

function Get-HealthDashDataSignature {
    <#
    .SYNOPSIS
        v0.1.1: Prüfsumme (SHA256) ALLER Eingaben des Dashboard-Modells: Messwerte (Datum, Name, Wert, Einheit),
        Marker-Konfiguration (Name, Einheit, Bereiche, Gruppe; berechnete Marker zusätzlich Formel und
        Pflichtmarker), persönliche Metriken, Vorbelastungen, eigene Ansichten (v2.39.0) und Stichtag.
        Gleiche Prüfsumme = Modell unverändert -> Update-HealthDashboardTab rechnet nicht neu.
        Beschreibungstexte der Marker zählen bewusst nicht (ohne Einfluss auf das Modell).
    #>
    param($Config, $Items, [datetime]$Today = (Get-Date).Date)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('T=').Append($Today.Date.ToString('yyyy-MM-dd')).Append("`n")
    foreach ($item in @($Items)) {
        if ($null -eq $item) { continue }
        [void]$sb.Append([string]$item.Date).Append('|').Append([string]$item.Name).Append('|').Append([string]$item.Value).Append('|').Append([string]$item.Unit).Append("`n")
    }
    foreach ($listName in @('Markers', 'CalculatedMarkers')) {
        [void]$sb.Append('#').Append($listName).Append("`n")
        foreach ($marker in @(Get-HdField $Config $listName)) {
            if ($null -eq $marker) { continue }
            $line = @([string]$marker.Name, [string]$marker.Unit, [string]$marker.RefMin, [string]$marker.RefMax, [string]$marker.OptimalMin, [string]$marker.OptimalMax,
                      [string]$marker.Group, [string]$marker.Formula, (@($marker.RequiredMarkers) -join ',')) -join '|'
            [void]$sb.Append($line).Append("`n")
        }
    }
    foreach ($section in @('Personal', 'GeneticPredispositions', 'DashboardViews')) {
        $value = Get-HdField $Config $section
        [void]$sb.Append('#').Append($section).Append('=').Append((ConvertTo-HdCanonicalText -Value $value)).Append("`n")
    }
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return ([System.BitConverter]::ToString($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($sb.ToString())))).Replace('-', '') }
    finally { $sha.Dispose() }
}

function Test-HdModelCurrent {
    <# v0.1.1: $true = vorhandenes Modell passt zur Prüfsumme (keine Neuberechnung). Ohne Prüfsumme oder mit -Force: $false. #>
    param($Model, [string]$ModelSignature, [string]$Signature, [switch]$Force)
    if ($Force -or $null -eq $Model -or [string]::IsNullOrEmpty($Signature) -or [string]::IsNullOrEmpty($ModelSignature)) { return $false }
    return ($Signature -eq $ModelSignature)
}

function Get-HealthDashExplainText {
    <# Text des Dialogs "Wie berechnet?" (Spec 4.6). #>
    param($Model)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('GESUNDHEITSSCORE')
    [void]$sb.AppendLine($script:HealthDashMeaningText)
    [void]$sb.AppendLine('')
    if ($Model.ScoreMode -eq 'Weighted') { [void]$sb.AppendLine($script:HealthDashWeightingText) }
    else { [void]$sb.AppendLine('Einfacher Mittelwert aller Marker (nur Prototyp-Vergleich).') }
    [void]$sb.AppendLine("Berücksichtigt werden gemessene Marker mit Bewertungsbereich und einem Wert der letzten $($script:HealthDashWindowDays) Tage (mindestens $($script:HealthDashMinMarkers) Marker aus $($script:HealthDashMinGroups) Bereichen). Nicht berücksichtigt: $($script:HealthDashScoreExcludeGroups -join ', ') und berechnete Marker.")
    if ($Model.Experimental) { [void]$sb.AppendLine('EXPERIMENTELL: Werte im Referenzbereich zählen mit 100 Punkten (nur Prototyp).') }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('FUNKTIONSBEREICHE (schlechtester zuerst)')
    if (@($Model.Score.Groups).Count -eq 0) { [void]$sb.AppendLine('  –') }
    foreach ($group in $Model.Score.Groups) { [void]$sb.AppendLine(('  {0}: {1} Punkte · {2} Marker' -f $group.Name, (Format-HdNumber $group.Score -MaxDecimals 1), $group.Count)) }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('VERGLEICH')
    $mainCount = $Model.Score.MarkerCount
    if ($Model.Comparison.Available) {
        [void]$sb.AppendLine("  Heute: $($Model.Score.Score) Punkte aus $mainCount Markern")
        [void]$sb.AppendLine("  $(Format-HdDate $Model.Comparison.PrevDate): $($Model.Comparison.PrevScore) Punkte aus $($Model.Comparison.PrevMarkerCount) Markern")
        if ($Model.Comparison.DataChanged) { [void]$sb.AppendLine('  Datengrundlage verändert (unterschiedliche Marker).') }
    } else { [void]$sb.AppendLine("  $($Model.Comparison.Text)") }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('NICHT BEWERTET')
    if (@($Model.Score.NotRated).Count -eq 0) { [void]$sb.AppendLine('  –') }
    foreach ($item in $Model.Score.NotRated) { [void]$sb.AppendLine("  $($item.Name): $($item.Reason)") }
    if ($Model.Predispositions.State -eq 'Active') {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('VORBELASTUNGEN')
        foreach ($entry in $Model.Predispositions.Entries) {
            [void]$sb.AppendLine("  $($entry.Name): $($entry.K) von $($entry.M) Markern bewertet")
            foreach ($item in $entry.NotRated) { [void]$sb.AppendLine("    - $($item.Name): $($item.Reason)") }
            if (@($entry.CalculatedNames).Count -gt 0) { [void]$sb.AppendLine("    Berechnete Marker (enthalten ihre Eingangswerte erneut): $(@($entry.CalculatedNames) -join ', ')") }
        }
    }
    return $sb.ToString()
}

# ---------------------------------------------------------------------------------
# Oberfläche (WinForms + GDI+) – Spec Teil A (Abschnitte 3–7, 9)
# Zustand im $script:-Scope (Event-Handler laufen in eigenem Scope, PS-5.1-Falle 9).
# Text: TextRenderer; Formen (Ring, Balken, Punkte, Pfeile): GDI+ mit Kantenglättung.
# Zeichenfunktionen erhalten Graphics + Modell als Parameter (ohne GUI testbar).
# ---------------------------------------------------------------------------------
$script:HealthDashThemeDefaults = @{
    Page = '#F4F4F2'; Card = '#FFFFFF'; Border = '#E2E1DC'; TextPrimary = '#0B0B0B'; TextSecondary = '#52514E'; TextMuted = '#6B6A65'
    Link = '#1C5CAB'; Accent = '#2A78D6'; StatusGood = '#0CA30C'; StatusWarning = '#FAB219'; StatusCritical = '#D03B3B'; StatusNone = '#9A9993'
    TrackGood = '#D6F0D6'; TrackWarning = '#FDEDC6'; TrackCritical = '#F6D6D6'; TrackNone = '#EBEAE6'; RangeRef = '#DCDBD6'; RangeOptimal = '#B7DDB7'
    FontFamily = 'Segoe UI'
}
$script:hdShowSparklines = $true

function Get-HdColor {
    <# Farbe aus $script:HealthUiTheme (gemeinsames Theme), sonst Dashboard-Standard. #>
    param([string]$Token)
    $hex = $null
    if ($script:HealthUiTheme -and $script:HealthUiTheme.ContainsKey($Token)) { $hex = [string]$script:HealthUiTheme[$Token] }
    if ([string]::IsNullOrWhiteSpace($hex)) { $hex = [string]$script:HealthDashThemeDefaults[$Token] }
    if ([string]::IsNullOrWhiteSpace($hex)) { $hex = '#000000' }
    return [System.Drawing.ColorTranslator]::FromHtml($hex)
}

function Get-HdLevelToken {
    <# Status-Ebene -> Farb-Token (Punkt/Füllung) bzw. Spur. #>
    param([string]$Level, [switch]$Track)
    $suffix = switch ($Level) { 'Good' { 'Good' } 'Warning' { 'Warning' } 'Critical' { 'Critical' } default { 'None' } }
    if ($Track) { return "Track$suffix" }
    return "Status$suffix"
}

function New-HdFonts {
    <# Schriften (Spec 9): Titel 16 · Kartentitel 11 · Kachelwert 20 · Hauptzahl 36 · Standard 9; "halbfett" = Segoe UI Semibold, sonst Fett. #>
    $family = [string]$script:HealthDashThemeDefaults['FontFamily']
    if ($script:HealthUiTheme -and $script:HealthUiTheme['FontFamily']) { $family = [string]$script:HealthUiTheme['FontFamily'] }
    $semi = $null
    $probe = New-Object System.Drawing.Font("$family Semibold", 9)
    if ($probe.Name -eq "$family Semibold") { $semi = "$family Semibold" }
    $probe.Dispose()
    $newSemi = { param([float]$Size) if ($semi) { New-Object System.Drawing.Font($semi, $Size) } else { New-Object System.Drawing.Font($family, $Size, [System.Drawing.FontStyle]::Bold) } }
    return @{
        Title = (& $newSemi 16); CardTitle = (& $newSemi 11); TileValue = (& $newSemi 20); Main = (& $newSemi 36)
        Normal = (New-Object System.Drawing.Font($family, 9)); NormalSemi = (& $newSemi 9); Score11 = (& $newSemi 11)
    }
}

function Get-HdTextFlags {
    param([ValidateSet('Left', 'Center', 'Right')][string]$Align = 'Left', [ValidateSet('Top', 'Middle', 'Bottom')][string]$VAlign = 'Middle')
    $flags = [System.Windows.Forms.TextFormatFlags]::NoPrefix -bor [System.Windows.Forms.TextFormatFlags]::SingleLine -bor [System.Windows.Forms.TextFormatFlags]::EndEllipsis -bor [System.Windows.Forms.TextFormatFlags]::NoPadding
    if ($Align -eq 'Center') { $flags = $flags -bor [System.Windows.Forms.TextFormatFlags]::HorizontalCenter }
    if ($Align -eq 'Right') { $flags = $flags -bor [System.Windows.Forms.TextFormatFlags]::Right }
    if ($VAlign -eq 'Middle') { $flags = $flags -bor [System.Windows.Forms.TextFormatFlags]::VerticalCenter }
    if ($VAlign -eq 'Bottom') { $flags = $flags -bor [System.Windows.Forms.TextFormatFlags]::Bottom }
    return $flags
}

function Invoke-HdDrawText {
    param($Graphics, [string]$Text, $Font, [string]$ColorToken, [int]$X, [int]$Y, [int]$Width, [int]$Height, [string]$Align = 'Left', [string]$VAlign = 'Middle')
    if ([string]::IsNullOrEmpty($Text) -or $Width -le 0) { return }
    $rect = New-Object System.Drawing.Rectangle($X, $Y, $Width, $Height)
    [System.Windows.Forms.TextRenderer]::DrawText($Graphics, $Text, $Font, $rect, (Get-HdColor $ColorToken), (Get-HdTextFlags -Align $Align -VAlign $VAlign))
}

function Get-HdTextWidth {
    param([string]$Text, $Font)
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $size = New-Object System.Drawing.Size(4000, 200)
    return [System.Windows.Forms.TextRenderer]::MeasureText($Text, $Font, $size, (Get-HdTextFlags)).Width
}

function Invoke-HdFillCircle {
    param($Graphics, [string]$ColorToken, [float]$CenterX, [float]$CenterY, [float]$Diameter)
    $brush = New-Object System.Drawing.SolidBrush((Get-HdColor $ColorToken))
    try { $Graphics.FillEllipse($brush, [float]($CenterX - $Diameter / 2), [float]($CenterY - $Diameter / 2), $Diameter, $Diameter) } finally { $brush.Dispose() }
}

function Invoke-HdFillRect {
    param($Graphics, [string]$ColorToken, [int]$X, [int]$Y, [int]$Width, [int]$Height)
    if ($Width -le 0 -or $Height -le 0) { return }
    $brush = New-Object System.Drawing.SolidBrush((Get-HdColor $ColorToken))
    try { $Graphics.FillRectangle($brush, $X, $Y, $Width, $Height) } finally { $brush.Dispose() }
}

function Invoke-HdDrawStatusMark {
    <#
        StatusMark (Spec 5): gezeichneter Punkt Ø 8 + 6 px Abstand + Wort in TextPrimary (Status nie nur über Farbe).
        Align Center: Punkt + Text gemeinsam zentriert in [X, X+Width]. Rückgabe: belegte Breite.
    #>
    param($Graphics, [string]$Level, [string]$Text, $Font, [int]$X, [int]$Y, [int]$Width, [int]$Height, [string]$Align = 'Left', [string]$TextToken = 'TextPrimary')
    if ([string]::IsNullOrEmpty($Text)) { return 0 }
    $textWidth = [Math]::Min((Get-HdTextWidth -Text $Text -Font $Font), [Math]::Max($Width - 14, 0))
    $total = 14 + $textWidth
    $startX = $X
    if ($Align -eq 'Center') { $startX = $X + [int](($Width - $total) / 2) }
    if ($Align -eq 'Right') { $startX = $X + $Width - $total }
    Invoke-HdFillCircle -Graphics $Graphics -ColorToken (Get-HdLevelToken $Level) -CenterX ($startX + 4) -CenterY ($Y + $Height / 2) -Diameter 8
    Invoke-HdDrawText -Graphics $Graphics -Text $Text -Font $Font -ColorToken $TextToken -X ($startX + 14) -Y $Y -Width ($textWidth + 2) -Height $Height
    return $total
}

function Invoke-HdDrawArrow {
    <# TrendArrow (Spec 5): gezeichnetes Dreieck 8 x 8 (Up/Down/Flat=rechts) in TextSecondary. #>
    param($Graphics, [string]$Direction, [int]$X, [int]$Y)
    if ([string]::IsNullOrEmpty($Direction)) { return }
    $points = switch ($Direction) {
        'Up'   { @((New-Object System.Drawing.Point($X, ($Y + 7))), (New-Object System.Drawing.Point(($X + 8), ($Y + 7))), (New-Object System.Drawing.Point(($X + 4), ($Y + 1)))) }
        'Down' { @((New-Object System.Drawing.Point($X, ($Y + 1))), (New-Object System.Drawing.Point(($X + 8), ($Y + 1))), (New-Object System.Drawing.Point(($X + 4), ($Y + 7)))) }
        default { @((New-Object System.Drawing.Point(($X + 1), $Y)), (New-Object System.Drawing.Point(($X + 1), ($Y + 8))), (New-Object System.Drawing.Point(($X + 7), ($Y + 4)))) }
    }
    $brush = New-Object System.Drawing.SolidBrush((Get-HdColor 'TextSecondary'))
    try { $Graphics.FillPolygon($brush, [System.Drawing.Point[]]$points) } finally { $brush.Dispose() }
}

function Invoke-HdDrawCardFrame {
    <# Karte (Spec 5): weiß, 1-px-Rahmen, optional Titel 11 pt halbfett bei (16, 16). #>
    param($Graphics, [int]$Width, [int]$Height, [string]$Title = '', $Fonts)
    Invoke-HdFillRect -Graphics $Graphics -ColorToken 'Card' -X 0 -Y 0 -Width $Width -Height $Height
    $pen = New-Object System.Drawing.Pen((Get-HdColor 'Border'), 1)
    try { $Graphics.DrawRectangle($pen, 0, 0, ($Width - 1), ($Height - 1)) } finally { $pen.Dispose() }
    if ($Title) { Invoke-HdDrawText -Graphics $Graphics -Text $Title -Font $Fonts.CardTitle -ColorToken 'TextPrimary' -X 16 -Y 16 -Width ($Width - 32) -Height 20 }
}

function Invoke-HdDrawMeter {
    <# MiniMeter (Spec 5): Spur in heller Statusstufe, Füllung in Statusfarbe (Score / 100); ohne Score nur graue Spur. #>
    param($Graphics, [int]$X, [int]$Y, [int]$Width, [int]$Height, $Score, [string]$Level)
    if ($null -eq $Score) { Invoke-HdFillRect -Graphics $Graphics -ColorToken 'TrackNone' -X $X -Y $Y -Width $Width -Height $Height; return }
    Invoke-HdFillRect -Graphics $Graphics -ColorToken (Get-HdLevelToken $Level -Track) -X $X -Y $Y -Width $Width -Height $Height
    $fill = [int][Math]::Round($Width * [Math]::Max(0, [Math]::Min(100, [double]$Score)) / 100)
    Invoke-HdFillRect -Graphics $Graphics -ColorToken (Get-HdLevelToken $Level) -X $X -Y $Y -Width $fill -Height $Height
}

function Get-HdRangeScale {
    <#
        Darstellungsbereich des Bereichsbalkens (Spec 4.5): min(RefMin, OptMin) .. max(RefMax, OptMax),
        je 25 % der Spanne erweitert; Untergrenze nie unter 0, wenn alle Bereichswerte >= 0. $null ohne Bereiche.
    #>
    param($Range)
    if ($null -eq $Range -or -not $Range.Rateable) { return $null }
    $values = @(@($Range.RefMin, $Range.RefMax, $Range.OptMin, $Range.OptMax) | Where-Object { $null -ne $_ })
    if ($values.Count -eq 0) { return $null }
    $low = ($values | Measure-Object -Minimum).Minimum; $high = ($values | Measure-Object -Maximum).Maximum
    $span = $high - $low
    if ($span -le 0) { $span = [Math]::Max([Math]::Abs($high), 1.0) }
    $displayLow = $low - 0.25 * $span; $displayHigh = $high + 0.25 * $span
    if (@($values | Where-Object { $_ -lt 0 }).Count -eq 0) { $displayLow = [Math]::Max(0.0, $displayLow) }
    return [PSCustomObject]@{ Low = [double]$displayLow; High = [double]$displayHigh }
}

function Invoke-HdDrawRangeBar {
    <# RangeBar (Spec 4.5): Track · Referenzbereich · Optimalbereich · eigener Wert (2-px-Strich + Punkt, außerhalb: Randmarke mit Pfeil). #>
    param($Graphics, [int]$X, [int]$Y, [int]$Width, [int]$Height, $Range, $Value)
    $scale = Get-HdRangeScale -Range $Range
    if ($null -eq $scale) { return $false }
    $toX = { param([double]$v) $X + ($v - $scale.Low) / ($scale.High - $scale.Low) * $Width }
    Invoke-HdFillRect -Graphics $Graphics -ColorToken 'TrackNone' -X $X -Y $Y -Width $Width -Height $Height
    $refLow = if ($null -ne $Range.RefMin) { $Range.RefMin } else { $Range.OptMin }
    $refHigh = if ($null -ne $Range.RefMax) { $Range.RefMax } else { $Range.OptMax }
    if ($null -ne $refLow -and $null -ne $refHigh) {
        $x1 = [int](& $toX $refLow); $x2 = [int](& $toX $refHigh)
        Invoke-HdFillRect -Graphics $Graphics -ColorToken 'RangeRef' -X $x1 -Y $Y -Width ([Math]::Max(1, $x2 - $x1)) -Height $Height
    }
    if ($Range.HasOptimal) {
        $o1 = [int](& $toX $Range.OptMin); $o2 = [int](& $toX $Range.OptMax)
        Invoke-HdFillRect -Graphics $Graphics -ColorToken 'RangeOptimal' -X $o1 -Y $Y -Width ([Math]::Max(1, $o2 - $o1)) -Height $Height
    }
    if ($null -ne $Value) {
        $raw = & $toX ([double]$Value)
        $outside = ''
        if ($raw -lt $X) { $outside = 'Left' } elseif ($raw -gt $X + $Width) { $outside = 'Right' }
        $markX = [int][Math]::Max($X, [Math]::Min($X + $Width, $raw))
        $pen = New-Object System.Drawing.Pen((Get-HdColor 'TextPrimary'), 2)
        try { $Graphics.DrawLine($pen, $markX, ($Y - 2), $markX, ($Y + $Height + 2)) } finally { $pen.Dispose() }
        Invoke-HdFillCircle -Graphics $Graphics -ColorToken 'TextPrimary' -CenterX $markX -CenterY ($Y + $Height / 2) -Diameter 6
        if ($outside) {
            $tip = if ($outside -eq 'Left') { $X - 7 } else { $X + $Width + 7 }
            $base = if ($outside -eq 'Left') { $X - 2 } else { $X + $Width + 2 }
            $points = @((New-Object System.Drawing.Point($base, ($Y))), (New-Object System.Drawing.Point($base, ($Y + $Height))), (New-Object System.Drawing.Point($tip, ($Y + [int]($Height / 2)))))
            $brush = New-Object System.Drawing.SolidBrush((Get-HdColor 'TextPrimary'))
            try { $Graphics.FillPolygon($brush, [System.Drawing.Point[]]$points) } finally { $brush.Dispose() }
        }
    }
    return $true
}

function Invoke-HdDrawSparkline {
    <# Verlauf (Spec 4.3): bis zu 12 Werte, Linie 2 px TextMuted, letzter Punkt Accent; < 2 Werte: nichts. #>
    param($Graphics, [int]$X, [int]$Y, [int]$Width, [int]$Height, $History)
    $points = @($History)
    if ($points.Count -lt 2) { return }
    $values = @($points | ForEach-Object { [double]$_.Value })
    $min = ($values | Measure-Object -Minimum).Minimum; $max = ($values | Measure-Object -Maximum).Maximum
    $span = $max - $min; if ($span -le 0) { $span = 1.0 }
    $coords = @()
    for ($i = 0; $i -lt $values.Count; $i++) {
        $px = $X + [int][Math]::Round($i * ($Width - 4) / ($values.Count - 1)) + 2
        $py = $Y + 3 + [int][Math]::Round(($Height - 6) * (1 - ($values[$i] - $min) / $span))
        $coords += , (New-Object System.Drawing.Point($px, $py))
    }
    $pen = New-Object System.Drawing.Pen((Get-HdColor 'TextMuted'), 2)
    try { $Graphics.DrawLines($pen, [System.Drawing.Point[]]$coords) } finally { $pen.Dispose() }
    Invoke-HdFillCircle -Graphics $Graphics -ColorToken 'Accent' -CenterX $coords[-1].X -CenterY $coords[-1].Y -Diameter 7
}

function Format-HdValueText {
    <# Anzeigewert einer Kachel/Zeile (Nachkommastellen nach Größe). #>
    param($Value)
    if ($null -eq $Value) { return '–' }
    $abs = [Math]::Abs([double]$Value)
    $decimals = if ($abs -ge 100) { 0 } elseif ($abs -ge 10) { 1 } else { 2 }
    return (Format-HdNumber $Value -MaxDecimals $decimals)
}

# ---------------------------------------------------------------------------------
# Zeichnen der Bereiche (Koordinaten nach Spec Abschnitt 4)
# ---------------------------------------------------------------------------------
function Invoke-HdPaintScoreCard {
    <# Score-Karte 360 x 280 (Spec 4.2, E1). #>
    param($Graphics, $Model, $Fonts, [int]$Width = 360, [int]$Height = 280)
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    Invoke-HdDrawCardFrame -Graphics $Graphics -Width $Width -Height $Height -Fonts $Fonts
    $score = $Model.Score
    if ($Model.Experimental) {
        $warn = @((New-Object System.Drawing.Point(8, 18)), (New-Object System.Drawing.Point(20, 18)), (New-Object System.Drawing.Point(14, 7)))
        $brush = New-Object System.Drawing.SolidBrush((Get-HdColor 'StatusWarning'))
        try { $Graphics.FillPolygon($brush, [System.Drawing.Point[]]$warn) } finally { $brush.Dispose() }
        Invoke-HdDrawText -Graphics $Graphics -Text 'EXPERIMENTELL' -Font $Fonts.NormalSemi -ColorToken 'TextSecondary' -X 24 -Y 4 -Width 160 -Height 16
    }
    # Ring Ø 160, Strich 14, Start 135°, Bogen 270°, runde Enden
    $ringX = [int](($Width - 160) / 2); $ringY = 16
    $arcRect = New-Object System.Drawing.Rectangle(($ringX + 7), ($ringY + 7), 146, 146)
    $trackPen = New-Object System.Drawing.Pen((Get-HdColor (Get-HdLevelToken $score.Level -Track)), 14)
    $trackPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round; $trackPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    try { $Graphics.DrawArc($trackPen, $arcRect, 135, 270) } finally { $trackPen.Dispose() }
    if ($null -ne $score.Score -and $score.Score -gt 0) {
        $fillPen = New-Object System.Drawing.Pen((Get-HdColor (Get-HdLevelToken $score.Level)), 14)
        $fillPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round; $fillPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
        try { $Graphics.DrawArc($fillPen, $arcRect, 135, [float](270 * [Math]::Min(100, $score.Score) / 100)) } finally { $fillPen.Dispose() }
    }
    $mainText = if ($null -ne $score.Score) { [string]$score.Score } else { '–' }
    Invoke-HdDrawText -Graphics $Graphics -Text $mainText -Font $Fonts.Main -ColorToken 'TextPrimary' -X $ringX -Y 58 -Width 160 -Height 64 -Align Center
    Invoke-HdDrawText -Graphics $Graphics -Text 'Gesundheitsscore' -Font $Fonts.Normal -ColorToken 'TextMuted' -X ($ringX + 20) -Y 122 -Width 120 -Height 16 -Align Center
    if ($null -ne $score.Score) { $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level $score.Level -Text $score.Status -Font $Fonts.Normal -X 16 -Y 184 -Width ($Width - 32) -Height 18 -Align Center }
    else { $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level 'None' -Text $score.Reason -Font $Fonts.Normal -X 16 -Y 184 -Width ($Width - 32) -Height 18 -Align Center -TextToken 'TextSecondary' }
    # Vergleich (Spec 12.4): Pfeil + Zahl ohne Wertung
    $comparison = $Model.Comparison
    if ($null -ne $score.Score) {
        if ($comparison.Available) {
            $textWidth = Get-HdTextWidth -Text $comparison.Text -Font $Fonts.Normal
            $total = [Math]::Min($Width - 32, 14 + $textWidth)
            $startX = [int](($Width - $total) / 2)
            Invoke-HdDrawArrow -Graphics $Graphics -Direction $comparison.Direction -X $startX -Y 208
            Invoke-HdDrawText -Graphics $Graphics -Text $comparison.Text -Font $Fonts.Normal -ColorToken 'TextSecondary' -X ($startX + 14) -Y 204 -Width ($total - 14) -Height 16
        } else {
            Invoke-HdDrawText -Graphics $Graphics -Text $comparison.Text -Font $Fonts.Normal -ColorToken 'TextSecondary' -X 16 -Y 204 -Width ($Width - 32) -Height 16 -Align Center
        }
    }
    $pen = New-Object System.Drawing.Pen((Get-HdColor 'Border'), 1)
    try { $Graphics.DrawLine($pen, 16, 228, ($Width - 16), 228) } finally { $pen.Dispose() }
    $longevity = $Model.Longevity
    $lvText = 'Longevity-Score  ' + $(if ($longevity.Available) { [string]$longevity.Score } else { '–' })
    Invoke-HdDrawText -Graphics $Graphics -Text $lvText -Font $Fonts.Score11 -ColorToken 'TextPrimary' -X 16 -Y 234 -Width 200 -Height 18
    if ($longevity.Available) { $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level $longevity.Level -Text $longevity.Status -Font $Fonts.Normal -X 200 -Y 234 -Width ($Width - 216) -Height 18 -Align Right }
    Invoke-HdDrawText -Graphics $Graphics -Text $longevity.Text -Font $Fonts.Normal -ColorToken 'TextMuted' -X 16 -Y 254 -Width ($Width - 32) -Height 16
}

function Get-HdTileStatusText {
    <# Kurzer Statustext für die schmale Kachel (Datum steht bereits in der Fußzeile). #>
    param($Tile)
    switch ($Tile.State) {
        'Invalid'      { return 'Ungültiger Wert' }
        'Conflict'     { return 'Konflikt: mehrere Werte' }
        'RangeInvalid' { return 'Bereichskonfiguration ungültig' }
        default        { return [string]$Tile.Status }
    }
}

function Invoke-HdPaintTile {
    <# Longevity-Kachel 237 x 132 (Spec 4.3, E2/E3): Name · Wert + Einheit · Status · Verlauf · Datum + Trend. #>
    param($Graphics, $Tile, $Fonts, [int]$Width = 237, [int]$Height = 132, [bool]$Focused = $false, [bool]$Hover = $false, [bool]$ShowSparkline = $true)
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    Invoke-HdFillRect -Graphics $Graphics -ColorToken 'Card' -X 0 -Y 0 -Width $Width -Height $Height
    $borderToken = if ($Hover) { 'TextMuted' } else { 'Border' }
    $pen = New-Object System.Drawing.Pen((Get-HdColor $borderToken), 1)
    try { $Graphics.DrawRectangle($pen, 0, 0, ($Width - 1), ($Height - 1)) } finally { $pen.Dispose() }
    if ($Focused) {
        $focusPen = New-Object System.Drawing.Pen((Get-HdColor 'Link'), 2)
        try { $Graphics.DrawRectangle($focusPen, 2, 2, ($Width - 5), ($Height - 5)) } finally { $focusPen.Dispose() }
    }
    $inner = $Width - 32
    Invoke-HdDrawText -Graphics $Graphics -Text $Tile.Title -Font $Fonts.NormalSemi -ColorToken 'TextSecondary' -X 16 -Y 12 -Width $inner -Height 16
    $hasValue = ($null -ne $Tile.Value -and $Tile.State -notin @('Invalid', 'Conflict', 'NoData', 'NoConfig'))
    $valueText = if ($hasValue) { Format-HdValueText $Tile.Value } else { '–' }
    $valueWidth = [Math]::Min($inner, (Get-HdTextWidth -Text $valueText -Font $Fonts.TileValue))
    Invoke-HdDrawText -Graphics $Graphics -Text $valueText -Font $Fonts.TileValue -ColorToken 'TextPrimary' -X 16 -Y 30 -Width ($valueWidth + 2) -Height 28 -VAlign Bottom
    if ($hasValue -and $Tile.Unit) { Invoke-HdDrawText -Graphics $Graphics -Text $Tile.Unit -Font $Fonts.Normal -ColorToken 'TextSecondary' -X (16 + $valueWidth + 4) -Y 30 -Width ($inner - $valueWidth - 4) -Height 24 -VAlign Bottom }
    $statusToken = if ($Tile.Level -eq 'None') { 'TextSecondary' } else { 'TextPrimary' }
    $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level $Tile.Level -Text (Get-HdTileStatusText -Tile $Tile) -Font $Fonts.Normal -X 16 -Y 60 -Width $inner -Height 16 -TextToken $statusToken
    if ($ShowSparkline) { Invoke-HdDrawSparkline -Graphics $Graphics -X 16 -Y 80 -Width $inner -Height 20 -History $Tile.History }
    Invoke-HdDrawText -Graphics $Graphics -Text $Tile.DateText -Font $Fonts.Normal -ColorToken 'TextMuted' -X 16 -Y 104 -Width 120 -Height 16
    if ($Tile.Trend -and $Tile.Trend.Direction) {
        $trendWidth = Get-HdTextWidth -Text $Tile.Trend.Text -Font $Fonts.Normal
        $trendX = 16 + $inner - $trendWidth
        Invoke-HdDrawArrow -Graphics $Graphics -Direction $Tile.Trend.Direction -X ($trendX - 12) -Y 108
        Invoke-HdDrawText -Graphics $Graphics -Text $Tile.Trend.Text -Font $Fonts.Normal -ColorToken 'TextSecondary' -X $trendX -Y 104 -Width ($trendWidth + 2) -Height 16
    }
}

function Get-HdPredispositionHeight {
    <# Inhaltshöhe der Vorbelastungen (Spec 4.4): je Block 108 px + 12 px Abstand. #>
    param($Predispositions)
    if ($Predispositions.State -ne 'Active') { return 0 }
    $count = @($Predispositions.Entries).Count
    if ($count -eq 0) { return 0 }
    return ($count * 108 + ($count - 1) * 12)
}

function Invoke-HdPaintPredispositions {
    <# Blöcke je aktiver Vorbelastung (Spec 4.4, E4): Kopf · Meter · Status + "k von m" · bis zu 3 Fokusmarker. #>
    param($Graphics, $Predispositions, $Fonts, [int]$Width = 528)
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    if ($Predispositions.State -ne 'Active') { return }
    $y = 0
    foreach ($entry in @($Predispositions.Entries)) {
        $scoreText = if ($null -ne $entry.Score) { [string]$entry.Score } else { '–' }
        Invoke-HdDrawText -Graphics $Graphics -Text $entry.Name -Font $Fonts.NormalSemi -ColorToken 'TextPrimary' -X 0 -Y $y -Width ([Math]::Min(440, $Width - 60)) -Height 18
        Invoke-HdDrawText -Graphics $Graphics -Text $scoreText -Font $Fonts.Score11 -ColorToken 'TextPrimary' -X ($Width - 60) -Y $y -Width 60 -Height 18 -Align Right
        Invoke-HdDrawMeter -Graphics $Graphics -X 0 -Y ($y + 22) -Width $Width -Height 8 -Score $entry.Score -Level $entry.Level
        $statusToken = if ($null -eq $entry.Score) { 'TextSecondary' } else { 'TextPrimary' }
        $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level $entry.Level -Text $entry.Status -Font $Fonts.Normal -X 0 -Y ($y + 34) -Width ($Width - 200) -Height 16 -TextToken $statusToken
        Invoke-HdDrawText -Graphics $Graphics -Text "$($entry.K) von $($entry.M) Markern ($($script:HealthDashWindowDays) Tage)" -Font $Fonts.Normal -ColorToken 'TextMuted' -X ($Width - 200) -Y ($y + 34) -Width 200 -Height 16 -Align Right
        $rowY = $y + 54
        foreach ($focus in @($entry.Focus)) {
            Invoke-HdDrawText -Graphics $Graphics -Text $focus.Name -Font $Fonts.Normal -ColorToken 'TextPrimary' -X 0 -Y $rowY -Width 220 -Height 18
            Invoke-HdDrawText -Graphics $Graphics -Text ("$(Format-HdValueText $focus.Value) $($focus.Unit)").Trim() -Font $Fonts.Normal -ColorToken 'TextSecondary' -X 228 -Y $rowY -Width 120 -Height 18
            $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level $focus.Level -Text $focus.Status -Font $Fonts.Normal -X 356 -Y $rowY -Width ([Math]::Min(172, $Width - 356)) -Height 18
            $rowY += 18
        }
        $y += 120
    }
}

function Invoke-HdPaintTestosterone {
    <# Karte Testosteron 544 x 376 (Spec 4.5): 3 Messwerte mit Bereichsbalken, berechneter Abschnitt ohne Bewertung, Legende. #>
    param($Graphics, $Testosterone, $Fonts, [int]$Width = 544, [int]$Height = 376)
    $Graphics.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    Invoke-HdDrawCardFrame -Graphics $Graphics -Width $Width -Height $Height -Title 'Testosteron' -Fonts $Fonts
    $inner = $Width - 32
    $y = 44
    foreach ($row in @($Testosterone.Rows)) {
        Invoke-HdDrawText -Graphics $Graphics -Text $row.Label -Font $Fonts.NormalSemi -ColorToken 'TextPrimary' -X 16 -Y $y -Width 230 -Height 18
        $valueText = if ($null -ne $row.Value -and $row.State -notin @('Invalid', 'Conflict')) { ("$(Format-HdValueText $row.Value) $($row.Unit) $($row.ExtraText)").Trim() } else { '–' }
        Invoke-HdDrawText -Graphics $Graphics -Text $valueText -Font $Fonts.NormalSemi -ColorToken 'TextPrimary' -X 250 -Y $y -Width ($inner - 234) -Height 18 -Align Right
        $drawn = $false
        if ($row.State -notin @('Invalid', 'Conflict', 'NoData', 'RangeInvalid') -and $row.Range) { $drawn = Invoke-HdDrawRangeBar -Graphics $Graphics -X 16 -Y ($y + 22) -Width $inner -Height 10 -Range $row.Range -Value $row.Value }
        if (-not $drawn -and $row.State -eq 'NoData') { Invoke-HdFillRect -Graphics $Graphics -ColorToken 'TrackNone' -X 16 -Y ($y + 22) -Width $inner -Height 10 }
        $statusToken = if ($row.Level -eq 'None') { 'TextSecondary' } else { 'TextPrimary' }
        $null = Invoke-HdDrawStatusMark -Graphics $Graphics -Level $row.Level -Text $row.Status -Font $Fonts.Normal -X 16 -Y ($y + 36) -Width ($inner - 170) -Height 16 -TextToken $statusToken
        Invoke-HdDrawText -Graphics $Graphics -Text $row.DateText -Font $Fonts.Normal -ColorToken 'TextMuted' -X ($Width - 186) -Y ($y + 36) -Width 170 -Height 16 -Align Right
        $y += 62
    }
    $pen = New-Object System.Drawing.Pen((Get-HdColor 'Border'), 1)
    try { $Graphics.DrawLine($pen, 16, 236, ($Width - 16), 236) } finally { $pen.Dispose() }
    $computed = $Testosterone.Computed
    if ($computed.Available) {
        Invoke-HdDrawText -Graphics $Graphics -Text "Berechnet aus Messwerten vom $(Format-HdDate $computed.Date)" -Font $Fonts.Normal -ColorToken 'TextMuted' -X 16 -Y 244 -Width $inner -Height 16
        Invoke-HdDrawText -Graphics $Graphics -Text 'Freies Testosteron (berechnet)' -Font $Fonts.Normal -ColorToken 'TextPrimary' -X 16 -Y 264 -Width 300 -Height 18
        Invoke-HdDrawText -Graphics $Graphics -Text "$(Format-HdNumber $computed.FreePgMl -MaxDecimals 0) pg/ml ($(Format-HdNumber $computed.FreePercent -MaxDecimals 1) %)" -Font $Fonts.NormalSemi -ColorToken 'TextPrimary' -X 316 -Y 264 -Width ($inner - 300) -Height 18 -Align Right
        Invoke-HdDrawText -Graphics $Graphics -Text 'Freier Androgen-Index (FAI)' -Font $Fonts.Normal -ColorToken 'TextPrimary' -X 16 -Y 284 -Width 300 -Height 18
        Invoke-HdDrawText -Graphics $Graphics -Text (Format-HdNumber $computed.Fai -MaxDecimals 0) -Font $Fonts.NormalSemi -ColorToken 'TextPrimary' -X 316 -Y 284 -Width ($inner - 300) -Height 18 -Align Right
        if ($computed.AlbuminEstimated) { Invoke-HdDrawText -Graphics $Graphics -Text "Albumin geschätzt ($(Format-HdNumber $computed.AlbuminGdl -MaxDecimals 1) g/dl)" -Font $Fonts.Normal -ColorToken 'TextSecondary' -X 16 -Y 304 -Width $inner -Height 16 }
    } else {
        $reason = if ($computed.Reason -eq 'Albumin ungültig – keine Berechnung') { $computed.Reason } else { "Nicht berechenbar: $($computed.Reason)" }
        Invoke-HdDrawText -Graphics $Graphics -Text $reason -Font $Fonts.Normal -ColorToken 'TextSecondary' -X 16 -Y 244 -Width $inner -Height 16
        if ($computed.Date) { Invoke-HdDrawText -Graphics $Graphics -Text "Testtag $(Format-HdDate $computed.Date)" -Font $Fonts.Normal -ColorToken 'TextMuted' -X 16 -Y 264 -Width $inner -Height 16 }
    }
    # Legende
    $lx = 16
    Invoke-HdFillRect -Graphics $Graphics -ColorToken 'RangeRef' -X $lx -Y 344 -Width 14 -Height 8; $lx += 20
    Invoke-HdDrawText -Graphics $Graphics -Text 'Referenzbereich' -Font $Fonts.Normal -ColorToken 'TextSecondary' -X $lx -Y 340 -Width 110 -Height 16; $lx += (Get-HdTextWidth 'Referenzbereich' $Fonts.Normal) + 18
    Invoke-HdFillRect -Graphics $Graphics -ColorToken 'RangeOptimal' -X $lx -Y 344 -Width 14 -Height 8; $lx += 20
    Invoke-HdDrawText -Graphics $Graphics -Text 'Optimalbereich' -Font $Fonts.Normal -ColorToken 'TextSecondary' -X $lx -Y 340 -Width 110 -Height 16; $lx += (Get-HdTextWidth 'Optimalbereich' $Fonts.Normal) + 18
    $legendPen = New-Object System.Drawing.Pen((Get-HdColor 'TextPrimary'), 2)
    try { $Graphics.DrawLine($legendPen, ($lx + 4), 341, ($lx + 4), 355) } finally { $legendPen.Dispose() }
    $lx += 14
    Invoke-HdDrawText -Graphics $Graphics -Text 'eigener Wert' -Font $Fonts.Normal -ColorToken 'TextSecondary' -X $lx -Y 340 -Width 110 -Height 16
}

# ---------------------------------------------------------------------------------
# Aufbau des Reiters, Aktualisierung, Dialog
# ---------------------------------------------------------------------------------
function Set-HdDoubleBuffered {
    <# Flackerfreies Zeichnen (DoubleBuffered ist protected -> per Reflection). #>
    param($Control)
    try {
        $property = [System.Windows.Forms.Control].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'NonPublic,Instance')
        $property.SetValue($Control, $true, $null)
    } catch { Write-Verbose "DoubleBuffered nicht gesetzt: $($_.Exception.Message)" }
}

function Update-HdPredispositionView {
    <# v2.38.0: Gemeinsamer Kachelbereich fuer Vorbelastungen/Auffaelligkeiten; Cache je Modell UND Ansicht. #>
    param([string]$View = 'Vorbelastungen')
    if (-not $script:hdControls) { return }
    $panel = $script:hdControls.ViewPanel
    if ($script:hdModel -and $panel.Tag -and $panel.Tag.View -eq $View -and [object]::ReferenceEquals($panel.Tag.Model, $script:hdModel)) { return }
    $panel.SuspendLayout()
    try {
        $oldControls = @($panel.Controls | ForEach-Object { $_ })
        $panel.Controls.Clear()
        foreach ($oldControl in $oldControls) {
            if ($script:hdToolTip) { $script:hdToolTip.SetToolTip($oldControl, $null) }
            $oldControl.Dispose()
        }
        $panel.Tag = $null
        $viewModel = if ($script:hdModel) { $script:hdModel.PredispositionTiles } else { $null }
        if ($View -eq 'Auffaelligkeiten' -and $script:hdModel) {
            $abnormal = $script:hdModel.Abnormalities
            $viewModel = [PSCustomObject]@{ State = 'Active'; Groups = @([PSCustomObject]@{ Name = "Aktuelle Auffälligkeiten ($($abnormal.Count))"; Tiles = @($abnormal.Tiles) }) }
        }
        if ($View -like 'Custom:*' -and $script:hdModel) {
            $customView = @($script:hdModel.CustomViews.Views | Where-Object { ('Custom:' + $_.Id) -eq $View })
            $viewModel = if ($customView.Count -eq 1) { [PSCustomObject]@{ State = 'Active'; Groups = @($customView[0]) } } else { $null }
        }
        $y = 0; $tabIndex = 0
        if (-not $viewModel -or $viewModel.State -ne 'Active') {
            $message = New-Object System.Windows.Forms.Label
            $message.Font = $script:hdFonts.Normal; $message.ForeColor = Get-HdColor 'TextSecondary'
            $message.AutoSize = $true; $message.Location = New-Object System.Drawing.Point(0, 4)
            $message.Text = if (-not $viewModel) { 'Vorbelastungen konnten nicht angezeigt werden.' } elseif ($viewModel.State -eq 'Disabled') { 'Vorbelastungen sind ausgeschaltet.' } else { 'Keine Vorbelastung aktiv.' }
            $panel.Controls.Add($message)
            if ($viewModel) {
                $settingsLink = New-Object System.Windows.Forms.LinkLabel
                $settingsLink.Font = $script:hdFonts.Normal; $settingsLink.AutoSize = $true
                $settingsLink.Location = New-Object System.Drawing.Point(0, 32)
                $settingsLink.Text = if ($viewModel.State -eq 'Disabled') { 'In „Persönliche Metriken“ einschalten' } else { 'In „Persönliche Metriken“ festlegen' }
                $settingsLink.LinkColor = Get-HdColor 'Link'; $settingsLink.ActiveLinkColor = Get-HdColor 'Link'; $settingsLink.VisitedLinkColor = Get-HdColor 'Link'
                $settingsLink.TabIndex = 0
                $settingsLink.Add_LinkClicked({ if ($script:hdOnOpenSettings) { try { & $script:hdOnOpenSettings } catch { Write-Warning "Vorbelastungen: $($_.Exception.Message)" } } })
                $panel.Controls.Add($settingsLink)
            }
            $y = 64
        } else {
            foreach ($group in @($viewModel.Groups)) {
                $heading = New-Object System.Windows.Forms.Label
                $heading.Text = [string]$group.Name; $heading.Font = $script:hdFonts.CardTitle
                $heading.ForeColor = Get-HdColor 'TextPrimary'; $heading.AutoSize = $false
                $heading.UseMnemonic = $false
                $headingHeight = [Math]::Max(24, [System.Windows.Forms.TextRenderer]::MeasureText($heading.Text, $heading.Font, (New-Object System.Drawing.Size(1120, 10000)), [System.Windows.Forms.TextFormatFlags]::WordBreak).Height + 8)
                $heading.Location = New-Object System.Drawing.Point(0, $y); $heading.Size = New-Object System.Drawing.Size(1120, $headingHeight)
                $panel.Controls.Add($heading)
                $y += $headingHeight + 8
                $tiles = @($group.Tiles)
                if ($tiles.Count -eq 0) {
                    $empty = New-Object System.Windows.Forms.Label
                    $empty.Text = if ($View -like 'Custom:*') { 'Dieser Ansicht sind keine Marker zugeordnet. Über Bearbeiten können Sie Marker hinzufügen.' } else { 'Dieser Vorbelastung sind keine Marker zugeordnet.' }
                    $empty.Font = $script:hdFonts.Normal; $empty.ForeColor = Get-HdColor 'TextMuted'; $empty.AutoSize = $true
                    $empty.Location = New-Object System.Drawing.Point(0, $y); $panel.Controls.Add($empty)
                    $y += 28
                }
                for ($i = 0; $i -lt $tiles.Count; $i++) {
                    $tile = $tiles[$i]
                    $button = New-Object System.Windows.Forms.Button
                    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $button.FlatAppearance.BorderSize = 0
                    $button.Text = ''; $button.BackColor = Get-HdColor 'Card'
                    $button.Size = New-Object System.Drawing.Size(268, 132)
                    $button.Location = New-Object System.Drawing.Point((($i % 4) * 284), ($y + [int][Math]::Floor($i / 4.0) * 148))
                    $button.TabIndex = $tabIndex; $tabIndex++
                    $button.Tag = @{ Tile = $tile; Hover = $false }
                    $button.AccessibleName = "$($tile.Title): $(Format-HdValueText $tile.Value) $($tile.Unit), $($tile.Status)"
                    Set-HdDoubleBuffered $button
                    $button.Add_Paint({
                        $tileButton = $args[0]
                        Invoke-HdPaintTile -Graphics $args[1].Graphics -Tile $tileButton.Tag['Tile'] -Fonts $script:hdFonts -Width $tileButton.Width -Height $tileButton.Height -Focused $tileButton.Focused -Hover ([bool]$tileButton.Tag['Hover']) -ShowSparkline $script:hdShowSparklines
                    })
                    $button.Add_GotFocus({ $args[0].Invalidate() })
                    $button.Add_LostFocus({ $args[0].Invalidate() })
                    $button.Add_MouseEnter({ $args[0].Tag['Hover'] = $true; $args[0].Invalidate() })
                    $button.Add_MouseLeave({ $args[0].Tag['Hover'] = $false; $args[0].Invalidate() })
                    $button.Add_Click({
                        $selectedTile = $args[0].Tag['Tile']
                        if ($selectedTile.State -ne 'NoConfig' -and $script:hdOnTileActivated) {
                            try { & $script:hdOnTileActivated $selectedTile.MarkerName } catch { Write-Warning "Kachel: $($_.Exception.Message)" }
                        }
                    })
                    $panel.Controls.Add($button)
                    if ($script:hdToolTip) { $script:hdToolTip.SetToolTip($button, "$($tile.LongName)`n$($tile.Status)") }
                }
                $y += [int][Math]::Ceiling($tiles.Count / 4.0) * 148 + 16
            }
        }
        $panel.Height = [Math]::Max(64, $y)
        $panel.Tag = @{ Model = $script:hdModel; View = $View }
    } finally { $panel.ResumeLayout($true) }
}

function Update-HdViewSelector {
    # v2.40.0: Anzeigenamen und stabile Kennungen getrennt halten.
    if (-not $script:hdControls) { return }
    $selector = $script:hdControls.ViewSelector
    $count = if ($script:hdModel) { [int]$script:hdModel.Abnormalities.Count } else { 0 }
    $choices = @(@{ Key = 'Standard'; Name = 'Standard' }, @{ Key = 'Vorbelastungen'; Name = 'Vorbelastungen' })
    if ($count -gt 0) { $choices += @{ Key = 'Auffaelligkeiten'; Name = "Aktuelle Auffälligkeiten ($count)" } }
    if ($script:hdModel) {
        foreach ($view in @($script:hdModel.CustomViews.Views)) { if ($view) { $choices += @{ Key = ('Custom:' + $view.Id); Name = $view.Name } } }
    }
    $previousBusy = $script:hdViewSelectorBusy; $script:hdViewSelectorBusy = $true
    $selector.BeginUpdate()
    try {
        $script:hdViewChoices = @($choices)
        $same = ($selector.Items.Count -eq $choices.Count)
        if ($same) { for ($i = 0; $i -lt $choices.Count; $i++) { if ([string]$selector.Items[$i] -cne $choices[$i].Name) { $same = $false; break } } }
        if (-not $same) { $selector.Items.Clear(); $selector.Items.AddRange([object[]]@($choices | ForEach-Object { $_.Name })) }
    } finally { $selector.EndUpdate(); $script:hdViewSelectorBusy = $previousBusy }
    $invalid = ($script:hdModel -and $script:hdModel.CustomViews.Error)
    $script:hdControls.NewView.Enabled = [bool]($script:hdModel -and -not $invalid)
    $script:hdControls.ViewError.Text = if ($invalid) { $script:hdModel.CustomViews.Error } else { '' }
    $script:hdControls.ViewError.Visible = [bool]$invalid
    Set-HealthDashboardView -Name $script:hdSelectedView
}

function Set-HealthDashboardView {
    param([string]$Name = 'Standard')
    if (-not $script:hdControls) { return }
    $controls = $script:hdControls; $selectedIndex = -1
    for ($i = 0; $i -lt @($script:hdViewChoices).Count; $i++) { if ($script:hdViewChoices[$i].Key -eq $Name) { $selectedIndex = $i; break } }
    if ($selectedIndex -lt 0) { $Name = 'Standard'; $selectedIndex = 0 }
    $changed = ($script:hdSelectedView -ne $Name); $script:hdSelectedView = $Name
    $controls.Content.SuspendLayout()
    try {
        if ($changed) { $controls.Content.AutoScrollPosition = New-Object System.Drawing.Point(0, 0) }
        $standard = ($Name -eq 'Standard'); $custom = ($Name -like 'Custom:*')
        foreach ($control in @($controls.ScoreCard, $controls.PreCard, $controls.TestoCard) + @($controls.Tiles)) { $control.Visible = $standard }
        $controls.EditView.Visible = $custom; $controls.DeleteView.Visible = $custom
        if (-not $standard) { Update-HdPredispositionView -View $Name }
        $controls.ViewPanel.Visible = (-not $standard)
        $height = if ($standard) { 760 } else { [Math]::Max(760, 72 + $controls.ViewPanel.Height + 16) }
        $controls.Content.AutoScrollMinSize = New-Object System.Drawing.Size(1152, $height)
        if ($controls.ViewSelector.SelectedIndex -ne $selectedIndex) {
            $previousBusy = $script:hdViewSelectorBusy; $script:hdViewSelectorBusy = $true
            try { $controls.ViewSelector.SelectedIndex = $selectedIndex }
            finally { $script:hdViewSelectorBusy = $previousBusy }
        }
    } finally { $controls.Content.ResumeLayout($true) }
}
function Initialize-HealthDashboardTab {
    <#
    .SYNOPSIS
        Baut den Reiter "Dashboard" in eine vorhandene TabPage (Spec Teil A).
    .PARAMETER ModelProvider
        { Get-HealthDashboardModel ... } - liefert das Modell; wird bei jeder Aktualisierung aufgerufen.
    .PARAMETER OnTileActivated
        { param($MarkerName) ... } - Klick/Enter/Leertaste auf eine Kachel (Haupt-Script: Einzelmarker-Analyse).
    .PARAMETER OnOpenSettings
        { ... } - Links in "Vorbelastungen" (Haupt-Script: Reiter "Persönliche Metriken").
    .PARAMETER SignatureProvider
        v0.1.1, optional: { Get-HealthDashDataSignature ... } - gleiche Prüfsumme = kein Neuberechnen.
        Ohne Angabe wird bei jeder Aktualisierung neu berechnet (Test-Starter).
    #>
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TabPage]$TabPage, [Parameter(Mandatory)][scriptblock]$ModelProvider,
        [scriptblock]$OnTileActivated = $null, [scriptblock]$OnOpenSettings = $null, $ToolTip = $null, [scriptblock]$SignatureProvider = $null
    )
    $script:hdSelectedView = 'Standard'
    $script:hdViewSelectorBusy = $false
    $script:hdModelProvider = $ModelProvider
    $script:hdSignatureProvider = $SignatureProvider
    $script:hdModelSignature = ''
    $script:hdOnTileActivated = $OnTileActivated
    $script:hdOnOpenSettings = $OnOpenSettings
    $script:hdToolTip = $ToolTip
    $script:hdModel = $null
    $script:hdFonts = New-HdFonts
    $TabPage.BackColor = Get-HdColor 'Page'

    $content = New-Object System.Windows.Forms.Panel
    $content.Dock = [System.Windows.Forms.DockStyle]::Fill
    $content.AutoScroll = $true
    $content.AutoScrollMinSize = New-Object System.Drawing.Size(1152, 760)
    $content.BackColor = Get-HdColor 'Page'
    $TabPage.Controls.Add($content)

    # Kopfzeile (Spec 4.1)
    $title = New-Object System.Windows.Forms.Label
    $title.Text = 'Gesundheitsstatus'; $title.Font = $script:hdFonts.Title; $title.ForeColor = Get-HdColor 'TextPrimary'
    $title.AutoSize = $true; $title.Location = New-Object System.Drawing.Point(16, 20)
    $content.Controls.Add($title)
    $link = New-Object System.Windows.Forms.LinkLabel
    $link.Text = 'Wie berechnet?'; $link.Font = $script:hdFonts.Normal; $link.AutoSize = $true
    $link.LinkColor = Get-HdColor 'Link'; $link.ActiveLinkColor = Get-HdColor 'Link'; $link.VisitedLinkColor = Get-HdColor 'Link'
    $link.TabIndex = 1
    $content.Controls.Add($link)
    $link.Location = New-Object System.Drawing.Point((1136 - $link.PreferredWidth), 28)
    $info = New-Object System.Windows.Forms.Label
    $info.Font = $script:hdFonts.Normal; $info.ForeColor = Get-HdColor 'TextMuted'; $info.AutoSize = $false
    $info.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $info.Size = New-Object System.Drawing.Size(744, 18); $info.Location = New-Object System.Drawing.Point(392, 48)
    $content.Controls.Add($info)

    # v2.40.0: Eigene Ansichten und Verwaltung neben der Auswahl.
    $viewLabel = New-Object System.Windows.Forms.Label
    $viewLabel.Text = 'Ansicht:'; $viewLabel.Font = $script:hdFonts.Normal; $viewLabel.ForeColor = Get-HdColor 'TextSecondary'
    $viewLabel.AutoSize = $true; $viewLabel.Location = New-Object System.Drawing.Point(392, 27)
    $content.Controls.Add($viewLabel)
    $viewSelector = New-Object System.Windows.Forms.ComboBox
    $viewSelector.Font = $script:hdFonts.Normal; $viewSelector.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    $viewSelector.Location = New-Object System.Drawing.Point(452, 23); $viewSelector.Size = New-Object System.Drawing.Size(268, 24)
    $viewSelector.Items.AddRange([object[]]@('Standard', 'Vorbelastungen')); $viewSelector.SelectedIndex = 0
    $viewSelector.TabIndex = 0; $viewSelector.AccessibleName = 'Dashboard-Ansicht'
    $content.Controls.Add($viewSelector)
    $newView = New-Object System.Windows.Forms.Button
    $newView.Text = '+ Eigene Ansicht…'; $newView.Location = New-Object System.Drawing.Point(728, 22); $newView.Size = New-Object System.Drawing.Size(148, 26)
    $newView.Font = $script:hdFonts.Normal; $newView.TabIndex = 1
    $newView.Add_Click({ Show-HdViewEditor })
    $editView = New-Object System.Windows.Forms.Button
    $editView.Text = 'Bearbeiten…'; $editView.Location = New-Object System.Drawing.Point(884, 22); $editView.Size = New-Object System.Drawing.Size(78, 26)
    $editView.Font = $script:hdFonts.Normal; $editView.TabIndex = 2; $editView.Visible = $false
    $editView.Add_Click({ if ($script:hdSelectedView -like 'Custom:*') { Show-HdViewEditor -Id $script:hdSelectedView.Substring(7) } })
    $deleteView = New-Object System.Windows.Forms.Button
    $deleteView.Text = 'Löschen…'; $deleteView.Location = New-Object System.Drawing.Point(968, 22); $deleteView.Size = New-Object System.Drawing.Size(72, 26)
    $deleteView.Font = $script:hdFonts.Normal; $deleteView.TabIndex = 3; $deleteView.Visible = $false
    $deleteView.Add_Click({ Show-HdDeleteViewConfirmation })
    $viewError = New-Object System.Windows.Forms.Label
    $viewError.Location = New-Object System.Drawing.Point(392, 48); $viewError.Size = New-Object System.Drawing.Size(744, 22)
    $viewError.Font = $script:hdFonts.Normal; $viewError.ForeColor = Get-HdColor 'TextSecondary'; $viewError.BackColor = Get-HdColor 'Page'
    $viewError.AutoEllipsis = $true; $viewError.Visible = $false
    $content.Controls.AddRange([System.Windows.Forms.Control[]]@($newView, $editView, $deleteView, $viewError))
    $viewError.BringToFront(); $link.TabIndex = 4
    $viewPanel = New-Object System.Windows.Forms.Panel
    $viewPanel.Location = New-Object System.Drawing.Point(16, 72); $viewPanel.Size = New-Object System.Drawing.Size(1120, 64)
    $viewPanel.BackColor = Get-HdColor 'Page'; $viewPanel.Visible = $false; $viewPanel.TabIndex = 2
    $content.Controls.Add($viewPanel)

    # Score-Karte (Spec 4.2)
    $scoreCard = New-Object System.Windows.Forms.Panel
    $scoreCard.Location = New-Object System.Drawing.Point(16, 72); $scoreCard.Size = New-Object System.Drawing.Size(360, 280)
    Set-HdDoubleBuffered $scoreCard
    $scoreCard.Add_Paint({ if ($script:hdModel) { Invoke-HdPaintScoreCard -Graphics $args[1].Graphics -Model $script:hdModel -Fonts $script:hdFonts } })
    $content.Controls.Add($scoreCard)

    # Kacheln (Spec 4.3, E3): fokussierbare, selbst gezeichnete Buttons
    $tiles = @()
    $positions = @(@(392, 72), @(645, 72), @(898, 72), @(392, 220), @(645, 220), @(898, 220))
    for ($i = 0; $i -lt 6; $i++) {
        $button = New-Object System.Windows.Forms.Button
        $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $button.FlatAppearance.BorderSize = 0
        $button.Text = ''
        $button.BackColor = Get-HdColor 'Card'
        $button.Location = New-Object System.Drawing.Point($positions[$i][0], $positions[$i][1]); $button.Size = New-Object System.Drawing.Size(237, 132)
        $button.TabIndex = 2 + $i
        $button.Tag = @{ Index = $i; Hover = $false }
        Set-HdDoubleBuffered $button
        $button.Add_Paint({
            $tileButton = $args[0]
            $tileIndex = [int]$tileButton.Tag['Index']
            if ($script:hdModel -and $tileIndex -lt @($script:hdModel.Tiles).Count) {
                Invoke-HdPaintTile -Graphics $args[1].Graphics -Tile @($script:hdModel.Tiles)[$tileIndex] -Fonts $script:hdFonts -Focused $tileButton.Focused -Hover ([bool]$tileButton.Tag['Hover']) -ShowSparkline $script:hdShowSparklines
            }
        })
        $button.Add_GotFocus({ $args[0].Invalidate() })
        $button.Add_LostFocus({ $args[0].Invalidate() })
        $button.Add_MouseEnter({ $args[0].Tag['Hover'] = $true; $args[0].Invalidate() })
        $button.Add_MouseLeave({ $args[0].Tag['Hover'] = $false; $args[0].Invalidate() })
        $button.Add_Click({
            $tileIndex = [int]$args[0].Tag['Index']
            if ($script:hdModel -and $script:hdOnTileActivated -and $tileIndex -lt @($script:hdModel.Tiles).Count) {
                $tile = @($script:hdModel.Tiles)[$tileIndex]
                try { & $script:hdOnTileActivated $tile.MarkerName } catch { Write-Warning "Kachel: $($_.Exception.Message)" }
            }
        })
        $content.Controls.Add($button)
        $tiles += , $button
    }

    # Vorbelastungen (Spec 4.4)
    $preCard = New-Object System.Windows.Forms.Panel
    $preCard.Location = New-Object System.Drawing.Point(16, 368); $preCard.Size = New-Object System.Drawing.Size(560, 376)
    Set-HdDoubleBuffered $preCard
    $preCard.Add_Paint({
        $title = 'Vorbelastungen'
        if ($script:hdModel -and $script:hdModel.Predispositions.State -eq 'Active') { $title = "Vorbelastungen ($(@($script:hdModel.Predispositions.Entries).Count) aktiv)" }
        Invoke-HdDrawCardFrame -Graphics $args[1].Graphics -Width 560 -Height 376 -Title $title -Fonts $script:hdFonts
    })
    $content.Controls.Add($preCard)
    $preScroll = New-Object System.Windows.Forms.Panel
    $preScroll.Location = New-Object System.Drawing.Point(16, 44); $preScroll.Size = New-Object System.Drawing.Size(528, 316)
    $preScroll.AutoScroll = $true; $preScroll.BackColor = Get-HdColor 'Card'; $preScroll.TabIndex = 9
    $preCard.Controls.Add($preScroll)
    $preCanvas = New-Object System.Windows.Forms.Panel
    $preCanvas.Location = New-Object System.Drawing.Point(0, 0); $preCanvas.Size = New-Object System.Drawing.Size(528, 10)
    $preCanvas.BackColor = Get-HdColor 'Card'
    Set-HdDoubleBuffered $preCanvas
    $preCanvas.Add_Paint({ if ($script:hdModel) { Invoke-HdPaintPredispositions -Graphics $args[1].Graphics -Predispositions $script:hdModel.Predispositions -Fonts $script:hdFonts -Width $args[0].Width } })
    $preScroll.Controls.Add($preCanvas)
    $preMessage = New-Object System.Windows.Forms.Label
    $preMessage.Font = $script:hdFonts.Normal; $preMessage.ForeColor = Get-HdColor 'TextSecondary'; $preMessage.AutoSize = $true
    $preMessage.Location = New-Object System.Drawing.Point(0, 4); $preMessage.Visible = $false
    $preScroll.Controls.Add($preMessage)
    $preLink = New-Object System.Windows.Forms.LinkLabel
    $preLink.Font = $script:hdFonts.Normal; $preLink.AutoSize = $true; $preLink.Location = New-Object System.Drawing.Point(0, 26); $preLink.Visible = $false
    $preLink.LinkColor = Get-HdColor 'Link'; $preLink.ActiveLinkColor = Get-HdColor 'Link'; $preLink.VisitedLinkColor = Get-HdColor 'Link'; $preLink.TabIndex = 8
    $preLink.Add_LinkClicked({ if ($script:hdOnOpenSettings) { try { & $script:hdOnOpenSettings } catch { Write-Warning "Vorbelastungen: $($_.Exception.Message)" } } })
    $preScroll.Controls.Add($preLink)

    # Testosteron (Spec 4.5)
    $testoCard = New-Object System.Windows.Forms.Panel
    $testoCard.Location = New-Object System.Drawing.Point(592, 368); $testoCard.Size = New-Object System.Drawing.Size(544, 376)
    Set-HdDoubleBuffered $testoCard
    $testoCard.Add_Paint({ if ($script:hdModel) { Invoke-HdPaintTestosterone -Graphics $args[1].Graphics -Testosterone $script:hdModel.Testosterone -Fonts $script:hdFonts } })
    $content.Controls.Add($testoCard)

    $link.Add_LinkClicked({ Show-HealthDashExplainDialog -Owner ($script:hdControls.Content.FindForm()) })
    $script:hdControls = @{ Content = $content; Info = $info; Link = $link; ScoreCard = $scoreCard; Tiles = $tiles; PreCard = $preCard; PreScroll = $preScroll; PreCanvas = $preCanvas; PreMessage = $preMessage; PreLink = $preLink; TestoCard = $testoCard; ViewSelector = $viewSelector; ViewPanel = $viewPanel; NewView = $newView; EditView = $editView; DeleteView = $deleteView; ViewError = $viewError }
    if ($ToolTip) {
        $ToolTip.SetToolTip($scoreCard, "$($script:HealthDashMeaningText)`n$($script:HealthDashWeightingText)")
        $ToolTip.SetToolTip($testoCard, 'Berechnete Werte ohne Bewertung. FAI: Index, keine Prozentangabe.')
        $ToolTip.SetToolTip($link, 'Rechenweg, Funktionsbereiche und nicht bewertete Marker')
    }
    $viewSelector.Add_SelectedIndexChanged({
        if ($script:hdViewSelectorBusy -or -not $args[0].SelectedItem) { return }
        if ($args[0].SelectedIndex -ge @($script:hdViewChoices).Count) { return }
        $viewName = $script:hdViewChoices[$args[0].SelectedIndex].Key
        Set-HealthDashboardView -Name $viewName
    })
    Update-HealthDashboardTab -Force
}

function Update-HealthDashboardTab {
    <#
        Modell neu berechnen und alle Bereiche neu zeichnen (z. B. beim Reiterwechsel).
        v0.1.1: Mit SignatureProvider nur, wenn sich die Eingaben geändert haben (oder -Force).
    #>
    param([switch]$Force)
    if (-not $script:hdControls -or -not $script:hdModelProvider) { return }
    $controls = $script:hdControls
    $signature = ''
    if ($script:hdSignatureProvider) {
        try { $signature = [string](& $script:hdSignatureProvider) }
        catch { $signature = ''; Write-Warning "Dashboard-Prüfsumme: $($_.Exception.Message)" }
    }
    if (Test-HdModelCurrent -Model $script:hdModel -ModelSignature $script:hdModelSignature -Signature $signature -Force:$Force) { return }
    $previousCursor = [System.Windows.Forms.Cursor]::Current
    [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::WaitCursor
    try { $script:hdModel = & $script:hdModelProvider }
    catch {
        $script:hdModel = $null
        $script:hdModelSignature = ''
        $controls.Info.Text = "Dashboard konnte nicht berechnet werden: $($_.Exception.Message)"
        Write-Warning "Dashboard: $($_.Exception.Message)"
        Update-HdViewSelector
        foreach ($control in @($controls.ScoreCard, $controls.PreCard, $controls.PreCanvas, $controls.TestoCard) + @($controls.Tiles)) { $control.Invalidate() }
        return
    }
    finally { [System.Windows.Forms.Cursor]::Current = $previousCursor }
    $script:hdModelSignature = $signature
    $model = $script:hdModel
    $infoText = 'Noch keine Bluttests erfasst'
    if ($model.LastTestDate) { $infoText = "Letzter Bluttest: $(Format-HdDate $model.LastTestDate) · $($model.Score.MarkerCount) Marker im Gesundheitsscore" }
    $controls.Info.Text = $infoText
    # Kacheln: Tooltip = Langname (+ Erklärung)
    $tiles = @($model.Tiles)
    for ($i = 0; $i -lt $controls.Tiles.Count; $i++) {
        $button = $controls.Tiles[$i]
        if ($i -lt $tiles.Count) {
            $tile = $tiles[$i]
            $button.AccessibleName = "$($tile.Title): $(Format-HdValueText $tile.Value) $($tile.Unit), $($tile.Status)"
            if ($script:hdToolTip) { $tipText = $tile.LongName; if ($tile.Tooltip) { $tipText += "`n$($tile.Tooltip)" }; $script:hdToolTip.SetToolTip($button, $tipText) }
        }
        $button.Invalidate()
    }
    # Vorbelastungen: Leerzustände als Text + Link, sonst gezeichnete Blöcke (Inhaltshöhe = Scrollhöhe)
    $pre = $model.Predispositions
    $controls.PreMessage.Visible = $false; $controls.PreLink.Visible = $false
    if ($pre.State -eq 'Disabled') {
        $controls.PreMessage.Text = 'Vorbelastungen sind ausgeschaltet.'; $controls.PreLink.Text = 'In „Persönliche Metriken“ einschalten'
        $controls.PreMessage.Visible = $true; $controls.PreLink.Visible = $true
    } elseif ($pre.State -eq 'NoneActive') {
        $controls.PreMessage.Text = 'Keine Vorbelastung aktiv.'; $controls.PreLink.Text = 'In „Persönliche Metriken“ festlegen'
        $controls.PreMessage.Visible = $true; $controls.PreLink.Visible = $true
    }
    $contentHeight = Get-HdPredispositionHeight -Predispositions $pre
    $canvasWidth = 528
    if ($contentHeight -gt $controls.PreScroll.ClientSize.Height) { $canvasWidth = 528 - [System.Windows.Forms.SystemInformation]::VerticalScrollBarWidth }
    $controls.PreCanvas.Size = New-Object System.Drawing.Size($canvasWidth, [Math]::Max(1, $contentHeight))
    $controls.PreCanvas.Visible = ($contentHeight -gt 0)
    foreach ($control in @($controls.ScoreCard, $controls.PreCard, $controls.PreCanvas, $controls.TestoCard)) { $control.Invalidate() }
    Update-HdViewSelector
}

function Clear-HealthDashboardTab {
    <# Reiter leeren und Schriften freigeben (Prototyp: "Design neu laden"; danach Initialize-HealthDashboardTab). #>
    param([Parameter(Mandatory)][System.Windows.Forms.TabPage]$TabPage)
    $old = @($TabPage.Controls | ForEach-Object { $_ })
    $TabPage.Controls.Clear()
    foreach ($control in $old) { $control.Dispose() }
    if ($script:hdFonts) { foreach ($font in @($script:hdFonts.Values)) { if ($font) { $font.Dispose() } } }
    $script:hdFonts = $null; $script:hdControls = $null; $script:hdModel = $null; $script:hdModelSignature = ''
}

function Show-HealthDashExplainDialog {
    <# Dialog "Wie berechnet?" (Spec 4.6): modal 480 x 420, Esc und "Schließen", schreibgeschützt, per Tastatur scrollbar. #>
    param($Owner = $null)
    if (-not $script:hdModel) { return }
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = 'Wie berechnet?'; $dialog.Size = New-Object System.Drawing.Size(480, 420)
    $dialog.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
    $dialog.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog; $dialog.MaximizeBox = $false; $dialog.MinimizeBox = $false; $dialog.ShowInTaskbar = $false
    $dialog.Font = $script:hdFonts.Normal; $dialog.BackColor = Get-HdColor 'Card'
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Multiline = $true; $textBox.ReadOnly = $true; $textBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $textBox.BackColor = Get-HdColor 'Card'; $textBox.ForeColor = Get-HdColor 'TextPrimary'; $textBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $textBox.Location = New-Object System.Drawing.Point(12, 12); $textBox.Size = New-Object System.Drawing.Size(440, 318)
    $textBox.Text = (Get-HealthDashExplainText -Model $script:hdModel) -replace "`r?`n", "`r`n"
    $textBox.TabIndex = 0
    $dialog.Controls.Add($textBox)
    $close = New-Object System.Windows.Forms.Button
    $close.Text = 'Schließen'; $close.Size = New-Object System.Drawing.Size(100, 28); $close.Location = New-Object System.Drawing.Point(352, 340)
    $close.DialogResult = [System.Windows.Forms.DialogResult]::Cancel; $close.TabIndex = 1
    $dialog.Controls.Add($close)
    $dialog.CancelButton = $close; $dialog.AcceptButton = $close
    $dialog.Add_Shown({ $args[0].Controls[0].SelectionLength = 0 })
    try { if ($Owner) { [void]$dialog.ShowDialog($Owner) } else { [void]$dialog.ShowDialog() } } finally { $dialog.Dispose() }
}
# ════ Ende DASHBOARD (v2.36.0) ═══════════════════════════════════════════════════════════

# =====================================================================
# v2.48.0: Zusaetzliche Mehrfacheingabe mit Pruefansicht, gemeinsames Datum
# =====================================================================
function Test-BtBloodFormEntries {
    # Reine Vorbereitung: kein Dateizugriff, keine Speicherung. Leere Zeilen ignorieren.
    param([string]$DateString, [object[]]$Rows, [object[]]$ExistingItems = $script:allHistoricalItems)
    $items = @(); $display = @(); $errors = @(); $index = 0
    foreach ($row in $Rows) {
        $index++
        $contents = @($row.MarkerQuery, $row.ValueText, $row.HivResult, $row.ApoeGenotype, $row.Note)
        if (@($contents | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -eq 0) { continue }
        try {
            $item = New-BtBloodValueItem -DateString $DateString -MarkerQuery $row.MarkerQuery -ValueText $row.ValueText -HivResult $row.HivResult -ApoeGenotype $row.ApoeGenotype -TestosteroneUnit $row.TestosteroneUnit -Note $row.Note -ExistingItems $ExistingItems -PendingItems $items
            $cfg = @($script:data.Config.Markers | Where-Object Name -eq $item.Name)[0]
            $hint = 'Keine Referenzbewertung vorhanden'; $warning = $false
            try {
                $score = Get-MarkerRangeScore -Value $item.Value -RefMin $cfg.RefMin -RefMax $cfg.RefMax -OptimalMin $cfg.OptimalMin -OptimalMax $cfg.OptimalMax
                if ($null -ne $score -and -not [double]::IsNaN($score) -and -not [double]::IsInfinity($score)) {
                    if ($score -lt 50) { $hint = 'Außerhalb der Norm – Eingabe bitte prüfen'; $warning = $true }
                    elseif ($score -ge 80) { $hint = 'Im Optimalbereich' }
                    else { $hint = 'Im Referenzbereich' }
                }
            } catch { $hint = 'Bewertungsbereiche bitte prüfen; der Wert kann gespeichert werden.' }
            $original = [string]$row.ValueText
            $stored = ([double]$item.Value).ToString('G', [cultureinfo]::CurrentCulture) + ' ' + $item.Unit
            if ($item.Name -eq 'HIV (Anti-HIV-1/2)') { $original = $row.HivResult; $stored = $row.HivResult }
            elseif ($item.Name -eq $script:ApoeMarkerName) { $original = $row.ApoeGenotype; $stored = $row.ApoeGenotype }
            elseif ($item.Name -eq 'Testosteron, gesamt') { $original += ' ' + $row.TestosteroneUnit }
            else { $original += ' ' + $item.Unit }
            $items += $item
            $display += [PSCustomObject]@{ Zeile = $index; Marker = $item.Name; Eingabe = $original.Trim(); Gespeichert = $stored.Trim(); Hinweis = $hint; Warnung = $warning; Notiz = $item.Note }
        } catch { $errors += [PSCustomObject]@{ Zeile = $index; Text = $_.Exception.Message } }
    }
    if ($items.Count -eq 0 -and $errors.Count -eq 0) { $errors += [PSCustomObject]@{ Zeile = 0; Text = 'Bitte mindestens einen Blutwert eingeben.' } }
    return [PSCustomObject]@{ IsValid = ($errors.Count -eq 0); Items = @($items); Anzeige = @($display); Fehler = @($errors) }
}

function Resize-BtBloodFormRows {
    $state = $script:btBloodForm
    if ($null -eq $state -or $state.Form.IsDisposed -or $state.Resizing) { return }
    $state.Resizing = $true
    $state.Flow.SuspendLayout()
    try {
        $width = [Math]::Max(840, $state.Flow.ClientSize.Width - 24)
        foreach ($row in $state.Rows) {
            $row.Panel.Width = $width
            $row.Note.Width = $width - 588
            $row.Remove.Left = $width - 30
            $row.Hint.Width = $width - 65
        }
    } finally {
        # Scroll-Ausdehnung erst aus allen angepassten Zeilen neu berechnen.
        $state.Flow.ResumeLayout($true)
        $state.Resizing = $false
    }
}

function Set-BtBloodFormMarkerList {
    param($Row, [string[]]$Names, [string]$Text)
    $previous = $Row.Busy; $Row.Busy = $true
    $Row.Marker.BeginUpdate()
    try {
        $Row.Marker.Items.Clear()
        if ($Names.Count -gt 0) { $Row.Marker.Items.AddRange([string[]]$Names) }
        $Row.Marker.SelectedIndex = -1; $Row.Marker.Text = $Text
        $Row.Marker.SelectionStart = $Text.Length
    } finally { $Row.Marker.EndUpdate(); $Row.Busy = $previous }
}

function Set-BtBloodFormInputMode {
    param($Row, [string]$Name)
    $previous = $Row.Busy; $Row.Busy = $true
    try {
        if ($Row.ResolvedName -ne $Name) {
            # Kein Ergebnis von einem anderen Marker unbemerkt uebernehmen.
            $Row.Value.Clear(); $Row.Choice.Items.Clear(); $Row.Choice.SelectedIndex = -1
            $Row.Unit.SelectedIndex = 0
            if ($Name -eq 'HIV (Anti-HIV-1/2)') { $Row.Choice.Items.AddRange([string[]]@('nicht reaktiv', 'reaktiv')) }
            elseif ($Name -eq $script:ApoeMarkerName) { $Row.Choice.Items.AddRange([string[]]$script:ApoeGenotypeOptions) }
        }
        $Row.ResolvedName = $Name
        $special = ($Name -eq 'HIV (Anti-HIV-1/2)' -or ($Name -and $Name -eq $script:ApoeMarkerName))
        $Row.Value.Visible = -not $special; $Row.Choice.Visible = $special
        $Row.Unit.Visible = ($Name -eq 'Testosteron, gesamt'); $Row.UnitText.Visible = -not $Row.Unit.Visible
        $Row.UnitText.Text = ''; $Row.Hint.ForeColor = [Drawing.Color]::DimGray
        $Row.Hint.Text = 'Marker auswählen oder einen eindeutigen Suchbegriff eingeben.'
        if ($Name) {
            $cfg = @($script:data.Config.Markers | Where-Object Name -eq $Name)[0]
            $Row.UnitText.Text = [string]$cfg.Unit
            if ($Name -eq $script:ApoeMarkerName) { $Row.UnitText.Text = ''; $Row.Hint.Text = 'Genotyp wählen | Referenz: kein E4-Allel (E2/E2, E2/E3, E3/E3)' }
            elseif ($Name -eq 'HIV (Anti-HIV-1/2)') { $Row.UnitText.Text = ''; $Row.Hint.Text = 'Ergebnis ausdrücklich aus dem Befund auswählen.' }
            elseif ($null -ne $cfg.RefMin -and [string]$cfg.RefMin -ne '' -and $null -ne $cfg.RefMax -and [string]$cfg.RefMax -ne '') { $Row.Hint.Text = "Referenz: $($cfg.RefMin) – $($cfg.RefMax) $($cfg.Unit)" }
            else { $Row.Hint.Text = 'Kein vollständiger Referenzbereich hinterlegt.' }
        }
    } finally { $Row.Busy = $previous }
}

function Complete-BtBloodFormMarker {
    param($Row)
    if ($Row.Busy) { return }
    $typed = $Row.Marker.Text.Trim()
    $name = Resolve-MarkerName -Query $typed
    $Row.Marker.DroppedDown = $false
    Set-BtBloodFormMarkerList -Row $Row -Names $script:btBloodForm.MarkerNames -Text $(if ($name) { $name } else { $typed })
    Set-BtBloodFormInputMode -Row $Row -Name $name
    $Row.Marker.BackColor = if ($typed -and -not $name) { [Drawing.Color]::MistyRose } else { [Drawing.SystemColors]::Window }
}

function Add-BtBloodFormRow {
    $state = $script:btBloodForm
    $panel = New-Object Windows.Forms.Panel; $panel.Size = New-Object Drawing.Size(920, 62); $panel.Margin = New-Object Windows.Forms.Padding(0, 0, 0, 4)
    $number = New-Object Windows.Forms.Label; $number.Location = New-Object Drawing.Point(0, 5); $number.Size = New-Object Drawing.Size(28, 20)
    $marker = New-Object Windows.Forms.ComboBox; $marker.Location = New-Object Drawing.Point(28, 2); $marker.Size = New-Object Drawing.Size(280, 24); $marker.DropDownStyle = 'DropDown'; $marker.MaxDropDownItems = 12; $marker.TabIndex = 0
    $value = New-Object Windows.Forms.TextBox; $value.Location = New-Object Drawing.Point(316, 2); $value.Size = New-Object Drawing.Size(130, 24); $value.TabIndex = 1
    $choice = New-Object Windows.Forms.ComboBox; $choice.Location = $value.Location; $choice.Size = $value.Size; $choice.DropDownStyle = 'DropDownList'; $choice.TabIndex = 1; $choice.Visible = $false
    $unit = New-Object Windows.Forms.ComboBox; $unit.Location = New-Object Drawing.Point(454, 2); $unit.Size = New-Object Drawing.Size(82, 24); $unit.DropDownStyle = 'DropDownList'; $unit.Items.AddRange([string[]]@('ng/dl', 'nmol/l')); $unit.SelectedIndex = 0; $unit.TabIndex = 2; $unit.Visible = $false
    $unitText = New-Object Windows.Forms.Label; $unitText.Location = New-Object Drawing.Point(454, 5); $unitText.Size = New-Object Drawing.Size(86, 20); $unitText.AutoEllipsis = $true
    $note = New-Object Windows.Forms.TextBox; $note.Location = New-Object Drawing.Point(545, 2); $note.Size = New-Object Drawing.Size(330, 24); $note.TabIndex = 3
    $remove = New-Object Windows.Forms.Button; $remove.Text = '×'; $remove.Size = New-Object Drawing.Size(28, 26); $remove.Top = 0; $remove.TabIndex = 4; $remove.AccessibleName = 'Zeile entfernen'
    $hint = New-Object Windows.Forms.Label; $hint.Location = New-Object Drawing.Point(28, 31); $hint.Height = 28; $hint.Font = $state.HintFont
    $panel.Controls.AddRange([Windows.Forms.Control[]]@($number, $marker, $value, $choice, $unit, $unitText, $note, $remove, $hint))
    $row = [PSCustomObject]@{ Panel = $panel; Number = $number; Marker = $marker; Value = $value; Choice = $choice; Unit = $unit; UnitText = $unitText; Note = $note; Remove = $remove; Hint = $hint; ResolvedName = ''; Busy = $false }
    foreach ($control in @($marker, $value, $choice, $unit, $note, $remove)) { $control.Tag = $row }
    Set-BtBloodFormMarkerList -Row $row -Names $state.MarkerNames -Text ''
    Set-BtBloodFormInputMode -Row $row -Name ''
    $marker.Add_TextUpdate({
        $row = $args[0].Tag
        if ($row.Busy) { return }
        $typed = $row.Marker.Text
        $hits = if ([string]::IsNullOrWhiteSpace($typed)) { $script:btBloodForm.MarkerNames } else { @(Get-MarkerMatches -Query $typed) }
        if (@($hits).Count -eq 0) { $hits = $script:btBloodForm.MarkerNames }
        Set-BtBloodFormMarkerList -Row $row -Names @($hits) -Text $typed
        Set-BtBloodFormInputMode -Row $row -Name (Resolve-MarkerName -Query $typed)
        if ($typed -and $row.Marker.Focused) { $row.Marker.DroppedDown = $true; $row.Marker.SelectionStart = $typed.Length }
        $script:btBloodForm.Dirty = $true
    })
    $marker.Add_SelectedIndexChanged({
        $row = $args[0].Tag
        if ($row.Busy -or $row.Marker.SelectedIndex -lt 0) { return }
        Set-BtBloodFormInputMode -Row $row -Name ([string]$row.Marker.SelectedItem)
        $row.Marker.BackColor = [Drawing.SystemColors]::Window
        $script:btBloodForm.Dirty = $true
    })
    $marker.Add_Leave({ Complete-BtBloodFormMarker -Row $args[0].Tag })
    $marker.Add_KeyDown({
        if ($args[1].KeyCode -eq [Windows.Forms.Keys]::Enter) {
            $args[1].Handled = $true; $args[1].SuppressKeyPress = $true
            $row = $args[0].Tag; Complete-BtBloodFormMarker -Row $row
            if ($row.ResolvedName) { if ($row.Choice.Visible) { $null = $row.Choice.Focus() } else { $null = $row.Value.Focus() } }
        } elseif ($args[1].KeyCode -eq [Windows.Forms.Keys]::Escape -and $args[0].DroppedDown) {
            $args[0].DroppedDown = $false; $args[1].Handled = $true; $args[1].SuppressKeyPress = $true
        }
    })
    foreach ($control in @($value, $note)) { $control.Add_TextChanged({ if (-not $args[0].Tag.Busy) { $script:btBloodForm.Dirty = $true } }) }
    foreach ($control in @($choice, $unit)) { $control.Add_SelectedIndexChanged({ if (-not $args[0].Tag.Busy) { $script:btBloodForm.Dirty = $true } }) }
    foreach ($control in @($value, $note)) {
        $control.Add_KeyDown({
            if ($args[1].KeyCode -eq [Windows.Forms.Keys]::Enter) {
                $args[1].Handled = $true; $args[1].SuppressKeyPress = $true
                $null = $script:btBloodForm.Form.SelectNextControl($args[0], $true, $true, $true, $true)
            }
        })
    }
    $remove.Add_Click({
        $row = $args[0].Tag
        $script:btBloodForm.Dirty = $true
        $script:btBloodForm.Flow.Controls.Remove($row.Panel); $script:btBloodForm.Rows.Remove($row); $row.Panel.Dispose()
        if ($script:btBloodForm.Rows.Count -eq 0) { Add-BtBloodFormRow }
        $index = 0; foreach ($remaining in $script:btBloodForm.Rows) { $index++; $remaining.Number.Text = "$index."; $remaining.Panel.TabIndex = $index - 1 }
        Resize-BtBloodFormRows
    })
    $panel.TabIndex = $state.Rows.Count; $number.Text = [string]($state.Rows.Count + 1) + '.'
    [void]$state.Rows.Add($row); $state.Flow.Controls.Add($panel)
    Resize-BtBloodFormRows
    $state.Flow.ScrollControlIntoView($panel); $null = $marker.Focus()
}

function Set-BtBloodFormReview {
    param([bool]$Visible)
    $state = $script:btBloodForm
    $state.Reviewing = $Visible; $state.EntryPanel.Visible = -not $Visible; $state.ReviewPanel.Visible = $Visible
    $state.Date.Enabled = -not $Visible; $state.Add.Visible = -not $Visible; $state.Back.Visible = $Visible
    $state.Primary.Text = if ($Visible) { "Alle $($state.Prepared.Count) Werte speichern" } else { 'Eingaben prüfen …' }
    if (-not $Visible) { $state.Prepared = @(); $state.Status.Text = 'Tab/Shift+Tab: Felder wechseln. Enter übernimmt die Markerauswahl.'; $state.Status.ForeColor = [Drawing.Color]::DimGray }
}

function Invoke-BtBloodFormPrimary {
    $state = $script:btBloodForm
    if ($state.Busy) { return }
    $state.Busy = $true; $state.Primary.Enabled = $false
    try {
        if (-not $state.Reviewing) {
            $models = @()
            foreach ($row in $state.Rows) {
                Complete-BtBloodFormMarker -Row $row
                $models += [PSCustomObject]@{
                    MarkerQuery = $row.Marker.Text; ValueText = $row.Value.Text; Note = $row.Note.Text
                    HivResult = $(if ($row.ResolvedName -eq 'HIV (Anti-HIV-1/2)') { $row.Choice.SelectedItem } else { '' })
                    ApoeGenotype = $(if ($row.ResolvedName -eq $script:ApoeMarkerName) { $row.Choice.SelectedItem } else { '' })
                    TestosteroneUnit = $row.Unit.SelectedItem
                }
            }
            $review = Test-BtBloodFormEntries -DateString $state.Date.Value.ToString('yyyy-MM-dd') -Rows $models
            if (-not $review.IsValid) {
                foreach ($problem in $review.Fehler) {
                    if ($problem.Zeile -gt 0) { $row = $state.Rows[$problem.Zeile - 1]; $row.Hint.Text = $problem.Text; $row.Hint.ForeColor = [Drawing.Color]::Firebrick }
                }
                $first = $review.Fehler[0]
                $state.Status.Text = if ($first.Zeile -gt 0) { "Zeile $($first.Zeile): $($first.Text)" } else { $first.Text }
                $state.Status.ForeColor = [Drawing.Color]::Firebrick
                if ($first.Zeile -gt 0) { $state.Flow.ScrollControlIntoView($state.Rows[$first.Zeile - 1].Panel) }
                return
            }
            $state.Prepared = @($review.Items); $state.Grid.Rows.Clear()
            foreach ($item in $review.Anzeige) {
                $index = $state.Grid.Rows.Add($item.Marker, $item.Eingabe, $item.Gespeichert, $item.Hinweis, $item.Notiz)
                if ($item.Warnung) { $state.Grid.Rows[$index].DefaultCellStyle.BackColor = [Drawing.Color]::LemonChiffon }
            }
            $state.ReviewTitle.Text = "$($review.Items.Count) Werte für $($state.Date.Value.ToString('dd.MM.yyyy')) – bitte Eingaben und Umrechnungen prüfen."
            $state.Status.Text = 'Hinweise außerhalb der Norm verhindern das Speichern nicht. Gespeichert wird erst mit dem rechten Button.'
            $state.Status.ForeColor = [Drawing.Color]::DimGray
            Set-BtBloodFormReview -Visible $true
            $null = $state.Back.Focus()
        } else {
            Save-Data -data @{ NewItems = @($state.Prepared) } -type Daily -dateString $state.Date.Value.ToString('yyyy-MM-dd') -ThrowOnError
            $state.Result = [PSCustomObject]@{ Date = $state.Date.Value.ToString('yyyy-MM-dd'); Items = @($state.Prepared) }
            $state.Dirty = $false; $state.Busy = $false
            $state.Form.DialogResult = [Windows.Forms.DialogResult]::OK
            $state.Form.Close()
        }
    } catch {
        $state.Status.Text = 'Nicht gespeichert: ' + $_.Exception.Message
        $state.Status.ForeColor = [Drawing.Color]::Firebrick
    } finally {
        $state.Busy = $false
        if (-not $state.Form.IsDisposed) { $state.Primary.Enabled = $true }
    }
}

function New-BtBloodValueForm {
    # Konstruktion getrennt vom modalen Anzeigen, damit die echte GUI isoliert pruefbar ist.
    param([datetime]$Date = (Get-Date))
    if ($null -ne $script:btBloodForm -and -not $script:btBloodForm.Form.IsDisposed) { throw 'Die Formulareingabe ist bereits geöffnet.' }
    Build-MarkerSearchIndex
    $form = New-Object Windows.Forms.Form; $form.Text = 'Mehrere Blutwerte erfassen'; $form.Size = New-Object Drawing.Size(1000, 660); $form.MinimumSize = New-Object Drawing.Size(920, 600); $form.StartPosition = 'CenterParent'; $form.MinimizeBox = $false
    $font = New-Object Drawing.Font('Segoe UI', 9); $form.Font = $font
    $hintFont = New-Object Drawing.Font('Segoe UI', 8)
    $dateLabel = New-Object Windows.Forms.Label; $dateLabel.Text = 'Datum des Bluttests:'; $dateLabel.Location = New-Object Drawing.Point(16, 18); $dateLabel.Size = New-Object Drawing.Size(135, 22)
    $picker = New-Object Windows.Forms.DateTimePicker; $picker.Format = 'Short'; $picker.Location = New-Object Drawing.Point(155, 15); $picker.Size = New-Object Drawing.Size(145, 24); $picker.Value = $Date; $picker.TabIndex = 0
    $info = New-Object Windows.Forms.Label; $info.Text = 'Ein gemeinsames Datum, eine Notiz je Zeile. Vollständig leere Zeilen werden ignoriert.'; $info.Location = New-Object Drawing.Point(16, 49); $info.Size = New-Object Drawing.Size(940, 22); $info.Anchor = 'Top,Left,Right'
    $entry = New-Object Windows.Forms.Panel; $entry.Location = New-Object Drawing.Point(16, 76); $entry.Size = New-Object Drawing.Size(952, 443); $entry.Anchor = 'Top,Bottom,Left,Right'; $entry.TabIndex = 1
    foreach ($column in @(@('Blutmarker',28,280),@('Wert / Auswahl',316,130),@('Einheit',454,82),@('Notiz',545,220))) {
        $label = New-Object Windows.Forms.Label; $label.Text = $column[0]; $label.Location = New-Object Drawing.Point($column[1],0); $label.Size = New-Object Drawing.Size($column[2],22); $entry.Controls.Add($label)
    }
    $flow = New-Object Windows.Forms.FlowLayoutPanel; $flow.Location = New-Object Drawing.Point(0, 25); $flow.Size = New-Object Drawing.Size(952, 418); $flow.Anchor = 'Top,Bottom,Left,Right'; $flow.AutoScroll = $true; $flow.FlowDirection = 'TopDown'; $flow.WrapContents = $false; $flow.TabIndex = 0; $entry.Controls.Add($flow)
    $review = New-Object Windows.Forms.Panel; $review.Location = $entry.Location; $review.Size = $entry.Size; $review.Anchor = $entry.Anchor; $review.Visible = $false
    $reviewTitle = New-Object Windows.Forms.Label; $reviewTitle.Location = New-Object Drawing.Point(0, 0); $reviewTitle.Size = New-Object Drawing.Size(950, 26); $reviewTitle.Anchor = 'Top,Left,Right'; $review.Controls.Add($reviewTitle)
    $grid = New-Object Windows.Forms.DataGridView; $grid.Location = New-Object Drawing.Point(0, 30); $grid.Size = New-Object Drawing.Size(952, 413); $grid.Anchor = 'Top,Bottom,Left,Right'; $grid.ReadOnly = $true; $grid.AllowUserToAddRows = $false; $grid.AllowUserToDeleteRows = $false; $grid.RowHeadersVisible = $false; $grid.AutoSizeColumnsMode = 'Fill'; $grid.AutoSizeRowsMode = 'AllCells'; $grid.DefaultCellStyle.WrapMode = 'True'; $grid.SelectionMode = 'FullRowSelect'; $grid.BackgroundColor = [Drawing.SystemColors]::Window
    foreach ($column in @(@('Marker',26),@('Eingabe',14),@('Gespeicherter Wert',17),@('Hinweis',23),@('Notiz',20))) { $c = New-Object Windows.Forms.DataGridViewTextBoxColumn; $c.HeaderText = $column[0]; $c.FillWeight = $column[1]; $c.SortMode = 'NotSortable'; $null = $grid.Columns.Add($c) }
    $review.Controls.Add($grid)
    $add = New-Object Windows.Forms.Button; $add.Text = '+ Weitere Zeile'; $add.Location = New-Object Drawing.Point(16, 532); $add.Size = New-Object Drawing.Size(150, 29); $add.Anchor = 'Bottom,Left'; $add.TabIndex = 2
    $status = New-Object Windows.Forms.Label; $status.Location = New-Object Drawing.Point(180, 527); $status.Size = New-Object Drawing.Size(786, 43); $status.Anchor = 'Bottom,Left,Right'; $status.Font = $hintFont
    $cancel = New-Object Windows.Forms.Button; $cancel.Text = 'Abbrechen'; $cancel.Location = New-Object Drawing.Point(16, 579); $cancel.Size = New-Object Drawing.Size(115, 30); $cancel.Anchor = 'Bottom,Left'; $cancel.TabIndex = 5
    $back = New-Object Windows.Forms.Button; $back.Text = 'Zurück zur Eingabe'; $back.Location = New-Object Drawing.Point(594, 579); $back.Size = New-Object Drawing.Size(150, 30); $back.Anchor = 'Bottom,Right'; $back.TabIndex = 3; $back.Visible = $false
    $primary = New-Object Windows.Forms.Button; $primary.Location = New-Object Drawing.Point(754, 579); $primary.Size = New-Object Drawing.Size(214, 30); $primary.Anchor = 'Bottom,Right'; $primary.TabIndex = 4
    $form.Controls.AddRange([Windows.Forms.Control[]]@($dateLabel,$picker,$info,$entry,$review,$add,$status,$cancel,$back,$primary))
    # Kein AcceptButton: Enter in einem Eingabefeld darf nicht speichern.
    $form.CancelButton = $cancel
    $state = [PSCustomObject]@{
        Form = $form; Date = $picker; EntryPanel = $entry; Flow = $flow; Rows = (New-Object Collections.ArrayList)
        MarkerNames = @($script:data.Config.Markers | ForEach-Object Name | Sort-Object -Unique)
        ReviewPanel = $review; Grid = $grid; ReviewTitle = $reviewTitle; Add = $add; Back = $back; Primary = $primary; Cancel = $cancel; Status = $status
        Font = $font; HintFont = $hintFont; Dirty = $false; Busy = $false; Resizing = $false; ResizePending = $false; Reviewing = $false; Prepared = @(); Result = $null
    }
    $form.Tag = $state; $script:btBloodForm = $state
    $picker.Add_ValueChanged({ $script:btBloodForm.Dirty = $true })
    $flow.Add_SizeChanged({
        Resize-BtBloodFormRows
        # Nach der laufenden WinForms-Groessenaenderung auch die Scroll-Ausdehnung
        # aktualisieren; innerhalb von SizeChanged kann sie noch die alte Breite haben.
        if ($args[0].IsHandleCreated -and -not $script:btBloodForm.ResizePending) {
            $script:btBloodForm.ResizePending = $true
            $null = $args[0].BeginInvoke([System.Action]{
                $state = $script:btBloodForm
                if ($null -eq $state -or $state.Form.IsDisposed) { return }
                $state.ResizePending = $false
                Resize-BtBloodFormRows
                $state.Flow.PerformLayout()
            })
        }
    })
    $add.Add_Click({ Add-BtBloodFormRow })
    $back.Add_Click({ Set-BtBloodFormReview -Visible $false })
    $primary.Add_Click({ Invoke-BtBloodFormPrimary })
    $cancel.Add_Click({ $script:btBloodForm.Form.Close() })
    $form.Add_FormClosing({
        $state = $args[0].Tag
        if ($state.Busy) { $args[1].Cancel = $true; return }
        if ($state.Dirty -and $null -eq $state.Result) {
            $answer = [Windows.Forms.MessageBox]::Show($args[0], 'Ungespeicherte Eingaben verwerfen?', 'Formular schließen', 'YesNo', 'Question', 'Button2')
            if ($answer -ne [Windows.Forms.DialogResult]::Yes) { $args[1].Cancel = $true }
        }
    })
    $form.Add_Disposed({ $args[0].Tag.HintFont.Dispose(); $args[0].Tag.Font.Dispose() })
    Set-BtBloodFormReview -Visible $false
    Add-BtBloodFormRow
    $state.Dirty = $false
    return $state
}

function Show-BloodValueEntryForm {
    param([Windows.Forms.IWin32Window]$Owner, [datetime]$Date = (Get-Date))
    $state = New-BtBloodValueForm -Date $Date
    try {
        if ($null -ne $Owner) { $null = $state.Form.ShowDialog($Owner) } else { $null = $state.Form.ShowDialog() }
        return $state.Result
    } finally { $state.Form.Dispose(); $script:btBloodForm = $null }
}


function Show-DataEntryForm {
    $form = New-Object System.Windows.Forms.Form; $form.Size = New-Object System.Drawing.Size(1200, 880); $form.StartPosition = "CenterScreen"; $form.Text = "Blutwert-Tracker - Version $Version"; $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    # Die bisherige Inhaltsflaeche bleibt auch bei kleinerem Fenster erreichbar.
    $navigationViewport = New-Object System.Windows.Forms.Panel
    $navigationViewport.Location = New-Object System.Drawing.Point(10, 40)
    $navigationViewport.Size = New-Object System.Drawing.Size(1160, 790)
    $navigationViewport.Anchor = 'Top, Bottom, Left, Right'
    $navigationViewport.AutoScroll = $true; $navigationViewport.TabIndex = 1
    $form.Controls.Add($navigationViewport)
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(0, 0)
    $tabControl.Size = New-Object System.Drawing.Size(1160, 790)
    $tabControl.Anchor = 'Top, Left'
    $tabControl.MinimumSize = New-Object System.Drawing.Size(1160, 790)
    $navigationViewport.Controls.Add($tabControl)
    $navigationViewport.Add_ClientSizeChanged({
        if ($args[0].Controls.Count -eq 0) { return }
        $content = $args[0].Controls[0]
        $size = New-Object System.Drawing.Size(([Math]::Max($content.MinimumSize.Width, $args[0].ClientSize.Width)), ([Math]::Max($content.MinimumSize.Height, $args[0].ClientSize.Height)))
        if ($content.Size -ne $size) { $content.Size = $size }
    })
    Initialize-BtActiveTabIndicator -TabControl $tabControl

    # v2.15.0/2.15.1: Form-globale Buttons in der Tab-Header-Zeile (aus JEDEM Tab sichtbar).
    # Y=8 damit die Buttons die obere Linie des TabControl-Rahmens nicht minimal überlappen.
    # BringToFront() damit sie optisch über dem TabControl liegen.
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Beenden"
    $closeButton.Size = New-Object System.Drawing.Size(90, 24)
    $closeButton.Location = New-Object System.Drawing.Point(($form.ClientSize.Width - $closeButton.Width - 20), 8)
    $closeButton.Anchor = "Top, Right"; $closeButton.TabIndex = 3
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
    $globalSettingsButton.Anchor = "Top, Right"; $globalSettingsButton.TabIndex = 2
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
        # v2.36.0: sichtbares Dashboard nachziehen (nur bei geaenderter Pruefsumme)
        if ($tabControl.SelectedTab -eq $tabPageDashboard) { try { Update-HealthDashboardTab } catch { Write-Warning "Dashboard: $($_.Exception.Message)" } }
    })
    $form.Controls.Add($globalSettingsButton)
    $globalSettingsButton.BringToFront()

    # v2.36.0: Reiter "Dashboard" (ERSTER Reiter). Inhalt wird im Form-Load aufgebaut, nachdem die
    # Daten geladen sind (siehe "v2.36.0: Reiter Dashboard aufbauen").
    $tabPageDashboard = New-Object System.Windows.Forms.TabPage; $tabPageDashboard.Text = "Dashboard"
    $tabPageDashboard.Tag = "Gesundheitsstatus auf einen Blick: Gesundheitsscore, Longevity-Werte, Vorbelastungen und Testosteron."
    $tabControl.TabPages.Add($tabPageDashboard)

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

    # v2.34.0: Reiter "Release-Notes" (letzter Reiter). Inhalt wird am Ende von Show-DataEntryForm
    # aufgebaut (siehe "TAB RELEASE-NOTES"), nachdem alle anderen Reiter fertig sind.
    $tabPageReleaseNotes = New-Object System.Windows.Forms.TabPage; $tabPageReleaseNotes.Text = "Release-Notes"
    $tabPageReleaseNotes.Tag = "Alle Versionen mit Datum und Änderungen – mit Suche und Filter nach Typ."
    $tabControl.TabPages.Add($tabPageReleaseNotes)
    # v2.34.0: Reiterköpfe dürfen "globale Einstellungen"/"Beenden" nicht überdecken (nur Darstellung)
    $navigationState = Initialize-BtCategoryNavigation -Form $form -Tabs $tabControl -SettingsButton $globalSettingsButton -Groups @(
        [PSCustomObject]@{ Name = 'Übersicht'; Pages = @($tabPageDashboard, $tabPage2) }
        [PSCustomObject]@{ Name = 'Auswertung'; Pages = @($tabPage1, $tabPage3, $tabPage5, $tabPageCustom) }
        [PSCustomObject]@{ Name = 'Daten & Profil'; Pages = @($tabPage4, $tabPage6) }
        [PSCustomObject]@{ Name = 'Info'; Pages = @($tabPageReleaseNotes) }
    )
    
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
    # v2.35.0: jede Spalte sortierbar (Klick: auf -> ab -> Standard); Standard = Bewertung, dann Bereich, Marker
    Register-BtGridSorting -Grid $cockpitGrid -Key 'cockpitGrid' -Modes @{ Group = 'Text'; Marker = 'Text'; Value = 'Auto'; Date = 'Date'; Risk = 'Number'; Unit = 'Text' } -RankColumns @{ Rating = 'SortOrder' } -DefaultColumn 'Rating' -TieBreak @('Group', 'Marker') -GroupColumn 'Group'
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

    $entryGroup = New-Object System.Windows.Forms.GroupBox; $entryGroup.Location = New-Object System.Drawing.Point(20, 20); $entryGroup.Size = New-Object System.Drawing.Size(420, 255); $entryGroup.Text = "Werte erfassen"; $tabPage1.Controls.Add($entryGroup)
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
    $apoeComboBox.Size = New-Object System.Drawing.Size(90, 20) # v2.47.0: Abstand zum Hinweistext
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
            $markerInfoLabel.Text = "Genotyp wählen  |  Referenz: kein E4-Allel (E2/E2, E2/E3, E3/E3)"
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
    $addButton = New-Object System.Windows.Forms.Button; $addButton.Location = New-Object System.Drawing.Point(110, 180); $addButton.Size = New-Object System.Drawing.Size(300, 30); $addButton.Text = "Einzelwert speichern"; $entryGroup.Controls.Add($addButton)
    $multiEntryButton = New-Object System.Windows.Forms.Button; $multiEntryButton.Location = New-Object System.Drawing.Point(110, 215); $multiEntryButton.Size = New-Object System.Drawing.Size(300, 30); $multiEntryButton.Text = 'Werte nach Blutbild speichern'; $entryGroup.Controls.Add($multiEntryButton)
    $markerMgmtGroup = New-Object System.Windows.Forms.GroupBox; $markerMgmtGroup.Location = New-Object System.Drawing.Point(20, 285); $markerMgmtGroup.Size = New-Object System.Drawing.Size(420, 220); $markerMgmtGroup.Text = "Blutmarker verwalten"; $tabPage1.Controls.Add($markerMgmtGroup)
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
    $warningThresholdGroup = New-Object System.Windows.Forms.GroupBox; $warningThresholdGroup.Location = New-Object System.Drawing.Point(20, 515); $warningThresholdGroup.Size = New-Object System.Drawing.Size(420, 70); $warningThresholdGroup.Text = "Schwellenwert für Warnungen"; $tabPage1.Controls.Add($warningThresholdGroup)
    $labelThreshold = New-Object System.Windows.Forms.Label; $labelThreshold.Text = "Warnen bei >"; $labelThreshold.Location = New-Object System.Drawing.Point(10, 30); $labelThreshold.Size = New-Object System.Drawing.Size(100, 20); $warningThresholdGroup.Controls.Add($labelThreshold)
    $thresholdTextBox = New-Object System.Windows.Forms.TextBox; $thresholdTextBox.Location = New-Object System.Drawing.Point(110, 28); $thresholdTextBox.Size = New-Object System.Drawing.Size(50, 20); $warningThresholdGroup.Controls.Add($thresholdTextBox)
    $labelPercent = New-Object System.Windows.Forms.Label; $labelPercent.Text = "% Abweichung / Jahr"; $labelPercent.Location = New-Object System.Drawing.Point(165, 30); $labelPercent.AutoSize = $true; $warningThresholdGroup.Controls.Add($labelPercent)
    $saveSettingsButton = New-Object System.Windows.Forms.Button; $saveSettingsButton.Location = New-Object System.Drawing.Point(290, 26); $saveSettingsButton.Size = New-Object System.Drawing.Size(110, 28); $saveSettingsButton.Text = "Speichern"; $warningThresholdGroup.Controls.Add($saveSettingsButton)
    # Beenden-Button wird auf $form-Ebene platziert (siehe Form-Setup unten, v2.15.0)
    $filterLabel = New-Object System.Windows.Forms.Label; $filterLabel.Text = "Verlauf anzeigen für:"; $filterLabel.Location = New-Object System.Drawing.Point(460, 20); $filterLabel.AutoSize = $true; $tabPage1.Controls.Add($filterLabel)
    $filterComboBox = New-Object System.Windows.Forms.ComboBox; $filterComboBox.Location = New-Object System.Drawing.Point(600, 18); $filterComboBox.Size = New-Object System.Drawing.Size(250, 20); $tabPage1.Controls.Add($filterComboBox)

    # ---- v2.32.0: Suche per Eingabe in "Verlauf anzeigen für" (wie im Feld "Blutmarker") ----
    $filterComboBox.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDown
    $filterComboBox.AutoCompleteMode = [System.Windows.Forms.AutoCompleteMode]::None
    $filterComboBox.AutoCompleteSource = [System.Windows.Forms.AutoCompleteSource]::None
    $filterComboBox.MaxDropDownItems = 15
    $toolTip.SetToolTip($filterComboBox, "Marker aus der Liste wählen ODER Namen/Kürzel tippen (z. B. 'HbA1c', 'Leuk', 'PREVENT'). Enter übernimmt den Treffer, Esc stellt die bisherige Auswahl wieder her.")
    $script:filterBusy          = $false   # verhindert Event-Rekursion und Neuzeichnen während des Tippens
    $script:filterFullList      = @()      # vollständige Liste (inkl. "--- Bitte auswählen ---")
    $script:filterLastSelection = $null    # zuletzt angezeigter Marker

    function Set-FilterComboItems {
        param([string[]]$Items)
        $previousBusy = $script:filterBusy
        $script:filterBusy = $true
        $filterComboBox.BeginUpdate()
        $filterComboBox.Items.Clear()
        if ($Items -and $Items.Count -gt 0) { $filterComboBox.Items.AddRange([object[]]$Items) }
        $filterComboBox.EndUpdate()
        $script:filterBusy = $previousBusy
    }

    function Get-FilterSearchItems {
        return @($script:filterFullList | Where-Object { $_ -and $_ -ne "--- Bitte auswählen ---" })
    }

    function Restore-FilterSelection {
        # Volle Liste wiederherstellen und $Name auswählen - ohne den Verlauf neu zu zeichnen
        param([string]$Name)
        $script:filterBusy = $true
        try {
            if ($filterComboBox.DroppedDown) { $filterComboBox.DroppedDown = $false }
            Set-FilterComboItems -Items $script:filterFullList
            $restoreIndex = if ($Name) { $filterComboBox.Items.IndexOf($Name) } else { -1 }
            if ($restoreIndex -lt 0 -and $filterComboBox.Items.Count -gt 0) { $restoreIndex = 0 }
            $filterComboBox.SelectedIndex = $restoreIndex
            $filterComboBox.BackColor = [System.Drawing.SystemColors]::Window
        } finally {
            $script:filterBusy = $false
        }
    }

    function Commit-FilterSelection {
        # Eingabe auflösen; $true = übernommen (oder leer), $false = kein eindeutiger Treffer
        $typedText = "$($filterComboBox.Text)".Trim()
        if ([string]::IsNullOrWhiteSpace($typedText)) { Restore-FilterSelection -Name $script:filterLastSelection; return $true }
        $target = Resolve-ListSelection -Query $typedText -Items (Get-FilterSearchItems)
        if (-not $target) { return $false }
        Restore-FilterSelection -Name $target
        if ($target -ne $script:filterLastSelection) {
            $script:filterLastSelection = $target
            Update-Chart
        }
        return $true
    }

    # Live-Filter während der Eingabe (TextUpdate feuert nur bei Nutzer-Tippen)
    $filterComboBox.Add_TextUpdate({
        if ($script:filterBusy) { return }
        $typedText = "$($filterComboBox.Text)"
        $script:filterBusy = $true
        try {
            if ([string]::IsNullOrWhiteSpace($typedText)) {
                Set-FilterComboItems -Items $script:filterFullList
                $filterComboBox.SelectedIndex = -1
                $filterComboBox.Text = ""
                $filterComboBox.BackColor = [System.Drawing.SystemColors]::Window
                return
            }
            $filterHits = @(Get-ListMatches -Query $typedText -Items (Get-FilterSearchItems))
            if ($filterHits.Count -eq 0) {
                Set-FilterComboItems -Items $script:filterFullList
                $filterComboBox.BackColor = [System.Drawing.Color]::MistyRose
            } else {
                Set-FilterComboItems -Items $filterHits
                $filterComboBox.BackColor = [System.Drawing.SystemColors]::Window
            }
            $filterComboBox.SelectedIndex = -1
            $filterComboBox.Text = $typedText
            if ($filterHits.Count -gt 0 -and -not $filterComboBox.DroppedDown) {
                $filterComboBox.DroppedDown = $true
                [System.Windows.Forms.Cursor]::Current = [System.Windows.Forms.Cursors]::Default
            }
            $filterComboBox.SelectionStart = $typedText.Length
            $filterComboBox.SelectionLength = 0
        } finally {
            $script:filterBusy = $false
        }
    })

    $filterComboBox.Add_KeyDown({
        $filterKeyArgs = $args[1]   # (Sender, KeyEventArgs) - ohne param(), Sender wird nicht benötigt
        if ($filterKeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            $filterKeyArgs.Handled = $true
            $filterKeyArgs.SuppressKeyPress = $true
            if (-not (Commit-FilterSelection)) { $filterComboBox.BackColor = [System.Drawing.Color]::MistyRose }
        } elseif ($filterKeyArgs.KeyCode -eq [System.Windows.Forms.Keys]::Escape) {
            $filterKeyArgs.Handled = $true
            $filterKeyArgs.SuppressKeyPress = $true
            Restore-FilterSelection -Name $script:filterLastSelection
        }
    })

    $filterComboBox.Add_Leave({
        if ($script:filterBusy) { return }
        if (-not (Commit-FilterSelection)) { Restore-FilterSelection -Name $script:filterLastSelection }
    })

    # Aufklappen per Pfeil nach einer Auswahl aus der gefilterten Liste: wieder alle Marker zeigen
    $filterComboBox.Add_DropDown({
        if ($script:filterBusy) { return }
        $currentSelection = [string]$filterComboBox.SelectedItem
        if ($currentSelection -and $filterComboBox.Items.Count -lt @($script:filterFullList).Count -and "$($filterComboBox.Text)" -eq $currentSelection) {
            $script:filterBusy = $true
            try {
                Set-FilterComboItems -Items $script:filterFullList
                $filterComboBox.SelectedIndex = $filterComboBox.Items.IndexOf($currentSelection)
            } finally {
                $script:filterBusy = $false
            }
        }
    })
    $chart = New-Object System.Windows.Forms.DataVisualization.Charting.Chart; $chart.Location = New-Object System.Drawing.Point(460, 50); $chart.Size = New-Object System.Drawing.Size(680, 570); $chart.Anchor = "Top, Left"; $tabPage1.Controls.Add($chart)
    # v2.15.1: $settingsButton entfernt - wurde zu globalem "globale Einstellungen"-Button
    # auf Form-Ebene (siehe Show-DataEntryForm). Bearbeiten/Drucken rücken an die frei
    # gewordene Position rechts im Chart-Bereich.
    $editButton = New-Object System.Windows.Forms.Button; $editButton.Text = "Bearbeiten"; $editButton.Size = New-Object System.Drawing.Size(90, 24); $editButton.Location = New-Object System.Drawing.Point(($chart.Right - $editButton.Width), 17); $tabPage1.Controls.Add($editButton)
    $chartPrintButton = New-Object System.Windows.Forms.Button; $chartPrintButton.Text = "Drucken"; $chartPrintButton.Size = New-Object System.Drawing.Size(90, 24); $chartPrintButton.Location = New-Object System.Drawing.Point(($editButton.Left - $chartPrintButton.Width - 8), 17); $tabPage1.Controls.Add($chartPrintButton)
    $chartExportButton = New-Object System.Windows.Forms.Button; $chartExportButton.Text = "Exportieren..."; $chartExportButton.Size = $chartPrintButton.Size; $chartExportButton.Location = New-Object System.Drawing.Point(($chartPrintButton.Left - $chartExportButton.Width - 8), $chartPrintButton.Top); $tabPage1.Controls.Add($chartExportButton)
    $chartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea; $chartArea.AxisX.Title = "Datum"; $chartArea.AxisX.LabelStyle.Format = "dd.MM.yy"; $chartArea.AxisY.Title = "Wert"; $chartArea.AxisY.MajorGrid.LineColor = [System.Drawing.Color]::LightGray; $chart.ChartAreas.Add($chartArea); $chart.Legends.Clear()
    $series = New-Object System.Windows.Forms.DataVisualization.Charting.Series; $series.Name = "Verlauf"; $series.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime; $series.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line; $series.BorderWidth = 3; $series.MarkerStyle = [System.Windows.Forms.DataVisualization.Charting.MarkerStyle]::Circle; $series.MarkerSize = 8; $chart.Series.Add($series)
    $trendSeries = New-Object System.Windows.Forms.DataVisualization.Charting.Series; $trendSeries.Name = "Trend"; $trendSeries.XValueType = [System.Windows.Forms.DataVisualization.Charting.ChartValueType]::DateTime; $trendSeries.ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line; $trendSeries.BorderWidth = 2; $trendSeries.Color = [System.Drawing.Color]::MediumPurple; $trendSeries.BorderDashStyle = "Dash"; $chart.Series.Insert(0, $trendSeries) # v2.26.1: Trend UNTER der Verlaufslinie (Messpunkte bleiben oben und per Maus erreichbar)
    $warningDisplayGroup = New-Object System.Windows.Forms.GroupBox; $warningDisplayGroup.Location = New-Object System.Drawing.Point($chart.Left, ($chart.Bottom + 8)); $warningDisplayGroup.Size = New-Object System.Drawing.Size($chart.Width, 122); $warningDisplayGroup.Text = "Hinweis/e"; $warningDisplayGroup.Anchor = "Top, Left"; $warningDisplayGroup.Visible = $false; $tabPage1.Controls.Add($warningDisplayGroup)
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
    $personalMetricsGroup = New-Object System.Windows.Forms.GroupBox; $personalMetricsGroup.Text = "Persönliche & biometrische Daten"; $personalMetricsGroup.Location = New-Object System.Drawing.Point(20, 20); $personalMetricsGroup.Size = New-Object System.Drawing.Size(500, 570); $personalMetricsGroup.Anchor = 'Top, Left'; $tabPage6.Controls.Add($personalMetricsGroup)
    
    $controlsYPos = 30; $labelWidth = 200; $controlWidth = 250; $yStep = 35
    
    # --- Name ---
    $labelName = New-Object System.Windows.Forms.Label; $labelName.Text = "Vor- und Nachname:"; $labelName.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelName.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelName)
    $textName = New-Object System.Windows.Forms.TextBox; $textName.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textName.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($textName); $controlsYPos += $yStep

    # --- Geburtsdatum (v2.29.0: ersetzt das feste Alter; Alter wird je Testtag berechnet) ---
    $labelGeburtsdatum = New-Object System.Windows.Forms.Label; $labelGeburtsdatum.Text = "Geburtsdatum (TT.MM.JJJJ):"; $labelGeburtsdatum.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelGeburtsdatum.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelGeburtsdatum)
    $textGeburtsdatum = New-Object System.Windows.Forms.TextBox; $textGeburtsdatum.Location = New-Object System.Drawing.Point(230, $controlsYPos); $textGeburtsdatum.Size = New-Object System.Drawing.Size(110, 20); $textGeburtsdatum.MaxLength = 10; $personalMetricsGroup.Controls.Add($textGeburtsdatum)
    $labelAlterAnzeige = New-Object System.Windows.Forms.Label; $labelAlterAnzeige.Location = New-Object System.Drawing.Point(348, ($controlsYPos + 3)); $labelAlterAnzeige.Size = New-Object System.Drawing.Size(140, 20); $labelAlterAnzeige.ForeColor = [System.Drawing.Color]::DimGray; $personalMetricsGroup.Controls.Add($labelAlterAnzeige); $controlsYPos += $yStep

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
    $checkRaucher = New-Object System.Windows.Forms.CheckBox; $checkRaucher.Text = "Raucher"; $checkRaucher.Location = New-Object System.Drawing.Point(230, $controlsYPos); $checkRaucher.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($checkRaucher); $controlsYPos += 25

    # --- v2.27.0: PREVENT-Angaben (Medikation, Diabetes, Anzeige im Cockpit) ---
    $checkBlutdrucksenker = New-Object System.Windows.Forms.CheckBox; $checkBlutdrucksenker.Text = "Nehme Blutdrucksenker"; $checkBlutdrucksenker.Location = New-Object System.Drawing.Point(230, $controlsYPos); $checkBlutdrucksenker.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($checkBlutdrucksenker); $controlsYPos += 25
    $checkLipidsenker = New-Object System.Windows.Forms.CheckBox; $checkLipidsenker.Text = "Nehme Lipidsenker (z. B. Statin)"; $checkLipidsenker.Location = New-Object System.Drawing.Point(230, $controlsYPos); $checkLipidsenker.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($checkLipidsenker); $controlsYPos += 25
    $checkDiabetes = New-Object System.Windows.Forms.CheckBox; $checkDiabetes.Text = "Diabetes diagnostiziert"; $checkDiabetes.Location = New-Object System.Drawing.Point(230, $controlsYPos); $checkDiabetes.Size = New-Object System.Drawing.Size($controlWidth, 20); $personalMetricsGroup.Controls.Add($checkDiabetes); $controlsYPos += $yStep
    $labelPreventAnzeige = New-Object System.Windows.Forms.Label; $labelPreventAnzeige.Text = "PREVENT-Anzeige (Cockpit):"; $labelPreventAnzeige.Location = New-Object System.Drawing.Point(20, $controlsYPos); $labelPreventAnzeige.Size = New-Object System.Drawing.Size($labelWidth, 20); $personalMetricsGroup.Controls.Add($labelPreventAnzeige)
    $comboPreventAnzeige = New-Object System.Windows.Forms.ComboBox; $comboPreventAnzeige.Location = New-Object System.Drawing.Point(230, $controlsYPos); $comboPreventAnzeige.Size = New-Object System.Drawing.Size($controlWidth, 20); $comboPreventAnzeige.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; [void]$comboPreventAnzeige.Items.AddRange(@("10 Jahre ASCVD (Standard)", "10 Jahre Herz-Kreislauf gesamt")); $comboPreventAnzeige.SelectedIndex = 0; $personalMetricsGroup.Controls.Add($comboPreventAnzeige); $controlsYPos += ($yStep + 20)
    
    # --- Speicherbutton ---
    $savePersonalButton = New-Object System.Windows.Forms.Button; $savePersonalButton.Text = "Persönliche Daten speichern"; $savePersonalButton.Location = New-Object System.Drawing.Point(230, $controlsYPos); $savePersonalButton.Size = New-Object System.Drawing.Size($controlWidth, 40); $personalMetricsGroup.Controls.Add($savePersonalButton)
    
    # --- Tooltips ---
    $toolTip.SetToolTip($textGroesse, "Ihre Körpergröße in Zentimetern. Wird zur Berechnung des BMI benötigt.")
    # v2.29.0: Geburtsdatum
    $toolTip.SetToolTip($textGeburtsdatum, "Geburtsdatum im Format TT.MM.JJJJ. Das Alter wird für jeden Bluttest am jeweiligen Testtag berechnet (vollendete Lebensjahre) - für PhenoAge, PREVENT und eGFR.")
    $toolTip.SetToolTip($textBlutdruck, "Geben Sie hier Ihren typischen systolischen Blutdruck (oberer Wert) in mmHg ein. Dieser Wert ist für die Risikoberechnung (PREVENT-Score) erforderlich.")
    # v2.27.0: PREVENT-Hilfetexte
    $toolTip.SetToolTip($checkBlutdrucksenker, "PREVENT: Werden blutdrucksenkende Medikamente eingenommen? Fließt in die Risikoberechnung ein.")
    $toolTip.SetToolTip($checkLipidsenker, "PREVENT: Werden Statine oder andere Lipidsenker eingenommen? Fließt in die Risikoberechnung ein.")
    $toolTip.SetToolTip($checkDiabetes, "PREVENT: Ärztlich diagnostizierter Diabetes. Unabhängig davon zählt ein HbA1c ab 6,5 % am Testtag als Diabetes.")
    $toolTip.SetToolTip($comboPreventAnzeige, "Welcher PREVENT-Wert im Risiko-Cockpit erscheint: 10-Jahres-Risiko für Herzinfarkt/Schlaganfall (ASCVD, Standard) oder für Herz-Kreislauf gesamt inkl. Herzinsuffizienz. Im Tab 'Einzelmarker-Analyse' sind beide wählbar.")

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
    # v2.35.0: jede Spalte sortierbar (Klick: auf -> ab -> Standard); Standard = feste Marker-Reihenfolge
    Register-BtGridSorting -Grid $longevityGrid -Key 'longevityGrid' -Modes @{ Marker = 'Text'; Value = 'Auto'; Ideal = 'Auto'; Score = 'Number'; Interpretation = 'Text' } -DefaultColumn '__Order' -TieBreak @('Marker')
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
        # v2.32.0: Grundlage für die Suche per Eingabe
        $script:filterFullList = @($filterComboBox.Items | ForEach-Object { [string]$_ })
        $script:filterLastSelection = [string]$filterComboBox.SelectedItem
        $filterComboBox.BackColor = [System.Drawing.SystemColors]::Window
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

        # v2.27.0: Welcher PREVENT-Wert erscheint im Cockpit? (Persönliche Metriken > PREVENT-Anzeige)
        $preventShowCvd = ([string]$script:data.Config.Personal.PreventAnzeige -eq 'CVD')
        $preventShownName  = if ($preventShowCvd) { "PREVENT-CVD-10Y" } else { "PREVENT-ASCVD-10Y" }
        $preventHiddenName = if ($preventShowCvd) { "PREVENT-ASCVD-10Y" } else { "PREVENT-CVD-10Y" }

        $preventScoreData = Get-CalculatedValuesForMarker -markerName $preventShownName
        if ($cutoffDate) { $preventScoreData = @($preventScoreData | Where-Object { [datetime]$_.Date -ge $cutoffDate }) }
        $preventScoreData = $preventScoreData | Sort-Object Date -Descending | Select-Object -First 1

        foreach ($marker in $allMarkers) {
            $latestItem = $null
            $isCalculated = $marker.PSObject.Properties.Name -contains "RequiredMarkers"
            if ($marker.Name -eq $preventHiddenName) { continue }   # v2.27.0: nur der gewählte PREVENT-Wert

            if ($isCalculated) {
                if ($marker.Name -eq $preventShownName) { $latestItem = $preventScoreData }
                else {
                    # v2.31.3: @() - ein Einzelergebnis hat in Windows PowerShell 5.1 kein .Count
                    $calculatedData = @(Get-CalculatedValuesForMarker -markerName $marker.Name)
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
                $row.Cells["Group"].Tag = $marker.Group   # v2.35.0: echter Name (Sortieren / Ausblenden)
                $row.Tag = $rowIndex                       # v2.35.0: Einfügereihenfolge (Gleichstand)
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
                
                if ($marker.Name -eq $preventShownName -and $latestItem) {
                    $row.Cells["Risk"].Value = $latestItem.Value
                    $riskVal = $latestItem.Value
                    $riskColor = [System.Drawing.Color]::White
                    # v2.27.0: PREVENT-Kategorien - ASCVD: 3 / 5 / 10 %, Herz-Kreislauf gesamt: 5 / 7,5 / 20 %
                    if ($preventShowCvd) { $riskLimits = @(5, 7.5, 20) } else { $riskLimits = @(3, 5, 10) }
                    if ($riskVal -lt $riskLimits[0]) { $riskColor = [System.Drawing.Color]::LightGreen }
                    elseif ($riskVal -lt $riskLimits[1]) { $riskColor = [System.Drawing.Color]::LightYellow }
                    elseif ($riskVal -lt $riskLimits[2]) { $riskColor = [System.Drawing.Color]::Orange }
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

        # v2.35.0: gewählte Sortierung bzw. Standard (Bewertung, Bereich, Marker); Funktionsbereich nur
        # in der ersten Zeile einer Gruppe (ersetzt Sort + Ausblend-Schleife bis v2.34.1)
        Invoke-BtGridApplySort -Grid $cockpitGrid
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
            @{ Name = "Longevity-Score"; Ideal = "> 80 Punkte"; Interpretation = "Mittelwert der Einzel-Scores (HbA1c, Insulin, Vitamin D, Homocystein, CRP/hs-CRP, NAD+, Albumin) am selben Testtag - ab 4 Werten." },
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
                    $score = [Math]::Round($latestItem.Value, 0, [MidpointRounding]::AwayFromZero)   # v2.31.0: kaufmännisch
                } elseif ($calculatedOnlyMarkers -contains $markerName) {
                    # Eigene Score-Berechnung für neue Longevity-Marker (0-100 Skala)
                    $score = Convert-ToLongevityScore -markerName $markerName -value $latestItem.Value -date $latestItem.Date
                } else {
                    # v2.30.0: Bewertung an den eigenen Bereichen des Markers (wie die Ampel im Cockpit)
                    $markerScore = Get-LongevityMarkerScore -MarkerName $markerName -Value ([double]$latestItem.Value)
                    if ($null -ne $markerScore) { $score = [Math]::Round($markerScore, 0, [MidpointRounding]::AwayFromZero) }   # v2.31.0: kaufmännisch (64,5 -> 65)
                }
                # v2.30.0: Spalte "Ideal" für Laborwerte aus dem Optimalbereich der Marker-Einstellungen
                $idealText = $markerInfo.Ideal
                if ($calculatedOnlyMarkers -notcontains $markerName) { $idealText = Get-OptimalRangeText -MarkerName $markerName -Fallback $markerInfo.Ideal }
                
                $rowIndex = $longevityGrid.Rows.Add($markerName, "$($latestItem.Value) $($latestItem.Unit)", $idealText, $score, $markerInfo.Interpretation)
                $row = $longevityGrid.Rows[$rowIndex]
                $row.Tag = $rowIndex   # v2.35.0: Standardreihenfolge (Einfügereihenfolge)
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

        # v2.35.0: gewählte Sortierung bzw. Standardreihenfolge wiederherstellen
        Invoke-BtGridApplySort -Grid $longevityGrid

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

    function Update-BirthDateAgeLabel {
        # v2.29.0: Heutiges Alter neben dem Geburtsdatum anzeigen (bzw. Hinweis)
        $birthPreview = ConvertTo-BtBirthDate -Text $textGeburtsdatum.Text
        if ($null -ne $birthPreview -and $birthPreview -le (Get-Date).Date) {
            $agePreview = Get-AgeAtDate -Personal @{ Geburtsdatum = $birthPreview.ToString('yyyy-MM-dd') }
            $labelAlterAnzeige.ForeColor = [System.Drawing.Color]::DimGray
            $labelAlterAnzeige.Text = "= $agePreview Jahre (heute)"
        } elseif ([string]::IsNullOrWhiteSpace($textGeburtsdatum.Text)) {
            $labelAlterAnzeige.ForeColor = [System.Drawing.Color]::Firebrick
            $fixedAgeText = [string]$script:data.Config.Personal.Age
            $labelAlterAnzeige.Text = if ($fixedAgeText) { "bisher fest: $fixedAgeText J." } else { "bitte eintragen" }
        } else {
            $labelAlterAnzeige.ForeColor = [System.Drawing.Color]::Firebrick
            $labelAlterAnzeige.Text = "Format TT.MM.JJJJ"
        }
    }

    function Load-PersonalMetrics {
        $personal = $script:data.Config.Personal
        $textName.Text = if ($personal.Name) { $personal.Name } else { "" }
        # v2.29.0: Geburtsdatum (Anzeige TT.MM.JJJJ); ohne Geburtsdatum Hinweis auf das bisherige feste Alter
        $birthLoaded = $null
        if ($personal.Geburtsdatum -is [datetime]) { $birthLoaded = $personal.Geburtsdatum.Date }
        elseif ($personal.Geburtsdatum) { $birthLoaded = ConvertTo-BtBirthDate -Text ([string]$personal.Geburtsdatum) }
        $textGeburtsdatum.Text = if ($birthLoaded) { $birthLoaded.ToString('dd.MM.yyyy') } else { "" }
        Update-BirthDateAgeLabel
        $comboGeschlecht.SelectedItem = $personal.Geschlecht
        $textGroesse.Text = $personal.Groesse
        $textGewicht.Text = $personal.Gewicht
        $textKFA.Text = $personal.KFA
        $textGrundumsatz.Text = $personal.Grundumsatz
        $comboWorkout.SelectedItem = $personal.WorkoutFreq
        $textBlutdruck.Text = $personal.SystBlutdruck
        $checkRaucher.Checked = $personal.Raucher
        # v2.27.0: PREVENT-Angaben
        $checkBlutdrucksenker.Checked = [bool]$personal.Blutdrucksenker
        $checkLipidsenker.Checked = [bool]$personal.Lipidsenker
        $checkDiabetes.Checked = [bool]$personal.DiabetesDiagnose
        $comboPreventAnzeige.SelectedIndex = if ([string]$personal.PreventAnzeige -eq 'CVD') { 1 } else { 0 }

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
            # v2.29.0: Geburtsdatum statt festem Alter - das Alter wird je Testtag berechnet
            $birthDate = ConvertTo-BtBirthDate -Text $textGeburtsdatum.Text
            if ($null -eq $birthDate) { throw "Bitte ein gültiges Geburtsdatum im Format TT.MM.JJJJ eingeben." }
            if ($birthDate -gt (Get-Date).Date) { throw "Das Geburtsdatum liegt in der Zukunft." }
            $birthIso = $birthDate.ToString('yyyy-MM-dd')
            $age = Get-AgeAtDate -Personal @{ Geburtsdatum = $birthIso }
            if ($age -lt 18) { throw "Das Alter muss mindestens 18 sein." }
            if ($age -gt 120) { throw "Bitte das Geburtsdatum prüfen (Alter über 120 Jahre)." }
            $script:data.Config.Personal.Geburtsdatum = $birthIso
            $script:data.Config.Personal.Age = [int]$age    # heutiges Alter (Kompatibilität älterer Versionen)
            $script:data.Config.Personal.Geschlecht = $comboGeschlecht.SelectedItem
            $script:data.Config.Personal.Groesse = Parse-Number $textGroesse.Text
            $script:data.Config.Personal.Gewicht = Parse-Number $textGewicht.Text
            $script:data.Config.Personal.KFA = Parse-Number $textKFA.Text
            $script:data.Config.Personal.Grundumsatz = Parse-Number $textGrundumsatz.Text
            $script:data.Config.Personal.WorkoutFreq = $comboWorkout.SelectedItem
            $script:data.Config.Personal.SystBlutdruck = Parse-Number $textBlutdruck.Text
            $script:data.Config.Personal.Raucher = $checkRaucher.Checked
            # v2.27.0: PREVENT-Angaben (Felder existieren dank Load-Config auch in Alt-Configs)
            $script:data.Config.Personal.Blutdrucksenker = $checkBlutdrucksenker.Checked
            $script:data.Config.Personal.Lipidsenker = $checkLipidsenker.Checked
            $script:data.Config.Personal.DiabetesDiagnose = $checkDiabetes.Checked
            $script:data.Config.Personal.PreventAnzeige = if ($comboPreventAnzeige.SelectedIndex -eq 1) { 'CVD' } else { 'ASCVD' }

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
            $msg = "INFO: Der AIP wird aus Triglyceriden und HDL in mmol/l berechnet (mg/dl-Werte werden intern umgerechnet). Risiko: < 0,11 niedrig, 0,11-0,21 mittel, > 0,21 erhöht."   # v2.28.0
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
        # v2.31.1: @() - bei genau EINEM Wert lieferte Sort-Object ein Einzelobjekt ohne .Count
        # (Windows PowerShell 5.1) -> "Letzter Wert" fehlte bei neu erfassten Markern.
        $markerData = @($markerData | Sort-Object Date)

        foreach ($item in $markerData) { $series.Points.AddXY([datetime]$item.Date, $item.Value) }
        
        $regression = Calculate-LinearRegression -dataPoints $markerData
        # v2.26.1 BUGFIX: '>' war in PowerShell eine Datei-Umleitung (kein Vergleich) -
        # die Trendlinie wurde dadurch nie gezeichnet und im Arbeitsverzeichnis entstand
        # eine Datei "1". @() macht Count auch bei Einzelobjekten PS-5.1-sicher.
        if ($regression -and @($markerData).Count -gt 1) {
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
            $warningThresholdGroup.Location = New-Object System.Drawing.Point(20, 665)
        } else {
            $markerMgmtGroup.Height = 220
            $addNewMarkerButton.Location = New-Object System.Drawing.Point(110, 175)
            $deleteExistingMarkerButton.Location = New-Object System.Drawing.Point(265, 175)
            $warningThresholdGroup.Location = New-Object System.Drawing.Point(20, 515)
        }
    })
    $tabControl.Add_MouseMove({
        $text = ''
        for ($i = 0; $i -lt $args[0].TabCount; $i++) {
            if ($args[0].GetTabRect($i).Contains($args[1].Location)) { $text = [string]$args[0].TabPages[$i].Tag; break }
        }
        if ($toolTip.GetToolTip($args[0]) -ne $text) { $toolTip.SetToolTip($args[0], $text) }
    })
    $navigationState.Refresh = {
        if ($null -eq $tabControl.SelectedTab) { return }
        # v2.36.0: Dashboard - rechnet nur bei geaenderter Pruefsumme neu
        if ($tabControl.SelectedTab -eq $tabPageDashboard) { try { Update-HealthDashboardTab } catch { Write-Warning "Dashboard: $($_.Exception.Message)" } }
        if ($tabControl.SelectedTab.Text -eq "Einzelmarker-Analyse") { Update-FilterDropdown; Update-Chart }
        if ($tabControl.SelectedTab.Text -eq "Risiko-Cockpit") { Update-Dashboard }
        if ($tabControl.SelectedTab.Text -eq "Korrelationen") { Update-CorrelationDropdowns }
        if ($tabControl.SelectedTab.Text -eq "Daten nach Bluttests") { Update-DataManagementTab }
        # Persoenliche Eingaben bleiben erhalten; Laden erfolgt bei Start/Import.
        if ($tabControl.SelectedTab.Text -eq "Longevity-Indizes") { Update-LongevityTab }
        if ($tabControl.SelectedTab.Text -eq "Custom Report") { Update-CustomReportTab }
    }
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
    $filterComboBox.Add_SelectedIndexChanged({
        if ($script:filterBusy) { return }   # v2.32.0: kein Neuzeichnen während Suche/Listenaufbau
        $selectedFilterItem = [string]$filterComboBox.SelectedItem
        if ($selectedFilterItem) {
            $script:filterLastSelection = $selectedFilterItem
            $filterComboBox.BackColor = [System.Drawing.SystemColors]::Window
        }
        Update-Chart
    })
    $chart.Add_MouseMove({
        param($sender, $e)
        $result = $chart.HitTest($e.X, $e.Y)
        # v2.26.1: Trendlinie liefert keinen Messwert -> kein Messwert-Tooltip fuer die Trend-Serie
        if ($result.ChartElementType -eq 'DataPoint' -and $result.Series -and $result.Series.Name -ne 'Trend') {
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

    $multiEntryButton.Tag = [PSCustomObject]@{ DatePicker = $datePicker; Filter = $filterComboBox }
    $multiEntryButton.Add_Click({
        $result = $null
        try {
            $result = Show-BloodValueEntryForm -Owner $args[0].FindForm() -Date $args[0].Tag.DatePicker.Value
            if ($null -eq $result) { return }
            $script:allHistoricalItems = [System.Collections.ArrayList]@(Load-AllHistoricalItems)
            Update-FilterDropdown
            $args[0].Tag.Filter.SelectedItem = $result.Items[-1].Name
            Update-Chart
        } catch {
            $message = if ($null -ne $result) { 'Die Werte wurden gespeichert, aber die Ansicht konnte nicht vollständig aktualisiert werden: ' } else { 'Die Formulareingabe konnte nicht geöffnet werden: ' }
            [System.Windows.Forms.MessageBox]::Show($message + $_.Exception.Message, 'Formulareingabe', 'OK', 'Error') | Out-Null
        }
    })
    $addButton.Add_Click({ 
        try { 
            $selectedDateString = $datePicker.Value.ToString("yyyy-MM-dd"); 
            # v2.46.0: UI-Auswahl abschliessen, fachliche Pruefung zentral vorbereiten.
            Commit-MarkerSelection
            $inputArguments = @{
                DateString = $selectedDateString; MarkerQuery = $markerComboBox.Text
                ValueText = $valueTextBox.Text; Note = $noteTextBox.Text
                HivResult = $hivComboBox.SelectedItem; ApoeGenotype = $apoeComboBox.SelectedItem
                TestosteroneUnit = $testoUnitCombo.SelectedItem
                ExistingItems = $script:allHistoricalItems
            }
            $newItem = New-BtBloodValueItem @inputArguments

            Save-Data -data @{ NewItems = @($newItem) } -type "Daily" -dateString $selectedDateString -ThrowOnError
            
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
    $textGeburtsdatum.Add_TextChanged({ Update-BirthDateAgeLabel })   # v2.29.0

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

        # v2.36.0: Reiter Dashboard aufbauen (Daten sind jetzt geladen). Ein Fehler zeigt einen
        # Hinweis im Reiter und stoppt das Programm nicht.
        try {
            Initialize-HealthDashboardTab -TabPage $tabPageDashboard -ToolTip $toolTip `
                -ModelProvider { Get-HealthDashboardModel -Config $script:data.Config -Items $script:allHistoricalItems -CalcProvider { param($n) @(Get-CalculatedValuesForMarker -markerName $n) } } `
                -SignatureProvider { Get-HealthDashDataSignature -Config $script:data.Config -Items $script:allHistoricalItems } `
                -OnTileActivated {
                    param($MarkerName)
                    # Kachel -> Einzelmarker-Analyse; der Reiterwechsel baut die Markerliste neu auf
                    Select-BtNavigationPage -TabControl $tabControl -TabPage $tabPage1
                    $markerIndex = $filterComboBox.Items.IndexOf([string]$MarkerName)
                    # v2.37.0: Auch konfigurierte Marker ohne Messdaten gezielt oeffnen.
                    if ($markerIndex -lt 0 -and (Get-HdMarkerConfig -Config $script:data.Config -Name $MarkerName)) {
                        $markerIndex = $filterComboBox.Items.Add([string]$MarkerName)
                        $script:filterFullList = @($filterComboBox.Items | ForEach-Object { [string]$_ })
                    }
                    if ($markerIndex -ge 0 -and $filterComboBox.SelectedIndex -ne $markerIndex) { $filterComboBox.SelectedIndex = $markerIndex }
                } `
                -OnOpenSettings { Select-BtNavigationPage -TabControl $tabControl -TabPage $tabPage6 }
        } catch {
            $dashErrorText = $_.Exception.Message
            Write-Warning "Initialize-HealthDashboardTab: $dashErrorText"
            $tabPageDashboard.Controls.Clear()
            $dashErrorLabel = New-Object System.Windows.Forms.Label
            $dashErrorLabel.Text = "Dashboard konnte nicht angezeigt werden:`n$dashErrorText"
            $dashErrorLabel.Location = New-Object System.Drawing.Point(20, 20); $dashErrorLabel.Size = New-Object System.Drawing.Size(1000, 60)
            $dashErrorLabel.ForeColor = [System.Drawing.Color]::DarkRed
            $tabPageDashboard.Controls.Add($dashErrorLabel)
        }

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

    # ---------- TAB RELEASE-NOTES (v2.34.0) ----------
    # Eigenständige Funktionen auf oberster Ebene; ein Fehler beim Aufbau wird im Reiter
    # angezeigt und stoppt das Programm nicht.
    try {
        $null = Initialize-ReleaseNotesTab -TabPage $tabPageReleaseNotes -Notes $script:ReleaseNotes -CurrentVersion $Version -ToolTip $toolTip
    } catch {
        $rnErrorText = $_.Exception.Message
        Write-Warning "Initialize-ReleaseNotesTab: $rnErrorText"
        $tabPageReleaseNotes.Controls.Clear()
        $rnErrorLabel = New-Object System.Windows.Forms.Label
        $rnErrorLabel.Text = "Release-Notes konnten nicht angezeigt werden:`n$rnErrorText"
        $rnErrorLabel.Location = New-Object System.Drawing.Point(20, 20); $rnErrorLabel.Size = New-Object System.Drawing.Size(1000, 60)
        $rnErrorLabel.ForeColor = [System.Drawing.Color]::DarkRed
        $tabPageReleaseNotes.Controls.Add($rnErrorLabel)
    }
    # v2.34.0: BETA-Start -> Release-Notes-Konsistenz prüfen (eine Warnung, Programm läuft weiter)
    $form.Add_Shown({ $null = Invoke-ReleaseNotesStartCheck })

    $workingArea = [System.Windows.Forms.Screen]::FromPoint([System.Windows.Forms.Cursor]::Position).WorkingArea
    $form.MinimumSize = New-Object System.Drawing.Size(([Math]::Min(800, $workingArea.Width)), ([Math]::Min(500, $workingArea.Height)))
    $form.Size = New-Object System.Drawing.Size(([Math]::Min($form.Width, $workingArea.Width)), ([Math]::Min($form.Height, $workingArea.Height)))
    try { [void]$form.ShowDialog() } finally { $form.Dispose() }
}

# ---------- Skript-Einstiegspunkt ----------
Show-DataEntryForm
