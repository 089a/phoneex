# maildir_extract.py

Extrahiert **Telefonnummern** und **E-Mail-Adressen** aus allen Mails eines Maildirs und exportiert jeden Fund mit Referenz auf die Quelldatei in eine CSV-Datei.

- Nur Python-Standardbibliothek, keine externen Abhängigkeiten
- Durchsucht rekursiv `cur/` und `new/` inkl. Maildir++-Unterordner (`.Sent`, `.Archive`, …)
- Dekodiert MIME-Header, Zeichensätze, Quoted-Printable/Base64, Text- und HTML-Teile
- Erkennt `tel:`- und `mailto:`-Links in HTML-Mails
- Liest die Mails nur, schreibt nie ins Maildir

---

## 1. Vorbereitung

### 1.1 Voraussetzungen

| Komponente | Anforderung |
|---|---|
| Python | ≥ 3.8 |
| Module | nur Standardbibliothek: `argparse`, `csv`, `email`, `html`, `importlib`, `os`, `re` |
| Rechte | Lesezugriff auf das Maildir, Schreibzugriff auf das Zielverzeichnis der CSV |
| Plattform | Linux, FreeBSD, macOS (alles mit POSIX-Dateisystem) |

Das Script prüft Python-Version und Module beim Start selbst. Fehlt etwas, bricht es mit **Exit-Code 2** ab, bevor eine Datei angelegt wird.

Python-Version prüfen:

```bash
python3 --version
```

Python installieren, falls nötig:

```bash
# Debian / Ubuntu
apt install python3 libpython3-stdlib
# FreeBSD
pkg install python3
```

### 1.2 Checkliste vor dem ersten Lauf

- [ ] Script ablegen, z. B. unter `/opt/maildir_extract/maildir_extract.py`, und ausführbar machen:
      `chmod 755 /opt/maildir_extract/maildir_extract.py`
- [ ] Pfad zum Maildir ermitteln. Typische Orte: `~/Maildir`, `/var/vmail/<domain>/<user>/Maildir`, `/var/mail/vhosts/<domain>/<user>`
- [ ] Prüfen, dass der ausführende Benutzer das Maildir lesen darf. Bei Dovecot/Postfix ist das meist `vmail`:
      `sudo -u vmail ls /var/vmail/example.de/user/Maildir/cur | head`
- [ ] Zielverzeichnis für die CSV anlegen und den Zugriff einschränken, denn die CSV enthält personenbezogene Daten:
      `install -d -m 700 /srv/export/maildir`
- [ ] Bei großen Maildirs oder produktiven Servern empfiehlt sich eine Kopie oder ein Snapshot, damit der Mailserver nicht belastet wird:
      `rsync -a --delete /var/vmail/example.de/user/Maildir/ /srv/work/Maildir-copy/`
- [ ] Testlauf auf einem kleinen Unterordner und Ergebnis prüfen (siehe 2.3)
- [ ] Datenschutz klären: Zweck, Rechtsgrundlage, Aufbewahrungsdauer und Löschung der CSV festlegen (DSGVO)

---

## 2. Verwendung

### 2.1 Aufruf

```bash
python3 maildir_extract.py <MAILDIR> [Optionen]
```

Beispiele:

```bash
# Standard: Ausgabe nach ./maildir_funde.csv
python3 maildir_extract.py ~/Maildir

# Eigene Ausgabedatei, deutsche Nummern ins internationale Format normalisieren
python3 maildir_extract.py ~/Maildir -o /srv/export/maildir/funde.csv --default-cc 49

# Gründlich: auch Text-Anhänge, alle Dateien, weniger strenger Telefonfilter
python3 maildir_extract.py ~/Maildir --attachments --all-files --loose

# Komma als Trennzeichen (z. B. für LibreOffice/Pandas)
python3 maildir_extract.py ~/Maildir -d ','
```

### 2.2 Optionen

| Option | Standard | Beschreibung |
|---|---|---|
| `MAILDIR` | – | Pfad zum Maildir (Pflicht) |
| `-o`, `--output` | `maildir_funde.csv` | CSV-Ausgabedatei |
| `-d`, `--delimiter` | `;` | CSV-Trennzeichen |
| `--default-cc CC` | – | Ländervorwahl für nationale Nummern, z. B. `49` macht aus `0821…` die Nummer `+49821…` |
| `--min-digits N` | `7` | Mindestanzahl Ziffern einer Telefonnummer |
| `--max-digits N` | `15` | Höchstanzahl Ziffern (E.164) |
| `--loose` | aus | Auch Nummern ohne führendes `+`/`0` erfassen. Bringt mehr Fehltreffer. |
| `--attachments` | aus | Text-Anhänge (`text/*`) durchsuchen |
| `--all-files` | aus | Alle Dateien durchsuchen, nicht nur `cur/` und `new/` |
| `--all-occurrences` | aus | Jedes Vorkommen ausgeben statt einmal pro Mail |

### 2.3 Testlauf

```bash
python3 maildir_extract.py ~/Maildir/.Sent -o /tmp/test.csv --default-cc 49
head -20 /tmp/test.csv
```

### 2.4 Exit-Codes

| Code | Bedeutung |
|---|---|
| `0` | Erfolgreich. Einzelne unlesbare Mails werden als Warnung gemeldet und übersprungen. |
| `1` | Maildir-Pfad existiert nicht oder ist kein Verzeichnis |
| `2` | Python-Version zu alt, Modul fehlt oder ungültige Parameter |

Fortschritt und Zusammenfassung erscheinen auf **stderr**, z. B.:

```
... 1000 Mails, 3412 Funde
Fertig: 1873 Mails durchsucht, 6120 Funde, 2 Fehler -> funde.csv
```

---

## 3. Ausgabeformat

CSV in UTF-8 mit BOM, damit Excel Umlaute korrekt anzeigt. Eine Zeile pro Fund:

| Spalte | Inhalt |
|---|---|
| `typ` | `telefon` oder `email` |
| `wert` | Fund wie im Text, z. B. `0821 / 12 34 56` |
| `normalisiert` | Vereinheitlicht, z. B. `+49821123456` oder `info@firma.de` (Kleinschreibung) |
| `fundstelle` | `header:From`, `header:To`, `header:Subject`, `body`, `body(html)`, `body(html) (tel:-Link)`, `anhang:<name>` |
| `datei` | Absoluter Pfad der Mail-Datei |
| `datum` | `Date`-Header |
| `von` | `From`-Header |
| `betreff` | `Subject`-Header |
| `message_id` | `Message-ID`-Header |

Innerhalb einer Mail wird jeder Wert nur einmal ausgegeben, verglichen über `normalisiert`. Wer jedes Vorkommen braucht, nutzt `--all-occurrences`.

### Auswertungen

```bash
# Eindeutige Telefonnummern über alle Mails
awk -F';' '$1=="telefon"{print $3}' funde.csv | sort -u

# Eindeutige E-Mail-Adressen mit Häufigkeit
awk -F';' '$1=="email"{print $3}' funde.csv | sort | uniq -c | sort -rn

# In welchen Mails kommt eine bestimmte Nummer vor?
grep ';+49821123456;' funde.csv | cut -d';' -f5
```

---

## 4. Erkennungslogik und Grenzen

**Telefonnummern** (Standardmodus):

- Erkannt werden u. a. `+49 (0) 821-123 456`, `0821/123456`, `(089) 98765-43`, `0043 1 234 5678` und `0171 1234567`.
- Die Nummer muss mit `+`, `00`, `0` oder `(0` beginnen und 7 bis 15 Ziffern haben.
- Verworfen werden Datumsangaben (`06.10.2026`, `2026-10-06`), IPv4-Adressen und Ziffernketten ohne Trenner mit mehr als 12 Stellen (Bestell- und Kontonummern).
- Nummern über Zeilenumbrüche hinweg werden nicht erkannt.

**E-Mail-Adressen:** aus den Adress-Headern (From, To, Cc, Bcc, Reply-To, Sender, Return-Path, Delivered-To, Resent-*), aus dem Betreff und aus dem Body.

**Grenzen:**

- Es handelt sich um eine Heuristik. Fehltreffer (z. B. Rechnungsnummern mit führender 0) und übersehene Formate sind möglich, daher Ergebnisse stichprobenartig prüfen.
- Binäre Anhänge (PDF, DOCX, Bilder) werden nicht durchsucht.
- Zitierte Mailverläufe werden mit ausgewertet, weshalb dieselbe Nummer in vielen Mails auftauchen kann.

---

## 5. Automatisierung

### 5.1 Cron

```cron
# Täglich 02:15 als vmail, mit Datum im Dateinamen
15 2 * * * vmail /usr/bin/python3 /opt/maildir_extract/maildir_extract.py /var/vmail/example.de/user/Maildir -o /srv/export/maildir/funde_$(date +\%F).csv --default-cc 49 2>>/var/log/maildir_extract.log
```

Alte Exporte aufräumen, z. B. nach 30 Tagen:

```cron
30 2 * * * root find /srv/export/maildir -name 'funde_*.csv' -mtime +30 -delete
```

### 5.2 Jenkins (Declarative Pipeline)

```groovy
pipeline {
    agent { label 'mailhost' }
    parameters {
        string(name: 'MAILDIR', defaultValue: '/var/vmail/example.de/user/Maildir', description: 'Pfad zum Maildir')
        string(name: 'DEFAULT_CC', defaultValue: '49', description: 'Ländervorwahl')
    }
    stages {
        stage('Check') {
            steps { sh 'python3 --version' }
        }
        stage('Extract') {
            steps {
                sh '''
                  python3 maildir_extract.py "$MAILDIR" \
                    -o "funde_${BUILD_NUMBER}.csv" \
                    --default-cc "$DEFAULT_CC"
                '''
            }
        }
    }
    post {
        success { archiveArtifacts artifacts: 'funde_*.csv', fingerprint: true }
        failure { echo 'Abbruch: Exit-Code 2 bedeutet, dass eine Voraussetzung fehlt (siehe Log).' }
    }
}
```

Der Jenkins-Agent braucht Lesezugriff auf das Maildir, z. B. über die Gruppe `vmail`. Archivierte CSVs enthalten personenbezogene Daten, deshalb Zugriff auf den Job einschränken.

---

## 6. Fehlerbehebung

| Symptom | Ursache / Lösung |
|---|---|
| `FEHLER: Python 3.8 oder neuer erforderlich` | Neueres Python installieren oder explizit aufrufen, z. B. `python3.11` |
| `FEHLER: Folgende Python-Module wurden nicht gefunden` | Standardbibliothek unvollständig, z. B. bei Minimal-Images. Paket nachinstallieren (siehe 1.1). |
| `0 Mails durchsucht` | Falscher Pfad, oder das Maildir hat keine `cur/`/`new/`-Ordner. Dann `--all-files` verwenden. |
| `WARNUNG: …: Permission denied` | Script als Besitzer des Maildirs ausführen (`sudo -u vmail …`) |
| Umlaute in Excel kaputt | CSV per Doppelklick öffnen (BOM ist gesetzt) oder in Excel über *Daten → Aus Text/CSV* mit UTF-8 importieren |
| Zu viele Fehltreffer | `--min-digits 8` setzen und `--loose` weglassen |
| Nummern fehlen | Mit `--loose` laufen lassen oder `--min-digits` senken, Formate prüfen |
