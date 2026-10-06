#!/usr/bin/env python3
"""
maildir_extract.py - Telefonnummern und E-Mail-Adressen aus einem Maildir extrahieren.

Durchsucht rekursiv alle Mails in einem Maildir (inkl. Maildir++-Unterordner wie
.Sent, .Archive ...), dekodiert Header und Text-/HTML-Teile und schreibt jeden Fund
mit Referenz auf die Quelldatei in eine CSV-Datei.

Aufruf:
    python3 maildir_extract.py /pfad/zum/Maildir
    python3 maildir_extract.py ~/Maildir -o funde.csv --default-cc 49
    python3 maildir_extract.py ~/Maildir --loose --attachments --all-files

Nur Python-Standardbibliothek, Python >= 3.8.
"""

import sys

# ---------------------------------------------------------------------------
# Voraussetzungen prüfen: Python-Version und benötigte Module.
# Fehlt etwas, bricht das Script mit Exit-Code 2 ab, bevor es irgendetwas tut.
# ---------------------------------------------------------------------------

MIN_PYTHON = (3, 8)
REQUIRED_MODULES = ["argparse", "csv", "email", "email.header", "email.utils",
                    "html", "importlib", "os", "re"]

if sys.version_info < MIN_PYTHON:
    sys.stderr.write("FEHLER: Python %d.%d oder neuer erforderlich, gefunden: %s\n"
                     % (MIN_PYTHON[0], MIN_PYTHON[1], sys.version.split()[0]))
    sys.exit(2)


def _check_modules(names):
    import importlib
    missing = []
    for name in names:
        try:
            importlib.import_module(name)
        except ImportError as exc:
            missing.append((name, str(exc)))
    return missing


_missing = _check_modules(REQUIRED_MODULES)
if _missing:
    sys.stderr.write("FEHLER: Folgende Python-Module wurden nicht gefunden:\n")
    for _name, _err in _missing:
        sys.stderr.write("  - %s (%s)\n" % (_name, _err))
    sys.stderr.write("Bitte die vollständige Python-Standardbibliothek installieren "
                     "(z. B. Debian/Ubuntu: apt install python3 libpython3-stdlib, "
                     "FreeBSD: pkg install python3).\n")
    sys.exit(2)

import argparse  # noqa: E402
import csv  # noqa: E402
import email  # noqa: E402
import html  # noqa: E402
import os  # noqa: E402
import re  # noqa: E402
from email.header import decode_header, make_header  # noqa: E402
from email.utils import getaddresses  # noqa: E402

# ---------------------------------------------------------------------------
# Muster
# ---------------------------------------------------------------------------

EMAIL_RE = re.compile(
    r"(?<![\w.%+-])[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,24}(?![\w-])"
)

# Telefonnummern-Kandidaten: optionale Ländervorwahl (+49 / 0049), optional "(0)",
# optionale Vorwahl in Klammern, danach Ziffernblöcke getrennt durch
# Leerzeichen, '/', '-', '.'. Zeilenumbrüche werden bewusst nicht überbrückt.
PHONE_RE = re.compile(
    r"""
    (?<![\w@+])
    (
      (?:(?:\+|00)\d{1,3}(?:[ \t]?\(0\))?[ \t./-]{0,3})?   # Ländervorwahl, optional (0)
      (?:\(\d{1,6}\)[ \t]?)?                                # (0821)
      \d{1,6}
      (?:(?:[ \t]?[./-][ \t]?|[ \t])\d{1,8}){0,6}           # weitere Blöcke
    )
    (?![\w@])
    """,
    re.VERBOSE,
)

DATE_PATTERNS = [
    re.compile(r"\d{1,2}\.\d{1,2}\.\d{2,4}"),        # 06.10.2026
    re.compile(r"\d{4}-\d{1,2}-\d{1,2}"),            # 2026-10-06
    re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}"),          # 10/06/2026
    re.compile(r"\d{4}/\d{1,2}/\d{1,2}"),            # 2026/10/06
    re.compile(r"\d{1,3}(?:\.\d{1,3}){3}"),          # IPv4
]

TAG_RE = re.compile(r"<[^>]+>")
SCRIPT_STYLE_RE = re.compile(r"<(script|style)\b.*?</\1>", re.S | re.I)
BR_RE = re.compile(r"<\s*(br|/p|/div|/tr|/li)\b[^>]*>", re.I)
TEL_HREF_RE = re.compile(r"""href\s*=\s*["']?tel:([^"'\s>]+)""", re.I)

MAILDIR_SKIP = {"dovecot.index", "dovecot.index.log", "dovecot.index.cache",
                "dovecot-uidlist", "dovecot-keywords", "maildirfolder",
                "subscriptions", "courierimapuiddb", "courierimapkeywords"}

ADDR_HEADERS = ["From", "To", "Cc", "Bcc", "Reply-To", "Sender",
                "Return-Path", "Delivered-To", "Resent-From", "Resent-To"]


# ---------------------------------------------------------------------------
# Hilfsfunktionen
# ---------------------------------------------------------------------------

def decode_hdr(value):
    """MIME-kodierte Header (=?utf-8?...?=) in Klartext umwandeln."""
    if value is None:
        return ""
    try:
        return str(make_header(decode_header(str(value))))
    except Exception:
        return str(value)


def decode_payload(part):
    payload = part.get_payload(decode=True)
    if payload is None:
        return ""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except LookupError:
        return payload.decode("latin-1", errors="replace")


def html_to_text(raw):
    text = SCRIPT_STYLE_RE.sub(" ", raw)
    text = BR_RE.sub("\n", text)
    text = TAG_RE.sub(" ", text)
    return html.unescape(text)


def is_phone(candidate, min_digits, max_digits, loose):
    s = candidate.strip()
    digits = re.sub(r"\D", "", s)
    if not (min_digits <= len(digits) <= max_digits):
        return False
    for pat in DATE_PATTERNS:
        if pat.fullmatch(s):
            return False
    if not loose:
        # Konservativ: Telefonnummern beginnen mit +, 00, 0 oder (0...)
        if not re.match(r"(\+|0|\(0)", s):
            return False
        # Reine Ziffernketten ohne Trenner und ohne + sind meist IDs, Kontonummern o. ä.
        if s.isdigit() and len(s) > 12:
            return False
    return True


def normalize_phone(raw, default_cc=None):
    s = raw.strip()
    has_plus = s.startswith("+")
    s = s.replace("(0)", "")
    digits = re.sub(r"\D", "", s)
    if has_plus:
        return "+" + digits
    if digits.startswith("00"):
        return "+" + digits[2:]
    if default_cc and digits.startswith("0"):
        return "+" + str(default_cc) + digits[1:]
    return digits


def find_phones(text, args):
    for m in PHONE_RE.finditer(text):
        cand = m.group(1).rstrip(" \t./-")
        if is_phone(cand, args.min_digits, args.max_digits, args.loose):
            yield cand


def find_emails(text):
    for m in EMAIL_RE.finditer(text):
        yield m.group(0).strip(".")


def iter_mail_files(root, all_files):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        base = os.path.basename(dirpath)
        if not all_files and base not in ("cur", "new"):
            continue
        for fn in sorted(filenames):
            if fn in MAILDIR_SKIP or fn.startswith("."):
                continue
            yield os.path.join(dirpath, fn)


def iter_text_parts(msg, include_attachments):
    """Liefert (Fundstelle, Text, ist_html) für alle Textteile."""
    for part in msg.walk():
        if part.is_multipart():
            continue
        ctype = part.get_content_type()
        disp = (part.get("Content-Disposition") or "").lower()
        is_attachment = disp.startswith("attachment")
        if is_attachment and not include_attachments:
            continue
        if ctype not in ("text/plain", "text/html") and not (
                is_attachment and ctype.startswith("text/")):
            continue
        label = "body"
        if is_attachment:
            fname = decode_hdr(part.get_filename()) or "unbenannt"
            label = "anhang:" + fname
        elif ctype == "text/html":
            label = "body(html)"
        yield label, decode_payload(part), ctype == "text/html"


# ---------------------------------------------------------------------------
# Verarbeitung
# ---------------------------------------------------------------------------

def process_file(path, args):
    with open(path, "rb") as fh:
        msg = email.message_from_binary_file(fh)

    meta = {
        "datum": decode_hdr(msg.get("Date")),
        "von": decode_hdr(msg.get("From")),
        "betreff": decode_hdr(msg.get("Subject")),
        "message_id": decode_hdr(msg.get("Message-ID")).strip(),
    }

    seen = set()
    rows = []

    def add(typ, value, normalized, location):
        key = (typ, normalized)
        if key in seen and not args.all_occurrences:
            return
        seen.add(key)
        rows.append({"typ": typ, "wert": value, "normalisiert": normalized,
                     "fundstelle": location, **meta})

    # Adress-Header
    for hname in ADDR_HEADERS:
        values = [decode_hdr(v) for v in msg.get_all(hname, [])]
        if not values:
            continue
        for name, addr in getaddresses(values):
            addr = addr.strip("<> ")
            if EMAIL_RE.fullmatch(addr):
                add("email", addr, addr.lower(), "header:" + hname)

    # Betreff
    subject = meta["betreff"]
    for p in find_phones(subject, args):
        add("telefon", p, normalize_phone(p, args.default_cc), "header:Subject")
    for e in find_emails(subject):
        add("email", e, e.lower(), "header:Subject")

    # Body-Teile
    for label, text, is_html in iter_text_parts(msg, args.attachments):
        if is_html:
            for tel in TEL_HREF_RE.findall(text):
                tel = html.unescape(tel)
                add("telefon", tel, normalize_phone(tel, args.default_cc), label + " (tel:-Link)")
            for e in find_emails(text):  # erfasst auch mailto:-Links
                add("email", e, e.lower(), label)
            text = html_to_text(text)
        for p in find_phones(text, args):
            add("telefon", p, normalize_phone(p, args.default_cc), label)
        for e in find_emails(text):
            add("email", e, e.lower(), label)

    for r in rows:
        r["datei"] = os.path.abspath(path)
    return rows


FIELDS = ["typ", "wert", "normalisiert", "fundstelle", "datei",
          "datum", "von", "betreff", "message_id"]


def main():
    ap = argparse.ArgumentParser(
        description="Telefonnummern und E-Mail-Adressen aus einem Maildir in eine CSV exportieren.")
    ap.add_argument("maildir", help="Pfad zum Maildir (wird rekursiv durchsucht)")
    ap.add_argument("-o", "--output", default="maildir_funde.csv",
                    help="CSV-Ausgabedatei (Standard: maildir_funde.csv)")
    ap.add_argument("-d", "--delimiter", default=";",
                    help="CSV-Trennzeichen (Standard: ';' für Excel DE)")
    ap.add_argument("--default-cc", type=int, metavar="CC",
                    help="Ländervorwahl für nationale Nummern bei der Normalisierung, z. B. 49")
    ap.add_argument("--min-digits", type=int, default=7,
                    help="Mindestanzahl Ziffern einer Telefonnummer (Standard: 7)")
    ap.add_argument("--max-digits", type=int, default=15,
                    help="Höchstanzahl Ziffern einer Telefonnummer (Standard: 15)")
    ap.add_argument("--loose", action="store_true",
                    help="Weniger strenger Filter (auch Nummern ohne führendes +/0) – mehr Fehltreffer")
    ap.add_argument("--attachments", action="store_true",
                    help="Auch Text-Anhänge (text/*) durchsuchen")
    ap.add_argument("--all-files", action="store_true",
                    help="Alle Dateien durchsuchen, nicht nur cur/ und new/")
    ap.add_argument("--all-occurrences", action="store_true",
                    help="Jeden Fund ausgeben, statt pro Mail zu deduplizieren")
    args = ap.parse_args()

    if not os.path.isdir(args.maildir):
        sys.exit(f"Fehler: '{args.maildir}' ist kein Verzeichnis.")

    n_files = n_err = n_rows = 0
    with open(args.output, "w", newline="", encoding="utf-8-sig") as out:
        writer = csv.DictWriter(out, fieldnames=FIELDS, delimiter=args.delimiter)
        writer.writeheader()
        for path in iter_mail_files(args.maildir, args.all_files):
            n_files += 1
            try:
                rows = process_file(path, args)
            except Exception as exc:
                n_err += 1
                print(f"WARNUNG: {path}: {exc}", file=sys.stderr)
                continue
            writer.writerows(rows)
            n_rows += len(rows)
            if n_files % 1000 == 0:
                print(f"... {n_files} Mails, {n_rows} Funde", file=sys.stderr)

    print(f"Fertig: {n_files} Mails durchsucht, {n_rows} Funde, {n_err} Fehler -> {args.output}",
          file=sys.stderr)


if __name__ == "__main__":
    main()
