#!/usr/bin/env python3
"""Testy regresyjne terminarz-platnosci (uruchamiaj po generuj_fixtures.py).

Obejmuje: liczniki koszyków (rejestr-okno, rejestr-mieszany), pusty rejestr,
--okno 0, --z/--do złą datę, CSV z przecinkami, wielostronicowy PDF z
etykietą sekcji („— cd.”). Zero sieci, wszystko lokalnie, pliki testowe
w katalogu tymczasowym.
"""
import csv
import os
import re
import subprocess
import sys
import tempfile
from datetime import date, timedelta

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
NARZ = os.path.join(REPO, "terminarz_platnosci.py")
PY = sys.executable
NAGL = ["plik", "sprzedawca", "nip_sprzedawcy", "nr_faktury",
        "data_wystawienia", "termin_platnosci", "kwota_brutto", "komentarz"]

błędy = []


def uruchom(args):
    p = subprocess.run([PY, NARZ] + args, capture_output=True, text=True, cwd=REPO)
    return p.returncode, p.stdout, p.stderr


def sprawdz(nazwa, warunek, detale=""):
    status = "OK" if warunek else "BŁĄD"
    print(f"[{status}] {nazwa}" + (f" — {detale}" if not warunek else ""))
    if not warunek:
        błędy.append(nazwa)


def zapisz_csv(sciezka, wiersze, delimiter=";"):
    with open(sciezka, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=delimiter)
        w.writerow(NAGL)
        w.writerows(wiersze)


def liczniki(stdout):
    m = {k: int(v) for k, v in re.findall(
        r"(przeterminowane|w tygodniu|w następnym tygodniu): (\d+)", stdout)}
    return m


def main():
    tmp = tempfile.mkdtemp(prefix="tp-testy-")
    fix_okno = os.path.join(REPO, "testy", "rejestr-okno.csv")
    fix_mieszany = os.path.join(REPO, "testy", "rejestr-mieszany.csv")
    fix_duzy = os.path.join(REPO, "testy", "rejestr-duzy.csv")
    for f in (fix_okno, fix_mieszany, fix_duzy):
        if not os.path.exists(f):
            sys.exit(f"brak fixture'u {f} — uruchom najpierw testy/generuj_fixtures.py")

    # 1. regresja: rejestr-okno -> 1 przeterminowane / 4 w tygodniu / 2 następny
    rc, out, err = uruchom(["przeglad", "--rejestr", fix_okno])
    c = liczniki(out)
    sprawdz("rejestr-okno rc=0", rc == 0, f"rc={rc}, err={err.strip()}")
    sprawdz("rejestr-okno 1/4/2", (c.get("przeterminowane"), c.get("w tygodniu"),
                                  c.get("w następnym tygodniu")) == (1, 4, 2), str(c))

    # 2. regresja: rejestr-mieszany -> 2/1/2, 10 wierszy w pliku
    with open(fix_mieszany, encoding="utf-8-sig") as f:
        n_mieszany = sum(1 for _ in csv.reader(f)) - 1
    rc, out, err = uruchom(["przeglad", "--rejestr", fix_mieszany])
    c = liczniki(out)
    sprawdz("rejestr-mieszany rc=0", rc == 0, f"rc={rc}, err={err.strip()}")
    sprawdz("rejestr-mieszany n=10", n_mieszany == 10, f"n={n_mieszany}")
    sprawdz("rejestr-mieszany 2/1/2", (c.get("przeterminowane"), c.get("w tygodniu"),
                                      c.get("w następnym tygodniu")) == (2, 1, 2), str(c))
    wiersze_n = len(re.findall(r"^  \d{4}-\d{2}-\d{2}  ", out, re.M))
    sprawdz("rejestr-mieszany 5 wierszy w tabeli (2+1+2)",
            wiersze_n == 5, f"tabela: {wiersze_n}")
    sprawdz("rejestr-mieszany pominięte jawne", "pominięte:" in out and "zła data: 1" in out
            and "zła kwota: 3" in out)

    # 3. pusty rejestr -> rc=2
    pusty = os.path.join(tmp, "pusty.csv")
    zapisz_csv(pusty, [])
    rc, out, err = uruchom(["przeglad", "--rejestr", pusty])
    sprawdz("pusty rejestr rc=2", rc == 2, f"rc={rc}, err={err.strip()}")

    # 4. --okno 0 -> rc=1
    rc, out, err = uruchom(["czytaj", "--rejestr", fix_okno, "--okno", "0"])
    sprawdz("--okno 0 rc=1", rc == 1 and "--okno musi być ≥ 1" in err,
            f"rc={rc}, err={err.strip()}")

    # 5. zła data --z / --do -> rc=1, ludzki komunikat, bez tracebacku
    rc, out, err = uruchom(["czytaj", "--rejestr", fix_okno, "--z", "26.10"])
    sprawdz("--z 26.10 rc=1", rc == 1, f"rc={rc}")
    sprawdz("--z 26.10 komunikat", "zła data: 26.10 (oczekuję RRRR-MM-DD)" in err,
            err.strip())
    sprawdz("--z 26.10 bez tracebacku", "ValueError" not in err and "Traceback" not in err)
    rc, out, err = uruchom(["czytaj", "--rejestr", fix_okno, "--do", "jutrze"])
    sprawdz("--do 'jutrze' rc=1 + komunikat", rc == 1
            and "zła data: jutrze (oczekuję RRRR-MM-DD)" in err
            and "Traceback" not in err, err.strip())
    # dobra data DD.MM.RRRR nadal działa
    rc, out, err = uruchom(["czytaj", "--rejestr", fix_okno, "--z", "01.01.2026"])
    sprawdz("--z DD.MM.RRRR nadal działa", rc == 0, err.strip())

    # 6. CSV z przecinkami -> rc=2 + komunikat (nie mylące „nie ma kolumn")
    przecinek = os.path.join(tmp, "przecinek.csv")
    zapisz_csv(przecinek, [], delimiter=",")
    rc, out, err = uruchom(["przeglad", "--rejestr", przecinek])
    sprawdz("CSV-przecinek rc=2", rc == 2, f"rc={rc}, err={err.strip()}")
    sprawdz("CSV-przecinek komunikat", "rejestr wygląda na CSV z przecinkami" in err
            and "separator to średnik" in err, err.strip())
    sprawdz("CSV-przecinek bez mylącego 'nie ma kolumn'", "nie ma kolumn" not in err)

    # 7. przed/po oknem przy --z w przyszłości (rejestr-duzy: offsets -3..+14)
    z2 = (date.today() + timedelta(days=2)).isoformat()
    do2 = (date.today() + timedelta(days=6)).isoformat()
    rc, out, err = uruchom(["czytaj", "--rejestr", fix_duzy, "--z", z2, "--do", do2])
    sprawdz("--z przyszłość rc=0", rc == 0, f"rc={rc}, err={err.strip()}")
    m_przed = re.search(r"przed oknem \(do [\d-]+\): (\d+)", out)
    m_poza = re.search(r"poza oknem \(po [\d-]+\): (\d+)", out)
    sprawdz("--z przyszłość: koszyk 'przed oknem' (offsets 0..1 → 4 wiersze)",
            m_przed and m_przed.group(1) == "4", out[:0] + (m_przed.group(0) if m_przed else ""))
    sprawdz("--z przyszłość: koszyk 'poza oknem' (offset 7 + 25 następnych → 26)",
            m_poza and m_poza.group(1) == "26", m_poza.group(0) if m_poza else "")
    sprawdz("--z przyszłość: przed oknem ma osobną sumę w PODSUMOWANIU",
            re.search(r"przed oknem: 4 \(suma", out) is not None)

    # 8. wielostronicowy PDF (40 wierszy) — etykieta sekcji „— cd.” na stronach 2+
    pdf_out = os.path.join(tmp, "duzy.pdf")
    rc, out, err = uruchom(["pdf", "--rejestr", fix_duzy, "--wyjscie", pdf_out])
    sprawdz("pdf duzy rc=0", rc == 0, f"rc={rc}, err={err.strip()}")
    sprawdz("pdf duzy wielostronicowy", "wielostronicowy" in out, out.strip())
    if os.path.exists(pdf_out):
        from pypdf import PdfReader
        tekst = "\n".join((pg.extract_text() or "") for pg in PdfReader(pdf_out).pages)
        strony_tekst = [(pg.extract_text() or "") for pg in PdfReader(pdf_out).pages]
        sprawdz("pdf duzy >1 strona", len(strony_tekst) > 1, f"stron: {len(strony_tekst)}")
        # pierwsza strona ma zwykłe etykiety sekcji, strony 2+ mają „— cd.”
        sprawdz("pdf str.1 ma 'PRZETERMINOWANE' bez '— cd.'",
                "PRZETERMINOWANE" in strony_tekst[0]
                and "PRZETERMINOWANE — cd." not in strony_tekst[0])
        cd = [i + 1 for i, t in enumerate(strony_tekst) if "— cd." in t]
        sprawdz("pdf strony 2+ mają '— cd.'", len(cd) >= 1 and cd[0] >= 2,
                f"strony z '— cd.': {cd}")
        # każda strona z „— cd.” ma też powtórzony nagłówek tabeli
        for nr in cd:
            sprawdz(f"pdf str.{nr}: nagłówek tabeli pod etykietą",
                    "nr faktury" in strony_tekst[nr - 1])

    print()
    if błędy:
        print(f"NIEZALICZONE: {len(błędy)}: " + "; ".join(błędy))
        sys.exit(1)
    print("Wszystkie testy zaliczone.")


if __name__ == "__main__":
    main()
