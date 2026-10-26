#!/usr/bin/env python3
"""terminarz-platnosci — rejestr faktur -> „co płacimy w tym tygodniu / co jest przeterminowane”.

Wejście: rejestr CSV (średnik, UTF-8) albo XLSX z kolumnami:
plik;sprzedawca;nip_sprzedawcy;nr_faktury;data_wystawienia;termin_platnosci;kwota_brutto;komentarz
(układ rejestru z narzędzia faktury-do-rejestru).

Komendy:
  czytaj   --rejestr R [--z D] [--do D] [--okno N]   podgląd w terminalu
  przeglad --rejestr R [--okno N]                    podgląd w terminalu (domyślne okno)
  pdf      --rejestr R --wyjscie P.pdf [--okno N]    PDF: 1 strona do ~25 wierszy,
                                                     przy większych wielostronicowy
                                                     z powtórzonym nagłówkiem tabeli

Koszyki:
  PRZETERMINOWANE        termin < dzisiaj (sort: od najstarszych)
  PRZED OKNEM            termin ≥ dzisiaj, ale przed startem okna
                         (tylko przy --z z przyszłości) — osobna suma
  W TYM TYGODNIU         przez najbliższe 7 dni, od dnia startu okna (default: od dziś)
  ZAPEWNE NASTĘPNY TYDZIEŃ  od 8. dnia do końca okna (--okno, default 14)

Kody wyjścia: 0 = OK, 1 = błąd, 2 = nic nie zrobiono (rejestr pusty / nic do pokazania).
Całość lokalnie, zero sieci. To nie jest system księgowy: po prostu terminarz.
"""
import argparse
import csv
import os
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

KOLUMNY = ["plik", "sprzedawca", "nip_sprzedawcy", "nr_faktury",
           "data_wystawienia", "termin_platnosci", "kwota_brutto", "komentarz"]
FORMATY_DATY = ("%Y-%m-%d", "%d.%m.%Y", "%d.%m.%y")
PUSTE_DATY = ("brak", "-", "x", "n/d", "nd", "")


def die(msg, code=1):
    print(f"BŁĄD: {msg}", file=sys.stderr)
    sys.exit(code)


# ============================================================ PARSOWANIE
def parsuj_date(s):
    """ISO, DD.MM.RRRR, DD.MM.RR. Rzuca ValueError przy złej dacie-kalendarzu."""
    for f in FORMATY_DATY:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    raise ValueError(s)


def parsuj_termin(s):
    """Zwraca (date|None, powód_pominięcia|None). Puste/„brak” = bez terminu."""
    s = (s or "").strip()
    if s.lower() in PUSTE_DATY:
        return None, "bez terminu"
    try:
        return parsuj_date(s), None
    except ValueError:
        return None, "zła data"


def _kropki_tysiece(czesc_calkowita):
    """True = „1.234.567” (kropki co 3 cyfry od prawej, bez zera wiodącego)."""
    segs = czesc_calkowita.split(".")
    if len(segs) < 2 or not all(seg.isdigit() for seg in segs):
        return False
    if len(segs[0]) not in (1, 2, 3):
        return False
    return all(len(seg) == 3 for seg in segs[1:])


def parsuj_kwota(s):
    """Polski format: „1 845,00 zł”, „1845,00 złoty”, „2 411,2”, „2829,00”;
    też kropka dziesiętna („1845.5”). Nigdy ciche zgadywanie: formaty
    niejednoznaczne (US-style, „1.234”, „1,234”) rzucają ValueError z powodem —
    wiersz idzie do licznika pominiętych z powodem, nigdy po cichu."""
    t = (s or "").strip()
    if not t:
        raise ValueError("pusta kwota")
    t = t.replace("\u00a0", " ").replace("\u202f", " ")
    # najpierw pełne słowa waluty („złoty” zawiera „zł” — kolejność ma znaczenie)
    t = (t.replace("złotych", "").replace("złotego", "")
          .replace("złoty", "").replace("złote", "")
          .replace("zł", "").replace("PLN", "").replace("pln", "").strip())
    znak = ""
    if t.startswith(("+", "-")):
        znak = "-" if t.startswith("-") else ""
        t = t[1:].strip()
    if "." in t and "," in t:
        # kropka i przecinek razem: rozpoznajemy TYLKO polski wariant
        # „1.234,56” (kropki-tysiące co 3 cyfry, przecinek + 2 cyfry groszy).
        # Reszta (np. US „1,234.56”) = odrzucenie, nie zgadywanie.
        cal, po = t.rsplit(",", 1)
        cal = cal.replace(" ", "")
        if len(po) == 2 and po.isdigit() and _kropki_tysiece(cal):
            t = cal.replace(".", "") + "." + po
        else:
            raise ValueError("podejrzany format kwoty (może US-style)")
    elif "," in t:
        po = t.rsplit(",", 1)[1]
        if len(po) == 3 and po.isdigit():
            raise ValueError("podejrzany format kwoty (przecinek-tysiące czy grosze?)")
        t = t.replace(" ", "").replace(".", "").replace(",", ".")
    elif "." in t:
        # kropka bez przecinka: grosza („1845.5”, „2829.00”) albo tysiące („1.234”)
        po = t.rsplit(".", 1)[1]
        if len(po) == 3 and po.isdigit():
            raise ValueError("podejrzany format kwoty (kropka-tysiące czy grosze?)")
        t = t.replace(" ", "")
    else:
        t = t.replace(" ", "")
    d = Decimal(znak + t)
    if not d.is_finite():
        raise ValueError("podejrzany format kwoty")
    if abs(d) > Decimal("10000000"):
        raise ValueError("kwota podejrzana (> 10 mln)")
    return d


def kwota_str(d, waluta=True):
    q = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    s = f"{q:,.2f}".replace(",", " ").replace(".", ",")
    return f"{s} zł" if waluta else s


# ============================================================ WEJŚCIE
def _wiersz_z_wiersza(raw):
    w = {k: "" for k in KOLUMNY}
    for i, k in enumerate(KOLUMNY):
        v = raw[i] if i < len(raw) else ""
        if isinstance(v, datetime):
            v = v.date().isoformat()
        elif hasattr(v, "isoformat"):  # openpyxl: datetime.date
            v = v.isoformat()
        w[k] = ("" if v is None else str(v)).strip()
    return w


def wczytaj_rejestr(sciezka):
    if not os.path.exists(sciezka):
        die(f"nie ma pliku rejestru: {sciezka}")
    ext = os.path.splitext(sciezka)[1].lower()
    wiersze, naglowek = [], []
    if ext in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            die("brak openpyxl: pip3 install openpyxl")
        wb = load_workbook(sciezka, read_only=True, data_only=True)
        ws = wb.active
        it = ws.iter_rows(values_only=True)
        try:
            naglowek = [("" if c is None else str(c).strip()) for c in next(it)]
        except StopIteration:
            die(f"rejestr pusty: {sciezka}", 2)
        for raw in it:
            wiersze.append(_wiersz_z_wiersza(raw))
    elif ext == ".csv":
        with open(sciezka, newline="", encoding="utf-8-sig") as f:
            r = csv.reader(f, delimiter=";")
            try:
                naglowek = [c.strip() for c in next(r)]
            except StopIteration:
                die(f"rejestr pusty: {sciezka}", 2)
            for raw in r:
                wiersze.append(_wiersz_z_wiersza(raw))
    else:
        die(f"nieznany format rejestru: {sciezka} (obsługa: .csv, .xlsx)")
    braki = [k for k in ("termin_platnosci", "kwota_brutto") if k not in naglowek]
    if braki:
        die(f"rejestr nie ma kolumn: {', '.join(braki)} (nagłówek: {';'.join(naglowek)})")
    idx = {k: naglowek.index(k) for k in KOLUMNY if k in naglowek}
    for i, w in enumerate(wiersze):
        w["nr_wiersza"] = i + 2  # jak w Excelu (nagłówek = wiersz 1)
    return wiersze


def klasyfikuj(wiersze, dzis, start, koniec):
    """Dzieli wiersze na koszyki + liczniki pominiętych."""
    kosz = {"przeterminowane": [], "przed_oknem": [], "w_tyg": [],
            "nastepny": [], "poza": []}
    bez_terminu, zle_daty, zle_kwoty = [], [], []
    for w in wiersze:
        termin, powod = parsuj_termin(w["termin_platnosci"])
        try:
            kwota = parsuj_kwota(w["kwota_brutto"])
        except InvalidOperation:
            w["powod_kwota"] = "zła kwota"
            zle_kwoty.append(w)
            continue
        except ValueError as e:
            w["powod_kwota"] = str(e)
            zle_kwoty.append(w)
            continue
        w["kwota"] = kwota
        if termin is None:
            if powod == "zła data":
                zle_daty.append(w)
            else:
                bez_terminu.append(w)
            continue
        w["termin"] = termin
        if termin < dzis:
            kosz["przeterminowane"].append(w)
        elif termin < start:
            kosz["przed_oknem"].append(w)  # po dziś, ale przed startem (--z z przyszłości)
        elif termin <= start + timedelta(days=7) and termin <= koniec:
            kosz["w_tyg"].append(w)
        elif termin <= koniec:
            kosz["nastepny"].append(w)
        else:
            kosz["poza"].append(w)
    kosz["przeterminowane"].sort(key=lambda w: w["termin"])
    for k in ("w_tyg", "nastepny"):
        kosz[k].sort(key=lambda w: (w["termin"], w["nr_faktury"]))
    return kosz, bez_terminu, zle_daty, zle_kwoty


def suma(wiersze):
    return sum((w["kwota"] for w in wiersze), Decimal("0"))


def dni_opis(termin, dzis):
    d = (termin - dzis).days
    if d == 0:
        return "dzisiaj"
    if d > 0:
        return f"za {d} {'dzień' if d == 1 else 'dni'}"
    dd = -d
    return f"{dd} {'dzień' if dd == 1 else 'dni'} po terminie"


def linia_wiersza(w, dzis):
    return (f"  {w['termin'].isoformat()}  ({dni_opis(w['termin'], dzis):>20})  "
            f"{(w['nr_faktury'] or '(bez nr)')[:28]:<28}  "
            f"{(w['sprzedawca'] or '(bez sprzedawcy)')[:38]:<38}  {kwota_str(w['kwota']):>14}")


def wydruk(kosz, dzis, start, koniec, bez_terminu, zle_daty, zle_kwoty, rejestr):
    print(f"TERMINARZ PŁATNOŚCI — na dzień {dzis.isoformat()} (rejestr: {rejestr})")
    print(f"okno płatności: {start.isoformat()}..{koniec.isoformat()}  "
          f"(przeterminowane: wszystko przed {dzis.isoformat()})")
    p, t, n = kosz["przeterminowane"], kosz["w_tyg"], kosz["nastepny"]
    if p:
        print("\nPRZETERMINOWANE:")
        for w in p:
            print(linia_wiersza(w, dzis))
    if t:
        print(f"\nW TYM TYGODNIU (do {(start + timedelta(days=7)).isoformat()}):")
        for w in t:
            print(linia_wiersza(w, dzis))
    if n:
        print(f"\nZAPEWNE NASTĘPNY TYDZIEŃ ({(start + timedelta(days=8)).isoformat()}"
              f"..{koniec.isoformat()}):")
        for w in n:
            print(linia_wiersza(w, dzis))
    if kosz["przed_oknem"]:
        print(f"\nprzed oknem (do {start.isoformat()}): {len(kosz['przed_oknem'])} — "
              f"suma {kwota_str(suma(kosz['przed_oknem']))}")
    if kosz["poza"]:
        print(f"\npoza oknem (po {koniec.isoformat()}): {len(kosz['poza'])} — "
              f"suma {kwota_str(suma(kosz['poza']))}")
    pom = []
    if bez_terminu:
        pom.append(f"bez terminu: {len(bez_terminu)} "
                   f"(wiersze: {', '.join(str(w['nr_wiersza']) for w in bez_terminu)})")
    if zle_daty:
        pom.append(f"zła data: {len(zle_daty)} "
                   f"(wiersze: {', '.join(str(w['nr_wiersza']) for w in zle_daty)})")
    if zle_kwoty:
        det = ", ".join(f"{w['nr_wiersza']} — {w['powod_kwota']}" for w in zle_kwoty)
        pom.append(f"zła kwota: {len(zle_kwoty)} (wiersze: {det})")
    if pom:
        print("pominięte: " + "; ".join(pom))
    print("\nPODSUMOWANIE:")
    print(f"  przeterminowane: {len(p)} (suma {kwota_str(suma(p))})")
    print(f"  w tygodniu: {len(t)} (suma {kwota_str(suma(t))})")
    print(f"  w następnym tygodniu: {len(n)} (suma {kwota_str(suma(n))})")
    if kosz["przed_oknem"]:
        print(f"  przed oknem: {len(kosz['przed_oknem'])} "
              f"(suma {kwota_str(suma(kosz['przed_oknem']))})")
    if kosz["poza"]:
        print(f"  poza oknem: {len(kosz['poza'])} "
              f"(suma {kwota_str(suma(kosz['poza']))})")
    print(f"  do zapłaty łącznie w oknie: {len(t) + len(n)} "
          f"(suma {kwota_str(suma(t) + suma(n))})")
    print(f"  wraz z przeterminowanymi: {len(p) + len(t) + len(n)} "
          f"(suma {kwota_str(suma(p) + suma(t) + suma(n))})")


# ============================================================ FONTS (jak wystaw.py)
def znajdz_font():
    kandydaci = []
    env_warn = []
    for zm in ("TERMINARZ_FONT", "FDR_FONT"):
        env = os.environ.get(zm)
        if env:
            if os.path.exists(env):
                kandydaci.append(env)
            else:
                env_warn.append(f"{zm}={env} — plik nie istnieje")
    for komunikat in env_warn:
        print(f"UWAGA: {komunikat} — użyję innego fontu z listy.", file=sys.stderr)
    try:
        import matplotlib
        kandydaci.append(os.path.join(matplotlib.get_data_path(), "fonts", "ttf", "DejaVuSans.ttf"))
    except ImportError:
        pass
    kandydaci += [
        # Windows (Arial/Segoe UI/Tahoma mają polskie znaki)
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/segoeui.ttf",
        "C:/Windows/Fonts/tahoma.ttf",
        "/System/Library/Fonts/Supplemental/Arial Unicode.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "/usr/local/share/fonts/DejaVuSans.ttf",
        os.path.expanduser("~/Library/Fonts/Arial Unicode.ttf"),
    ]
    for k in kandydaci:
        if k and os.path.exists(k):
            return k
    die("brak fontu z polskimi znakami. Ustaw TERMINARZ_FONT=/ścieżka/do/DejaVuSans.ttf "
        "(Linux: fonts-dejavu; macOS: Arial Unicode jest w systemie).", 1)


# ============================================================ PDF
def pdf_tabela(pdf, tytul, kolor_nagl, wiersze, dzis, opis_kol2, nagl_kol2):
    pdf.set_text_color(*kolor_nagl)
    pdf.set_font_size(11)
    pdf.cell(0, 7, f"{tytul}  ({len(wiersze)} poz.)", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    szer_kol = [22, 24, 34, 80, 30]
    pdf.set_fill_color(232, 232, 232)
    nagl = ["termin", nagl_kol2, "nr faktury", "sprzedawca", "kwota"]

    def naglowek_tabeli():
        # powtarzany na każdej stronie przy wielostronicowym PDF
        pdf.set_font_size(7.5)
        for w_, t_ in zip(szer_kol, nagl):
            pdf.cell(w_, 5, t_, border=1, fill=True, align="L")
        pdf.ln()
        pdf.set_font_size(8.5)

    naglowek_tabeli()
    for w in wiersze:
        if pdf.will_page_break(5.5):
            pdf.add_page()
            naglowek_tabeli()
        kom2 = dni_opis(w["termin"], dzis)
        pdf.set_text_color(*kolor_nagl) if w["termin"] < dzis else pdf.set_text_color(0, 0, 0)
        pdf.cell(szer_kol[0], 5.5, w["termin"].isoformat(), border=1)
        pdf.cell(szer_kol[1], 5.5, kom2, border=1)
        pdf.set_text_color(0, 0, 0)
        pdf.cell(szer_kol[2], 5.5, (w["nr_faktury"] or "(bez nr)")[:24], border=1)
        pdf.cell(szer_kol[3], 5.5, (w["sprzedawca"] or "(bez sprzedawcy)")[:52], border=1)
        pdf.cell(szer_kol[4], 5.5, kwota_str(w["kwota"], waluta=False), border=1, align="R")
        pdf.ln()
    if pdf.will_page_break(8):
        pdf.add_page()  # nie zostawiaj sumy samej na dole strony
    pdf.set_font_size(9)
    pdf.cell(0, 6, f"Suma: {kwota_str(suma(wiersze))}", align="R",
             new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)


def cmd_pdf(a):
    try:
        from fpdf import FPDF
    except ModuleNotFoundError:
        die("brak fpdf2: pip3 install fpdf2")
    if a.okno < 1:
        die(f"--okno musi być ≥ 1 (podano: {a.okno})")
    wiersze = wczytaj_rejestr(a.rejestr)
    dzis = date.today()
    start, koniec = dzis, dzis + timedelta(days=a.okno)
    kosz, bez_terminu, zle_daty, zle_kwoty = klasyfikuj(wiersze, dzis, start, koniec)
    p, t, n = kosz["przeterminowane"], kosz["w_tyg"], kosz["nastepny"]
    if not (p or t or n):
        print(f"OK: nic do pokazania (brak przeterminowanych i nic w oknie "
              f"{start.isoformat()}..{koniec.isoformat()}; PDF nie został wygenerowany")
        return 2
    font = znajdz_font()
    pdf = FPDF(format="A4")
    pdf.add_font("tp", "", font)
    pdf.set_margins(10, 12, 10)
    pdf.set_auto_page_break(True, margin=15)
    pdf.add_page()
    pdf.set_font("tp", size=10)
    pdf.set_font_size(17)
    pdf.cell(0, 10, "Terminarz płatności", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font_size(8.5)
    pdf.cell(0, 5, f"na dzień {dzis.isoformat()} · okno płatności "
             f"{start.isoformat()}..{koniec.isoformat()} · rejestr: "
             f"{os.path.basename(a.rejestr)}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)
    if p:
        pdf_tabela(pdf, "PRZETERMINOWANE", (178, 20, 20), p, dzis,
                   "dni po terminie", "po terminie")
    if t:
        pdf_tabela(pdf, "W TYM TYGODNIU", (20, 90, 40), t, dzis,
                   "dni do terminu", "do terminu")
    if n:
        pdf_tabela(pdf, "ZAPEWNE NASTĘPNY TYDZIEŃ", (120, 90, 10), n, dzis,
                   "dni do terminu", "do terminu")
    pdf.set_font_size(10)
    pdf.cell(0, 6, f"Do zapłaty łącznie w oknie: {kwota_str(suma(t) + suma(n))}",
             new_x="LMARGIN", new_y="NEXT")
    if p:
        pdf.cell(0, 6, f"Wraz z przeterminowanymi: {kwota_str(suma(p) + suma(t) + suma(n))}",
                 new_x="LMARGIN", new_y="NEXT")
    pom = []
    if bez_terminu:
        pom.append(f"bez terminu: {len(bez_terminu)}")
    if zle_daty:
        pom.append(f"zła data: {len(zle_daty)}")
    if zle_kwoty:
        pom.append(f"zła kwota: {len(zle_kwoty)}")
    if kosz["przed_oknem"]:
        pom.append(f"przed oknem (do {start.isoformat()}): {len(kosz['przed_oknem'])} "
                   f"({kwota_str(suma(kosz['przed_oknem']))})")
    if kosz["poza"]:
        pom.append(f"poza oknem (po {koniec.isoformat()}): {len(kosz['poza'])} "
                   f"({kwota_str(suma(kosz['poza']))})")
    if pom:
        pdf.set_font_size(7.5)
        pdf.cell(0, 5, "pominięte / poza oknem: " + "; ".join(pom),
                 new_x="LMARGIN", new_y="NEXT")
    pdf.set_font_size(7)
    pdf.set_text_color(110, 110, 110)
    pdf.ln(2)
    pdf.cell(0, 5, "Wygenerowano lokalnie (terminarz-platnosci v0.2). Rejestr roboczy — "
             "potwierdź kwoty u źródła (faktura/bank), to nie jest system księgowy.",
             new_x="LMARGIN", new_y="NEXT")
    kat = os.path.dirname(os.path.abspath(a.wyjscie))
    if kat:
        os.makedirs(kat, exist_ok=True)
    pdf.output(a.wyjscie)
    strony = getattr(pdf, "pages_count", None) or pdf.page
    print(f"OK: terminarz PDF -> {a.wyjscie} "
          f"(przeterminowane: {len(p)}, w tym tygodniu: {len(t)}, następny tydzień: {len(n)}; "
          f"do zapłaty w oknie: {kwota_str(suma(t) + suma(n))})")
    if strony > 1:
        w_wierszy = len(p) + len(t) + len(n)
        print(f"UWAGA: PDF wielostronicowy: {w_wierszy} wierszy (stron: {strony}) — "
              f"nagłówek tabeli powtórzony na każdej stronie.")
    return 0


# ============================================================ CLI
def cmd_czytaj(a):
    wiersze = wczytaj_rejestr(a.rejestr)
    dzis = date.today()
    start = parsuj_date(a.z) if a.z else dzis
    if a.okno < 1 and not a.do:
        die(f"--okno musi być ≥ 1 (podano: {a.okno})")
    koniec = parsuj_date(a.do) if a.do else start + timedelta(days=a.okno)
    if koniec < start:
        die(f"--do ({a.do}) jest wcześniejsze niż --z ({a.z})")
    kosz, bez_terminu, zle_daty, zle_kwoty = klasyfikuj(wiersze, dzis, start, koniec)
    wydruk(kosz, dzis, start, koniec, bez_terminu, zle_daty, zle_kwoty, a.rejestr)
    if kosz["przeterminowane"] or kosz["w_tyg"] or kosz["nastepny"]:
        return 0
    print(f"OK: nic do pokazania (rejestr: {len(wiersze)} pozycji, "
          f"wszystkie pominięte albo poza oknem)")
    return 2


def cmd_przeglad(a):
    a.z = a.do = None
    return cmd_czytaj(a)


def main():
    p = argparse.ArgumentParser(
        prog="terminarz_platnosci.py",
        description="rejestr faktur -> co płacimy w tym tygodniu / co jest przeterminowane",
        epilog="Całość lokalnie, zero sieci. Nie jest systemem księgowym.")
    p.add_argument("--version", action="version", version="%(prog)s 0.2")
    sub = p.add_subparsers(dest="komenda", required=True)

    def dodaj(kom, opis):
        sp = sub.add_parser(kom, help=opis)
        sp.add_argument("--rejestr", required=True,
                        help="rejestr CSV (średnik, UTF-8) albo XLSX")
        return sp

    sp = dodaj("czytaj", "podgląd w terminalu z opcjonalnym oknem")
    sp.add_argument("--z", help="start okna RRRR-MM-DD (default: dzisiaj)")
    sp.add_argument("--do", help="koniec okna RRRR-MM-DD (default: dzisiaj+okno)")
    sp.add_argument("--okno", type=int, default=14, help="długość okna w dniach (default 14)")
    sp.set_defaults(func=cmd_czytaj)

    sp = dodaj("przeglad", "podgląd w terminalu (tekst, bez PDF)")
    sp.add_argument("--okno", type=int, default=14, help="długość okna w dniach (default 14)")
    sp.set_defaults(func=cmd_przeglad)

    sp = dodaj("pdf", "jednostronicowy PDF: przeterminowane / ten tydzień / następny")
    sp.add_argument("--wyjscie", required=True, help="plik wynikowy P.pdf")
    sp.add_argument("--okno", type=int, default=14, help="długość okna w dniach (default 14)")
    sp.set_defaults(func=cmd_pdf)

    a = p.parse_args()
    return a.func(a)


if __name__ == "__main__":
    try:
        sys.exit(main())
    except BrokenPipeError:
        sys.exit(1)
    except KeyboardInterrupt:
        sys.exit(130)
