#!/usr/bin/env python3
"""Generuje fikcyjne fixture'y rejestru do pomiaru terminarz-platnosci.

Daty liczone od dnia uruchomienia (przeszłość / dzisiaj / ten tydzień /
następny tydzień / poza oknem), więc pomiar zawsze jest aktualny.

Tworzy:
  przyklad-rejestr.csv            (repo root — przykładowy wiersz)
  testy/rejestr-okno.csv          (8 wierszy: okno + bez terminu)
  testy/rejestr-mieszany.csv      (7 wierszy: formaty dat/kwot + złe wiersze)
Fikcyjne dane, poprawne formatowo. Nic nie wysyła, zero sieci.
"""
import csv
import os
from datetime import date, timedelta

DZIS = date.today()
NAGL = ["plik", "sprzedawca", "nip_sprzedawcy", "nr_faktury",
        "data_wystawienia", "termin_platnosci", "kwota_brutto", "komentarz"]
SPRA = [("Przykladowa Firma sp. z o.o.", "5252445514"),
        ("Hurtownia Testowa S.A.", "5842159772"),
        ("Usługi Krak-Bud sp. z o.o.", "6762211846"),
        ("Kancelaria Przykładska", "1231067433")]


def d(offset):
    return (DZIS + timedelta(days=offset)).isoformat()


def zapisz(sciezka, wiersze):
    with open(sciezka, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(NAGL)
        for r in wiersze:
            w.writerow([r[k] for k in NAGL])
    print(f"OK: {sciezka} ({len(wiersze)} wierszy)")


def wiersz(n, sprzedawca, nip, termin, kwota, offset_wyst=-14, komentarz="OK"):
    return {"plik": f"fikcja-{n}.pdf", "sprzedawca": sprzedawca, "nip_sprzedawcy": nip,
            "nr_faktury": n, "data_wystawienia": d(offset_wyst),
            "termin_platnosci": termin, "kwota_brutto": kwota, "komentarz": komentarz}


s1, s2, s3, s4 = SPRA

# -- przykład (1 wiersz, w oknie) --------------------------------------------
zapisz("przyklad-rejestr.csv", [wiersz("FV/2026/10/014", s1[0], s1[1], d(5), "2829,00")])

# -- rejestr-okno: 1 przeterminowana, 4 w tygodniu, 2 następny, 1 bez terminu --
zapisz("testy/rejestr-okno.csv", [
    wiersz("FV/2026/10/001", s2[0], s2[1], d(-6),  "2 411,20", offset_wyst=-20),
    wiersz("FV/2026/10/002", s3[0], s3[1], d(0),   "1 580,00", offset_wyst=-12),
    wiersz("FV/2026/10/003", s4[0], s4[1], d(2),   "920,5",    offset_wyst=-10),
    wiersz("FV/2026/10/004", s1[0], s1[1], d(4),   "1 234,50", offset_wyst=-8),
    wiersz("FV/2026/10/005", s2[0], s2[1], d(7),   "6 000,00", offset_wyst=-6),
    wiersz("FV/2026/10/006", s3[0], s3[1], d(9),   "845,00",   offset_wyst=-5),
    wiersz("FV/2026/10/007", s4[0], s4[1], d(12),  "3 100,00", offset_wyst=-4),
    wiersz("FV/2026/10/008", s1[0], s1[1], "brak", "500,00",   offset_wyst=-3,
           komentarz="brak terminu w PDF źródłowym"),
])

# -- rejestr-mieszany: formaty dat/kwot + zła data, zła kwota, pusty termin ----
zapisz("testy/rejestr-mieszany.csv", [
    wiersz("FV/2026/10/011", s1[0], s1[1], d(-14), "5 230,00", offset_wyst=-28),
    wiersz("FV/2026/10/012", s2[0], s2[1], d(-1),  "1 000,00", offset_wyst=-15),
    wiersz("FV/2026/10/013", s3[0], s3[1], d(1),   "1845.5",   offset_wyst=-13),  # kropka dziesiętna
    wiersz("FV/2026/10/014", s4[0], s4[1], d(13),  "7 999,99", offset_wyst=-2),
    wiersz("FV/2026/10/015", s1[0], s1[1], "31.02.2026", "100,00", offset_wyst=-2,  # zła data
           komentarz="31.02 nie istnieje w kalendarzu"),
    wiersz("FV/2026/10/016", s2[0], s2[1], d(3), "dwieście złotych", offset_wyst=-2,  # zła kwota
           komentarz="kwota słownie — rejestr nie powinien tak wyglądać"),
    wiersz("FV/2026/10/017", s3[0], s3[1], "", "3 412,80", offset_wyst=-2,  # pusty termin
           komentarz="brak terminu w PDF źródłowym"),
])
