# terminarz-platnosci

Z rejestru faktur robi PDF: **co płacimy w tym tygodniu / co jest
przeterminowane** — jedna strona przy typowym rejestrze (do ~25 wierszy);
przy większych wielostronicowy z powtórzonym nagłówkiem tabeli.

## Skąd wziąć to narzędzie

```bash
git clone https://github.com/martin0ne/terminarz-platnosci
cd terminarz-platnosci && pip3 install fpdf2 openpyxl
```

Rejestr wejściowy generuje
[faktury-do-rejestru](https://github.com/martin0ne/faktury-do-rejestru):
`czytaj → rejestr.csv → tutaj`. Cały łańcuch w dwóch komendach:

```bash
python3 faktury_do_rejestru.py czytaj --katalog faktury/ --rejestr rejestr.csv
python3 terminarz_platnosci.py przeglad --rejestr rejestr.csv
```
(przetestowane wklejeniem — obie komendy na świeżym katalogu)

Wejście to rejestr z narzędzia
[faktury-do-rejestru](https://github.com/martin0ne/faktury-do-rejestru)
(CSV ze średnikiem albo XLSX), więc nie trzeba niczego przepisywać drugi raz.

Całość lokalnie, zero sieci, zero chmury, zero kont.

| Komenda | Co robi |
|---|---|
| `czytaj` | wczytuje rejestr, drukuje w terminalu koszyki: przeterminowane + okno płatności |
| `przeglad` | to samo co `czytaj` z domyślnym oknem (tekst, bez PDF) |
| `pdf` | PDF z sekcjami PRZETERMINOWANE / W TYM TYGODNIU / ZAPEWNE NASTĘPNY TYDZIEŃ (jedna strona do ~25 wierszy; większe — wielostronicowe z powtórzonym nagłówkiem + ostrzeżenie w terminalu) |

## Instalacja

```bash
pip3 install fpdf2 openpyxl
```

Do `pdf` potrzebny font z polskimi znakami (na macOS (Arial Unicode) i Windows
(Arial/Segoe UI) jest w systemie, na Linuxie `fonts-dejavu`. Można wskazać
własny: `TERMINARZ_FONT=/ścieżka/do/DejaVuSans.ttf`.

## Użycie

```bash
# podgląd w terminalu: przeterminowane + najbliższe 14 dni
python3 terminarz_platnosci.py przeglad --rejestr rejestr.csv

# węższe okno
python3 terminarz_platnosci.py czytaj --rejestr rejestr.csv --z 2026-10-26 --do 2026-11-09

# PDF (jedna strona przy typowym rejestrze, wielostronicowy przy większym)
python3 terminarz_platnosci.py pdf --rejestr rejestr.csv --wyjscie terminarz.pdf
python3 terminarz_platnosci.py pdf --rejestr rejestr.xlsx --wyjscie terminarz.pdf --okno 21
```

Rejestr: CSV (średnik + UTF-8) albo XLSX (pierwszy arkusz, nagłówek w pierwszym
wierszu), kolumny jak w faktury-do-rejestru:
`plik;sprzedawca;nip_sprzedawcy;nr_faktury;data_wystawienia;termin_platnosci;kwota_brutto;komentarz`.

## Jak to działa

- **Koszyki**: przeterminowane = termin < dzisiaj; „w tym tygodniu" = 7 dni od
  dnia startu okna (default: od dziś, czyli dziś + 7 dni); „zapewne następny
  tydzień" = od 8. dnia do końca okna (`--okno`, default 14). To okna kroczące,
  nie kalendarzowe tygodnie; żeby wynik nie zależał od tego, jaki jest dzień
  tygodnia.
- **Sort**: przeterminowane od najstarszych (najbardziej zaległe na górze),
  reszta od najbliższych terminów.
- **Okno od --z**: gdy start okna (`--z`) jest w przyszłości, wiersze między
  dzisiejszą datą a startem okna dostają osobny koszyk „przed oknem (do …)"
  z osobną sumą — rozdzielny od „poza oknem (po …)".
- **Daty**: ISO `RRRR-MM-DD`, `DD.MM.RRRR`, `DD.MM.RR`; walidacja kalendarza
  (31.02 to błąd, nie zgadujemy). Zła data w `--z`/`--do` = jasny komunikat
  i kod 1, nie traceback. Daty puste i „brak" pomijamy z licznikiem —
  „pominięte" jest jawnie wypisywane, nigdy po cichu.
- **Separator CSV**: średnik. CSV z przecinkami (eksport z Excela) jest
  rozpoznawany po nagłówku i odrzucany z podpowiedzią (kod 2), nie z mylącym
  „nie ma kolumn".
- **Kwoty**: polski format („1 845,00 zł”, „2 411,2”, „2829,00 złoty”), też
  kropka dziesiętna („1845.5”). Nigdy ciche zgadywanie: formaty niejednoznaczne
  są **odrzucane z powodem** — „1,234.56” (może US-style), „1.234” (kropka-
  tysiące czy grosze?), „1,234” (przecinek-tysiące czy grosze?), kwota > 10 mln
  zł („kwota podejrzana”). Gdy kropka i przecinek występują razem, czytelny
  jest tylko polski wariant „1.234,56”. Sumy liczone na `Decimal` z
  zaokrągleniem half-up (0,5 grosza idzie w górę).
- **PDF wielostronicowy**: przy powtórzonym nagłówku tabeli na stronach 2+
  wraca też etykieta sekcji („PRZETERMINOWANE — cd.”), żeby widać było,
  w którym koszyku jesteśmy.
- **XLSX**: czytany przez openpyxl, pierwszy arkusz, `data_only=True`
  (czyta wyniki formuł, nie formuły).
- **Zero AI, zero sieci**: to proste reguły na pliku, deterministycznie.

### Tabela pomiarów (uaktualniana przy zmianie logiki)

| Zbiór | n | przeterminowane | w tygodniu | następny tydzień | bez terminu | zła data | zła kwota | poza oknem |
|---|---|---|---|---|---|---|---|---|
| `testy/rejestr-okno.csv` | 8 | 1 | 4 | 2 | 1 | 0 | 0 | 0 |
| `testy/rejestr-mieszany.csv` | 10 | 2 | 1 | 2 | 1 | 1 | 3 | 0 |
| `testy/rejestr-duzy.csv` | 40 | 3 | 12 | 25 | 0 | 0 | 0 | 0 |

Pomiar z dnia 2026-10-26 (daty w fixture'ach są policzone względem tej daty).
`rejestr-mieszany` zawiera 3 odrzucenia kwot z powodem: „dwieście złotych"
(słownie), „1,234.56" (może US-style), „1.234" (kropka-tysiące czy grosze?);
przypadek „1845,00 złoty" jest liczony poprawnie (następny tydzień).
**Brak pomiaru = brak deklaracji.** Fixtures wygenerujesz skryptem
`python3 testy/generuj_fixtures.py` (daty liczone od dnia uruchomienia, więc
pomiar zawsze jest aktualny). Fikcyjne dane, poprawne formatowo.

## Czego to NIE robi

- **Nie deduplikuje**: rejestr przyjmuje 1:1 — jeśli masz duplikat, PDF pokaże
  go dwa razy (narzędzie tylko czyta, niczego nie zapisuje).
- **Nie wysyła niczego i nie planuje płatności**: to terminarz do oglądania;
  przelew robisz w banku, nie tu.
- **Nie liczy odsetek, nie pilnuje kompensat, walut obcych ani kursów** —
  kwota z rejestru jest brana jak jest.
- **Nie naprawia rejestru**: zła data albo zła kwota = wiersz pominięty
  z licznikiem i numerem wiersza, nigdy „ciche zgadywanie".
- **Nie zaprasza do sieci**: zero telemetrii, zero chmury, zero kont.

## RODO i dane

- Rejestr faktur to dane o transakcjach Twojej firmy; ten skrypt przetwarza je
  **wyłącznie lokalnie** (dysk → PDF/druk) i niczego nie wysyła.
- `.gitignore` wyklucza `*.csv`, `*.xlsx`, `*.pdf` **bez żadnych wyjątków** —
  Twoje dane nie powinny lądować w żadnym repozytorium. Fixtures do pomiarów
  są fikcyjne i w repo jest tylko skrypt, który je generuje
  (`testy/generuj_fixtures.py`), nie same pliki.
- Zabezpiecz rejestr i wygenerowany PDF jak teczkę z fakturami (dysk + backup).
- **To narzędzie nie jest doradcą prawnym ani księgowym.** Decyzje o
  przetwarzaniu danych i o płatnościach podejmujesz Ty.

## Struktura repo

```
terminarz_platnosci.py    # CLI (czytaj / przeglad / pdf) + font resolver
testy/generuj_fixtures.py # generuje rejestr-okno.csv + rejestr-mieszany.csv + rejestr-duzy.csv (daty od dnia uruchomienia, lokalne, nie w gicie)
testy/testy.py            # testy regresyjne (uruchom generuj_fixtures.py, potem: python3 testy/testy.py)
LICENSE                   # MIT
```

## Licencja

MIT. Rób co chcesz, bez gwarancji. Zanim poprawisz terminowe decyzje: sprawdź
wynik na swoim rejestrze i porównaj z własnym przeglądem ręcznym.
