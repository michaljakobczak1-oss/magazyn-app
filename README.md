# Magazyn – aplikacja do zarządzania sprzętem

Prosta aplikacja webowa: rejestr sprzętu, rezerwacje z kontrolą kolizji terminów, stany magazynowe, protokoły wydania/przyjęcia (PDF), wielu użytkowników z rolami.

## Funkcje

- Rejestr sprzętu: kod, numer projektu, nazwa, wymiary, zdjęcie, miejsce w magazynie, ilość sztuk
- Rezerwacje na termin – system blokuje rezerwację, jeśli w danym terminie brak wolnych sztuk (uwzględnia inne rezerwacje i wydany sprzęt)
- Stany magazynowe i dostępność "na dziś" per sprzęt
- Oznaczanie wydania i zwrotu przez obsługę (z zapisem kto i kiedy)
- Protokoły wydania (WZ) i przyjęcia (PZ) jako PDF – z kodem, miejscem w magazynie i zdjęciem
- Logowanie, hasła hashowane, role: admin (zarządza sprzętem i kontami) i użytkownik

## Uruchomienie – Docker (zalecane na serwerze)

```bash
docker compose up -d --build
```

Aplikacja: http://adres-serwera:5000

## Uruchomienie – bez Dockera

```bash
pip install -r requirements.txt
python app.py                # tryb prosty, port 5000
# lub produkcyjnie:
gunicorn -b 0.0.0.0:5000 -w 2 app:app
```

## Przykładowe dane (opcjonalnie)

```bash
python seed.py                          # bez Dockera
docker compose exec magazyn python seed.py   # z Dockerem
```

Doda 3 przykładowe produkty (CC0039, CC0041, CC0043) z placeholderami zdjęć – prawdziwe zdjęcia podmienisz przez „Edytuj".

## Pierwsze logowanie

Login: **admin**, hasło: **admin123**

**Po pierwszym logowaniu zmień hasło admina** (zakładka Użytkownicy) i ustaw zmienną środowiskową `SECRET_KEY` na losowy ciąg (w `docker-compose.yml` lub w środowisku).

## Dane

- Baza: `data/magazyn.db` (SQLite – wystarczy kopiować ten plik jako backup)
- Zdjęcia: `static/uploads/`

Oba katalogi są montowane jako wolumeny w docker-compose, więc dane przetrwają aktualizację kontenera.
