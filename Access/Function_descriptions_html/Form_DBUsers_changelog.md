# Muutosloki

- **Tiedosto:** `Form_DBUsers.cls`
- **P‰iv‰m‰‰r‰:** 2026-11-10

## Kriittiset muutokset

- `WhosOn`: `While ... And i <= 32` korvattu `Do While`-silmukalla, jossa indeksi tarkistetaan ennen taulukon lukua. Aiemmin 32 tavun nimi aiheutti virheen 9 (VBA ei oikosulje `And`-ehtoa).
- `.LACCDB`-polku muodostetaan `InStrRev`-funktiolla, joten polun pisteet eiv‰t riko tiedostonime‰.

## Siivous ja optimointi

- Kommentti korjattu: lukitustiedoston tietue on 64 tavua (ei bitti‰).
- Muuttuja `X` nimetty uudelleen `sTiedosto`.
