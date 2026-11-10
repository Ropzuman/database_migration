# Muutosloki

- **Tiedosto:** `KAANNOS.bas`
- **P‰iv‰m‰‰r‰:** 2026-11-10

## Kriittiset muutokset

- `Kaanna`: puuttuva sulkeva `}` ei en‰‰ aiheuta virhett‰ tai v‰‰r‰‰ tulosta.
- `{POS}` ilman kuvausta k‰sitell‰‰n oikein (kuvaus tyhj‰).
- Positio, jossa on alle kolme osaa, ei kaada funktiota vaan palauttaa `[ERR: Not found]`.
- Seuraavan viittauksen alkusulku `{` s‰ilyy tuloksessa.

## Siivous ja optimointi

- Kuvausteksti haetaan kerran muuttujaan `sKuvaus` (aiemmin kolme toistuvaa `Mid$`-lauseketta).
- Palautusarvon kommentti korjattu.
