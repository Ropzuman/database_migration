# Muutosloki

## Tiedosto

APIKoodit.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `wu_GetUserName` ja `GetOpenFileName` -declarit yhtenäistetty `PtrSafe`-muotoon.
- OPENFILENAME-rakenteen osoitinkentät varmistettu 64-bit-kelpoisiksi.
- API-kutsuissa käytetyt kahva- ja osoitintyypit pidetty `LongPtr`-muodossa.

## Siivous ja optimointi

- Deklaraatiot harmonisoitu moduulin sisällä.
- Muu lukituskaavion liiketoimintalogiikka jätetty koskematta.
