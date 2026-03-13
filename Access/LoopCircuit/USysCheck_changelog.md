# Muutosloki

## Tiedosto

USysCheck.bas

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- `wu_GetUserName` ja `GetOpenFileName` -declarit päivitetty `PtrSafe`-muotoon kaikissa haaroissa.
- OPENFILENAME-rakenteen osoitinkentät (`lCustData`, `lpfnHook`) varmistettu 64-bit-yhteensopiviksi.

## Siivous ja optimointi

- API-yhteensopivuus vakioitu ilman lomakelogiikan muutoksia.
- Ylläpidettävyys parani deklarointien yhdenmukaistuksen kautta.
