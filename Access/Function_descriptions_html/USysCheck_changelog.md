# Muutosloki

## Tiedosto

USysCheck.bas

## Päivämäärä

2026-11-10

## Kriittiset muutokset

- API-aliakset `GetUserNameW`/`GetComputerNameW` vaihdettu `…A`-muotoon: `ByVal String` -puskuri on VBA:ssa ANSI, joten W-versio tuotti typistyneen tai sotkuisen nimen.
- Merkkijonon päättäminen
: `InStr(NBuffer & vbNullChar, vbNullChar)` estää `Left$`-virheen, jos nollamerkkiä ei löydy.
- `OpenRecordset("UsysUsers", dbOpenDynaset, dbAppendOnly)` toimii myös, kun `UsysUsers` on linkitetty taulu (aiemmin `dbOpenTable`).

## Siivous ja optimointi

- `DB` ja `Taulu` tyypitetty `DAO.Database`/`DAO.Recordset` (aiemmin `Object`).
- Virheenkäsittelijä kirjaa virheen Immediate-ikkunaan (`Debug.Print`), vaikka kirjaus pysyy hiljaisena käyttäjälle.

## Aiemmat muutokset (2026-03-13)

- `api_GetUserName` ja `api_GetComputerName` -declarit päivitetty `PtrSafe`-muotoon.
- 64-bit-käännöstä rikkova ei-PtrSafe-Declare poistettu käytöstä.
