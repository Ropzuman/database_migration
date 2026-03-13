# ACCESS 64-bit Declare Migration Summary

## Tiedosto

Access-kansion yhteenveto 64-bit `Declare`-korjauksista

## Päivämäärä

2026-03-13

## Kriittiset muutokset

- Koko aktiivinen `Access`-kansio auditoitiin `Declare`-lauseiden osalta.
- Kaikki ei-`PtrSafe`-declarit korjattiin `PtrSafe`-muotoon aktiivisissa lähdetiedostoissa.
- `GetUserName`/`GetComputerName`-kutsujen `nSize` pidettiin oikein `Long`-tyyppisenä (DWORD), ei `LongPtr`.
- OPENFILENAME-rakenteissa varmistettiin 64-bit-osoitinkentät (`lCustData`, `lpfnHook`) yhteensopiviksi.
- `Sleep`- ja `GetCursorPos`-declarit yhtenäistettiin 64-bit-käännöksen vaatimusten mukaisiksi.

## Siivous ja optimointi

- Muutokset rajattiin API-deklarointeihin ja UDT-osoitinkenttiin; liiketoimintalogiikkaa ei muutettu.
- Jokaiselle muokatulle tiedostolle luotiin oma `[OriginalFileName]_changelog.md` samaan kansioon.
- Markdown-muotoilu normalisoitiin (otsikot ja listojen välit), jotta dokumentit läpäisevät lint-tarkistuksen.

## Vaikutusalue

- Dokumentaatio päivitetty moduuleihin: `DOCUMENTS`, `FunctionDiagrams`, `Function_descriptions_html`, `instru3`, `LoopCircuit`, `Lukituskaavio`, `MAINEQ`, `PIPE`.
- `_ARCHIVE` jätettiin tarkoituksella koskematta (vain read-only auditointi).

## Suositeltu varmistus

- Aja Accessissa `Debug -> Compile` koko projektille.
- Jos compile pysähtyy, korjaa seuraava ilmoitettu rivi prioriteettijärjestyksessä:
  1. API-declare allekirjoitus
  2. Viitekirjasto (DAO/ActiveX)
  3. Tyyppiristiriita (`Long` vs `LongPtr`)
