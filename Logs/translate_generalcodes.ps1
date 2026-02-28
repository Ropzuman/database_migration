# Translate English comment headers in MAINEQ/GeneralCodes.bas and other large files
$ae = [char]0xe4; $AE = [char]0xc4; $oe = [char]0xf6; $OE = [char]0xd6; $ao = [char]0xe5

function Translate-File {
    param([string]$path, [hashtable]$subs)
    $content = Get-Content $path -Encoding UTF8 -Raw
    $changed = $false
    foreach ($kv in $subs.GetEnumerator()) {
        if ($content.Contains($kv.Key)) {
            $content = $content.Replace($kv.Key, $kv.Value)
            $changed = $true
        }
    }
    if ($changed) {
        [System.IO.File]::WriteAllText($path, $content, [System.Text.Encoding]::UTF8)
        Write-Host "Updated: $path"
    } else {
        Write-Host "No changes: $path"
    }
}

# ============ MAINEQ/GeneralCodes.bas ============
$path = "c:\database_migration\Access\MAINEQ\GeneralCodes.bas"
Translate-File $path @{
    "' Module: GeneralCodes" = "' Moduuli: GeneralCodes"
    "' Purpose: General utility functions for calculations, database queries, and UI" = "' Tarkoitus: Yleiset apufunktiot laskentaan, tietokantakyselyihin ja k${ae}ytt${oe}liittym${ae}${ae}n"
    "' Updated: 2025-11-11 - Added error handling, DAO typing, replaced custom`r`n'                       Replace() with VBA built-in, comprehensive comments" = "' P${ae}ivitetty: 2025-11-11 - Lis${ae}tty virheenk${ae}sittely, DAO-tyypitys, korvattu mukautettu`r`n'                       Replace() VBA:n sis${ae}isell${ae}, kattavat kommentit"
    "' Function: IsLoaded" = "' Funktio: IsLoaded"
    "' Purpose: Check if a form is currently open in Form or Datasheet view" = "' Tarkoitus: Tarkistaa onko lomake auki Form- tai Datasheet-n${ae}kym${ae}ss${ae}"
    "'   strFormName - Name of the form to check" = "'   strFormName - Tarkistettavan lomakkeen nimi"
    "' Rerturns: True if form is open and not in design view" = "' Palauttaa: True jos lomake on auki eik${ae} suunnittelun${ae}kym${ae}ss${ae}"
    "' Updated: 2025-11-11 - Added error handling and comments" = "' P${ae}ivitetty: 2025-11-11 - Lis${ae}tty virheenk${ae}sittely ja kommentit"
    "    ' Check if form is open (not closed)" = "    ' Tarkistetaan onko lomake auki (ei suljettu)"
    "        ' Check if form is not in design view" = "        ' Tarkistetaan ettei lomake ole suunnittelun${ae}kym${ae}ss${ae}"
    "    ' Form doesn't exist or error occurred - return False" = "    ' Lomaketta ei ole tai tapahtui virhe - palautetaan False"
    "' Function: HaeViimPaiva" = "' Funktio: HaeViimPaiva"
    "' Purpose: Extract the most recent revision date from multi-line revision text" = "' Tarkoitus: Poimii uusimman revision p${ae}iv${ae}m${ae}${ae}r${ae}n monirivisen revision tekstist${ae}"
    "'   Revision - Multi-line revision history string (lines separated by vbCrLf)" = "'   Revision - Monirivisinen revisiohistoria (rivit eroteltu vbCrLf:ll${ae})"
    "' Returns: Date portion of the most recent revision (last line)" = "' Palauttaa: Uusimman revision (viimeinen rivi) p${ae}iv${ae}m${ae}${ae}r${ae}osa"
    "' Notes: Assumes format `"REV DATE/MAKER/...`" with vbCrLf between entries" = "' Huomiot: Olettaa muodon `"REV PVM/TEKIJ${AE}/...`" vbCrLf-erotuksella"
    "' Updated: 2025-11-11 - Added error handling and detailed comments" = "' P${ae}ivitetty: 2025-11-11 - Lis${ae}tty virheenk${ae}sittely ja yksityiskohtaiset kommentit"
    "  ' Extract date portion (between space and first slash)" = "  ' Poimitaan p${ae}iv${ae}m${ae}${ae}r${ae}osa (v${ae}lilyönnin ja ensimm${ae}isen vinoviivan v${ae}lill${ae})"
    "' NOTE: Custom Replace() function REMOVED 2025-11-11" = "' HUOMIO: Mukautettu Replace()-funktio POISTETTU 2025-11-11"
    "' The custom Replace() function below has been removed because VBA has provided" = "' Alla oleva mukautettu Replace()-funktio on poistettu, koska VBA on tarjonnut"
    "' a built-in Replace() function since VBA 6.0 (Office 2000+)." = "' sis${ae}isen Replace()-funktion VBA 6.0:sta l${ae}htien (Office 2000+)."
    "' VBA Built-in Replace() Syntax:" = "' VBA:n sis${ae}isen Replace()-funktion syntaksi:"
    "'   Replace(expression, find, replace, [start], [count], [compare])" = "'   Replace(lauseke, etsit${ae}v${ae}, korvaava, [alku], [m${ae}${ae}r${ae}], [vertailu])"
    "' The built-in version is:" = "' Sis${ae}inen versio on:"
    "'   - Faster (compiled vs. interpreted VBA)" = "'   - Nopeampi (k${ae}${ae}nnetty C vs. tulkittu VBA)"
    "'   - Consistent with other VBA string functions" = "'   - Yhdenmukainen muiden VBA-merkkijonofunktioiden kanssa"
    "'   - Supports optional parameters for advanced control" = "'   - Tukee valinnaisia parametreja laajennettuun hallintaan"
    "' Original custom function behavior:" = "' Alkuper${ae}isen mukautetun funktion toiminta:"
    "' Equivalent using VBA built-in:" = "' Vastaava VBA:n sis${ae}isell${ae}:"
    "' REMOVED 2025-11-11: Custom Replace() function" = "' POISTETTU 2025-11-11: Mukautettu Replace()-funktio"
    "'   ' Custom implementation removed - using VBA built-in" = "'   ' Mukautettu toteutus poistettu - k${ae}ytet${ae}${ae}n VBA:n sis${ae}ist${ae}"
    "' Function: Optiot" = "' Funktio: Optiot"
    "' Purpose: Retrieve concatenated motor options for a given drive" = "' Tarkoitus: Hakee moottorin optiot yhdistettyn${ae} tietyll${ae} k${ae}ytölle"
    "'   Drives_ID - Drive ID to look up options for" = "'   Drives_ID - K${ae}ytön ID jonka optiot haetaan"
    "' Returns: Formatted string like `"+Option1 +Option2 +Option3`" or empty string" = "' Palauttaa: Muotoiltu merkkijono kuten `"+Optio1 +Optio2 +Optio3`" tai tyhj${ae}"
    "' Updated: 2025-11-11 - Added error handling, improved comments, fixed CurrentDB" = "' P${ae}ivitetty: 2025-11-11 - Lis${ae}tty virheenk${ae}sittely, parannettu kommentit, korjattu CurrentDB"
    "    ' Remove trailing `" +`"" = "    ' Poistetaan loppuosa `" +`""
    "' Function: Positiot" = "' Funktio: Positiot"
    "' Purpose: Retrieve customer positions for a given project element" = "' Tarkoitus: Hakee asiakkaan positiot tietyll${ae} projektielementille"
    "'   LaiteNr - Project element identifier" = "'   LaiteNr - Projektielementin tunniste"
    "' Returns: Formatted string like `"Pos: 01-M-01 / 01 and 01-M-02 / 01`"" = "' Palauttaa: Muotoiltu jono kuten `"Pos: 01-M-01 / 01 ja 01-M-02 / 01`""
    "' Notes: Joins MAINEQ and DRIVES tables to build position strings" = "' Huomiot: Yhdist${ae}${ae} MAINEQ- ja DRIVES-taulut positiomerkkijonojen rakentamiseksi"
    "' Build SQL query to join MAINEQ and DRIVES tables" = "' Rakennetaan SQL-kysely MAINEQ- ja DRIVES-taulujen yhdist${ae}miseksi"
    "    ' Remove trailing `" and `"" = "    ' Poistetaan loppuosa `" ja `""
    "' Function: Vaihekulma" = "' Funktio: Vaihekulma"
    "' Purpose: Calculate phase angle from power factor (cos ph)" = "' Tarkoitus: Lasketaan vaihekulma tehokertoimen (cos ph) perusteella"
    "' Purpose: Calculate phase angle from power factor (cos " + [char]0x03c6 + ")" = "' Tarkoitus: Lasketaan vaihekulma tehokertoimen (cos " + [char]0x03c6 + ") perusteella"
    "'   Cosfii - Power factor (cos ph)" = "'   Cosfii - Tehokerroin (cos ph)"
    "'   Cosfii - Power factor (cos " + [char]0x03c6 + ")" = "'   Cosfii - Tehokerroin (cos " + [char]0x03c6 + ")"
    "' Returns: Phase angle in radians" = "' Palauttaa: Vaihekulma radiaaneina"
    "' Notes: Uses arctangent mathematical formula" = "' Huomiot: K${ae}ytt${ae}${ae} arktangenttia matemaattisena kaavana"
    "' Calculate phase angle using: arctan(-cos" + [char]0x03c6 + " / sqrt(-cos" + [char]0x03c6 + [char]0x00b2 + " + 1)) + " + [char]0x03c0 + "/2" = "' Lasketaan vaihekulma: arctan(-cos" + [char]0x03c6 + " / sqrt(-cos" + [char]0x03c6 + [char]0x00b2 + " + 1)) + " + [char]0x03c0 + "/2"
    "' Function: MotKaapUh" = "' Funktio: MotKaapUh"
    "' Purpose: Calculate motor cable voltage drop percentage" = "' Tarkoitus: Lasketaan moottorin kaapelijohtimen j${ae}nnitteenalenema prosentteina"
    "'   Resist - Cable resistance (" + [char]0x03a9 + "/km)" = "'   Resist - Kaapelin resistanssi (" + [char]0x03a9 + "/km)"
    "'   React - Cable reactance (" + [char]0x03a9 + "/km)" = "'   React - Kaapelin reaktanssi (" + [char]0x03a9 + "/km)"
    "'   Virta - Current (A)" = "'   Virta - Virta (A)"
    "'   Voltage - Voltage (V)" = "'   Voltage - J${ae}nnite (V)"
    "'   Pituus - Cable length (m)" = "'   Pituus - Kaapelin pituus (m)"
    "' Returns: Formatted voltage drop percentage string (e.g., `"2.35 %`")" = "' Palauttaa: Muotoiltu j${ae}nnitteenalenema-prosenttimerkkijono (esim. `"2.35 %`")"
    "' Notes: Already has error handling (only function that did)" = "' Huomiot: Sis${ae}lt${ae}${ae} jo virheenk${ae}sittelyn (ainoa funktio jolla oli)"
    "' Updated: 2025-11-11 - Enhanced comments, standardized error handling" = "' P${ae}ivitetty: 2025-11-11 - Parannettu kommentit, standardoitu virheenk${ae}sittely"
    "' Calculate phase angle" = "' Lasketaan vaihekulma"
    "' Calculate voltage drop: " + [char]0x221a + "3 * I * (R*L*cos" + [char]0x03c6 + " + X*L*sin" + [char]0x03c6 + ")" = "' Lasketaan j${ae}nnitteenalenema: " + [char]0x221a + "3 * I * (R*L*cos" + [char]0x03c6 + " + X*L*sin" + [char]0x03c6 + ")"
    "' Convert to percentage of voltage" = "' Muunnetaan j${ae}nnitteen prosenteiksi"
    "' Format as percentage with 1-2 decimal places" = "' Muotoillaan prosenteiksi 1-2 desimaalin tarkkuudella"
    "' Function: LisaaNo" = "' Funktio: LisaaNo"
    "' Purpose: Add a number to a string and pad with leading zeros" = "' Tarkoitus: Lis${ae}${ae} luku merkkijonoon ja t${ae}ytt${ae}${ae} etunollilla"
    "'   Tieto - Original numeric string (e.g., `"001`")" = "'   Tieto - Alkuper${ae}inen numeerinen merkkijono (esim. `"001`")"
    "'   Lisays - Number to add (e.g., 100)" = "'   Lisays - Lis${ae}tt${ae}v${ae} luku (esim. 100)"
    "' Returns: Padded result string (e.g., `"101`")" = "' Palauttaa: Etunollilla t${ae}ytetty tulos (esim. `"101`")"
    "' Notes: Preserves original string length with zero-padding" = "' Huomiot: S${ae}ilytt${ae}${ae} alkuper${ae}isen merkkijonon pituuden etunollilla"
    "' Example: LisaaNo(`"001`", 100) = `"101`"" = "' Esimerkki: LisaaNo(`"001`", 100) = `"101`""
    "' Updated: 2025-11-11 - Added error handling and detailed comments" = "' P${ae}ivitetty: 2025-11-11 - Lis${ae}tty virheenk${ae}sittely ja yksityiskohtaiset kommentit"
    "  Pit = Len(Tieto)          ' Get original length" = "  Pit = Len(Tieto)          ' Haetaan alkuper${ae}inen pituus"
    "  No = Val(Tieto)           ' Convert to number" = "  No = Val(Tieto)           ' Muunnetaan numeroksi"
    "  No = No + Lisays          ' Add the increment" = "  No = No + Lisays          ' Lis${ae}t${ae}${ae}n kasvattaja"
    "  LisaaNo = CStr(No)        ' Convert back to string" = "  LisaaNo = CStr(No)        ' Muunnetaan takaisin merkkijonoksi"
    "    ' Pad with leading zeros to maintain original length" = "    ' T${ae}ytet${ae}${ae}n etunollilla alkuper${ae}isen pituuden s${ae}ilytt${ae}miseksi"
    "'Example usage in query:" = "'Esimerkki k${ae}yt${oe}st${ae} kyselyss${ae}:"
    "'   Field: LisaaNo([FieldName], 100)" = "'   Kent${ae}n arvo: LisaaNo([KentanNimi], 100)"
}

Write-Host "Done."
