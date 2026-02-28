# Translate MAINEQ/GeneralCodes.bas English comments - v2 (simple sequential replaces)
$ae = [char]0xe4; $oe = [char]0xf6; $AE = [char]0xc4

function r { param($s) $s.Replace('${ae}',$ae).Replace('${oe}',$oe) }

$path = "c:\database_migration\Access\MAINEQ\GeneralCodes.bas"
$c = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
$orig = $c

# Function header keywords
$c = $c.Replace("' Module: GeneralCodes", "' Moduuli: GeneralCodes")
$c = $c.Replace("' Purpose: General utility functions for calculations, database queries, and UI", (r "' Tarkoitus: Yleiset apufunktiot laskentaan, tietokantakyselyihin ja k${ae}ytt${oe}liittym${ae}${ae}n"))
$c = $c.Replace("' Function: IsLoaded", "' Funktio: IsLoaded")
$c = $c.Replace("' Purpose: Check if a form is currently open in Form or Datasheet view", (r "' Tarkoitus: Tarkistaa onko lomake auki Form- tai Datasheet-n${ae}kym${ae}ss${ae}"))
$c = $c.Replace("'   strFormName - Name of the form to check", (r "'   strFormName - Tarkistettavan lomakkeen nimi"))
$c = $c.Replace("' Rerturns: True if form is open and not in design view", (r "' Palauttaa: True jos lomake on auki eik${ae} suunnittelun${ae}kym${ae}ss${ae}"))
$c = $c.Replace("    ' Check if form is open (not closed)", "    ' Tarkistetaan onko lomake auki (ei suljettu)")
$c = $c.Replace("        ' Check if form is not in design view", (r "        ' Tarkistetaan ett${ae}i lomake ole suunnittelun${ae}kym${ae}ss${ae}"))
$c = $c.Replace("    ' Form doesn't exist or error occurred - return False", "    ' Lomaketta ei ole tai tapahtui virhe - palautetaan False")
$c = $c.Replace("' Function: HaeViimPaiva", "' Funktio: HaeViimPaiva")
$c = $c.Replace("' Purpose: Extract the most recent revision date from multi-line revision text", (r "' Tarkoitus: Poimii uusimman revision p${ae}iv${ae}m${ae}${ae}r${ae}n monirivisen revision tekstist${ae}"))
$c = $c.Replace("'   Revision - Multi-line revision history string (lines separated by vbCrLf)", (r "'   Revision - Monirivisinen revisiohistoria (rivit eroteltu vbCrLf:ll${ae})"))
$c = $c.Replace("' Returns: Date portion of the most recent revision (last line)", (r "' Palauttaa: Uusimman revision (viimeinen rivi) p${ae}iv${ae}m${ae}${ae}r${ae}osa"))
$c = $c.Replace("  ' Extract date portion (between space and first slash)", (r "  ' Poimitaan p${ae}iv${ae}m${ae}${ae}r${ae}osa (v${ae}lilyonnin ja ensimm${ae}isen vinoviivan v${ae}lill${ae})"))
$c = $c.Replace("' NOTE: Custom Replace() function REMOVED 2025-11-11", "' HUOMIO: Mukautettu Replace()-funktio POISTETTU 2025-11-11")
$c = $c.Replace("' The custom Replace() function below has been removed because VBA has provided", "' Alla oleva mukautettu Replace()-funktio on poistettu, koska VBA on tarjonnut")
$c = $c.Replace("' a built-in Replace() function since VBA 6.0 (Office 2000+).", (r "' sis${ae}isen Replace()-funktion VBA 6.0:sta l${ae}htien (Office 2000+)."))
$c = $c.Replace("' VBA Built-in Replace() Syntax:", (r "' VBA:n sis${ae}isen Replace()-funktion syntaksi:"))
$c = $c.Replace("' The built-in version is:", (r "' Sis${ae}inen versio on:"))
$c = $c.Replace("'   - Faster (compiled vs. interpreted VBA)", "   ' - Nopeampi (k${ae}${ae}nnetty vs. tulkittu VBA)".Replace('${ae}',$ae))
$c = $c.Replace("'   - Consistent with other VBA string functions", "   ' - Yhdenmukainen muiden VBA-merkkijonofunktioiden kanssa")
$c = $c.Replace("'   - Supports optional parameters for advanced control", "   ' - Tukee valinnaisia parametreja laajennettuun hallintaan")
$c = $c.Replace("' Original custom function behavior:", (r "' Alkuper${ae}isen mukautetun funktion toiminta:"))
$c = $c.Replace("' Equivalent using VBA built-in:", (r "' Vastaava VBA:n sis${ae}isell${ae}:"))
$c = $c.Replace("' REMOVED 2025-11-11: Custom Replace() function", "' POISTETTU 2025-11-11: Mukautettu Replace()-funktio")
$c = $c.Replace("' Function: Optiot", "' Funktio: Optiot")
$c = $c.Replace("' Purpose: Retrieve concatenated motor options for a given drive", (r "' Tarkoitus: Hakee moottorin optiot yhdistettyn${ae} tietyll${ae} k${ae}yt${oe}lle"))
$c = $c.Replace("'   Drives_ID - Drive ID to look up options for", (r "'   Drives_ID - K${ae}yt${oe}n ID jonka optiot haetaan"))
$c = $c.Replace("    ' Remove trailing `" +`"", "    ' Poistetaan loppuosa `" +`"")
$c = $c.Replace("' Function: Positiot", "' Funktio: Positiot")
$c = $c.Replace("' Purpose: Retrieve customer positions for a given project element", (r "' Tarkoitus: Hakee asiakkaan positiot tietyll${ae} projektielementille"))
$c = $c.Replace("'   LaiteNr - Project element identifier", "   ' LaiteNr - Projektielementin tunniste")
$c = $c.Replace("' Notes: Joins MAINEQ and DRIVES tables to build position strings", (r "' Huomiot: Yhdist${ae}${ae} MAINEQ- ja DRIVES-taulut positiomerkkijonojen rakentamiseksi"))
$c = $c.Replace("' Build SQL query to join MAINEQ and DRIVES tables", (r "' Rakennetaan SQL-kysely MAINEQ- ja DRIVES-taulujen yhdist${ae}miseksi"))
$c = $c.Replace("    ' Remove trailing `" and `"", "    ' Poistetaan loppuosa `" ja `"")
$c = $c.Replace("' Function: Vaihekulma", "' Funktio: Vaihekulma")
$c = $c.Replace("' Returns: Phase angle in radians", "' Palauttaa: Vaihekulma radiaaneina")
$c = $c.Replace("' Notes: Uses arctangent mathematical formula", (r "' Huomiot: K${ae}ytt${ae}${ae} arktangenttia matemaattisena kaavana"))
$c = $c.Replace("' Calculate phase angle", "' Lasketaan vaihekulma")
$c = $c.Replace("' Convert to percentage of voltage", (r "' Muunnetaan j${ae}nnitteen prosenteiksi"))
$c = $c.Replace("' Format as percentage with 1-2 decimal places", "' Muotoillaan prosenteiksi 1-2 desimaalin tarkkuudella")
$c = $c.Replace("' Function: MotKaapUh", "' Funktio: MotKaapUh")
$c = $c.Replace("' Purpose: Calculate motor cable voltage drop percentage", (r "' Tarkoitus: Lasketaan moottorin kaapelin j${ae}nnitteenalenema prosentteina"))
$c = $c.Replace("'   Virta - Current (A)", "   ' Virta - Virta (A)")
$c = $c.Replace("'   Voltage - Voltage (V)", (r "'   Voltage - J${ae}nnite (V)"))
$c = $c.Replace("'   Pituus - Cable length (m)", "   ' Pituus - Kaapelin pituus (m)")
$c = $c.Replace("' Notes: Already has error handling (only function that did)", (r "' Huomiot: Sis${ae}lt${ae}${ae} jo virheenk${ae}sittelyn (ainoa funktio jolla oli)"))
$c = $c.Replace("' Updated: 2025-11-11 - Enhanced comments, standardized error handling", (r "' P${ae}ivitetty: 2025-11-11 - Parannettu kommentit, standardoitu virheenk${ae}sittely"))
$c = $c.Replace("' Function: LisaaNo", "' Funktio: LisaaNo")
$c = $c.Replace("' Purpose: Add a number to a string and pad with leading zeros", (r "' Tarkoitus: Lis${ae}${ae} luku merkkijonoon ja t${ae}ytt${ae}${ae} etunollilla"))
$c = $c.Replace("' Notes: Preserves original string length with zero-padding", (r "' Huomiot: S${ae}ilytt${ae}${ae} alkuper${ae}isen merkkijonon pituuden etunollilla"))
$c = $c.Replace("' Updated: 2025-11-11 - Added error handling, improved comments, fixed CurrentDB", (r "' P${ae}ivitetty: 2025-11-11 - Lis${ae}tty virheenk${ae}sittely, parannettu kommentit, korjattu CurrentDB"))
$c = $c.Replace("  Pit = Len(Tieto)          ' Get original length", (r "  Pit = Len(Tieto)          ' Haetaan alkuper${ae}inen pituus"))
$c = $c.Replace("  No = Val(Tieto)           ' Convert to number", "  No = Val(Tieto)           ' Muunnetaan numeroksi")
$c = $c.Replace('  No = No + Lisays          ' + "'" + ' Add the increment', (r "  No = No + Lisays          ' Lis${ae}t${ae}${ae}n kasvattaja"))
$c = $c.Replace("  LisaaNo = CStr(No)        ' Convert back to string", "  LisaaNo = CStr(No)        ' Muunnetaan takaisin merkkijonoksi")
$c = $c.Replace("    ' Pad with leading zeros to maintain original length", (r "    ' T${ae}ytet${ae}${ae}n etunollilla alkuper${ae}isen pituuden s${ae}ilytt${ae}miseksi"))
$c = $c.Replace("'Example usage in query:", (r "' Esimerkki k${ae}yt${oe}st${ae} kyselyss${ae}:"))

if ($c -ne $orig) {
    [System.IO.File]::WriteAllText($path, $c, [System.Text.Encoding]::UTF8)
    Write-Host "UPDATED: GeneralCodes.bas"
    # Count changes
    $changes = 0
    $lines1 = $orig -split "`n"
    $lines2 = $c -split "`n"
    for ($i = 0; $i -lt [Math]::Min($lines1.Count, $lines2.Count); $i++) {
        if ($lines1[$i] -ne $lines2[$i]) { $changes++ }
    }
    Write-Host "Lines changed: ~$changes"
} else {
    Write-Host "NO CHANGES"
}
