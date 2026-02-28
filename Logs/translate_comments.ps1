# Comprehensive English-to-Finnish comment translation script
# Run with: powershell -ExecutionPolicy Bypass -File translate_comments.ps1

$ae = [char]0xe4  # ä
$AE = [char]0xc4  # Ä
$oe = [char]0xf6  # ö
$OE = [char]0xd6  # Ö

$totalChanges = 0
$changedFiles = [System.Collections.Generic.List[string]]::new()

function Apply-Replacements {
    param([string]$filePath, [hashtable]$replacements)
    
    $lines = Get-Content $filePath -Encoding UTF8 -ErrorAction SilentlyContinue
    if ($null -eq $lines) { return }
    
    $changed = $false
    $newLines = for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $trimmed = $line.TrimStart()
        $matched = $false
        foreach ($kv in $replacements.GetEnumerator()) {
            if ($trimmed -eq $kv.Key) {
                $indent = $line.Substring(0, $line.Length - $line.TrimStart().Length)
                $line = $indent + $kv.Value
                $changed = $true
                $matched = $true
                break
            }
        }
        $line
    }
    
    if ($changed) {
        $newLines | Set-Content $filePath -Encoding UTF8
        $script:totalChanges++
        $script:changedFiles.Add($filePath)
    }
}

# ================== DOCUMENTS/GlobalVBAs.vba ==================
$f = "c:\database_migration\Access\DOCUMENTS\GlobalVBAs.vba"
Apply-Replacements $f @{
    "' Extracts the original author name from a multi-line revision string."                            = "' Poimii alkuper${ae}isen tekij${ae}nimen monirivisen revisiojonon uit."
    "' Parses backward to find the first (oldest) revision entry."                                      = "' Selaa taaksep${ae}in l${oe}yt${ae}${ae}kseen ensimm${ae}isen (vanhimman) revisiorivin."
    "' @param Revisio: Revision string with format `"Rev Date/Author/Checker/...`" separated by vbCrLf" = "' @param Revisio: Revisiojonon muoto `"Rev Pvm/Tekij${ae}/Tarkastaja/...`" eroteltu vbCrLf:ll${ae}"
    "' @return Author name from the first revision, or empty string if Null"                            = "' @return Tekij${ae}nimi ensimm${ae}isest${ae} revisiosta tai tyhj${ae} merkkijono jos Null"
    "' Extracts the revision mark (e.g., `"A`", `"B`", `"0`") from revision string."                    = "' Poimii revisiomerkinn${ae}n (esim. `"A`", `"B`", `"0`") revisiojonosta."
    "' @param Revisio: Revision string with format `"Rev Date/Author/...`""                             = "' @param Revisio: Revisiojonon muoto `"Rev Pvm/Tekij${ae}/...`""
    "' @return Revision mark before the first space, or empty string if Null"                           = "' @return Revisiomerkki ennen ensimm${ae}ist${ae} v${ae}lily${oe}nti${ae} tai tyhj${ae} merkkijono jos Null"
    "' Extracts the date from the first (oldest) revision entry."                                       = "' Poimii p${ae}iv${ae}m${ae}${ae}r${ae}n ensimm${ae}isest${ae} (vanhimmasta) revisiorivist${ae}."
    "' Parses backward through multi-line revision string to find original date."                       = "' Selaa taaksep${ae}in monirivisen revisiojonon l${ae}pi l${oe}yt${ae}${ae}kseen alkuper${ae}isen p${ae}iv${ae}m${ae}${ae}r${ae}n."
    "' @param Revisio: Revision string with format `"Rev Date/Author/...`" separated by vbCrLf"         = "' @param Revisio: Revisiojonon muoto `"Rev Pvm/Tekij${ae}/...`" eroteltu vbCrLf:ll${ae}"
    "' @return Date string from the first revision"                                                     = "' @return P${ae}iv${ae}m${ae}${ae}r${ae}merkkijono ensimm${ae}isest${ae} revisiosta"
}

# ================== DOCUMENTS/Form_DISTRIBUTION.cls ==================
$f = "c:\database_migration\Access\DOCUMENTS\Form_DISTRIBUTION.cls"
Apply-Replacements $f @{
    "' @param Ehto - Distribution ID filter (if empty, generates report for all documents)" = "' @param Ehto - Jakelun ID-suodatin (tyhj${ae}n${ae} generoi raportin kaikille asiakirjoille)"
}

# ================== Function_descriptions_html/Form_FrmMUOKKAUS.cls ==================
$f = "c:\database_migration\Access\Function_descriptions_html\Form_FrmMUOKKAUS.cls"
Apply-Replacements $f @{
    "'   Generates both Word (.docx) and HTML documents from database templates."       = "'   Generoi sek${ae} Word (.docx) ett${ae} HTML-asiakirjoja tietokantapohjista."
    "'   Handles loops (PIIRIT) and equipment (MOOTTORIT) descriptions."                = "'   K${ae}sittelee piiri- (PIIRIT) ja laitesanalliskuvauksia (MOOTTORIT)."
    "'   2. Form reads data from TIEDOT_PIIRIT or TIEDOT_MOOTTORIT tables"              = "'   2. Lomake lukee tiedot tauluista TIEDOT_PIIRIT tai TIEDOT_MOOTTORIT"
    "'      - {LOOPTAG} inserts hyperlinked loop references"                            = "'      - {LOOPTAG} lis${ae}${ae} hyperlinkatut piiriviittaukset"
    "'   4. For Word: Opens template, replaces bookmarks, saves as new document"        = "'   4. Wordille: Avaa pohjan, korvaa kirjanmerkit, tallentaa uutena asiakirjana"
    "'   5. For HTML: Generates HTML file with CSS styling and hyperlinks"              = "'   5. HTML:lle: Generoi HTML-tiedoston CSS-tyylein ja hyperlinkein"
    "' Template Bookmarks (Word):"                                                      = "' Word-pohjan kirjanmerkit:"
    "'   - Queries: LUOPIIRIT (creates new loop descriptions)"                          = "'   - Kyselyt: LUOPIIRIT (luo uusia piirikuvauksia)"
    "'   Creates HTML file with loop/motor function description, including:"            = "'   Luo HTML-tiedoston piiri-/moottoritoimintokuvaukselle, sis${ae}lt${ae}en:"
    "'   Converts text field content to HTML format:"                                   = "'   Muuntaa tekstikent${ae}n sis${ae}ll${oe}n HTML-muotoon:"
    "'   Searches for {POS Description} patterns and converts to hyperlinks"            = "'   Etsii {POS Kuvaus} -rakenteita ja muuntaa ne hyperlinkeiksi"
    "'   Reads configuration from SETTINGS table including template paths,"             = "'   Lukee konfiguraation SETTINGS-taulusta sis${ae}lt${ae}en pohjatiedostopolut,"
    "' Tarkoitus: Fill Word template bookmarks with loop/equipment data"                = "' Tarkoitus: T${ae}ytt${ae}${ae} Word-pohjan kirjanmerkit piiri-/laitetiedoilla"
    "'   Replaces Word template bookmarks with data from current record."               = "'   Korvaa Word-pohjan kirjanmerkit nykyisen tietueen tiedoilla."
    "'   Handles headers (_REV, _STAT, _AUTH, etc.) and content fields"                 = "'   K${ae}sittelee otsikkokent${ae}t (_REV, _STAT, _AUTH jne.) ja sist${ae}lt${oe}kent${ae}t"
    "'   (_LTAG, _DESCR, _FUNCT, etc.). Generates alarm lists and equipment lists."     = "'   (_LTAG, _DESCR, _FUNCT jne.). Generoi h${ae}lytyslistan ja laitelistan."
    "'   Replaces bookmark text in document. Handles long text (>255 chars) separately" = "'   Korvaa kirjanmerkin tekstin asiakirjassa. K${ae}sittelee pitk${ae}n tekstin (>255 merkki${ae}) erikseen"
    "'   Replaces bookmark text and converts vbCrLf to Word paragraphs."                = "'   Korvaa kirjanmerkin tekstin ja muuntaa vbCrLf Word-kappaleiksi."
}

# ================== Function_descriptions_html/Form_FrmASETUKSET.cls ==================
$f = "c:\database_migration\Access\Function_descriptions_html\Form_FrmASETUKSET.cls"
Apply-Replacements $f @{
    "' Event: Command2_Click"                                                  = "' Tapahtuma: Command2_Click"
    "' Purpose: Link all required tables from external databases"              = "' Tarkoitus: Linkitet${ae}${ae}n kaikki tarvittavat taulut ulkoisista tietokannoista"
    "' Sub: Linkkaa (Link)"                                                    = "' Aliohjelma: Linkkaa (Linkitys)"
    "' Purpose: Link external table to current database"                       = "' Tarkoitus: Linkitet${ae}${ae}n ulkoinen taulu nykyiseen tietokantaan"
    "' Parameters:"                                                            = "' Parametrit:"
    "' Description:"                                                           = "' Kuvaus:"
    "'   Drops existing table if present, creates new linked table definition" = "'   Poistaa olemassa olevan taulun jos l${oe}ytyy, luo uuden linkitetyn taulun"
}

# ================== Function_descriptions_html/Form_PIIRIT subform.cls ==================
$f = "c:\database_migration\Access\Function_descriptions_html\Form_PIIRIT subform.cls"
Apply-Replacements $f @{
    "' Sub: LisaaTeksti (Add Text)"                                              = "' Aliohjelma: LisaaTeksti (Lis${ae}${ae} teksti${ae})"
    "' Purpose: Insert loop reference at cursor position"                        = "' Tarkoitus: Lis${ae}t${ae}${ae}n piiriviittaus kursorin kohtaan"
    "' Description:"                                                             = "' Kuvaus:"
    "'   Inserts text in {brackets} at current cursor position or end of field." = "'   Lis${ae}${ae} tekstin {hakasulkeissa} kursorin kohtaan tai kent${ae}n loppuun."
    "'   Handles null values and maintains cursor position after insertion."     = "'   K${ae}sittelee null-arvot ja s${ae}ilytt${ae}${ae} kursorin paikan lis${ae}${ae}misen j${ae}lkeen."
    "' Append to end"                                                            = "' Lis${ae}t${ae}${ae}n kent${ae}n loppuun"
    "' Insert at cursor position"                                                = "' Lis${ae}t${ae}${ae}n kursorin kohtaan"
}

# ================== Function_descriptions_html/Form_MOTORS subform.cls ==================
$f = "c:\database_migration\Access\Function_descriptions_html\Form_MOTORS subform.cls"
Apply-Replacements $f @{
    "'   Inserts text in {brackets} at current cursor position or end of field." = "'   Lis${ae}${ae} tekstin {hakasulkeissa} kursorin kohtaan tai kent${ae}n loppuun."
    "'   Handles null values and maintains cursor position after insertion."     = "'   K${ae}sittelee null-arvot ja s${ae}ilytt${ae}${ae} kursorin paikan lis${ae}${ae}misen j${ae}lkeen."
    "' Append to end"                                                            = "' Lis${ae}t${ae}${ae}n kent${ae}n loppuun"
    "' Insert at cursor position"                                                = "' Lis${ae}t${ae}${ae}n kursorin kohtaan"
}

# ================== Function_descriptions_html/GeneralCodes.bas ==================
$f = "c:\database_migration\Access\Function_descriptions_html\GeneralCodes.bas"
Apply-Replacements $f @{
    "'   Provides utility functions for equipment revision editing, table display," = "'   Tarjoaa apufunktiot laiterevision muokkaukseen, taulun n${ae}ytt${oe}${oe}n,"
    "'   Checks if active table is an equipment table, opens corresponding record"  = "'   Tarkistaa onko aktiivinen taulu laitenimike-taulu, avaa t${ae}m${ae}n  tietueen"
    "'   Translates mode codes:"                                                    = "'   K${ae}${ae}nt${ae}${ae} tilakoodit:"
    "'   Other codes pass through unchanged. Returns `"-`" for Null input."         = "'   Muut koodit l${ae}pik${ae}yt${ae}v${ae}t muuttumattomina. Palauttaa `"-`" Null-sy${oe}tteelle."
}

# ================== Function_descriptions_html/KAANNOS.bas ==================
$f = "c:\database_migration\Access\Function_descriptions_html\KAANNOS.bas"
Apply-Replacements $f @{
    "'   Processes text containing equipment references in braces {xx-xx-xx description}" = "'   K${ae}sittelee tekstin, joka sis${ae}lt${ae}${ae} laiteviittauksia hakasulkeissa {xx-xx-xx kuvaus}"
    "'   and translates them to actual equipment names from MAINEQ or Loops tables."      = "'   ja k${ae}${ae}nt${ae}${ae} ne todellisiksi laite- tai piirinimiksi MAINEQ- tai Loops-tauluista."
    "'   Validates references and marks deleted or missing items with error tags."        = "'   Validoi viittaukset ja merkitsee poistetut tai puuttuvat kohteet virhetunnisteilla."
    "'   Parses text for {POS description} patterns where:"                               = "'   J${ae}sentyyy tekstin {POS kuvaus} -rakenteille miss${ae}:"
    "' No references to translate"                                                        = "' Ei viittauksia k${ae}${ae}nnett${ae}v${ae}ksi"
    "' Initialize output with text before first reference"                                = "' Alustetaan tulostus tekstill${ae} ennen ensimm${ae}ist${ae} viittausta"
    "' Check for errors and mark accordingly"                                             = "' Tarkistetaan virheet ja merkitaan vastaavasti"
}

# ================== instru3/Form_CopyLoops.cls ==================
$f = "c:\database_migration\Access\instru3\Form_CopyLoops.cls"
Apply-Replacements $f @{
    "'   - Displays available loops from source database"          = "'   - N${ae}ytt${ae}${ae} saatavilla olevat piirit l${ae}hdetietokannasta"
    "'   - Handles duplicate loop prevention"                      = "'   - Estaa duplikaattipiirit"
    "'   - Provides detailed status feedback"                      = "'   - Tarjoaa yksityiskohtaista tilatietoa"
    "'   3. Iterate through all devTbl* tables in source database" = "'   3. K${ae}yd${ae}${ae}n l${ae}pi kaikki devTbl*-taulut l${ae}hdetietokannassa"
    "'   - Prevents duplicate loops (checks LOOPS table first)"    = "'   - Estaa duplikaattipiirit (tarkistaa LOOPS-taulun ensin)"
    "'   - Provides detailed status messages in TTiedot textbox"   = "'   - N${ae}ytt${ae}${ae} yksityiskohtaiset viestit TTiedot-tekstikent${ae}ss${ae}"
    "' Iterate through selected loops in subform"                  = "' K${ae}yd${ae}${ae}n l${ae}pi valitut piirit alilomakkeessa"
    "' Iterate through all records in source device table"         = "' K${ae}yd${ae}${ae}n l${ae}pi kaikki l${ae}hdelaitteen taulun tietueet"
    "'   - Form_Unload event will handle cleanup"                  = "'   - Form_Unload-tapahtuma k${ae}sittelee siivousty${oe}t"
    "'   - Creates temporary linked table for loop browsing"       = "'   - Luo v${ae}liaikaisen linkitetyn taulun piirien selaamista varten"
    "'   - Provides responsive UI for better loop viewing"         = "'   - Tarjoaa reagoivan k${ae}ytt${oe}liittym${ae}n piirien katseluun"
}

# ================== instru3/Form_Linkkien vaihto.cls ==================
$f = "c:\database_migration\Access\instru3\Form_Linkkien vaihto.cls"
Apply-Replacements $f @{
    "'   - Extracts database filename from current link path"      = "'   - Poimii tietokannan tiedostonimen nykyisest${ae} linkkipolusta"
    "'   - Prevents links to wrong directory after database moves" = "'   - Est${ae}${ae} linkit v${ae}${ae}r${ae}${ae}n hakemistoon tietokannan siirron j${ae}lkeen"
    "'   - Prevents duplicate relinking of same table"             = "'   - Est${ae}${ae} saman taulun uudelleen linkitt${ae}misen"
}

# ================== instru3/Form_SizingOut.cls ==================
$f = "c:\database_migration\Access\instru3\Form_SizingOut.cls"
Apply-Replacements $f @{
    "'   - Converts decimal commas to periods (Finnish ${ae} international format)" = "'   - Muuntaa desimaalipilkut pisteiksi (suomalainen ${ae} kansainv${ae}linen muoto)"
    "'   - Opens output folder in Windows Explorer after export"                    = "'   - Avaa tulostekansioin Windowsin Resurssienhallinnassa viennin j${ae}lkeen"
    "'   - Handles null values appropriately"                                       = "'   - K${ae}sittelee null-arvot asianmukaisesti"
    "'   4. Iterate through all records:"                                           = "'   4. K${ae}yd${ae}${ae}n l${ae}pi kaikki tietueet:"
    "'      - Handle null values as empty strings"                                  = "'      - K${ae}sitell${ae}${ae}n null-arvot tyhj${ae}n${ae} merkkijonoina"
    "'   - Handles Finnish decimal format (comma) conversion"                       = "'   - K${ae}sittelee suomalaisen desimaaliformaatin (pilkku) muunnoksen"
    "'   - Opens output folder for user convenience"                                = "'   - Avaa tulostekansioin k${ae}ytt${ae}j${ae}n mukavuuden vuoksi"
    "' Handle null values"                                                          = "' K${ae}sitell${ae}${ae}n null-arvot"
    "' Tarkoitus: Initialize form with default values"                              = "' Tarkoitus: Alustetaan lomake oletusarvoilla"
}

# ================== instru3/general.bas ==================
$f = "c:\database_migration\Access\instru3\general.bas"
Apply-Replacements $f @{
    "'   Provides utility functions for:"                                                    = "'   Tarjoaa apufunktioita:"
    "' Tarkoitus: Converts decimal comma to decimal point (Finnish to international format)" = "' Tarkoitus: Muuntaa desimaalipilkun desimaalipiste (suomalainen kansainv${ae}liseen muotoon)"
    "'   - Returns empty string if input is null or empty"                                   = "'   - Palauttaa tyhj${ae}n merkkijonon jos sy${oe}te on null tai tyhj${ae}"
    "' Handle null/empty input"                                                              = "' K${ae}sitell${ae}${ae}n null/tyhj${ae} sy${oe}te"
    "' Tarkoitus: Extracts revision number from user notes based on date"                    = "' Tarkoitus: Poimitaan revisionumero k${ae}ytt${ae}j${ae}n muistiinpanoista p${ae}iv${ae}m${ae}${ae}r${ae}n perusteella"
    "'   - Parses date from UdNote string (format: `"something:MM/DD/YYYY|something`")"      = "'   - J${ae}sentyyy p${ae}iv${ae}m${ae}${ae}r${ae} UdNote-merkkijonosta (muoto: `"jotain:KK/PP/VVVV|jotain`")"
    "'   - Returns first revision where BeforeDate > parsed date"                            = "'   - Palauttaa ensimm${ae}isen revision miss${ae} BeforeDate > j${ae}sennetty p${ae}iv${ae}m${ae}${ae}r${ae}"
    "' Handle null input"                                                                    = "' K${ae}sitell${ae}${ae}n null-sy${oe}te"
    "' Tarkoitus: Checks if a loop exists in the system"                                     = "' Tarkoitus: Tarkistetaan onko piiri olemassa j${ae}rjestelm${ae}ss${ae}"
    "'   - Queries qrysolvalve for matching AreaCode and LoopNo"                             = "'   - Kysyy qrysolvalve-kyselyst${ae} vastaavaa AreaCode- ja LoopNo-arvoille"
    "'   - Returns simple existence flag (not boolean for backward compatibility)"           = "'   - Palauttaa yksinkertaisen olemassaolon lipun (ei boolean taaksep${ae}inyhteensopivuuden vuoksi)"
}

# ================== LoopCircuit/Form_DBUsers.cls ==================
$f = "c:\database_migration\Access\LoopCircuit\Form_DBUsers.cls"
Apply-Replacements $f @{
    "' Reads the .LACCDB lock file to determine who's currently logged on" = "' Lukee .LACCDB-lukitustiedoston m${ae}${ae}ritt${ae}${ae}kseen k${ae}ytt${ae}j${ae}t jotka ovat kirjautuneet sis${ae}${ae}n"
    "' Extract machine name (null-terminated)"                             = "' Poimitaan koneennimi (null-p${ae}${ae}tteinen)"
    "' Extract user name (null-terminated)"                                = "' Poimitaan k${ae}ytt${ae}j${ae}nimi (null-p${ae}${ae}tteinen)"
}

# ================== LoopCircuit/Form_Tee Kuvat.cls ==================
$f = "c:\database_migration\Access\LoopCircuit\Form_Tee Kuvat.cls"
Apply-Replacements $f @{
    "' Initialize log textbox"                       = "' Alustetaan lokitekstikentt${ae}"
    "' Execute multiple queries from TKyselyt field" = "' Ajetaan useita kyselyit${ae} TKyselyt-kent${ae}st${ae}"
}

# ================== LoopCircuit/General.bas ==================
$f = "c:\database_migration\Access\LoopCircuit\General.bas"
Apply-Replacements $f @{
    "' Log error or handle silently" = "' Kirjataan virhe tai k${ae}sitell${ae}${ae}n hiljaisesti"
}

# ================== MAINEQ/DataToACAD.bas ==================
$f = "c:\database_migration\Access\MAINEQ\DataToACAD.bas"
Apply-Replacements $f @{
    "' Huomiot: Handles legacy naming convention with asterisk markers"                         = "' Huomiot: K${ae}sittelee vanhan nimeamiskonvention t${ae}htimerkkeineen"
    "'   1. Reads configuration from common table"                                              = "'   1. Lukee konfiguraation yhteisest${ae} taulusta"
    "'   2. Resets/initializes output .txt files"                                               = "'   2. Nollaa/alustaa .txt-tulostiedostot"
    "'   3. Generates non-loop-based lists"                                                     = "'   3. Generoi ei-piiripohjaisia listoja"
    "'   4. Generates loop-based lists (if applicable)"                                         = "'   4. Generoi piiripohjaisia listoja (tarvittaessa)"
    "'   5. Closes all files properly"                                                          = "'   5. Sulkee kaikki tiedostot asianmukaisesti"
    "'--- Reset/Initialize all output .txt files with opening parenthesis ---"                  = "'--- Nollataan/alustetaan kaikki .txt-tulostiedostot avaavalla sululla ---"
    "' Initialize non-loop-based files"                                                         = "' Alustetaan ei-piiripohjaiset tiedostot"
    "' Initialize loop-based files"                                                             = "' Alustetaan piiripohjaiset tiedostot"
    "' Tarkoitus: Generate LISP lists from tables/queries that don't require loop ID filtering" = "' Tarkoitus: Generoidaan LISP-listat tauluista/kyselyist${ae} ilman piiri-ID-suodatusta"
    "' Huomiot: Handles both single tables and wildcard table groups (e.g., `"CIRCUIT*`")"      = "' Huomiot: K${ae}sittelee sek${ae} yksitt${ae}iset taulut ett${ae} jokerimerkkitauluryhmiat (esim. `"CIRCUIT*`")"
    "'--- Handle wildcard table names (e.g., `"CIRCUIT*`") ---"                                 = "'--- K${ae}sitell${ae}${ae}n jokerimerkkiset taulunnimet (esim. `"CIRCUIT*`") ---"
    "' Loop through all tables matching the prefix"                                             = "' K${ae}yd${ae}${ae}n l${ae}pi kaikki etuliitteen vastaavat taulut"
    "'--- Handle single table/query names ---"                                                  = "'--- K${ae}sitell${ae}${ae}n yksitt${ae}iset taulu-/kyselynimet ---"
}

# ================== MAINEQ/Form_Revisiointi.cls ==================
$f = "c:\database_migration\Access\MAINEQ\Form_Revisiointi.cls"
Apply-Replacements $f @{
    "' Compares two tables and marks changed records with revision information" = "' Vertaa kahta taulua ja merkitsee muuttuneet tietueet revisiomerkinn${ae}ill${ae}"
    "' Compare values between the two tables (Nz handles nulls)"                = "' Vertaillaan arvoja kahden taulun v${ae}lill${ae} (Nz k${ae}sittelee null-arvot)"
}

# ================== MAINEQ/GeneralCodes.bas ==================
$f = "c:\database_migration\Access\MAINEQ\GeneralCodes.bas"
Apply-Replacements $f @{
    "'   - More robust (handles edge cases better)"                                      = "'   - Luotettavampi (k${ae}sittelee reunatapaukset paremmin)"
    "' If any code calls this function, it will now use the VBA built-in automatically." = "' Jos jokin koodi kutsuu t${ae}t${ae} funktiota, se k${ae}ytt${ae}${ae} nyt automaattisesti VBA:n sis${ae}ist${ae} funktiota."
    "' Handle null input"                                                                = "' K${ae}sitell${ae}${ae}n null-sy${oe}te"
}

# ================== PIPE/Form_DBUsers.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_DBUsers.cls"
Apply-Replacements $f @{
    "'   - Displays machine name and username for each connection" = "'   - N${ae}ytt${ae}${ae} koneennimen ja k${ae}ytt${ae}j${ae}nimen jokaiselle yhteydelle"
    "'   3. Open lock file in binary mode"                         = "'   3. Avataan lukitustiedosto binaaritilassa"
    "'   - Opens file with shared read access"                     = "'   - Avaa tiedoston jaetussa lukutilassa"
    "'   - Handles missing file gracefully"                        = "'   - K${ae}sittelee puuttuvan tiedoston asiallisesti"
    "' Extract machine name (null-terminated)"                     = "' Poimitaan koneennimi (null-p${ae}${ae}tteinen)"
    "' Extract user name (null-terminated)"                        = "' Poimitaan k${ae}ytt${ae}j${ae}nimi (null-p${ae}${ae}tteinen)"
}

# ================== PIPE/Form_frmOpenPIPELINE.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_frmOpenPIPELINE.cls"
Apply-Replacements $f @{
    "'   - Displays list of pipeline segments from PIPELINEDATA"      = "'   - N${ae}ytt${ae}${ae} putkil${ae}n segmenttilistan PIPELINEDATA-taulusta"
    "'   - Opens selected drawing and zooms to pipeline block"        = "'   - Avaa valitun piirustuksen ja zoomaa putkilinjablokkiin"
    "'   - Auto-opens if only 2 segments (one is header row)"         = "'   - Avautuu automaattisesti jos vain 2 segmentti${ae} (toinen on otsikkorivi)"
    "' Tarkoitus: Initialize form and populate pipeline segment list" = "' Tarkoitus: Alustetaan lomake ja t${ae}ytet${ae}${ae}n putkilinjasegmenttilistaus"
}

# ================== PIPE/Form_Linkkien vaihto.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_Linkkien vaihto.cls"
Apply-Replacements $f @{
    "'   - Extracts database filename from current link path"      = "'   - Poimii tietokannan tiedostonimen nykyisest${ae} linkkipolusta"
    "'   - Prevents links to wrong directory after database moves" = "'   - Est${ae}${ae} linkit v${ae}${ae}r${ae}${ae}n hakemistoon tietokannan siirron j${ae}lkeen"
    "'   - Prevents duplicate relinking of same table"             = "'   - Est${ae}${ae} saman taulun uudelleen linkitt${ae}misen"
    "' Initialize counter"                                         = "' Alustetaan laskuri"
}

# ================== PIPE/Form_TYÖKALUT.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_TYOKALUT.cls"
$fAlt = "c:\database_migration\Access\PIPE\Form_TYÖKALUT.cls"
foreach ($target in @($f, $fAlt)) {
    if (Test-Path $target) {
        Apply-Replacements $target @{
            "'   - Batch processes .DWG files from specified folder"                     = "'   - K${ae}sittelee DWG-tiedostot erin${ae} m${ae}${ae}ritetyst${ae} kansiosta"
            "'   - Opens specialized editor forms (FlowPickNo, PipeFromTo, PipeToOther)" = "'   - Avaa erikoistuneet muokkauslomakkeet (FlowPickNo, PipeFromTo, PipeToOther)"
            "' Note: Searches for UP079D blocks where LOOPNO attribute contains `"AA`""  = "' Huomio: Etsii UP079D-blokkeja joiden LOOPNO-attribuutti sis${ae}lt${ae}${ae} `"AA`""
        }
    }
}

# ================== PIPE/Form_USysFlowPickNo.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_USysFlowPickNo.cls"
Apply-Replacements $f @{
    "'   - Connects to running AutoCAD instance"                                            = "'   - Yhdist${ae}${ae} k${ae}ynniss${ae} olevaan AutoCAD-instanssiin"
    "' Tarkoitus: Initialize form - connect to AutoCAD and create flow block selection set" = "' Tarkoitus: Alustetaan lomake - yhdistet${ae}${ae}n AutoCADiin ja luodaan virtausblokin valintajoukko"
    "'   - Closes form if AutoCAD not running"                                              = "'   - Sulkee lomakkeen jos AutoCAD ei ole k${ae}ynniss${ae}"
    "'   - Closes form if no flow blocks found"                                             = "'   - Sulkee lomakkeen jos virtausblokkeja ei l${oe}ydy"
    "' Initialize with first block"                                                         = "' Alustetaan ensimm${ae}isell${ae} blokilla"
    "'   - Iterates through all blocks automatically"                                       = "'   - K${ae}y l${ae}pi kaikki blokit automaattisesti"
    "'   - Sets EiLinjaa flag if pipeline not found for any block"                          = "'   - Asettaa EiLinjaa-lipun jos putkilinjaa ei l${oe}ydy millekk${ae}${ae}n blokille"
    "'   - Returns concatenated string (e.g., `"AREA01L001`")"                              = "'   - Palauttaa yhdistetytn merkkijonon (esim. `"AREA01L001`")"
    "'   - Handles both DEP and DEPA attribute tags"                                        = "'   - K${ae}sittelee sek${ae} DEP- ett${ae} DEPA-attribuuttitunnisteet"
    "'   - Searches for PIPEREF, PIPELINENO, or PIPE_LINE attribute"                        = "'   - Etsii PIPEREF-, PIPELINENO- tai PIPE_LINE-attribuuttia"
    "'   - Sets module-level oACAD and oDOC variables"                                      = "'   - Asettaa moduulitason oACAD- ja oDOC-muuttujat"
    "'   - Searches for PIPELINE, BAHPIPEL, ARAPIPEL blocks"                                = "'   - Etsii PIPELINE-, BAHPIPEL- ja ARAPIPEL-blokkeja"
    "'   - Deletes temporary selection set after use"                                       = "'   - Poistaa v${ae}liaikaisen valintajoukon k${ae}yt${oe}n j${ae}lkeen"
    "'   - Sets EiLinjaa=True if no pipeline found"                                         = "'   - Asettaa EiLinjaa=True jos putkilinjaa ei l${oe}ydy"
    "'   - Uses Nz to handle null values"                                                   = "'   - K${ae}ytt${ae}${ae} Nz:t${ae} null-arvojen k${ae}sittelyyn"
}

# ================== PIPE/Form_USysPipeFromTo.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_USysPipeFromTo.cls"
Apply-Replacements $f @{
    "'   - Connects to running AutoCAD instance"                                            = "'   - Yhdist${ae}${ae} k${ae}ynniss${ae} olevaan AutoCAD-instanssiin"
    "'   - Filters and displays only pipeline blocks (PIPELINE, PIPELINE_DATA, METSO_PIPE)" = "'   - Suodattaa ja n${ae}ytt${ae}${ae} vain putkilinjablokit (PIPELINE, PIPELINE_DATA, METSO_PIPE)"
    "' Tarkoitus: Initialize form, connect to AutoCAD, filter pipeline blocks"              = "' Tarkoitus: Alustetaan lomake, yhdistet${ae}${ae}n AutoCADiin, suodatetaan putkilinjablokit"
}

# ================== PIPE/Form_USysPipeToOther.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_USysPipeToOther.cls"
Apply-Replacements $f @{
    "'   - Connects to running AutoCAD instance"                                = "'   - Yhdist${ae}${ae} k${ae}ynniss${ae} olevaan AutoCAD-instanssiin"
    "' Tarkoitus: Initialize form, connect to AutoCAD, filter equipment blocks" = "' Tarkoitus: Alustetaan lomake, yhdistet${ae}${ae}n AutoCADiin, suodatetaan laiteblokit"
}

# ================== PIPE/Form_Venttiiliblokkien vaihto.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_Venttiiliblokkien vaihto.cls"
Apply-Replacements $f @{
    "'   - Reads valve records from MANUALVALVES table"               = "'   - Lukee venttiilitietueet MANUALVALVES-taulusta"
    "'   - Opens each .DWG file containing valves"                    = "'   - Avaa jokaisen venttiilit sis${ae}lt${ae}v${ae}n .DWG-tiedoston"
    "'   - Preserves attributes, layer, rotation (handles mirroring)" = "'   - S${ae}ilytt${ae}${ae} attribuutit, tason ja kiertokulman (k${ae}sittelee peilauksen)"
    "'   - Updates database with new Handle and block name"           = "'   - P${ae}ivitt${ae}${ae} tietokantaan uuden handlen ja blokin nimen"
    "' Handle rotation (account for mirroring)"                       = "' K${ae}sitell${ae}${ae}n kiertokulma (otetaan huomioon peilaus)"
    "' Update database with new handle and block name"                = "' P${ae}ivitet${ae}${ae}n tietokanta uudella handlella ja blokin nimell${ae}"
}

# ================== PIPE/Form_zFunc.cls ==================
$f = "c:\database_migration\Access\PIPE\Form_zFunc.cls"
Apply-Replacements $f @{
    "'   2. Iterate through all records"                 = "'   2. K${ae}yd${ae}${ae}n l${ae}pi kaikki tietueet"
    "'   - Improves database normalization"              = "'   - Parantaa tietokannan normalisointia"
    "'   - Prevents empty strings in database"           = "'   - Est${ae}${ae} tyhji${ae} merkkijonoja tietokannassa"
    "'   1. Iterate through all TableDefs in database"   = "'   1. K${ae}yd${ae}${ae}n l${ae}pi kaikki TableDefit tietokannassa"
    "'   - Populates KaikkiTaulukot combo box RowSource" = "'   - T${ae}ytt${ae}${ae} KaikkiTaulukot-yhdistelm${ae}listan RowSource-ominaisuuden"
}

# ================== PIPE/Koodit.bas ==================
$f = "c:\database_migration\Access\PIPE\Koodit.bas"
Apply-Replacements $f @{
    "'   Provides essential functionality for the PIPE database:" = "'   Tarjoaa keskeist${ae} toiminnallisuutta PIPE-tietokannalle:"
    "'   Handle - AutoCAD block handle (hex string)"              = "'   Handle - AutoCAD-blokin k${ae}sittelytunnus (heksadesimaalimerkkijono)"
    "'   - Handles missing drawings gracefully"                   = "'   - K${ae}sittelee puuttuvat piirustukset asiallisesti"
    "' Normalize drawing name for consistent comparison"          = "' Normalisoidaan piirustuksen nimi yhtenev${ae}${ae} vertailua varten"
    "'   POIMI(`"AREA-123-VALVE`", 2) returns `"123`""            = "'   POIMI(`"AREA-123-VALVE`", 2) palauttaa `"123`""
    "'   - Returns Null for null or empty input"                  = "'   - Palauttaa Null null- tai tyhj${ae}lle sy${oe}tteelle"
}

# ================== Excel/Moduulit/AcadDATA/Koodit.bas ==================
$f = "c:\database_migration\Excel\Moduulit\AcadDATA\Koodit.bas"
Apply-Replacements $f @{
    "' Initialize headers if clearing worksheet"                                                      = "' Alustetaan otsikot jos ty${oe}jj${ae}rjestelm${ae} tyhjennet${ae}${ae}n"
    "' Set appropriate formats per column"                                                            = "' Asetetaan sopivat muotoilut kullekin sarakkeelle"
    "' Initialize tag-to-column cache"                                                                = "' Alustetaan tagi-sarake-v${ae}limuisti"
    "' Fallback: If specific names were requested and the selection is empty, re-select by type only" = "' Varasuunnitelma: Jos nimett${ae} pyydettiin mutta valinta on tyhj${ae}, valitaan uudelleen pelk${ae}ll${ae} tyypill${ae}"
    "' Normalize checks for both interface TypeName and AcDb* values"                                 = "' Normalisoidaan tarkistukset sek${ae} rajapinnan TypeName- ett${ae} AcDb*-arvoille"
    "' Late binding: InsertionPoint returns a Variant array (x,y,z). Retrieve then index."            = "' My${oe}h${ae}inen sidonta: InsertionPoint palauttaa Variant-taulukon (x,y,z). Haetaan ja indeksoidaan."
    "' Fallback: attempt property access directly"                                                    = "' Varasuunnitelma: yritet${ae}${ae}n ominaisuuden k${ae}ytt${oe}${ae} suoraan"
    "' Handle text entities only when requested"                                                      = "' K${ae}sitell${ae}${ae}n tekstientiteetit vain pyydett${ae}ess${ae}"
    "' After flushing rows, coerce coordinate columns to numbers (handles any text leftovers)"        = "' Rivien huuhtelun j${ae}lkeen pakotetaan koordinaattisarakkeet luvuiksi (k${ae}sittelee teksti${ae} j${ae}${ae}nteist${ae})"
    "' Loop through each attribute in the block"                                                      = "' K${ae}yd${ae}${ae}n l${ae}pi kaikki blokin attribuutit"
}

# ================== Excel/Moduulit/Listojen kyselyt/Module2.bas ==================
$f = "c:\database_migration\Excel\Moduulit\Listojen kyselyt\Module2.bas"
Apply-Replacements $f @{
    "' Initialize all document info variables" = "' Alustetaan kaikki asiakirjatietomuuttujat"
}

Write-Host ""
Write-Host "=== TRANSLATION COMPLETE ==="
Write-Host "Files modified: $($changedFiles.Count)"
Write-Host "Total change operations: $totalChanges"
foreach ($cf in $changedFiles) {
    Write-Host "  - $($cf.Replace('c:\database_migration\',''))"
}
