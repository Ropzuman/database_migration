# Koodin kommenttien kaantaminen suomeksi - instru3-moduulit
# Korjataan myos merkkikoodausvirheet (kÄ -> kä jne.)

$ae = [char]0xe4  # ä
$oe = [char]0xf6  # ö
$Ae = [char]0xc4  # Ä
$Oe = [char]0xd6  # Ö
$aa = [char]0xe5  # å

function Fix-File($path, $replacements) {
    if (-not (Test-Path $path)) { Write-Host "EI LOYDY: $path"; return }
    $c = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $before = $c.Length
    foreach ($r in $replacements) {
        $c = $c.Replace($r[0], $r[1])
    }
    [System.IO.File]::WriteAllText($path, $c, [System.Text.Encoding]::UTF8)
    Write-Host "OK: $path ($before -> $($c.Length) merkkia)"
}

# ============================================================
# general.bas
# ============================================================
$ae_ = $ae; $oe_ = $oe; $Ae_ = $Ae; $Oe_ = $Oe
Fix-File "c:\database_migration\Access\instru3\general.bas" @(
    # Otsikkokommentti
    , @("' Updated 2025-11-11: Added VBA7/64-bit support for GetOpenFileName API",
        "' P" + $ae + "ivitetty 2025-11-11: Lis" + $ae + "tty VBA7/64-bit-tuki GetOpenFileName API:lle")
    # Tyyppirakenne-kommentit
    , @("' Updated for 64-bit (window handle)",
        "' P" + $ae + "ivitetty 64-bittiseksi (ikkunakahva)")
    , @("' Updated for 64-bit (instance handle)",
        "' P" + $ae + "ivitetty 64-bittiseksi (instanssikahva)")
    , @("' Updated for 64-bit" + [char]13 + [char]10,
        "' P" + $ae + "ivitetty 64-bittiseksi" + [char]13 + [char]10)
    , @("' Updated for 64-bit`n",
        "' P" + $ae + "ivitetty 64-bittiseksi`n")
    , @("' Updated for 64-bit (callback pointer)",
        "' P" + $ae + "ivitetty 64-bittiseksi (takaisinkutsuosoitin)")
    # PilkkuPiste - korjataan koodausvirheet ja kaannetaan englanti
    , @("suomalainen kansainv" + $Ae + "liseen muotoon",
        "suomalainen kansainv" + $ae + "liseen muotoon")
    , @("'   Luku - Variant containing number with comma or point decimal separator",
        "'   Luku - Variant joka sis" + $ae + $ae + "  luvun pilkulla tai pisteell" + $ae + " desimaalierottimena")
    , @("' Returns:" + [char]13 + [char]10 + "'   String with decimal point format (e.g., ""3,14"" becomes ""3.14"")",
        "' Palauttaa:" + [char]13 + [char]10 + "'   Merkkijono desimaalipiste-muodossa (esim. ""3,14"" muuttuu muotoon ""3.14"")")
    , @("'   - Used for international number format conversion" + [char]13 + [char]10 + "'   - Commonly used before exporting to CSV or external systems",
        "'   - K" + $ae + "ytet" + $ae + $ae + "n kansainv" + $ae + "liseen numeromuunnokseen" + [char]13 + [char]10 + "'   - Yleisesti k" + $ae + "ytetty ennen CSV- tai ulkoisten j" + $ae + "rjestelmien vienti" + $ae)
    , @("' K" + $Ae + "sitell" + $Ae + $Ae + "n null/tyhjä syöte",
        "' K" + $ae + "sitell" + $ae + $ae + "n null/tyhj" + $ae + " sy" + $oe + "te")
    , @("' K" + $Ae + "sitell" + $Ae + $Ae + "n null/tyhjä sy" + $Oe + "te",
        "' K" + $ae + "sitell" + $ae + $ae + "n null/tyhj" + $ae + " sy" + $oe + "te")
    , @("' Find and replace comma with period",
        "' Etsit" + $ae + $ae + "n ja korvataan pilkku pisteell" + $ae)
    , @("PilkkuPiste = Luku  ' No comma found, return as-is",
        "PilkkuPiste = Luku  ' Pilkkua ei l" + $oe + "ydy, palautetaan sellaisenaan")
    , @("Dim Osoitin As Long  ' Position of comma in string",
        "Dim Osoitin As Long  ' Pilkun sijainti merkkijonossa")
    # UdNoteToRev - koodausvirheet ja englanti
    , @("k" + $Ae + "ytt" + $Ae + "j" + $Ae + "n muistiinpanoista p" + $Ae + "iv" + $Ae + "m" + $Ae + $Ae + "r" + $Ae + "n perusteella",
        "k" + $ae + "ytt" + $ae + "j" + $ae + "n muistiinpanoista p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae + "n perusteella")
    , @("'   UdNote - Variant containing user note string with format ""text:date|moretext""",
        "'   UdNote - Variant joka sis" + $ae + $ae + "  muistiinpanomerkkijonon muodossa ""teksti:pvm|lis" + $ae + "teksti""")
    , @("' Returns:" + [char]13 + [char]10 + "'   Variant - Revision code from _Revisions table or Null if not found",
        "' Palauttaa:" + [char]13 + [char]10 + "'   Variant - Revisionumero _Revisions-taulusta tai Null jos ei l" + $oe + "ydy")
    , @("'   - J" + $Ae + "sentyyy p" + $Ae + "iv" + $Ae + "m" + $Ae + $Ae + "r" + $Ae,
        "'   - J" + $ae + "sennet" + $ae + $ae + "n p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae)
    , @("'   - Looks up corresponding revision in _Revisions table",
        "'   - Etsii vastaavan revision _Revisions-taulusta")
    , @("'   - Palauttaa ensimm" + $Ae + "isen revision miss" + $Ae + " BeforeDate > j" + $Ae + "sennetty p" + $Ae + "iv" + $Ae + "m" + $Ae + $Ae + "r" + $Ae,
        "'   - Palauttaa ensimm" + $ae + "isen revision miss" + $ae + " BeforeDate > j" + $ae + "sennetty p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae)
    , @("'   - Used for historical revision tracking",
        "'   - K" + $ae + "ytet" + $ae + $ae + "n historialliseen revisioneurantaan")
    , @("Dim Paiva As String  ' Date string extracted from note",
        "Dim Paiva As String  ' Muistiinpanosta poimittu p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae + "merkkijono")
    , @("Dim Os As Long  ' Position marker for string parsing",
        "Dim Os As Long  ' Sijaintimuuttuja merkkijonojen j" + $ae + "sennyst" + $ae + " varten")
    , @("Dim VP As Date  ' Parsed date value",
        "Dim VP As Date  ' J" + $ae + "sennelty p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae + "arvo")
    , @("Dim RevTaul As DAO.Recordset  ' _Revisions table recordset",
        "Dim RevTaul As DAO.Recordset  ' _Revisions-taulun tietueet")
    , @("' K" + $Ae + "sitell" + $Ae + $Ae + "n null-sy" + $Oe + "te",
        "' K" + $ae + "sitell" + $ae + $ae + "n null-sy" + $oe + "te")
    , @("' Parse date from note string (format: ""text:date|moretext"")",
        "' J" + $ae + "sennnet" + $ae + $ae + "n p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae + " muistiinpanomerkkijonosta (muoto: ""teksti:pvm|lis" + $ae + "teksti"")")
    , @("' Extract date portion between : and |",
        "' Poimitaan p" + $ae + "iv" + $ae + "m" + $ae + $ae + "r" + $ae + "osa : ja | v" + $ae + "lilt" + $ae)
    # EtsiLoop - koodausvirheet ja englanti
    , @("olemassa j" + $Ae + "rjestelm" + $Ae + "ss" + $Ae,
        "olemassa j" + $ae + "rjestelm" + $ae + "ss" + $ae)
    , @("'   Looppi - String containing loop number",
        "'   Looppi - Merkkijono joka sis" + $ae + $ae + "  silmukkanumeron")
    , @("' Returns:" + [char]13 + [char]10 + '   String - "1" if loop exists, "" (empty) if not found',
        "' Palauttaa:" + [char]13 + [char]10 + '   String - "1" jos silmukka on olemassa, "" (tyhj' + $ae + ') jos ei l' + $oe + 'ydy')
    , @("'   - Kysyy qrysolvalve-kyselyst" + $Ae + " vastaavaa",
        "'   - Kysyy qrysolvalve-kyselyst" + $ae + " vastaavaa")
    , @("taaksep" + $Ae + "inyhteensopivuuden vuoksi",
        "taaksep" + $ae + "inyhteensopivuuden vuoksi")
    , @("'   - Used for validation before creating new loops",
        "'   - K" + $ae + "ytet" + $ae + $ae + "n validointiin ennen uusien silmukoiden luomista")
    , @("Dim Taul As DAO.Recordset  ' Query results recordset",
        "Dim Taul As DAO.Recordset  ' Kyselytulokset-tietueet")
    , @("' Query for matching loop",
        "' Kysely vastaavalle silmukalle")
    , @("EtsiLoop = """"  ' Not found",
        "EtsiLoop = """"  ' Ei l" + $oe + "ydy")
    , @('EtsiLoop = "1"  ' + "' Found",
        'EtsiLoop = "1"  ' + "' L" + $oe + "ytyi")
)

# ============================================================
# USysCheck.bas
# ============================================================
Fix-File "c:\database_migration\Access\instru3\USysCheck.bas" @(
    , @("' Get network username via Windows API",
        "' Haetaan verkkok" + $ae + "ytt" + $ae + "j" + $ae + "nimi Windows API:lla")
    , @("' Get computer name via Windows API",
        "' Haetaan tietokoneen nimi Windows API:lla")
    , @("' Write login record to tracking table",
        "' Kirjoitetaan kirjautumistietue seurantatauluun")
    , @("' Cleanup",
        "' Siivotaan")
)

Write-Host "`nValmis."
