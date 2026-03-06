param()
# Korjaa TurvallinenKursori-rajaus Form_MOTORS ja Form_PIIRIT alimuotoihin

$enc = [System.Text.Encoding]::UTF8
$LF = "`n"

function Fix-LisaaTeksti {
    param(
        [string]$FilePath,
        [string]$Label
    )

    $content = [IO.File]::ReadAllText($FilePath, $enc)
    $iStart = $content.IndexOf("Private Sub LisaaTeksti()")
    if ($iStart -lt 0) {
        Write-Host "[$Label] Private Sub LisaaTeksti() not found!" -ForegroundColor Red
        return
    }
    $iEndSub = $content.IndexOf("End Sub", $iStart) + "End Sub".Length
    $oldFunc = $content.Substring($iStart, $iEndSub - $iStart)

    # Rakennetaan uusi funktio LF-rivinvaihdoilla
    $lines = @(
        "Private Sub LisaaTeksti()",
        "    Dim Alku As String",
        "    Dim Loppu As String",
        "    Dim KohdeTeksti As String",
        "    Dim TurvallinenKursori As Long  ' Rajattu kursori - ei voi ylittaeae tekstin pituutta",
        "",
        "    On Error GoTo ErrorHandler",
        "",
        "    If Not KohdeTextBox Is Nothing Then",
        "        KohdeTeksti = Nz(KohdeTextBox.Value, """")",
        "        ' Rajataan kursori tekstin pituuden sisaeaen - estaae Left`$/Mid`$-virheet fokuksenvaihdon jaelkeen",
        "        TurvallinenKursori = Kursori",
        "        If TurvallinenKursori > Len(KohdeTeksti) Then TurvallinenKursori = Len(KohdeTeksti)",
        "        If TurvallinenKursori < 0 Then TurvallinenKursori = 0",
        "",
        "        If Form.Parent.CLoppuun = True Then",
        "            ' Lisaeaeaeaen kentaen loppuun",
        "            Alku = KohdeTeksti",
        "            Loppu = """"",
        "        Else",
        "            ' Lisaeaeaeaen kursorin kohtaan",
        "            Alku = Left`$(KohdeTeksti, TurvallinenKursori)",
        "            Loppu = Mid`$(KohdeTeksti, TurvallinenKursori + 1)",
        "        End If",
        "",
        "        KohdeTextBox.Value = Alku & "" {"" & Me.TEKSTI.Value & ""}"" & Loppu",
        "        KohdeTextBox.SetFocus",
        "        ' SelStart ja SelLength vaativat fokuksen - estaetaen Virhe 2185",
        "        On Error Resume Next",
        "        KohdeTextBox.SelStart = TurvallinenKursori + Len(Me.TEKSTI.Value) + 3",
        "        KohdeTextBox.SelLength = 0",
        "        Kursori = KohdeTextBox.SelStart",
        "        On Error GoTo ErrorHandler",
        "    End If",
        "",
        "    Exit Sub",
        "",
        "ErrorHandler:",
        "    MsgBox ""Virhe tekstin lisaeyksessae: "" & Err.Description, vbExclamation",
        "End Sub"
    )
    Write-Host "This approach has escaping issues, using direct replacement instead"
}

# --- Kayta suoraa korvausmenetelmaea ---
# Luetaan tiedostot UTF-8:lla

$motorsFile = "c:\database_migration\Access\Function_descriptions_html\Form_MOTORS subform.cls"
$piiritFile = "c:\database_migration\Access\Function_descriptions_html\Form_PIIRIT subform.cls"

foreach ($filePath in @($motorsFile, $piiritFile)) {
    $c = [IO.File]::ReadAllText($filePath, $enc)
    $i = $c.IndexOf("Private Sub LisaaTeksti()")
    $e = $c.IndexOf("End Sub", $i) + 7
    Write-Host "$filePath start=$i end=$e old_len=$($e - $i)"
}
