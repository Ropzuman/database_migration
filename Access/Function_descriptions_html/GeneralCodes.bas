Option Compare Database
Option Explicit

'================================================================================
' Moduuli: GeneralCodes
' Tarkoitus: Yleiset apufunktiot ja yhteiset muuttujat
' Päivitetty: 2025-11-13 — VBA7/64-bit tuki lisätty
'             2026-03-03 — Kommentit suomeksi
'
' Kuvaus:
'   Tarjoaa apufunktiot laitteen revisiomuokkaukseen, taulunselaukseen
'   sekä tila- ja moodi-käännöksiin. Sisältää lomakkeiden välistä
'   kommunikaatiota varten julkisia muuttujia.
'
' Riippuvuudet:
'   - DAO.Recordset
'   - Lomakkeet: USysRevision, frmTables
'   - Taulut: MAINEQ, DRIVES, PUMPS, GEARS, TANKS
'================================================================================

' Julkiset muuttujat lomakkeiden välistä tiedonsiirtoa varten
Public oTaulu As DAO.Recordset
Public PaluuTaulu As Object
Public KohdeTextBox As TextBox
Public Kursori As Long

'================================================================================
' Funktio: MuutaRev
' Tarkoitus: Avaa revisiomuokkauslomake laitetauluille
' Palauttaa: Ei paluuarvoa
'
' Kuvaus:
'   Tarkistaa, onko aktiivinen taulu laitetaulu, ja avaa vastaavan tietueen
'   USysRevision-lomakkeessa muokkausta varten.
'================================================================================
Public Function MuutaRev()
    Dim Taul As String
    
    On Error GoTo ErrorHandler
    
    Taul = Application.CurrentObjectName
    
    If Taul = "MAINEQ" Or Taul = "DRIVES" Or Taul = "PUMPS" Or Taul = "GEARS" Or Taul = "TANKS" Then
        Set PaluuTaulu = Screen.ActiveDatasheet
        Set oTaulu = CurrentDb.OpenRecordset("SELECT * FROM " & Taul & " WHERE ID=" & Screen.ActiveDatasheet("ID").Value)
        DoCmd.OpenForm "USysRevision"
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Virhe revisiolomakkeen avauksessa: " & Err.Description, vbExclamation
End Function

'================================================================================
' Funktio: NaytaTables
' Tarkoitus: Avaa taulujenselauslomake
' Palauttaa: Ei paluuarvoa
'================================================================================
Public Function NaytaTables()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmTables"
    Exit Function
ErrorHandler:
    MsgBox "Virhe taulunselauslomakkeen avauksessa: " & Err.Description, vbExclamation
End Function

'================================================================================
' Funktio: Moodit
' Tarkoitus: Kääntää pilkulla eroteltujen moodikolmien koodit kuvaavaksi tekstiksi
' Parametrit:
'   Tieto — Pilkulla eroteltu moodi-koodi lista (A, M, E, L)
' Palauttaa: Moodi kuvaukset rivinvaihdolla eroteltuina, tai "-" jos Null
'
' Kuvaus:
'   Kääntää moodikolmit:
'     A = AUTO, M = MANUAL, E = EXTERNAL, L = LOCAL
'   Muut koodit pass-through-käännös. Palauttaa "-" Null-syötteen yhteydessä.
'================================================================================
Function Moodit(Tieto As Variant) As Variant
    Dim Tiedot As Variant
    Dim i As Integer
    Dim T As String
    
    On Error GoTo ErrorHandler
    
    If IsNull(Tieto) Then
        Moodit = "-"
        Exit Function
    End If
    
    Tiedot = Split(Tieto, ",")
    
    For i = 0 To UBound(Tiedot)
        Select Case UCase$(Trim$(Tiedot(i)))
            Case "A"
                T = "AUTO"
            Case "M"
                T = "MANUAL"
            Case "E"
                T = "EXTERNAL"
            Case "L"
                T = "LOCAL"
            Case Else
                T = UCase$(Trim$(Tiedot(i)))
        End Select
        
        If i > 0 Then
            Moodit = Moodit & vbCrLf & T
        Else
            Moodit = T
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    Moodit = "-"
End Function
