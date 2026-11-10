Option Compare Database
Option Explicit

'==============================================================================
' Lomake: MoottTilaus
' Tarkoitus: Vie lomakkeella suodatetut moottorit Excel-liitteeksi
'            moottorien ostoehdotukseen.
' P‰ivitetty: 2026-11-10 - Kovakoodattu Santa Fe -mallipolku korvattu
'             tiedostonvalinnalla, virheenk‰sittely ja Excelin siivous,
'             DROP TABLE vain jos taulu on olemassa, kuollut Command3 poistettu
'==============================================================================

' Muistetaan viimeksi valittu mallitiedosto lomakkeen elini‰n ajan
Private MalliPolku As String

' Kysyy k‰ytt‰j‰lt‰ Excel-mallitiedoston; palauttaa "" jos valinta peruttiin
Private Function ValitseMalli() As String
    Dim fd As Object  ' FileDialog-objekti

    If MalliPolku <> "" Then
        If Dir$(MalliPolku) <> "" Then
            ValitseMalli = MalliPolku
            Exit Function
        End If
    End If

    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker = 3
    With fd
        .Title = "Valitse moottorien ostoehdotuksen liitteen Excel-malli"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel-tiedostot", "*.xls;*.xlsx;*.xlsm;*.xlt;*.xltx"
        If .Show = -1 Then
            ValitseMalli = .SelectedItems(1)
            MalliPolku = ValitseMalli
        End If
    End With
    Set fd = Nothing
End Function

' Luo Excel-liitteen lomakkeen suodatetuista riveist‰.
' Rivit kirjoitetaan silmukassa solu kerrallaan, jotta mallipohjan muotoilut s‰ilyv‰t.
Private Sub Command0_Click()
    On Error GoTo ErrorHandler

    Dim PWB As Object  ' Excel.Workbook (mallipohja)
    Dim UWB As Object  ' Excel.Workbook (valmis liite)
    Dim XL As Object   ' Excel.Application
    Dim Siv As Object  ' Valmiin liitteen taulukko
    Dim rs As DAO.Recordset  ' Lomakkeen suodatetut rivit
    Dim i As Long  ' Juokseva numero
    Dim Rivi As Long  ' Excelin kirjoitusrivi
    Dim K As Long  ' Sarakelaskuri
    Dim Polku As String  ' Mallitiedoston polku
    Dim virhenumero As Long  ' Virheen numero k‰sittelij‰ss‰
    Dim virhekuvaus As String  ' Virheen kuvaus k‰sittelij‰ss‰
    Dim Valmis As Boolean  ' Onnistuiko liitteen luonti

    Polku = ValitseMalli()
    If Polku = "" Then Exit Sub

    Set rs = Me.RecordsetClone
    If rs.EOF And rs.BOF Then
        MsgBox "Lomakkeella ei ole viet‰vi‰ rivej‰.", vbInformation, "Ostoehdotuksen liite"
        GoTo Cleanup
    End If

    Set XL = CreateObject("Excel.Application")
    ' Piilotetaan Excel ajon ajaksi - nopeuttaa k‰yttˆ‰ huomattavasti
    XL.Visible = False
    XL.ScreenUpdating = False
    XL.Calculation = -4135  ' xlCalculationManual

    Set PWB = XL.Workbooks.Open(Polku, , True)  ' Vain luku -tilassa
    PWB.Sheets.Copy
    Set UWB = XL.ActiveWorkbook
    PWB.Close False
    Set PWB = Nothing
    Set Siv = UWB.ActiveSheet

    i = 1
    Rivi = 6
    rs.MoveFirst
    Do While Not rs.EOF
        Siv.Cells(Rivi, 1).Value = i
        Siv.Cells(Rivi, 1).VerticalAlignment = -4160  ' xlTop
        Siv.Cells(Rivi, 2).Value = rs.Fields(0).Value
        Siv.Cells(Rivi, 2).VerticalAlignment = -4160
        ' Sarakkeeseen 3 yhdistet‰‰n kaksi kentt‰‰
        Siv.Cells(Rivi, 3).Value = Nz(rs.Fields(1).Value, "") & Nz(rs.Fields(2).Value, "")
        For K = 4 To 10
            Siv.Cells(Rivi, K).Value = rs.Fields(K - 1).Value
            Siv.Cells(Rivi, K).VerticalAlignment = -4160
        Next K
        Rivi = Rivi + 1
        i = i + 1
        rs.MoveNext
    Loop

    With Siv.Range(Siv.Cells(6, 1), Siv.Cells(Rivi - 1, 10)).Borders
        .LineStyle = 1   ' xlContinuous
        .Weight = 1      ' xlHairline
        .ColorIndex = -4105  ' xlAutomatic
    End With
    Valmis = True

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    If Not XL Is Nothing Then
        XL.Calculation = -4105  ' xlCalculationAutomatic
        XL.ScreenUpdating = True
        If Valmis Then
            ' N‰ytet‰‰n tulos k‰ytt‰j‰lle
            XL.Visible = True
            AppActivate XL.Caption
        Else
            ' Virheess‰ tai keskeytyksess‰ suljetaan piilotettu Excel-prosessi
            If Not UWB Is Nothing Then UWB.Close False
            XL.Quit
        End If
    End If
    Set Siv = Nothing
    Set PWB = Nothing
    Set UWB = Nothing
    Set XL = Nothing
    If virhenumero <> 0 Then
        MsgBox "Virhe Excel-raportin luonnissa: " & virhekuvaus, vbCritical, "Virhe " & virhenumero
    End If
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    Resume Cleanup
End Sub

' Rakentaa moottorioptioiden aputaulun uudelleen ja asettaa valmistajasuodattimen
Private Sub Form_Load()
    On Error GoTo ErrorHandler

    Dim Tbl As DAO.TableDef  ' Taulujen l‰pik‰ynti
    Dim Olemassa As Boolean  ' Onko aputaulu olemassa

    For Each Tbl In CurrentDb.TableDefs
        If StrComp(Tbl.Name, "TblMotorsOptions", vbTextCompare) = 0 Then
            Olemassa = True
            Exit For
        End If
    Next Tbl
    If Olemassa Then CurrentDb.Execute "DROP TABLE TblMotorsOptions", dbFailOnError
    CurrentDb.Execute "MakeTblMotorsOptions", dbFailOnError

    If Valm.ListCount > 0 Then
        Valm.Value = Valm.ItemData(0)
        Valm_AfterUpdate
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Virhe moottorioptioiden aputaulun p‰ivityksess‰: " & Err.Description, vbCritical, "Virhe " & Err.Number
End Sub

Private Sub Valm_AfterUpdate()
  ' Heittomerkit tuplataan, jotta valmistajan nimi ei riko suodatinta
  Me.Filter = "Manuf ='" & Replace(CStr(Nz(Valm.Value, "")), "'", "''") & "'"
  Me.FilterOn = True
End Sub
