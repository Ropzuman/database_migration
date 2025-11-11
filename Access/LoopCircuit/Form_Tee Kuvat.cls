Option Compare Database
Option Explicit

' Updated 2025-10-22: Explicit DAO, better error handling, removed dead code, optimized

Private Ajastin As Long
Private min As Long
Private sec As Long
Private OSMODE As Long
Private LokiPituus As Long
Private oAcad As Object ' AcadApplication (late binding for compatibility)
Public oDoc As Object ' AcadDocument

' Wait for specified time (hundredths of a second)
Private Sub Odota(Aika As Integer)
    Dim odotus As Double
    odotus = CDbl(Timer) + CDbl((Aika / 100))
    Do While odotus > Timer
        DoEvents
    Loop
End Sub

' Update status bar with current file and elapsed time
Private Sub Tila(Tiedosto As String)
    sec = CLng(Timer - Ajastin)
    min = sec \ 60
    sec = sec Mod 60
    SysCmd acSysCmdSetStatus, "Käsittelee tiedostoa: " & Tiedosto & "  Aika: " & min & " min  " & sec & " sek"
End Sub

Private Sub Form_Load()
    ' Initialize log textbox
    On Error Resume Next
    Loki.Value = ""
    Loki.Locked = False
    Loki.Enabled = True
    On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
End Sub

Private Sub FormClose_Click()
    DoCmd.Close
End Sub

' Extract text attributes from DWG blocks
Private Sub HaeTekstit_Click()
    Dim Tied As String
    Dim Polku As String
    Dim Nimi As String
    Dim Taulu As DAO.Recordset
    Dim oEnt As Object ' AcadEntity
    Dim db As DAO.Database
    
    On Error GoTo ErrHandler
    
    Ajastin = Timer
    Set oAcad = CreateObject("AutoCAD.Application")
    
    Polku = PohjaHakem.Value
    If Right(Polku, 1) <> "\" Then Polku = Polku & "\"
    
    Tied = Dir(Polku & "*.dwg")
    
    Set db = CurrentDb
    db.Execute "DELETE * FROM Blokit WHERE PATH='" & Replace(PohjaHakem.Value, "'", "''") & "'"
    
    Set Taulu = db.OpenRecordset("Blokit", dbOpenDynaset)
    
    Do While Tied <> ""
        Nimi = Left(Tied, Len(Tied) - 4)
        
        ' Skip BASE* and BLOCKMAP* files
        If UCase(Left(Tied, 4)) <> "BASE" And UCase(Left(Tied, 8)) <> "BLOCKMAP" Then
            Set oDoc = oAcad.Documents.Open(Polku & Tied)
            Tila Tied
            
            For Each oEnt In oDoc.ModelSpace
                If oEnt.ObjectName = "AcDbAttributeDefinition" Then
                    Taulu.AddNew
                    Taulu.Fields("Block").Value = Nimi
                    Taulu.Fields("ATAG").Value = oEnt.TagString
                    Taulu.Fields("VALUE").Value = oEnt.TextString
                    Taulu.Fields("Path").Value = Polku
                    Taulu.Update
                End If
            Next
            
            oDoc.Close False
        End If
        
        Tied = Dir
    Loop
    
Cleanup:
    On Error Resume Next
    If Not Taulu Is Nothing Then
        Taulu.Close
        Set Taulu = Nothing
    End If
    Set db = Nothing
    If Not oAcad Is Nothing Then
        oAcad.Quit
        Set oAcad = Nothing
    End If
    On Error GoTo 0

    MsgBox "Valmis", vbInformation, "Päivitä tekstit"
    Exit Sub

ErrHandler:
    MsgBox "Virhe: " & Err.Description, vbCritical, "HaeTekstit"
    Resume Cleanup
End Sub

' Extract text from selected blocks in active AutoCAD document
Private Sub HaeValitutTekstit_Click()
    Dim Tied As String
    Dim Polku As String
    Dim Nimi As String
    Dim Taulu As DAO.Recordset
    Dim Blokki As Object ' AcadBlockReference
    Dim Joukko As Object ' AcadSelectionSet
    Dim BlockArray As Variant
    Dim i As Long
    Dim j As Long
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim DWGName As String
    Dim Hakemisto As String
    Dim Rivi As Long
    Dim db As DAO.Database
    
    On Error GoTo ErrHandler
    
    FilterType(0) = 0
    FilterData(0) = "INSERT"
    
    Ajastin = Timer
    
    Set oAcad = GetObject(, "AutoCAD.Application")
    Polku = PohjaHakem.Value
    If Right(Polku, 1) <> "\" Then Polku = Polku & "\"
    
    Set oDoc = oAcad.ActiveDocument
    DWGName = Left(oDoc.Name, Len(oDoc.Name) - 4)
    Hakemisto = oDoc.Path
    
    Set Joukko = oDoc.ActiveSelectionSet
    Joukko.Clear
    
    Joukko.Select 4, , , FilterType, FilterData ' acSelectionSetPrevious = 4
    
    If Joukko.Count = 0 Then
        MsgBox "Kuvasta ei löytynyt valittuja blokkeja!", vbCritical, "Tuo DATA"
        Exit Sub
    End If
    
    Set db = CurrentDb
    Set Taulu = db.OpenRecordset("Blokit", dbOpenDynaset)
    
    For i = 0 To Joukko.Count - 1
        If Joukko(i).EntityName = "AcDbBlockReference" Then
            Set Blokki = Joukko(i)
            
            db.Execute "DELETE * FROM Blokit WHERE Block='" & Replace(Blokki.Name, "'", "''") & "'"
            
            If Blokki.HasAttributes Then
                BlockArray = Blokki.GetAttributes
                For j = 0 To UBound(BlockArray)
                    Taulu.AddNew
                    Taulu.Fields("Block").Value = Blokki.Name
                    Taulu.Fields("ATAG").Value = BlockArray(j).TagString
                    Taulu.Fields("VALUE").Value = BlockArray(j).TextString
                    Taulu.Fields("Path").Value = Hakemisto
                    Taulu.Update
                Next j
                Rivi = Rivi + 1
            End If
        End If
    Next i
    
Cleanup:
    On Error Resume Next
    If Not Taulu Is Nothing Then
        Taulu.Close
        Set Taulu = Nothing
    End If
    Set db = Nothing
    Set oAcad = Nothing
    On Error GoTo 0

    MsgBox "Valmis - Käsitelty " & Rivi & " blokkia.", vbInformation, "Päivitä tekstit"
    Exit Sub

ErrHandler:
    MsgBox "Virhe: " & Err.Description, vbCritical, "HaeValitutTekstit"
    Resume Cleanup
End Sub

' Generate DWG files from database data
Private Sub TeeKuvat_Click()
    Dim db As DAO.Database
    Dim Kuvat As DAO.Recordset
    Dim DbBlokit As DAO.Recordset
    Dim Tiedot As DAO.Recordset
    Dim FSO As Object
    Dim Tiedosto As String
    Dim Blokki As Object
    Dim Attribuutit As Variant
    Dim Paikka As Long
    Dim BNimi As String
    Dim i As Long
    Dim Sql As String
    Dim OtsTaulukko As String
    Dim BlockTaulukko As String
    Dim BlockTaulukkoTyhjennys As String
    Dim Vastaus As Integer
    Dim KuvaNimi As String
    Dim Os As Long
    Dim Os2 As Long
    Dim Kysely As String
    Dim Burst As Boolean
    
    On Error GoTo ErrHandler
    
    Set db = CurrentDb
    Ajastin = Timer
    LokiPituus = 0
    OtsTaulukko = TOtsTaulukko.Value
    Burst = TBurst.Value
    BlockTaulukko = TBlockTaulukko.Value
    BlockTaulukkoTyhjennys = TBlockTaulukko.Value

    Vastaus = MsgBox("Päivitetäänkö " & BlockTaulukkoTyhjennys & "-taulukko?", vbYesNoCancel, "Tee kuvat")

    If Vastaus = vbCancel Then Exit Sub
    
    If Vastaus = vbYes Then
        db.Execute "DELETE * FROM " & BlockTaulukkoTyhjennys
        db.Execute "DELETE * FROM Kuviin_Device_Info_Tbl"
        
        ' Execute multiple queries from TKyselyt field
        Os = 1
        Do
            Os2 = InStr(Os, TKyselyt.Value, vbCrLf)
            If Os2 = 0 Then
                Kysely = Mid(TKyselyt.Value, Os)
                db.Execute Kysely
                LisaaLokiin "- Ajettiin kysely " & Kysely
                Exit Do
            Else
                Kysely = Mid(TKyselyt.Value, Os, Os2 - Os)
                db.Execute Kysely
                LisaaLokiin "- Ajettiin kysely '" & Kysely & "'"
                Os = Os2 + 2
            End If
        Loop

        LisaaLokiin "- Päivitettiin " & BlockTaulukkoTyhjennys & "-taulukko"
    End If
    
    ' Validate folders
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If IsNull(KuvaHakem.Value) Then
        KuvaHakem.Value = PohjaHakem.Value
        LisaaLokiin "- Kuvahakemistoa ei määritelty, käytetään pohjahakemistoa"
    End If
    
    If Not FSO.FolderExists(PohjaHakem.Value) Then
        LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & PohjaHakem.Value & ")"
        MsgBox "Hakemistoa: " & vbCrLf & PohjaHakem.Value & vbCrLf & " ei löytynyt!", vbCritical, "Tee kuvat"
        GoTo Cleanup
    End If
    
    If Not FSO.FolderExists(KuvaHakem.Value) Then
        LisaaLokiin "Virhe: Hakemistoa ei löytynyt (" & KuvaHakem.Value & ")"
        MsgBox "Hakemistoa: " & vbCrLf & KuvaHakem.Value & vbCrLf & " ei löytynyt!", vbCritical, "Tee kuvat"
        GoTo Cleanup
    End If
    
    ' Ensure paths end with backslash
    If Right(PohjaHakem.Value, 1) <> "\" Then PohjaHakem.Value = PohjaHakem.Value & "\"
    If Right(KuvaHakem.Value, 1) <> "\" Then KuvaHakem.Value = KuvaHakem.Value & "\"
    
    Set Kuvat = db.OpenRecordset("SELECT * FROM " & OtsTaulukko)
    
    If Kuvat.EOF Then
        LisaaLokiin "Virhe: Taulukosta (" & OtsTaulukko & ") ei löytynyt kuvia"
        MsgBox "Taulukosta """ & OtsTaulukko & """ ei löytynyt tietoja!", vbCritical, "Tee kuvat"
        GoTo Cleanup
    End If
    
    Set oAcad = CreateObject("AutoCAD.Application")
    
    Do While Not Kuvat.EOF
        LisaaLokiin "=== Käsitellään: " & Kuvat.Fields("FileID").Value & " ==="
        Tiedosto = PohjaHakem.Value & Kuvat.Fields("BaseDWG").Value
        If LCase(Right(Tiedosto, 4)) <> ".dwg" Then
            Tiedosto = Tiedosto & ".dwg"
        End If
        
        If FSO.FileExists(Tiedosto) Then
            ' LisaaLokiin "- Avataan pohjakuva: " & Tiedosto
            Tila Kuvat.Fields("FileID").Value
            Set oDoc = oAcad.Documents.Open(Tiedosto)
            OSMODE = oDoc.GetVariable("OSMODE")
            oDoc.SetVariable "OSMODE", 0
            
            If HaeIPoints(Kuvat.Fields("BaseDWG").Value) Then
                ' LisaaLokiin "- Haettu insertointipaikat kuvalle: " & Kuvat.Fields("BaseDWG").Value
                If VaihdaOtsikkotiedot(Kuvat) Then
                    ' LisaaLokiin "- Vaihdettu otsikkotiedot kuvalle: " & Kuvat.Fields("FileID").Value
                    Sql = "SELECT Block, Blocknumber FROM " & BlockTaulukko & _
                          " WHERE FileID='" & Replace(Kuvat.Fields("FileID").Value, "'", "''") & _
                          "' GROUP BY Block, Blocknumber HAVING (Blocknumber Is Not Null);"
                    Set DbBlokit = db.OpenRecordset(Sql)
                    
                    ' LisaaLokiin "- Blokkeja insertointia varten: " & DbBlokit.RecordCount
                    
                    Do While Not DbBlokit.EOF
                        ' LisaaLokiin "  - Insertoitava blokki: " & DbBlokit.Fields("Block").Value & " paikkaan #" & DbBlokit.Fields("Blocknumber").Value
                        If FSO.FileExists(PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg") Then
                            Sql = "SELECT * FROM " & BlockTaulukko & _
                                  " WHERE FileID='" & Replace(Kuvat.Fields("FileID").Value, "'", "''") & _
                                  "' AND Block='" & Replace(DbBlokit.Fields("Block").Value, "'", "''") & _
                                  "' AND Blocknumber='" & DbBlokit.Fields("Blocknumber").Value & _
                                  "' ORDER BY AttrTag;"
                            Set Tiedot = db.OpenRecordset(Sql)
                            
                            Paikka = CLng(DbBlokit.Fields("Blocknumber").Value)
                            BNimi = PohjaHakem.Value & DbBlokit.Fields("Block").Value & ".dwg"
                            
                            ' LisaaLokiin "    Inserting block: " & DbBlokit.Fields("Block").Value & " at position " & Paikka
                            Set Blokki = oDoc.ModelSpace.InsertBlock(Paikat(Paikka).Pisteet, BNimi, 1, 1, 1, 0)
                            ' LisaaLokiin "    Block inserted successfully"
                            
                            Attribuutit = Blokki.GetAttributes
                            For i = 0 To UBound(Attribuutit)
                                Tiedot.FindFirst "AttrTag = '" & Replace(Attribuutit(i).TagString, "'", "''") & "'"
                                If Not Tiedot.NoMatch Then
                                    Attribuutit(i).TextString = IIf(IsNull(Tiedot.Fields("Teksti").Value), "", Tiedot.Fields("Teksti").Value)
                                ' Else
                                '     LisaaLokiin "Virhe: Blokin " & DbBlokit.Fields("Block").Value & " attribuutille (" & Attribuutit(i).TagString & ") ei löytynyt arvoa"
                                End If
                            Next i
                            
                            If Burst Then
                                oDoc.SendCommand "burst" & vbCr & "last" & vbCr & vbCr
                            End If
                            
                            Tiedot.Close
                            Set Tiedot = Nothing
                        ' Else
                        '     LisaaLokiin "Virhe: Blokkia " & DbBlokit.Fields("Block").Value & " ei löytynyt"
                        End If
                        DbBlokit.MoveNext
                    Loop
                    
                    DbBlokit.Close
                    Set DbBlokit = Nothing
                ' Else
                '     LisaaLokiin "VIRHE: VaihdaOtsikkotiedot epäonnistui kuvalle: " & Kuvat.Fields("FileID").Value
                End If
            ' Else
            '     LisaaLokiin "VIRHE: HaeIPoints epäonnistui kuvalle: " & Kuvat.Fields("BaseDWG").Value
            End If
            
            KuvaNimi = KuvaHakem.Value & Kuvat.Fields("FileID").Value
            If Right(LCase(KuvaNimi), 4) <> ".dwg" Then KuvaNimi = KuvaNimi & ".dwg"
            
            If FSO.FileExists(KuvaNimi) Then
                FSO.DeleteFile KuvaNimi
            End If
            
            oDoc.SetVariable "OSMODE", OSMODE
            oDoc.SaveAs KuvaNimi
            oDoc.Close
            LisaaLokiin "Valmis: " & KuvaNimi
        ' Else
        '     LisaaLokiin "Virhe: Määriteltyä tiedostoa ei löytynyt (" & Tiedosto & ")"
        End If
        
        Kuvat.MoveNext
    Loop
    
    MsgBox "Kuvat tehty", vbOKOnly, "Valmis"
    
Cleanup:
    On Error Resume Next
    DoCmd.Hourglass False
    SysCmd acSysCmdClearStatus
    
    If Not Tiedot Is Nothing Then Tiedot.Close
    If Not DbBlokit Is Nothing Then DbBlokit.Close
    If Not Kuvat Is Nothing Then Kuvat.Close
    Set Tiedot = Nothing
    Set DbBlokit = Nothing
    Set Kuvat = Nothing
    Set db = Nothing
    Set FSO = Nothing
    Set Blokki = Nothing
    If Not oAcad Is Nothing Then
        oAcad.Quit
        Set oAcad = Nothing
    End If
    On Error GoTo 0
    Exit Sub

ErrHandler:
    MsgBox "Virhe: " & Err.Description, vbCritical, "TeeKuvat"
    Resume Cleanup
End Sub

' Extract insertion points from placement block
Private Function HaeIPoints(ByVal Kuva As String) As Boolean
    Dim Virhe As Boolean
    Dim Blokki As Object
    Dim iPaikka As Variant
    Dim Attribuutit As Variant
    Dim Numero As Long
    Dim i As Long
    Dim Joukko As Object
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim Paikkablokki As String
    
    On Error GoTo ErrHandler
    
    Virhe = False
    ReDim Paikat(1 To 1)
    
    Paikkablokki = TPaikkaBlokki.Value
    FilterType(0) = 2
    FilterData(0) = Paikkablokki
    
    Set Joukko = oDoc.SelectionSets.Add("TempSet_IPoints_" & Format(Now, "hhnnss"))
    
    On Error Resume Next
    Joukko.Select 5, , , FilterType, FilterData ' acSelectionSetAll = 5 (select all with filter)
    On Error GoTo ErrHandler
    
    If Joukko.Count > 0 Then
        Set Blokki = Joukko(0)
        
        If Blokki.HasAttributes Then
            Attribuutit = Blokki.GetAttributes
            For i = 0 To UBound(Attribuutit)
                Numero = CLng(Attribuutit(i).TextString)
                If UBound(Paikat) < Numero Then
                    ReDim Preserve Paikat(1 To Numero)
                End If
                iPaikka = Attribuutit(i).InsertionPoint
                Paikat(Numero).Pisteet(0) = iPaikka(0)
                Paikat(Numero).Pisteet(1) = iPaikka(1)
                Paikat(Numero).Pisteet(2) = iPaikka(2)
            Next i
        ' Else
        '     LisaaLokiin "Virhe: Insertointipaikat sisältyvästä blokista ei löytynyt attribuutteja (Kuva:" & Kuva & ")"
        '     Virhe = True
        End If
    ' Else
    '     LisaaLokiin "Virhe: Kuvasta " & Kuva & " ei löytynyt insertointipaikkoja sisältävää blokkia (" & Paikkablokki & ")"
    '     Virhe = True
    End If
    
    On Error Resume Next
    Joukko.Delete
    On Error GoTo 0
    
    Set Blokki = Nothing
    HaeIPoints = Not Virhe
    Exit Function

ErrHandler:
    ' LisaaLokiin "Virhe HaeIPoints: " & Err.Description
    HaeIPoints = False
End Function

' Update title block attributes from database record
Private Function VaihdaOtsikkotiedot(Taulu As DAO.Recordset) As Boolean
    Dim Attribuutit As Variant
    Dim Virhe As Boolean
    Dim i As Long
    Dim j As Long
    Dim b As Long
    Dim Blokki As Object
    Dim tagi As String
    Dim tagiValue As String
    Dim Loytyi As Boolean
    Dim FilterType(0) As Integer
    Dim FilterData(0) As Variant
    Dim Titleblokki As String
    Dim Joukko As Object
    
    On Error GoTo ErrHandler
    
    Virhe = False
    Titleblokki = TTitleBlokki.Value
    FilterType(0) = 2
    FilterData(0) = Titleblokki
    
    Set Joukko = oDoc.SelectionSets.Add("TempSet_Title_" & Format(Now, "hhnnss"))
    
    On Error Resume Next
    Joukko.Select 5, , , FilterType, FilterData ' acSelectionSetAll = 5 (select all with filter)
    On Error GoTo ErrHandler
    
    If Joukko.Count > 0 Then
        For b = 0 To Joukko.Count - 1
            Set Blokki = Joukko(b)
            
            If Not Blokki Is Nothing Then
                If Blokki.HasAttributes Then
                    Attribuutit = Blokki.GetAttributes
                    
                    For i = 0 To UBound(Attribuutit)
                        tagi = Attribuutit(i).TagString
                        tagiValue = Attribuutit(i).TextString
                        Loytyi = False
                        
                        For j = 0 To Taulu.Fields.Count - 1
                            ' Match either by tag value or tag name
                            If Taulu.Fields(j).Name = tagiValue Or Taulu.Fields(j).Name = tagi Then
                                Attribuutit(i).TextString = IIf(IsNull(Taulu.Fields(j).Value), "", Taulu.Fields(j).Value)
                                Loytyi = True
                                Exit For
                            End If
                        Next j
                        
                        If Not Loytyi Then
                            ' LisaaLokiin "- Taulukosta ei löytynyt tietoa otsikkotaulun attribuutille: " & tagiValue
                        End If
                    Next i
                ' Else
                '     LisaaLokiin "Virhe: Otsikkoblokista ei löytynyt attribuutteja (Kuva: " & Taulu.Fields("BaseDWG").Value & ")"
                '     Virhe = True
                End If
            ' Else
            '     LisaaLokiin "Virhe: Kuvasta " & Taulu.Fields("BaseDWG").Value & " ei löytynyt otsikkotietoja sisältävää blokkia (" & Titleblokki & ")"
            '     Virhe = True
            End If
        Next b
    End If
    
    On Error Resume Next
    Joukko.Delete
    On Error GoTo 0
    
    Set Blokki = Nothing
    VaihdaOtsikkotiedot = Not Virhe
    Exit Function

ErrHandler:
    ' LisaaLokiin "Virhe VaihdaOtsikkotiedot: " & Err.Description
    VaihdaOtsikkotiedot = False
End Function

' Add entry to log
Private Sub LisaaLokiin(Tieto As String)
    On Error Resume Next
    If LokiPituus > 0 Then
        Loki.Value = Loki.Value & vbCrLf & Tieto
    Else
        Loki.Value = Tieto
        LokiPituus = 1
    End If
    On Error GoTo 0
End Sub

' Browse for folder
Private Sub ValHakem_Click()
    Dim Hakemisto As String
    Hakemisto = PoimiHakemisto(IIf(IsNull(PohjaHakem.Value), "", PohjaHakem.Value))
    If Hakemisto <> "" Then PohjaHakem.Value = Hakemisto
End Sub

Private Sub ValHakem2_Click()
    Dim Hakemisto As String
    Hakemisto = PoimiHakemisto(IIf(IsNull(KuvaHakem.Value), "", KuvaHakem.Value))
    If Hakemisto <> "" Then KuvaHakem.Value = Hakemisto
End Sub

' Extract directory from file selection dialog
Private Function PoimiHakemisto(AHakem As String) As String
    Dim OpenFile As OPENFILENAME
    #If VBA7 Then
        Dim lReturn As LongPtr
    #Else
        Dim lReturn As Long
    #End If
    Dim Filtteri As String
    Dim TNimi As String
    Dim Otsikko As String
    
    On Error GoTo ErrHandler
    
    Filtteri = "AutoCAD Drawing (*.dwg)" & Chr(0) & "*.DWG" & Chr(0)
    Otsikko = "Valitse AutoCAD piirustus saadaksesi hakemiston"
    
    With OpenFile
        .lStructSize = Len(OpenFile)
        .hwndOwner = Me.Hwnd
        .hInstance = 0
        .lpstrFilter = Filtteri
        .nFilterIndex = 1
        .lpstrFile = String(257, 0)
        .nMaxFile = Len(.lpstrFile) - 1
        .lpstrFileTitle = .lpstrFile
        .nMaxFileTitle = .nMaxFile
        .lpstrInitialDir = AHakem
        .lpstrTitle = Otsikko
        .flags = 0
    End With
    
    lReturn = GetOpenFileName(OpenFile)
    
    If lReturn = 0 Then
        PoimiHakemisto = AHakem
    Else
        TNimi = Mid(OpenFile.lpstrFileTitle, 1, InStr(OpenFile.lpstrFileTitle, Chr(0)) - 1)
        PoimiHakemisto = Mid(OpenFile.lpstrFile, 1, InStr(OpenFile.lpstrFile, Chr(0)) - 1 - Len(TNimi))
    End If
    Exit Function

ErrHandler:
    PoimiHakemisto = AHakem
End Function

