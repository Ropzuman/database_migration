VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Formi 
   Caption         =   "Jonotulostus"
   ClientHeight    =   7824
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7680
   OleObjectBlob   =   "Formi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Formi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PFile As Integer
Dim Ajastin As Long
Dim SDocMode As Boolean
Dim TmpHakemisto As String
Dim fso As FileSystemObject

Public acadname As String
Public AcadVer As String

Private Const BIF_RETURNONLYFSDIRS = &H1

Private Type BrowseInfo
    hwndOwner As LongPtr
    pidlRoot As LongPtr
    sDisplayName As String
    sTitle As String
    ulFlags As Long
    lpfn As LongPtr
    lParam As LongPtr
    iImage As Long
End Type

Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
    
Private Declare PtrSafe Function FindWindow Lib "User32" Alias _
    "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As LongPtr ' <-- KORJATTU Long -> LongPtr

Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias _
    "SHBrowseForFolderA" (lpBrowseInfo As BrowseInfo) As LongPtr

Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" _
    (ByVal lItem As LongPtr, ByVal sDir As String) As Boolean

Private Sub AsetaGS_Click()

End Sub

Private Sub CommandButton1_Click()
    'Formi.Hide
    KirjoitaTiedot
    Unload Formi
      
    ActiveDocument.SetVariable "FILEDIA", 1 'MRU 20111017
    
End Sub

Private Sub FPath_Change()
    If SamaPolku.Value = True Then
        PltTied.Value = FPath.Value
    End If
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub GS_Settings_Click()
    FormAsetukset.show
End Sub

Private Sub Label11_Click()
    select_dwg.Value = True
End Sub

Private Sub Label12_Click()
    select_lst.Value = True
End Sub

Private Sub Label13_Click()
    select_dxf.Value = True
End Sub

Private Sub Label16_Click()
    If SamaPolku.Value = True Then
        SamaPolku.Value = False
    ElseIf SamaPolku.Value = False Then
        SamaPolku.Value = True
    End If
End Sub

Private Sub Label17_Click()
    If PlotStyleFromPicture.Value = True Then
        PlotStyleFromPicture.Value = False
    Else
        PlotStyleFromPicture.Value = True
    End If
End Sub

Private Sub Label18_Click()
    If SHXKommentit.Value = True Then
        SHXKommentit.Value = False
    Else
        SHXKommentit.Value = True
    End If
End Sub

Private Sub Label19_Click()
    If Lineweight.Value = True Then
        Lineweight.Value = False
    Else
        Lineweight.Value = True
    End If
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Luo_lst_Click()
Dim Paate As String
Dim tiedosto As String
Dim intI As Integer
Dim CheckPaate As String
Dim Response

If Right(FPath.Value, 1) <> "\" Then
    FPath.Value = FPath.Value & "\"
End If

CheckPaate = UCase(Right(FName.Value, 4))
If CheckPaate = ".LST" Then
    MsgBox "Tulostuslistaa ei luoda lst-tiedostoista. Kirjoita suodatustieto kohtaan [Tiedosto tai jokerimerkki].", vbOKOnly, "Tulostuslistan luonti"
    FName.SetFocus
    With FName
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    Exit Sub
End If

tiedosto = Dir(FPath.Value & FName.Value)

Open FPath.Value & "tulostus.lst" For Output As #1

Do While tiedosto <> ""
    Paate = UCase(Right(tiedosto, 4))
    If Paate = ".DWG" Or Paate = ".DXF" Then
        'MsgBox Tiedosto, vbOKOnly, Tiedostolistaus
        Print #1, tiedosto
'       If AvaaTiedosto(FPath.Value & Tiedosto) = True Then 'Tiedoston avaus onnistui
'           Tila Tiedosto
'           TulostaDOC
'           PFile = PFile + 1
'           If Kaikki = False Then  'Lopetetaan t‰h‰n jos on kysymys testitulostuksesta
'               Exit Do
'           End If
'       End If
    ElseIf Paate = ".LST" Then
'       LueTiedosto FPath.Value & Tiedosto, Kaikki
'       If Kaikki = False Then  'Lopetetaan t‰h‰n jos on kysymys testitulostuksesta
'           Exit Do
'       End If
    End If

    tiedosto = Dir
    intI = intI + 1
Loop

Close #1

Response = MsgBox("Tulostustiedosto luotu " & intI & " kuvan tulostusta varten." & Chr(13) & "Avataanko luotu tiedosto?", vbYesNo)
If Response = vbYes Then    ' User chose Yes.
    Avaa_lista
End If

FName.Value = "tulostus.lst"
Nayta_lista.Enabled = True
End Sub

Private Sub Nayta_lista_Click()
    Avaa_lista
End Sub

Private Sub Avaa_lista()

    If Right(FPath.Value, 1) <> "\" Then
        FPath.Value = FPath.Value & "\"
    End If
    
    If Right(FName.Value, 5) = "*.lst" Then
        MyFinalPathFile = FPath.Value & "tulostus.lst"
    ElseIf Right(FName.Value, 3) = "lst" Then
        MyFinalPathFile = FPath.Value & FName.Value
    Else
        MyFinalPathFile = FPath.Value & "tulostus.lst"
    End If
    
    If Len(Dir(MyFinalPathFile, vbNormal)) = 0 Then
        MsgBox "Tiedostoa " & MyFinalPathFile & " ei ole olemassa.", vbOKOnly
    Else
        Shell "Notepad.exe" & " " & MyFinalPathFile, vbMaximizedFocus
    End If
End Sub


Private Sub PltTied_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tmpPath As String

    If PltTiedCheck.Value = True And SamaPolku = False Then
        
        'Application.FileDialog(msoFileDialogOpen).show
        
        tmpPath = BrowseForDirectory
        
        If tmpPath <> "" Then
            If Len(Dir(tmpPath, vbDirectory)) = 0 Then
                PltTied.Value = tmpPath
            ElseIf Right(tmpPath, 1) <> "\" Then
                PltTied.Value = tmpPath & "\"
            End If
        End If
    End If
End Sub

Private Sub SamaPolku_Click()
    If SamaPolku.Value = False Then
        PltTied.BackColor = &H80000005
        PltTied.Enabled = True
    Else
        PltTied.Enabled = False
        PltTied.BackColor = &H80000013
        PltTied.Value = FPath.Value
    End If
End Sub

Private Sub PltTiedCheck_Click()
    If PltTiedCheck.Value = False Then
        PltTied.Enabled = False
        PltTied.BackColor = &H80000013
    Else
'        PltTied.Value = BrowseForDirectory
        PltTied.BackColor = &H80000005
        PltTied.Enabled = True
    End If
End Sub

Private Sub select_dwg_Click()
    FName.Value = "*.dwg"
    If Len(Dir(FPath.Value & "tulostus.lst", vbNormal)) = 0 Then
        Nayta_lista.Enabled = False
    Else
        Nayta_lista.Enabled = True
    End If
    Luo_lst.Enabled = True
End Sub

Private Sub select_dxf_Click()
    FName.Value = "*.dxf"
    If Len(Dir(FPath.Value & "tulostus.lst", vbNormal)) = 0 Then
        Nayta_lista.Enabled = False
    Else
        Nayta_lista.Enabled = True
    End If
    Luo_lst.Enabled = True
End Sub

Private Sub select_lst_Click()
    FName.Value = "*.lst"
    'Tiedostosisalto = Dir(FPath.Value & "*.dwg")
    Luo_lst.Enabled = False
End Sub

Private Sub UserForm_Initialize()
    'Avaus
    Dim i As Integer
    Dim Tulostimet As Variant
    Dim Tyylit As Variant
    Dim Koot As Variant
    Dim oTiedosto As Scripting.TextStream
    Dim Tiedot As String
    Dim TulostimetArray() As String
    Dim Lˆytyy As Boolean
    
    Set fso = New FileSystemObject
    
    AcadVer = ActiveDocument.Application.Version
    acadname = ActiveDocument.Application.Caption
    'strip the dwg from the caption
    acadname = RTrim(Left(acadname, InStr(1, acadname, " -")))
    
    Label15.Caption = "Version 2018.9.12 / JV‰" & vbNewLine & "12.9.2018"
    
    pc3Path = GetSetting("Jonotulostus", "Asetukset", "pc3Path", "C:\Data\Tools")
    TulostimetArray() = Split(Preferences.Files.PrinterConfigPath, ";")
    Lˆytyy = False
    For i = LBound(TulostimetArray) To UBound(TulostimetArray)
        If UCase(TulostimetArray(i)) = UCase(pc3Path) Then
            Lˆytyy = True
            Exit For
        End If
    Next i
    
    If Lˆytyy = False Then
        Preferences.Files.PrinterConfigPath = pc3Path & ";" & Preferences.Files.PrinterConfigPath
    End If
    
    'Laitetaan pysty ja vaaka vaihtoehdot
    Orientation.AddItem "Portrait (Pysty)"
    Orientation.AddItem "Landscape (Vaaka)"
    Orientation.AddItem "Portrait (Pysty, k‰‰nteinen)"
    Orientation.AddItem "Landscape (Vaaka, k‰‰nteinen)"
  
    'Laitetaan kopioita vaihtoehdot
    For i = 1 To 30
        Copies.AddItem i
    Next i
  
    'Laitetaan tulostimet
    Tulostimet = ActiveDocument.Layouts(0).GetPlotDeviceNames()
    For i = LBound(Tulostimet) To UBound(Tulostimet)
        If Not Tulostimet(i) Like "DWF*" _
            And Not Tulostimet(i) Like "Publish*" _
            And Not Tulostimet(i) = "Fax" _
            And Not Tulostimet(i) Like "*OneNote*" _
            And Not Tulostimet(i) Like "*XPS*" _
            And Not Tulostimet(i) Like "HP ePrint*" _
            Then
            Plotter.AddItem Tulostimet(i)
        End If
    Next i
  
    'Laitetaan plottaustyylit
    Tyylit = ActiveDocument.Layouts(0).GetPlotStyleTableNames()
    For i = LBound(Tyylit) To UBound(Tyylit)
        Pens.AddItem Tyylit(i)
    Next i
    
    'Haetaan edelliset tai oletusasetukset rekisterist‰
    FPath.Value = GetSetting("Jonotulostus", "Asetukset", "FPath", "")
    FName.Value = GetSetting("Jonotulostus", "Asetukset", "FName", "")
    Plotter.Value = GetSetting("Jonotulostus", "Asetukset", "Plotter", "DWG To PDF.pc3")
    Orientation.Value = GetSetting("Jonotulostus", "Asetukset", "Orientation", "Landscape (Vaaka)")
    Copies.Value = GetSetting("Jonotulostus", "Asetukset", "Copies", "1")
    Pens.Value = GetSetting("Jonotulostus", "Asetukset", "Pens", "monochrome.ctb")
    Size.Value = GetSetting("Jonotulostus", "Asetukset", "Size", "ISO_A3_(297.00_x_420.00_MM)")
    PltTied.Value = GetSetting("Jonotulostus", "Asetukset", "PltTied", "")
    SamaPolku.Value = GetSetting("Jonotulostus", "Asetukset", "SamaPolku", "true")
    SHXKommentit.Value = GetSetting("Jonotulostus", "Asetukset", "SHXKommentit", "false")
    Lineweight.Value = GetSetting("Jonotulostus", "Asetukset", "Lineweight", "false")
    PlotStyleFromPicture.Value = GetSetting("Jonotulostus", "Asetukset", "PlotStyleFromPicture", "false")
    
End Sub

Private Sub KirjoitaTiedot()
    SaveSetting "Jonotulostus", "Asetukset", _
        "FPath", FPath.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "FName", FName.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "Plotter", Plotter.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "Orientation", Orientation.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "Copies", Copies.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "Pens", Pens.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "Size", Size.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "PltTied", PltTied.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "SamaPolku", SamaPolku.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "SHXKommentit", SHXKommentit.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "Lineweight", Lineweight.Value
     SaveSetting "Jonotulostus", "Asetukset", _
        "PlotStyleFromPicture", PlotStyleFromPicture.Value
        
End Sub

Private Sub Plotter_Change()
    Dim Koot As Variant
    Dim KokoID As Integer
    
    ActiveDocument.Layouts(0).ConfigName = Plotter.Value
  
    Koot = ActiveDocument.Layouts(0).GetCanonicalMediaNames()
    For i = 0 To Size.ListCount - 1
        Size.RemoveItem 0
    Next i

    For i = LBound(Koot) To UBound(Koot)
        If Koot(i) Like "ISO_A[0-4]_*" Or _
        Koot(i) Like "A[0-5]" Or _
        Koot(i) Like "User*" Or _
        Koot(i) Like "ISO_full_bleed_*A*" Then
            Size.AddItem Koot(i)
        End If
        If Koot(i) = "A4" Or Koot(i) = "ISO_A3_(297.00_x_420.00_MM)" Then
            KokoID = Size.ListCount - 1
        End If
    Next i

    Size.ListIndex = KokoID
    'If Size.ListCount > 0 Then Size.ListIndex = 0
    
    If Plotter.Value = "Generic PostScript Printer" Or Plotter.Value = "DWG To PDF.pc3" Or _
        Plotter.Value Like "AutoCAD PDF*" Then
        PltTiedCheck.Value = True
        Label16.Enabled = True
        SamaPolku.Enabled = True
            If SamaPolku.Value = True Then
                PltTied.Enabled = False
                PltTied.BackColor = &H80000013
                PltTied.Value = FPath.Value
            End If
    Else
        PltTiedCheck.Value = False
        PltTied.Enabled = False
        PltTied.BackColor = &H80000013
        Label16.Enabled = False
        SamaPolku.Enabled = False
    End If
    
End Sub

Private Sub Tulosta_Click()
    Set oBatTiedosto = Nothing
    Kay_Lapi_Tiedostot True 'Tulostaa kaikki dokumentit
'    YhdistaPSjamuunnaPDF
End Sub

Private Sub TulostaTest_Click()
'    Application.Visible = False
    Kay_Lapi_Tiedostot False 'Tulostaa yhden testiksi
'    Application.Visible = True
End Sub

Private Sub Kay_Lapi_Tiedostot(Kaikki As Boolean)
Dim Paate As String
Dim TiedostoNimi As String
Dim tiedosto As String
Dim erotin As String
Dim GSPath As String
Dim gsbinary As String
Dim YhdKansio As String
Dim PDFasetus As String
'Dim LueTiedosto As String
Dim i As Integer
Dim BackPlot As Variant
Dim Buffer() As Byte
Dim MyPreference As AcadPreferencesOutput
 
    'Hakee asetukset rekisterist‰
    GSPath = GetSetting("Jonotulostus", "Asetukset", "GSPath", "C:\Data\Tools\gs\gs9.21")
    If GSPath = "" Then
        MsgBox "Tarkista GS-asetukset!", vbOKOnly
        Exit Sub
    End If

    gsbinary = GetSetting("Jonotulostus", "Asetukset", "GSBinary", "gswin64c")
    If gsbinary = "" Then
        MsgBox "Tarkista GS-asetukset!", vbOKOnly
        Exit Sub
    End If

    Ajastin = Timer
    
    If FPath.Value = "" Or FPath.Value = "Hakemistoa ei ole valittu" Or FName.Value = "" Then
        MsgBox "Ole hyv‰ ja valitse polku ja tiedosto ensin.", vbCritical, "Plot"
        Exit Sub
    End If
    
    If MsgBox("T‰m‰ sulkee kaikki avoimet dokumentit. Oletko varma ett‰ n‰in voidaan tehd‰?", vbOKCancel, "Plot Files") = vbCancel Then
        Exit Sub
    End If
    
    If Plotter.Value = "Generic PostScript Printer" Then
        If MsgBox("'Generic PostScript Printer' on vanhentunut tapa tehd‰ PDF-tiedostoja. K‰yt‰ 'DWG To PDF.pc3' sen sijaan." & vbCrLf & vbCrLf & "Jatketaanko?", vbOKCancel, "Plot Files") = vbCancel Then
            Exit Sub
        End If
    End If
    
    If Size.Value <> "A4" And Plotter.Value = "Generic PostScript Printer" Then
        If MsgBox("Paperikoko ei ole A4 vaan " & Size.Value & ". Haluatko varmasti jatkaa tulostusta?", vbOKCancel, "Plot Files") = vbCancel Then
            Exit Sub
        End If
    End If
    
    'Application.Visible = False
    'Varmistetaan ett‰ ollaan monen dokumentin tilassa
    SDocMode = Application.Preferences.System.SingleDocumentMode
    Application.Preferences.System.SingleDocumentMode = False
    
    'Ei luoda plot.log -tiedostoa
    Application.Preferences.Output.AutomaticPlotLog = False
    Application.ActiveDocument.SendCommand "_-PLOTSTAMP" & vbCr & "_LOG" & vbCr & "_NO" & vbCr & "PLOT.LOG" & vbCr & "" & vbCr
    
    'EPDFSHX arvo talteen
    If Val(Left(Application.Version, 4)) < 23 Then
        PDFasetus = "EPDFSHX"
    Else
        PDFasetus = "PDFSHX"
    End If
    
    EPDFSHX = Application.ActiveDocument.GetVariable(PDFasetus)
        
    'Tulostetaan tekstit kommenteiksi
    If SHXKommentit.Value = True Then
        ActiveDocument.SetVariable PDFasetus, 1
    ElseIf SHXKommentit.Value = False Then
        ActiveDocument.SetVariable PDFasetus, 0
    End If
        
    'Tiedoston nimi on m‰‰ritelty
    Formi.MousePointer = fmMousePointerHourGlass
    PFile = 1

    Ajastin = Timer
    
    Randomize
    TmpHakemisto = Application.Preferences.Files.TempFilePath & "PDF_" & Int((Format(Now(), "yyyymmddhhmmss") - 1 + 1) * Rnd + 1) & "\"

    If Len(Dir(TmpHakemisto, vbDirectory)) = 0 Then
        MkDir TmpHakemisto
    End If

    Set oBatTiedosto = fso.CreateTextFile(TmpHakemisto & "tiedostot.bat", True)
    Set oBatTiedosto = Nothing
    
    If Right(FPath.Value, 1) <> "\" Then
        FPath.Value = FPath.Value & "\"
    End If
    
    tiedosto = Dir(FPath.Value & FName.Value)

    If PltTied.Value = "" Then
        YhdKansio = TmpHakemisto
    Else
        If Right(PltTied.Value, 1) <> "\" Then
            PltTied.Value = PltTied.Value & "\"
        End If
        YhdKansio = PltTied.Value
    End If
   
    If PltTied.Value = "" And PltTiedCheck.Value = True Then
        PDFHakemisto = TmpHakemisto
    ElseIf PltTied.Value <> "" And PltTiedCheck.Value = True Then
        PDFHakemisto = PltTied.Value
    End If
    
    Do While tiedosto <> ""
        Paate = UCase(Right(tiedosto, 4))
        If Plotter.Value = "Generic PostScript Printer" Then
            TiedostoNimi = (Left(tiedosto, Len(tiedosto) - Len(Paate))) & ".ps"
        ElseIf Plotter.Value = "DWG To PDF.pc3" Or Plotter.Value Like "AutoCAD PDF*" Then
            TiedostoNimi = (Left(tiedosto, Len(tiedosto) - Len(Paate))) & ".pdf"
        Else
            TiedostoNimi = (Left(tiedosto, Len(tiedosto) - Len(Paate))) & ".plt"
        End If
       
        If Paate = ".DWG" Or Paate = ".DXF" Then
            If AvaaTiedosto(FPath.Value & tiedosto) = True Then 'Tiedoston avaus onnistui
                Tila TiedostoNimi
          
                If Plotter.Value = "Generic PostScript Printer" Then
'                If Palokortit.Value = False Then
'                TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".ps"
                    Set oBatTiedosto = fso.CreateTextFile(TmpHakemisto & "tiedostot.bat", True)
                    oBatTiedosto.WriteLine "TYPE " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34) & Chr(32) & Chr(62) & Chr(62) & Chr(32) & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
                    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34)
                    Set oBatTiedosto = Nothing
                ElseIf Plotter.Value = "DWG To PDF.pc3" Or Plotter.Value Like "AutoCAD PDF*" Then
                    Set oBatTiedosto = fso.CreateTextFile(TmpHakemisto & "tiedostot.bat", True)
                    oBatTiedosto.WriteLine "ECHO " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34) & Chr(32) & Chr(62) & Chr(62) & Chr(32) & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
                    Set oBatTiedosto = Nothing
                End If
          
                TulostaDOC TiedostoNimi, True
                PFile = PFile + 1
                If Kaikki = False Then  'Lopetetaan t‰h‰n jos on kysymys testitulostuksesta
                    Exit Do
                End If
            End If
            
        ElseIf Paate = ".LST" Then
'            LueTiedosto FPath.Value & Tiedosto, Kaikki
            TiedostoNimi = LueTiedosto(FPath.Value & tiedosto, Kaikki)
            If Kaikki = False Then  'Lopetetaan t‰h‰n jos on kysymys testitulostuksesta
                Exit Do
            End If
        End If
      
        tiedosto = Dir
    Loop
    
    'Suljetaan viimeinen dokumentti, jotta se ei j‰‰ vahingossa auki
    For i = 0 To Application.Documents.Count - 1
        Application.Documents(0).Close False
    Next i
    
    Application.Documents.Add 'Tehd‰‰n uusi dokumentti jotta voidaan siirty‰ single tilaan. Single tilaan ei voi siirty‰, jollei yht‰‰n dokumenttia ole auki
    Application.Preferences.System.SingleDocumentMode = SDocMode 'Siirryt‰‰n tilaan, joka oli ennen ohjelman k‰ynnistyst‰
        
    If Plotter.Value = "Generic PostScript Printer" Or Plotter.Value = "DWG To PDF.pc3" Or _
    Plotter.Value Like "AutoCAD PDF*" Then
    
    'SIIRTO
    Set oBatTiedosto = fso.CreateTextFile(TmpHakemisto & "gsrun.bat", True)
    oBatTiedosto.WriteLine "CD " & TmpHakemisto
    oBatTiedosto.WriteLine "C:"
    'oBatTiedosto.WriteLine "gswin64c @" & Chr(34) & TmpHakemisto & "parametrit.arg" & Chr(34) & " -sOutputFile#" & Chr(34) & PDFHakemisto & "Valmis.pdf" & Chr(34) & " -f " & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
    
    If Plotter.Value = "Generic PostScript Printer" Then
        erotin = " -f "
    ElseIf Plotter.Value = "DWG To PDF.pc3" Or Plotter.Value Like "AutoCAD PDF*" Then
        erotin = " @"
    End If
    
    oBatTiedosto.WriteLine "START /B /WAIT " & gsbinary & " @" & Chr(34) & TmpHakemisto & "parametrit.arg" & Chr(34) & " -sOutputFile#" & Chr(34) & PDFHakemisto & "Valmis.pdf" & Chr(34) & erotin & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
    oBatTiedosto.WriteLine ":Loppu"
    oBatTiedosto.WriteLine "IF NOT EXIST " & Chr(34) & PDFHakemisto & "Valmis.pdf" & Chr(34) & " GOTO Loppu"
    oBatTiedosto.WriteLine Chr(34) & PDFHakemisto & "Valmis.pdf" & Chr(34)
    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & "parametrit.arg" & Chr(34)
    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
    oBatTiedosto.WriteLine "DEL /Q " & Chr(34) & TmpHakemisto & "*.pdf" & Chr(34)
    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & "alku.bat" & Chr(34)
    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & "gsrun.bat" & Chr(34)
    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & "tiedostot.bat" & Chr(34)
    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & "yhdista.bat" & Chr(34)

    Set oTiedosto = Nothing
    oBatTiedosto.Close
    Set oBatTiedosto = Nothing

    Set oTP = fso.CreateTextFile(TmpHakemisto & "parametrit.arg", True)
    oTP.WriteLine "-I" & GSPath & "\bin;" & GSPath & "\lib;C:\Data\Tools\gs\fonts"
'    oTP.WriteLine "-I" & Chr(34) & "C:\Program Files\gs\gs8.61\bin" & Chr(34) & ";" & Chr(34) & "C:\Program Files\gs\gs8.61\lib" & Chr(34) & ";" & Chr(34) & "C:\Program Files\gs\fonts" & Chr(34) & ""
'    oTP.WriteLine "-dCompatibilityLevel#1.3"
    oTP.WriteLine "-sDEVICE#pdfwrite"
    oTP.WriteLine "-dPDFSETTINGS#/printer"
'    oTP.WriteLine "-dDOPDFMARKS"
'    oTP.WriteLine "-sPAPERSIZE#a3"
'    oTP.WriteLine "-dShowAnnots"
    If SHXKommentit.Value = True Then
        oTP.WriteLine "-dPrinted=false"
    End If
    
    If Plotter.Value = "Generic PostScript Printer" Then
        oTP.WriteLine "-sPAPERSIZE#" & LCase(Size.Value)
    End If
    
'    oTP.WriteLine "-dAutoRotatePages=/None"
'    oTP.WriteLine "-c " & Chr(34) & "<</Orientation " & Orientation.ListIndex & ">> setpagedevice" & Chr(34)
    oTP.WriteLine "-q"
    oTP.WriteLine "-dSAFER"
    oTP.WriteLine "-dNOPAUSE"
    oTP.WriteLine "-dBATCH"

    Set oTP = Nothing

    Set oBatTiedostoAlku = fso.CreateTextFile(TmpHakemisto & "alku.bat", True)
    oBatTiedostoAlku.WriteLine "CHCP 1252" 'Sallii ‰‰kkˆset tiedostonimiss‰
    oBatTiedostoAlku.WriteLine "set PATH=%PATH%;" & GSPath & "\bin;" & GSPath & "\lib;C:\Data\Tools\gs\fonts"
'    oBatTiedostoAlku.WriteLine "set PATH=%PATH%;" & Chr(34) & "C:\Program Files\gs\gs8.61\bin" & Chr(34) & ";" & Chr(34) & "C:\Program Files\gs\gs8.61\lib" & Chr(34) & ";" & Chr(34) & "C:\Program Files\gs\fonts" & Chr(34) & ""
    oBatTiedostoAlku.WriteLine ":Alku"
    oBatTiedostoAlku.WriteLine "IF NOT EXIST " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34) & " GOTO Alku"

    oBatTiedostoAlku.Close
    Set oBatTiedostoAlku = Nothing
    
    Open TmpHakemisto & "alku.bat" For Binary Access Read As #1
    Open TmpHakemisto & "tiedostot.bat" For Binary Access Read As #2
    Open TmpHakemisto & "gsrun.bat" For Binary Access Read As #3
    Open TmpHakemisto & "yhdista.bat" For Binary Access Write As #4
    
    ReDim Buffer(1 To LOF(1))
    Get #1, , Buffer
    Put #4, , Buffer

    ReDim Buffer(1 To LOF(2))
    Get #2, , Buffer
    Put #4, , Buffer

    ReDim Buffer(1 To LOF(3))
    Get #3, , Buffer
    Put #4, , Buffer

    Close #1, #2, #3, #4
    
    Shell TmpHakemisto & "yhdista.bat", vbHide
    'Shell TmpHakemisto & "yhdista.bat", vbNormalFocus

    End If 'If Plotter.Value = "Generic PostScript Printer" Or Plotter.Value = "DWG To PDF.pc3" Then
    
    'SIIRTO LOPPUU
    
    TilaTieto.Caption = ""
    MsgBox "Tulostus suoritettu", vbOKOnly, "Valmis"
    Formi.MousePointer = fmMousePointerDefault
'    Application.Visible = True
 
'    Formi.Hide
    KirjoitaTiedot
'    Unload Formi
    
    ActiveDocument.SetVariable "FILEDIA", 1 'MRU 20111017

    'Palautetaan EPDF:n arvo
    ActiveDocument.SetVariable PDFasetus, EPDFSHX
End Sub

Public Function LueTiedosto(LstTiedosto As String, Kaikki As Boolean) As String
Dim oTiedosto As Scripting.TextStream
Dim Rivi As String
Dim i As Integer
Dim Paate As String
Dim Alku As String
Dim tiedosto As String
Dim TiedostoArray As String
Dim nro As Integer
Dim TiedostoNimi As String

    Set oBatTiedosto = fso.OpenTextFile(TmpHakemisto & "tiedostot.bat", ForWriting)

    Set oTiedosto = fso.OpenTextFile(LstTiedosto, 1)
    
    Do While Not oTiedosto.AtEndOfStream
        DoEvents
        Rivi = oTiedosto.ReadLine
        Alku = Mid(Rivi, 1, 1)
        If Alku = " " Or Alku = ";" Or Rivi = "" Then
            'Ei mit‰‰n
        ElseIf Alku = "@" Then
            LueTiedosto Mid(Rivi, 2), Kaikki
        Else
            'If Len(Rivi) > 50 Then
            '    TiedostoNimi = Mid(Rivi, 40)
            '    tiedosto = FPath.Value & Mid(Rivi, 40)
            'Else
                TiedostoNimi = Rivi
                tiedosto = FPath.Value & Rivi
            'End If
            
            nro = 0

            If Plotter.Value = "Generic PostScript Printer" Or Plotter.Value = "DWG To PDF.pc3" _
            Or Plotter.Value Like "AutoCAD PDF*" Then
'            If Palokortit.Value = False Then
                If Plotter.Value = "Generic PostScript Printer" Then
                    TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".ps"
                ElseIf Plotter.Value = "DWG To PDF.pc3" Or Plotter.Value Like "AutoCAD PDF*" Then
                    TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".pdf"
                End If
       
                If Plotter.Value = "Generic PostScript Printer" Then
                    oBatTiedosto.WriteLine "TYPE " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34) & Chr(32) & Chr(62) & Chr(62) & Chr(32) & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
                    oBatTiedosto.WriteLine "DEL " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34)
                ElseIf Plotter.Value = "DWG To PDF.pc3" Or Plotter.Value Like "AutoCAD PDF*" Then
                    oBatTiedosto.WriteLine "ECHO " & Chr(34) & TmpHakemisto & TiedostoNimi & Chr(34) & Chr(32) & Chr(62) & Chr(62) & Chr(32) & Chr(34) & TmpHakemisto & "Valmis.ps" & Chr(34)
                End If
            Else
                TiedostoNimi = (Left(TiedostoNimi, Len(TiedostoNimi) - 4)) & ".plt"

                If Right(UCase(tiedosto), 4) <> ".DWG" Then
                    tiedosto = tiedosto & ".DWG"
                End If
            End If
      
            If AvaaTiedosto(tiedosto) Then
                Tila fso.GetFileName(tiedosto)

                TulostaDOC TiedostoNimi, True
            
                PFile = PFile + 1
                
                If Kaikki = False Then  'Lopetetaan t‰h‰n jos on kysymys testitulostuksesta
                    Exit Do
                End If

            End If
        End If
    Loop

LueTiedosto = TiedostoNimi

End Function

Private Sub TulostaDOC(tiedosto As String, ToPDF As Boolean)
Dim Tulostettava As Variant
Dim Origo(0 To 1) As Double
Dim Layout(1) As String
Dim booOk As Boolean
Dim YhdKansio As String
Dim BackPlot As Variant

'    Layout(0) = "Model"
    Layout(0) = Application.ActiveDocument.ActiveLayout.Name  ' Muutos: Tulostetaan aktiivinen layout / Jal 10.2.2005
    Origo(0) = 0
    Origo(1) = 0
    
    If Layout(0) = "Model" Then
        Set Tulostettava = Application.ActiveDocument.ModelSpace.Layout
    Else
        Set Tulostettava = Application.ActiveDocument.ActiveLayout
    End If
    
    Tulostettava.ConfigName = Plotter.Value
    Tulostettava.PlotType = acExtents
    Tulostettava.PlotOrigin = Origo
    Tulostettava.StandardScale = acScaleToFit
    Tulostettava.UseStandardScale = True
    If Lineweight.Value = True Then
        Tulostettava.PlotWithLineweights = True
    Else
        Tulostettava.PlotWithLineweights = False
    End If
    Tulostettava.PlotWithPlotStyles = True
    If PlotStyleFromPicture.Value = False Then
        Tulostettava.StyleSheet = Pens.Value
    End If
    Tulostettava.CanonicalMediaName = Size.Value
    Tulostettava.CenterPlot = True
    Tulostettava.PlotRotation = Orientation.ListIndex  '0=ac0degrees (pysty), 1=ac90degrees (vaaka), 2=ac180degrees (pysty) upside-down, 3=ac270degrees (vaaka) upside-down
    Tulostettava.RefreshPlotDeviceInfo
    
    ActiveDocument.Regen acActiveViewport
'    ActiveDocument.Application.ZoomExtents
    ActiveDocument.Plot.QuietErrorMode = True
    ActiveDocument.Plot.SetLayoutsToPlot Layout
    ActiveDocument.Plot.NumberOfCopies = Copies.Value
    
    BackPlot = ActiveDocument.GetVariable("BACKGROUNDPLOT")
    ThisDrawing.SetVariable "BACKGROUNDPLOT", 0
       
    If Plotter.Value = "Generic PostScript Printer" Or Plotter.Value = "DWG To PDF.pc3" _
    Or Plotter.Value Like "AutoCAD PDF*" Then
        booOk = ActiveDocument.Plot.PlotToFile(TmpHakemisto & tiedosto, Plotter.Value)
    Else
        ActiveDocument.Plot.PlotToDevice
    End If
    
    ActiveDocument.SetVariable "BACKGROUNDPLOT", BackPlot

End Sub

Private Function AvaaTiedosto(Avattava As String) As Boolean
'Avaa tiedoston AutoCADiss‰.
Dim i As Integer

    'AppActivate acadname, True
    If fso.FileExists(Avattava) Then
        For i = 0 To Application.Documents.Count - 1
            Application.Documents(0).Close False   'Suljetaan documentit (Multi Document tilassa AutoCAD 2002 ei herjaa vaikkei documentteja olisikaan ja vaikka documentteja ei olisi talletettu)
        Next i
        
        Application.Documents.Open (Avattava)  'Avataan uusi dokumentti AutoCadiin (ACAD 2002 osaa avata myˆs DXF documentit ilman kyselyj‰)
        AvaaTiedosto = True
    Else
        AvaaTiedosto = False
        MsgBox "Tiedostoa ei lˆydy: " & Avattava, vbCritical, "Virhe!"
    End If
End Function

Private Sub UserForm_Terminate()
    KirjoitaTiedot
    Set fso = Nothing
End Sub

Private Sub Valitse_Click()
    Dim objFile As FileDialogs
    Dim strFilter As String
    Dim strDir As String
    Dim strFileName As String

    Set fso = New FileSystemObject
    
    strFilter = "AutoCAD Drawing (*.dwg)" & Chr(124) & "*.DWG" & Chr(124) _
             & "AutoCAD DXF (*.dxf)" & Chr(124) & "*.DXF" & Chr(124) _
             & "List Of Drawings (*.lst)" & Chr(124) & "*.LST" & Chr(124) _
             & "All Types (*.dwg;*.dxf;*.lst)" & Chr(124) & "*.DWG;*.LST;*.DXF" & Chr(124)
    
    'Tiedostop‰‰tteen mukaan s‰‰det‰‰n filtteri oikeanlaiseksi
    If FName.Value = "*.dwg" Then
        FiltteriIndeksi = 1
    ElseIf FName.Value = "*.dxf" Then
        FiltteriIndeksi = 2
    ElseIf FName.Value = "*.lst" Then
        FiltteriIndeksi = 3
    Else
        FiltteriIndeksi = 4
    End If
    
    Set objFile = New FileDialogs
    objFile.OwnerHwnd = ThisDrawing.HWND

    objFile.Title = "Valitse AutoCAD piirustus tai lista (*.dwg, *.dxf tai *.lst)"
    objFile.StartInDir = FPath.Value
    objFile.Filter = strFilter
    objFile.SelectedFilter = FiltteriIndeksi
    'return a valid filename
    strFileName = objFile.ShowOpen
    If Not strFileName = vbNullString Then
        'use this space to perform operation
        FName.Value = fso.GetFileName(strFileName)
        FPath.Value = fso.GetParentFolderName(strFileName) & "\"
        Nayta_lista.Enabled = True
    End If
    Set objFile = Nothing
    
End Sub

Private Sub Valitse_H_Click()
    Dim tmpPath As String
    
    tmpPath = BrowseForDirectory
    
    If tmpPath <> "" Then
        If Len(Dir(tmpPath, vbDirectory)) = 0 Then
            FPath.Value = tmpPath
        ElseIf Right(tmpPath, 1) <> "\" Then
            FPath.Value = tmpPath & "\"
        End If
    End If
    
    PltTied.Value = FPath.Value
    
End Sub

' Let the user browse for a directory. Return the
' selected directory. Return an empty string if
' the user cancels.
Public Function BrowseForDirectory() As String
    Dim browse_info As BrowseInfo
    Dim item As LongPtr ' <-- KORJATTU Long -> LongPtr
    Dim dir_name As String

    'modified for MS Access/VBA
    With browse_info
'        .hWndOwner = Application.hWndAccessApp
'        .hWndOwner = Me.Handle
        .pidlRoot = 0
        .sDisplayName = Space$(260)
        .sTitle = "Valitse hakemisto"
        .ulFlags = 1 ' Return directory name.
        .lpfn = 0
        .lParam = 0
        .iImage = 0
    End With

    item = SHBrowseForFolder(browse_info)
    If item Then
        dir_name = Space$(260)
        If SHGetPathFromIDList(ByVal item, ByVal dir_name) Then
            dir_name = Trim(dir_name)
            BrowseForDirectory = Mid(dir_name, 1, Len(dir_name) - 1)
        Else
            BrowseForDirectory = ""
        End If
    Else
        BrowseForDirectory = ""
    End If
    
End Function

Private Sub Tila(tiedosto As String)
    sec = CInt(Timer - Ajastin)
    Min = sec \ 60
    sec = sec Mod 60

    TilaTieto.Caption = "Printing File: " & PFile & " - " & tiedosto & vbCrLf & "Time: " & Min & " min  " & sec & " sec"
    
    Formi.Repaint
End Sub

