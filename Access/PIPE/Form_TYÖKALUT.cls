Option Compare Database
Option Explicit

'================================================================================
' Lomake: TYÖKALUT
' Tarkoitus: AutoCAD P&I -tietojen tuonti- ja päivitystyökalupakki PIPE-tietokantaan
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lisätty
'
' Kuvaus:
'   Keskuslomake P&I-kaaviotietojen tuontiin ja päivitykseen AutoCADista:
'   - Käsittelee .DWG-tiedostot eräajona valitusta kansiosta
'   - Tuo putkilinja-, venttiili-, kenttälaite- ja instrumenttisilmukkatiedot
'   - Kolmivaiheinen työnkulku per tietotyyppi:
'     1. Lue AutoCADista (Command1x_Click)
'     2. Lisää väliaikaistaulukkoon (Command2x_Click)
'     3. Päivitä tietokantaan (Command3x_Click)
'   - Avaa erikoiseditorilomakkeet (FlowPickNo, PipeFromTo, PipeToOther)
'
' Riippuvuudet:
'   - AutoCAD-sovellus (COM-automaatio)
'   - DAO.Recordset tietokantaoperaatioihin
'   - Kyselytoiminnot: *_APPEND, *_UPDATE, DELETE-kyselyt
'   - LPolku-otsikkoruutu: Määrittelee .DWG-tiedostojen alakansion
'
' Tietotyypit:
'   - PIPELINEDATA: Putkilinjablokit (PIPELINE_DATA)
'   - MANVALVEDATA: Käsiventtiilit (MANUALVALVE_DATA)
'   - INSTVALVEDATA: Instrumenttiventtiilit (INSTVALVE_DATA)
'   - FIELDINSTDATA: Kenttälaitteet (FIELDINST_DATA)
'   - INSTLOOPDATA: Instrumenttisilmukat (UP079D, UP008C, UP085D, UP08C)
'
' Kontrollit:
'   - Command11-13: Putkilinjatyönkulku
'   - Command21-23: Käsiventtiilityönkulku
'   - Command31-33: Instrumenttiventtiilityönkulku
'   - Command41-43: Kenttälaitetyönkulku
'   - Command51-53: Instrumenttisilmukkatyönkulku
'   - Command18, 19, 40: Avaa editorilomakkeet
'   - BPipelineToManval: Linkitä putkilinjat käsiventtiileihin
'================================================================================

Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
Dim oDOC As Object  ' AcadDocument - muutettu late binding (64-bit)
Dim Tied As String  ' Current .DWG filename during batch processing
Dim Joukko As Object  ' AcadSelectionSet - muutettu late binding (64-bit)

'===== Putkilinja käsiventtiileihin =====
'================================================================================
' Tapahtuma: BPipelineToManval_Click
' Tarkoitus: Avaa lomake putkilinjaviittausten linkittämiseen käsiventtiileihin
'================================================================================
Private Sub BPipelineToManval_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "Pipeline To Manualvalve") = vbYes Then
    DoCmd.OpenForm "USysPipeToOther", , , , , , "MANUALVALVE_DATA"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe PipeToOther-lomakkeen avaamisessa (käsiventtiilit): " & Err.Description, vbExclamation
End Sub

'===== Putkilinjatyönkulku (Commands 11-13) =====
'================================================================================
' Tapahtuma: Command11_Click
' Tarkoitus: Lukee putkilinjatiedot AutoCAD-piirustuksista
'================================================================================
Private Sub Command11_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "PIPELINEDATA", "PIPELINE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Putkilinjojen luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Virhe putkilinjojen tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command12_Click
' Tarkoitus: Append pipeline data to staging table
'================================================================================
Private Sub Command12_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "PIPELINES_APPEND"
  MsgBox "Putkilinjojen lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe putkilinjojen lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command13_Click
' Tarkoitus: Päivittää tietokantaan putkilinjatiedot väliaikaistaulukosta
'================================================================================
Private Sub Command13_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "DELETE * From UsysPIPEUPD"
    CurrentDb.Execute "PIPELINES_UPDATE_START"
    CurrentDb.Execute "PIPELINES_UPDATE"
    MsgBox "Putkilinjojen päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe putkilinjojen päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Editorilomakkeet (Commands 18-19, 40) =====
'================================================================================
' Tapahtuma: Command18_Click
' Tarkoitus: Avaa FlowPickNo-editorilomake
'================================================================================
Private Sub Command18_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "FlowPickNo") = vbYes Then
    DoCmd.OpenForm "USysFlowpickNo"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe FlowPickNo-lomakkeen avaamisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command19_Click
' Tarkoitus: Avaa putkilinjan FROM/TO-editorilomake
'================================================================================
Private Sub Command19_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "Pipeline From/To") = vbYes Then
    DoCmd.OpenForm "USysPipeFromTo"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe PipeFromTo-lomakkeen avaamisessa: " & Err.Description, vbExclamation
End Sub

'===== Käsiventtiilityönkulku (Commands 21-23) =====
'================================================================================
' Tapahtuma: Command21_Click
' Tarkoitus: Lukee käsiventtiilitiedot AutoCAD-piirustuksista
'================================================================================
Private Sub Command21_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "MANVALVEDATA", "MANUALVALVE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Käsiventtiilien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Virhe käsiventtiilien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command22_Click
' Tarkoitus: Append manual valve data to staging table
'================================================================================
Private Sub Command22_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "MANVALVE_APPEND"
  MsgBox "Käsiventtiilien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe käsiventtiilien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command23_Click
' Tarkoitus: Päivittää tietokantaan käsiventtiilidata väliaikaistaulukosta
'================================================================================
Private Sub Command23_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "MANVALVE_UPDATE"
    MsgBox "Käsiventtiilien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe käsiventtiilien päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Instrumenttiventtiilityönkulku (Commands 31-33) =====
'================================================================================
' Tapahtuma: Command31_Click
' Tarkoitus: Lukee instrumenttiventtiilitiedot AutoCAD-piirustuksista
' Huom: Searches for UP079D blocks where LOOPNO attribute contains "AA"
'================================================================================
Private Sub Command31_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedotByBlockAndAttribute "INSTVALVEDATA", "UP079D", "LOOPNO", "AA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Instrumenttiventtiilien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Virhe instrumenttiventtiilien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command32_Click
' Tarkoitus: Append instrument valve data to staging table
'================================================================================
Private Sub Command32_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "INSTVALVE_APPEND"
  MsgBox "Instrumenttiventtiilien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttiventtiilien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command33_Click
' Tarkoitus: Päivittää tietokantaan instrumenttiventtiilidata väliaikaistaulukosta
'================================================================================
Private Sub Command33_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "INSTVALVE_UPDATE"
    MsgBox "Instrumenttiventtiilien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttiventtiilien päivityksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command40_Click
' Tarkoitus: Avaa lomake putkilinjaviittausten linkittämiseen instrumenttisilmukoihin
'================================================================================
Private Sub Command40_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADissä?", vbYesNo, "Pipeline To Loops") = vbYes Then
    DoCmd.OpenForm "USysPipeToOther", , , , , , "UP079D,UP008C,UP085D,UP08C,"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe PipeToOther-instrumenttipiirilomakkeen avaamisessa: " & Err.Description, vbExclamation
End Sub

'===== Kenttälaitetyönkulku (Commands 41-43) =====
'================================================================================
' Tapahtuma: Command41_Click
' Tarkoitus: Lukee kenttälaitetiedot AutoCAD-piirustuksista (UP080D-blokit)
'================================================================================
Private Sub Command41_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "FIELDINSTDATA", "UP080D"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Kenttäinstrumenttien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Virhe kenttäinstrumenttien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command42_Click
' Tarkoitus: Append field instrument data to staging table
'================================================================================
Private Sub Command42_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "FIELDINST_APPEND"
  MsgBox "Kenttäinstrumenttien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe kenttäinstrumenttien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command43_Click
' Tarkoitus: Päivittää tietokantaan kenttälaitteen data väliaikaistaulukosta
'================================================================================
Private Sub Command43_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYesNo Then
    CurrentDb.Execute "FIELDINST_UPDATE"
    MsgBox "Kenttäinstrumenttien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe kenttäinstrumenttien päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Instrumenttisilmukkatyönkulku (Commands 51-53) =====
'================================================================================
' Tapahtuma: Command51_Click
' Tarkoitus: Lukee instrumenttisilmukkatiedot AutoCAD-piirustuksista
'================================================================================
Private Sub Command51_Click()
  On Error GoTo ErrorHandler
  
  Tied = ""
  LueTiedot "INSTLOOPDATA", "UP079D"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Instrumenttipiirien luku valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Virhe instrumenttipiirien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command52_Click
' Tarkoitus: Append instrument loop data to staging table
'================================================================================
Private Sub Command52_Click()
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "INSTLOOP_APPEND"
  MsgBox "Instrumenttipiirien lisäys valmis", vbInformation
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttipiirien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command53_Click
' Tarkoitus: Päivittää tietokantaan instrumenttisilmukkata väliaikaistaulukosta
'================================================================================
Private Sub Command53_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "Päivitä tiedot") = vbYes Then
    CurrentDb.Execute "INSTLOOP_UPDATE"
    MsgBox "Instrumenttipiirien päivitys valmis", vbInformation
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttipiirien päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Apufunktiot =====
'================================================================================
' Funktio: FieldExists
' Tarkoitus: Tarkistaa, onko kenttä olemassa recordsetissä
' Parametrit:
'   RS - Tarkistettava recordset
'   KentänNimi - Kenttänimi, jota etsitään
' Palautusarvo: True jos kenttä löytyy, False muuten
'================================================================================
Private Function FieldExists(RS As DAO.Recordset, FieldName As String) As Boolean
  Dim fld As DAO.Field
  On Error Resume Next
  Set fld = RS.Fields(FieldName)
  FieldExists = (Err.Number = 0)
  On Error GoTo 0
End Function

'================================================================================
' Aliohjelma: LueTiedot
' Tarkoitus: Lukee blokkitiedot AutoCAD-piirustuksista ja tallentaa Access-taulukkoon
' Parametrit:
'   Taulukko - Access-kohdontaulukon nimi
'   Blokki - Blokkityyppi/tyypit, jotka poimitaan (pilkuilla erotettu)
'================================================================================
Private Sub LueTiedot(Taulukko As String, Blokki As String)
  Dim i As Integer, j As Integer
  Dim oBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
    PoimiJoukko Blokki
    
    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOC.Path
      Tiedot.Fields("ImpFileID") = oDOC.Name
      Tiedot.Fields("BLOCK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.Close
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.SysCmd acSysCmdClearStatus
  Set oACAD = Nothing
  Set oDOC = Nothing
  Exit Sub
  
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.Close
    Set Tiedot = Nothing
  End If
  MsgBox "Virhe tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: LueTiedotByAttribute
' Tarkoitus: Lukee blokkitiedot AutoCAD-piirustuksista, suodatettuna attribuuttinimien mukaan
' Parametrit:
'   Taulukko - Access-kohdontaulukon nimi
'   AttribNames - Attribuuttitagi/tagit, joiden mukaan suodatetaan (pilkuilla erotettu)
'================================================================================
Private Sub LueTiedotByAttribute(Taulukko As String, AttribNames As String)
  Dim i As Integer, j As Integer
  Dim oBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
    PoimiJoukkoByAttribute AttribNames
    
    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOC.Path
      Tiedot.Fields("ImpFileID") = oDOC.Name
      Tiedot.Fields("BLOCK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.Close
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.SysCmd acSysCmdClearStatus
  Set oACAD = Nothing
  Set oDOC = Nothing
  Exit Sub
  
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.Close
    Set Tiedot = Nothing
  End If
  MsgBox "Virhe tietojen lukemisessa attribuuteilla: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: LueTiedotByBlockAndAttribute
' Tarkoitus: Lukee blokit tietyllä nimellä ja attribuuttiarvolla
' Parametrit:
'   Taulukko - Access-kohdontaulukon nimi
'   BlockName - Etsittävän blokin nimi
'   AttribName - Tarkistettavan attribuutin tagi
'   ContainsValue - Arvo, jolla attribuutin arvon täytyy alkaa
'================================================================================
Private Sub LueTiedotByBlockAndAttribute(Taulukko As String, BlockName As String, AttribName As String, ContainsValue As String)
  Dim i As Integer, j As Integer
  Dim oBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim Attrib As Variant
  Dim Tiedot As DAO.Recordset
  Dim fld As DAO.Field
  
  On Error GoTo ErrorHandler
  
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  
  Do While SeuraavaKuva
    Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
    PoimiJoukkoByBlockAndAttributeValue BlockName, AttribName, ContainsValue
    
    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      Attrib = oBlock.GetAttributes
      
      Tiedot.AddNew
      Tiedot.Fields("PATH") = oDOC.Path
      Tiedot.Fields("ImpFileID") = oDOC.Name
      Tiedot.Fields("BLOCK") = oBlock.EffectiveName
      Tiedot.Fields("Handles") = oBlock.Handle
      
      For j = 0 To UBound(Attrib)
        If Attrib(j).TextString <> "" And FieldExists(Tiedot, "[" & Attrib(j).TagString & "]") Then
          Set fld = Tiedot.Fields("[" & Attrib(j).TagString & "]")
          fld.Value = IIf(Len(Attrib(j).TextString) > fld.Size, Left$(Attrib(j).TextString, fld.Size), Attrib(j).TextString)
        End If
      Next j
      
      Tiedot.Update
    Next i
    
    Tiedot.Close
    Set Tiedot = Nothing
    Joukko.Delete
    Set Joukko = Nothing
  Loop
  
  Application.SysCmd acSysCmdClearStatus
  Set oACAD = Nothing
  Set oDOC = Nothing
  Exit Sub
  
ErrorHandler:
  Application.SysCmd acSysCmdClearStatus
  If Not Tiedot Is Nothing Then
    Tiedot.Close
    Set Tiedot = Nothing
  End If
  MsgBox "Virhe suodatettujen tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Funktio: Yhdista
' Tarkoitus: YhdistÃ¤Ã¤ kÃ¤ynnissÃ¤ olevaan AutoCAD-instanssiin
' Palautusarvo: True jos yhteys onnistui, False muuten
'================================================================================
Private Function Yhdista() As Boolean
  On Error Resume Next
  
  ' Yhdistetään käynnissä olevaan AutoCADiin
  Set oACAD = GetObject(, "AutoCAD.Application")
  If Err <> 0 Then
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Yhdista = False
    Exit Function
  Else
    Yhdista = True
  End If
  On Error GoTo 0
End Function

'================================================================================
' Funktio: SeuraavaKuva
' Tarkoitus: Avaa seuraava .DWG-tiedosto määritetystä kansiosta eräajoa varten
' Palautusarvo: True jos tiedosto avattu, False jos ei enää tiedostoja
'================================================================================
Private Function SeuraavaKuva() As Boolean
  Dim Polku As String
  
  On Error GoTo ErrorHandler
  
  If oACAD Is Nothing Then
    If Yhdista = False Then
      SeuraavaKuva = False
      Exit Function
    End If
  End If
  
  Polku = Left$(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))
  Polku = Polku & LPolku.Caption
    
  If Tied = "" Then
    Tied = Dir$(Polku & "*.dwg")
  Else
    Tied = Dir$
  End If
  
  If Tied = "" Then
    SeuraavaKuva = False
    Exit Function
  End If
  
  ' Suljetaan edellinen piirustus
  If Not oDOC Is Nothing Then
    oDOC.Close False
  End If
  
  ' Avataan seuraava piirustus
  Set oDOC = oACAD.Documents.Open(Polku & Tied)
  SeuraavaKuva = True
  Exit Function
  
ErrorHandler:
  MsgBox "Virhe piirustuksen avaamisessa: " & Err.Description, vbExclamation
  SeuraavaKuva = False
End Function

'================================================================================
' Aliohjelma: PoimiJoukko
' Tarkoitus: Luo suodatettu valintajoukko määritetyistä blokkityypeistä
' Parametrit:
'   Nimet - Blokin nimi/nimet, pilkuilla erotettu useampaa tyyppiä varten
'================================================================================
Private Sub PoimiJoukko(Nimet As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim Blokit As Variant
  Dim Poista() As Object  ' AcadEntity - muutettu late binding (64-bit)
  Dim L As Integer
  Dim EiPoisteta As Boolean
  
  On Error GoTo ErrorHandler
  
  ' Jäsennetään blokkinimet
  Blokit = Split(Nimet, ",")
  For i = 0 To UBound(Blokit)
    Blokit(i) = UCase$(Trim$(Blokit(i)))
  Next i
  
  ' Luodaan valintajoukko
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Suodatus: poistetaan blokit, jotka eivät ole listalla
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If UCase$(Joukko(j).EffectiveName) = Blokit(i) Then
          EiPoisteta = True
          Exit For
        End If
      Next i
      If EiPoisteta = False Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
  MsgBox "Virhe valintajoukon luomisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: PoimiJoukkoByAttribute
' Tarkoitus: Luo valintajoukko blokeista, joilla on tietyt attribuutit
' Parametrit:
'   AttribNames - Etsittävä attribuuttitagi/tagit, pilkuilla erotettu
'================================================================================
Private Sub PoimiJoukkoByAttribute(AttribNames As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim k As Integer
  Dim Attribs As Variant
  Dim AttribList As Variant
  Dim Poista() As Object  ' AcadEntity - muutettu late binding (64-bit)
  Dim L As Integer
  Dim HasAttribute As Boolean
  
  On Error GoTo ErrorHandler
  
  ' Jäsennetään etsittävät attribuuttinimet
  AttribList = Split(AttribNames, ",")
  For i = 0 To UBound(AttribList)
    AttribList(i) = UCase$(Trim$(AttribList(i)))
  Next i
  
  ' Luodaan valintajoukko kaikista blokeista
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Suodatus: pidetään vain blokit, joilla on määritetyt attribuutit
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasAttribute = False
      
      ' Tarkistetaan, onko blokilla attribuutteja
      If Joukko(j).HasAttributes Then
        Attribs = Joukko(j).GetAttributes
        
        ' Käydään läpi kaikki blokin attribuutit
        For k = 0 To UBound(Attribs)
          ' Tarkistetaan, täsmääkö attribuutti hakuun
          For i = 0 To UBound(AttribList)
            If UCase$(Attribs(k).TagString) = AttribList(i) Then
              HasAttribute = True
              Exit For
            End If
          Next i
          If HasAttribute Then Exit For
        Next k
      End If
      
      ' Jos blokilla ei ole attribuuttia, merkitään poistettavaksi
      If Not HasAttribute Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
  MsgBox "Virhe attribuuttisuodatetun valintajoukon luomisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Aliohjelma: PoimiJoukkoByBlockAndAttributeValue
' Tarkoitus: Luo valintajoukko tietyistä blokeista, joiden attribuutti sisältää arvon
' Parametrit:
'   BlockName - Etsittävän blokin nimi
'   AttribName - Tarkistettavan attribuutin tagi
'   ContainsValue - Arvo, jonka attribuutin täytyy sisältää (ei kirjainkokoherkkä)
'================================================================================
Private Sub PoimiJoukkoByBlockAndAttributeValue(BlockName As String, AttribName As String, ContainsValue As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim j As Integer
  Dim Attribs As Variant
  Dim Poista() As Object  ' AcadEntity - muutettu late binding (64-bit)
  Dim L As Integer
  Dim HasMatch As Boolean
  Dim AttribValue As String
  
  On Error GoTo ErrorHandler
  
  BlockName = UCase$(Trim$(BlockName))
  AttribName = UCase$(Trim$(AttribName))
  ContainsValue = UCase$(Trim$(ContainsValue))
  
  ' Luodaan valintajoukko kaikista blokeista
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Suodatus: pidetään vain vastaavat blokit, joiden attribuuttiarvo täsmää
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasMatch = False
      
      ' Tarkistetaan blokin nimi
      If UCase$(Joukko(j).EffectiveName) = BlockName Then
        ' Tarkistetaan, onko blokilla attribuutteja
        If Joukko(j).HasAttributes Then
          Attribs = Joukko(j).GetAttributes
          
          ' Käydään läpi kaikki attribuutit
          For i = 0 To UBound(Attribs)
            If UCase$(Attribs(i).TagString) = AttribName Then
              AttribValue = UCase$(Attribs(i).TextString)
              ' Tarkistetaan, alkaako attribuutin arvo määritetyllä arvolla
              If Left$(AttribValue, Len(ContainsValue)) = ContainsValue Then
                HasMatch = True
                Exit For
              End If
            End If
          Next i
        End If
      End If
      
      ' Jos blokki ei täsmää kriteereihin, merkitään poistettavaksi
      If Not HasMatch Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
MsgBox "Virhe suodatetun valintajoukon luomisessa: " & Err.Description, vbExclamation
End Sub
