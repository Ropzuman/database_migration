Option Compare Database
Option Explicit

' AutoCAD-vakio määritellään paikallisesti, koska tyyppikirjastoviittausta ei käytetä
Private Const acSelectionSetAll As Long = 5

'================================================================================
' Lomake: TYÖKALUT
' Tarkoitus: AutoCAD P&I -tietojen tuonti- ja päivitystyökalupakki PIPE-tietokantaan
' Yhteensopivuus: VBA7, DAO ja AutoCADin myöhäinen sidonta
'
' Kuvaus:
'   Keskuslomake P&I-kaaviotietojen tuontiin ja päivitykseen AutoCADista:
'   - Käsittelee DWG-tiedostot eräajona LPolku-otsikon määrittämästä kansiosta
'   - Tuo putkilinja-, venttiili-, kenttälaite- ja instrumenttisilmukkatiedot
'   - Kunkin tietotyypin työkalut (esimerkiksi Command11–13):
'     1. Lue AutoCADista välitauluun.
'     2. Lisää tiedot välitaulusta tietokannan tauluihin.
'     3. Päivitä tietokannan tiedot välitaulusta.
'   - Avaa erikoiseditorilomakkeet (FlowPickNo, PipeFromTo, PipeToOther)
'   - Haku korvaa välitaulun vain, kun osumia löytyy ja koko luku onnistuu.
'     Nollatuloksessa tai virheessä muutos pyritään perumaan; palautusvirhe ilmoitetaan.
'   - Välitaulun on tuettava DAO-tapahtumia.
'   - Avoinna olevat piirustukset jätetään auki ja niiden nykyiset tiedot luetaan.
'   - Lisäys- ja päivityskyselyistä näytetään viimeisen kyselyn muuttamien rivien määrä.
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
'   - INSTVALVEDATA: UP079D-blokit, joiden FUNCTION on HIC tai HS.
'     AREACODE saadaan samannimisestä attribuutista; LOOPNO:n viisi ensimmäistä
'     ja viimeistä merkkiä tallennetaan kenttiin VALVE_FUNCTION ja VALVE_NO.
'   - FIELDINSTDATA: UP080D-blokit; FUNCTION ja AREACODE kopioidaan attribuuteista.
'   - INSTLOOPDATA: UP079D-blokit, joiden FUNCTION ei ole HIC tai HS.
'     Myös blokit, joilla FUNCTION puuttuu tai on tyhjä, otetaan mukaan.
'   - FIELDINSTDATA- ja INSTLOOPDATA-tauluissa LOOPSYMB ja LOOPNO muodostetaan
'     LOOPNO-attribuutin viidestä ensimmäisestä ja viidestä viimeisestä merkistä.
'
' Kontrollit:
'   - Command11-13: Putkilinjatyökalut
'   - Command21-23: Käsiventtiilityökalut
'   - Command31-33: Instrumenttiventtiilityökalut
'   - Command41-43: Kenttälaitetyökalut
'   - Command51-53: Instrumenttisilmukkatyökalut
'   - Command18, 19, 40: Avaavat editorilomakkeet
'   - BPipelineToManval: Avaa putkilinjojen ja käsiventtiilien linkityslomakkeen
'================================================================================

Dim oACAD As Object  ' AutoCAD-sovellus myöhäisellä sidonnalla
Dim oDOC As Object  ' AutoCAD-piirustus myöhäisellä sidonnalla
Dim KuvaAvattuHaussa As Boolean
Dim Tied As String  ' Eräajossa käsiteltävän DWG-tiedoston nimi
Dim Joukko As Object  ' AutoCAD-valintajoukko myöhäisellä sidonnalla

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

'===== Putkilinjatyökalut (komennot 11–13) =====
'================================================================================
' Tapahtuma: Command11_Click
' Tarkoitus: Lukee putkilinjatiedot AutoCAD-piirustuksista
'================================================================================
Private Sub Command11_Click()
  Dim Lukumaara As Long
  On Error GoTo ErrorHandler
  
  Tied = ""
  Lukumaara = LueTiedot("PIPELINEDATA", "PIPELINE_DATA")
  NaytaLukumaara "Putkilinjojen", Lukumaara
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe putkilinjojen tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command12_Click
' Tarkoitus: Lisää välitaulun putkilinjatiedot tietokannan tauluihin
'================================================================================
Private Sub Command12_Click()
  On Error GoTo ErrorHandler
  
  NaytaKyselynTulos "Putkilinjojen lisäys", SuoritaKyselyt("PIPELINES_APPEND")
  
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
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf & "Haluatko jatkaa?", vbExclamation + vbYesNo + vbDefaultButton2, "Päivitä tiedot") = vbYes Then
    NaytaKyselynTulos "Putkilinjojen päivitys", SuoritaKyselyt("DELETE * FROM UsysPIPEUPD", "PIPELINES_UPDATE_START", "PIPELINES_UPDATE")
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe putkilinjojen päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Editorilomakkeet (komennot 18, 19 ja 40) =====
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

'===== Käsiventtiilityökalut (komennot 21–23) =====
'================================================================================
' Tapahtuma: Command21_Click
' Tarkoitus: Lukee käsiventtiilitiedot AutoCAD-piirustuksista
'================================================================================
Private Sub Command21_Click()
  Dim Lukumaara As Long
  On Error GoTo ErrorHandler
  
  Tied = ""
  Lukumaara = LueTiedot("MANVALVEDATA", "MANUALVALVE_DATA")
  NaytaLukumaara "Käsiventtiilien", Lukumaara
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe käsiventtiilien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command22_Click
' Tarkoitus: Lisää välitaulun käsiventtiilitiedot tietokannan tauluihin
'================================================================================
Private Sub Command22_Click()
  On Error GoTo ErrorHandler
  
  NaytaKyselynTulos "Käsiventtiilien lisäys", SuoritaKyselyt("MANVALVE_APPEND")
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe käsiventtiilien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command23_Click
' Tarkoitus: Päivittää tietokannan käsiventtiilitiedot välitaulusta
'================================================================================
Private Sub Command23_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbExclamation + vbYesNo + vbDefaultButton2, "Päivitä tiedot") = vbYes Then
    NaytaKyselynTulos "Käsiventtiilien päivitys", SuoritaKyselyt("MANVALVE_UPDATE")
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe käsiventtiilien päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Instrumenttiventtiilityökalut (komennot 31–33) =====
'================================================================================
' Tapahtuma: Command31_Click
' Tarkoitus: Lukee instrumenttiventtiilitiedot AutoCAD-piirustuksista
' Huom: Mukaan otetaan UP079D-blokit, joiden FUNCTION on HIC tai HS.
'================================================================================
Private Sub Command31_Click()
  Dim Lukumaara As Long
  On Error GoTo ErrorHandler
  
  Tied = ""
  Lukumaara = LueTiedotByBlockAndAttribute("INSTVALVEDATA", "UP079D", "FUNCTION", "HIC,HS")
  NaytaLukumaara "Instrumenttiventtiilien", Lukumaara
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttiventtiilien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command32_Click
' Tarkoitus: Lisää välitaulun instrumenttiventtiilitiedot tietokannan tauluihin
'================================================================================
Private Sub Command32_Click()
  On Error GoTo ErrorHandler
  
  NaytaKyselynTulos "Instrumenttiventtiilien lisäys", SuoritaKyselyt("INSTVALVE_APPEND")
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttiventtiilien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command33_Click
' Tarkoitus: Päivittää tietokannan instrumenttiventtiilitiedot välitaulusta
'================================================================================
Private Sub Command33_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbExclamation + vbYesNo + vbDefaultButton2, "Päivitä tiedot") = vbYes Then
    NaytaKyselynTulos "Instrumenttiventtiilien päivitys", SuoritaKyselyt("INSTVALVE_UPDATE")
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
    DoCmd.OpenForm "USysPipeToOther", , , , , , "UP079D"
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe PipeToOther-instrumenttipiirilomakkeen avaamisessa: " & Err.Description, vbExclamation
End Sub

'===== Kenttälaitetyökalut (komennot 41–43) =====
'================================================================================
' Tapahtuma: Command41_Click
' Tarkoitus: Lukee kenttälaitetiedot AutoCAD-piirustuksista (UP080D-blokit)
'================================================================================
Private Sub Command41_Click()
  Dim Lukumaara As Long
  On Error GoTo ErrorHandler
  
  Tied = ""
  Lukumaara = LueTiedot("FIELDINSTDATA", "UP080D")
  NaytaLukumaara "Kenttäinstrumenttien", Lukumaara
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe kenttäinstrumenttien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command42_Click
' Tarkoitus: Lisää välitaulun kenttälaitetiedot tietokannan tauluihin
'================================================================================
Private Sub Command42_Click()
  On Error GoTo ErrorHandler
  
  NaytaKyselynTulos "Kenttäinstrumenttien lisäys", SuoritaKyselyt("FIELDINST_APPEND")
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe kenttäinstrumenttien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command43_Click
' Tarkoitus: Päivittää tietokannan kenttälaitetiedot välitaulusta
'================================================================================
Private Sub Command43_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbExclamation + vbYesNo + vbDefaultButton2, "Päivitä tiedot") = vbYes Then
    NaytaKyselynTulos "Kenttäinstrumenttien päivitys", SuoritaKyselyt("FIELDINST_UPDATE")
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe kenttäinstrumenttien päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Instrumenttisilmukkatyökalut (komennot 51–53) =====
'================================================================================
' Tapahtuma: Command51_Click
' Tarkoitus: Lukee UP079D-blokit, joiden FUNCTION ei ole HIC tai HS
'================================================================================
Private Sub Command51_Click()
  Dim Lukumaara As Long
  On Error GoTo ErrorHandler
  
  Tied = ""
  Lukumaara = LueTiedotYhteinen("INSTLOOPDATA", "UP079D", "FUNCTION", "HIC,HS", 3)
  NaytaLukumaara "Instrumenttipiirien", Lukumaara
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttipiirien tietojen lukemisessa: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command52_Click
' Tarkoitus: Lisää välitaulun instrumenttisilmukkatiedot tietokannan tauluihin
'================================================================================
Private Sub Command52_Click()
  On Error GoTo ErrorHandler
  
  NaytaKyselynTulos "Instrumenttipiirien lisäys", SuoritaKyselyt("INSTLOOP_APPEND")
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttipiirien lisäyksessä: " & Err.Description, vbExclamation
End Sub

'================================================================================
' Tapahtuma: Command53_Click
' Tarkoitus: Päivittää tietokannan instrumenttisilmukkatiedot välitaulusta
'================================================================================
Private Sub Command53_Click()
  On Error GoTo ErrorHandler
  
  If MsgBox("Varoitus! Tämä ajaa kaavioissa olevat tiedot kannassa olevien tietojen päälle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia näihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbExclamation + vbYesNo + vbDefaultButton2, "Päivitä tiedot") = vbYes Then
    NaytaKyselynTulos "Instrumenttipiirien päivitys", SuoritaKyselyt("INSTLOOP_UPDATE")
  End If
  
  Exit Sub
ErrorHandler:
  MsgBox "Virhe instrumenttipiirien päivityksessä: " & Err.Description, vbExclamation
End Sub

'===== Apufunktiot =====
Private Sub NaytaKyselynTulos(Kuvaus As String, Maara As Long)
  If Maara = 0 Then
    MsgBox Kuvaus & ": yhtään riviä ei muuttunut.", vbExclamation
  Else
    MsgBox Kuvaus & " valmis. Muuttuneita rivejä: " & Maara & ".", vbInformation
  End If
End Sub

Private Function SuoritaKyselyt(Kysely1 As String, Optional Kysely2 As String = "", Optional Kysely3 As String = "") As Long
  Dim Tietokanta As DAO.Database
  Dim Tyoalue As DAO.Workspace
  Dim Kyselyt As Variant
  Dim Kysely As Variant
  Dim KyselynNimi As String
  Dim TapahtumaKaynnissa As Boolean
  Dim VirheNumero As Long
  Dim VirheKuvaus As String
  Dim VirheLahde As String

  On Error GoTo ErrorHandler
  Set Tyoalue = DBEngine.Workspaces(0)
  Set Tietokanta = CurrentDb
  If Not Tietokanta.Transactions Then
    Err.Raise vbObjectError + 1110, "SuoritaKyselyt", "Tietokanta ei tue tapahtumia."
  End If
  Kyselyt = Array(Kysely1, Kysely2, Kysely3)
  Tyoalue.BeginTrans
  TapahtumaKaynnissa = True
  For Each Kysely In Kyselyt
    KyselynNimi = CStr(Kysely)
    If KyselynNimi <> "" Then
      Tietokanta.Execute KyselynNimi, dbFailOnError
      SuoritaKyselyt = Tietokanta.RecordsAffected
    End If
  Next Kysely
  Tyoalue.CommitTrans
  TapahtumaKaynnissa = False
  Exit Function

ErrorHandler:
  VirheNumero = Err.Number
  VirheKuvaus = Err.Description
  VirheLahde = Err.Source
  If KyselynNimi <> "" Then VirheKuvaus = VirheKuvaus & vbCrLf & "Kysely: " & KyselynNimi
  On Error Resume Next
  If TapahtumaKaynnissa Then
    Err.Clear
    Tyoalue.Rollback
    If Err.Number <> 0 Then VirheKuvaus = VirheKuvaus & vbCrLf & "Muutosten palautus epäonnistui: " & Err.Description
  End If
  On Error GoTo 0
  Err.Raise VirheNumero, VirheLahde, VirheKuvaus
End Function

Private Sub AsetaTekstikentta(Tiedot As DAO.Recordset, Kentta As String, Arvo As String, Handle As String)
  Dim Kohde As DAO.Field
  Set Kohde = Tiedot.Fields(Kentta)
  If Kohde.Type = dbText And Len(Arvo) > Kohde.Size Then
    Err.Raise vbObjectError + 1104, "AsetaTekstikentta", "Arvo on liian pitkä kenttään " & Kentta & _
        " (pituus " & Len(Arvo) & ", enimmäispituus " & Kohde.Size & ", blokin handle " & Handle & ")."
  End If
  Kohde.Value = Arvo
End Sub

Private Sub NaytaLukumaara(Kuvaus As String, Lukumaara As Long)
  If Lukumaara = 0 Then
    MsgBox Kuvaus & " haussa ei löytynyt tietueita. Välitaulussa on edelleen aiemman haun tiedot; älä aja niiden perusteella lisäys- tai päivityskyselyjä.", vbExclamation
  Else
    MsgBox Kuvaus & " luku valmis. Välitauluun tallennettiin " & Lukumaara & " tietuetta.", vbInformation
  End If
End Sub

Private Function FieldExists(RS As DAO.Recordset, FieldName As String) As Boolean
  Dim fld As DAO.Field
  Dim VirheNumero As Long
  Dim VirheKuvaus As String
  Dim VirheLahde As String

  On Error Resume Next
  Err.Clear
  Set fld = RS.Fields(FieldName)
  VirheNumero = Err.Number
  VirheKuvaus = Err.Description
  VirheLahde = Err.Source
  On Error GoTo 0

  If VirheNumero = 0 Then
    FieldExists = True
  ElseIf VirheNumero <> 3265 Then
    Err.Raise VirheNumero, VirheLahde, VirheKuvaus
  End If
End Function

Private Function LueTiedot(Taulukko As String, Blokki As String) As Long
  LueTiedot = LueTiedotYhteinen(Taulukko, Blokki, "", "", 0)
End Function

Private Function LueTiedotByAttribute(Taulukko As String, AttribNames As String) As Long
  LueTiedotByAttribute = LueTiedotYhteinen(Taulukko, "", AttribNames, "", 1)
End Function

Private Function LueTiedotByBlockAndAttribute(Taulukko As String, BlockName As String, AttribName As String, AttributeValues As String) As Long
  LueTiedotByBlockAndAttribute = LueTiedotYhteinen(Taulukko, BlockName, AttribName, AttributeValues, 2)
End Function

Private Function LueTiedotYhteinen(Taulukko As String, Blokki As String, AttribName As String, AttributeValues As String, HakuTapa As Long) As Long
  Dim i As Long, j As Long
  Dim oBlock As Object
  Dim Attrib As Variant
  Dim LoopNo As String
  Dim Aluekoodi As String
  Dim OnkoVenttiilikartoitus As Boolean
  Dim OnkoPiirikartoitus As Boolean
  Dim Tietokanta As DAO.Database
  Dim Tyoalue As DAO.Workspace
  Dim Tiedot As DAO.Recordset
  Dim Lukumaara As Long
  Dim Kuvia As Long
  Dim TapahtumaKaynnissa As Boolean
  Dim TiedotVahvistettu As Boolean
  Dim VirheNumero As Long
  Dim VirheKuvaus As String
  Dim VirheLahde As String

  On Error GoTo ErrorHandler
  Set Tyoalue = DBEngine.Workspaces(0)
  Set Tietokanta = CurrentDb
  OnkoVenttiilikartoitus = (StrComp(Taulukko, "INSTVALVEDATA", vbTextCompare) = 0)
  OnkoPiirikartoitus = (StrComp(Taulukko, "INSTLOOPDATA", vbTextCompare) = 0 Or _
      StrComp(Taulukko, "FIELDINSTDATA", vbTextCompare) = 0)
  Set Tiedot = Tietokanta.OpenRecordset(Taulukko, dbOpenDynaset)
  If Not Tiedot.Transactions Then
    Err.Raise vbObjectError + 1100, "LueTiedotYhteinen", "Välitaulu ei tue tapahtumia; aiempia tietoja ei voida suojata."
  End If

  Tyoalue.BeginTrans
  TapahtumaKaynnissa = True
  Tietokanta.Execute "DELETE * FROM [" & Taulukko & "]", dbFailOnError

  Do While SeuraavaKuva()
    Kuvia = Kuvia + 1
    Select Case HakuTapa
      Case 0
        PoimiJoukko Blokki
      Case 1
        PoimiJoukkoByAttribute AttribName
      Case 2
        PoimiJoukkoByBlockAndAttributeValue Blokki, AttribName, AttributeValues
      Case 3
        PoimiJoukkoByBlockAndAttributeValue Blokki, AttribName, AttributeValues, True
      Case Else
        Err.Raise vbObjectError + 1101, "LueTiedotYhteinen", "Tuntematon hakutapa."
    End Select

    For i = 0 To Joukko.Count - 1
      Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & (i + 1) & "/" & Joukko.Count
      Set oBlock = Joukko(i)
      LoopNo = ""
      Aluekoodi = ""
      Tiedot.AddNew
      AsetaTekstikentta Tiedot, "PATH", oDOC.Path, oBlock.Handle
      AsetaTekstikentta Tiedot, "ImpFileID", oDOC.Name, oBlock.Handle
      AsetaTekstikentta Tiedot, "BLOCK", oBlock.EffectiveName, oBlock.Handle
      AsetaTekstikentta Tiedot, "Handles", oBlock.Handle, oBlock.Handle

      If oBlock.HasAttributes Then
        Attrib = oBlock.GetAttributes
        For j = LBound(Attrib) To UBound(Attrib)
          If OnkoVenttiilikartoitus Or OnkoPiirikartoitus Then
            Select Case UCase$(Attrib(j).TagString)
              Case "LOOPNO"
                LoopNo = Attrib(j).TextString
              Case "AREACODE"
                Aluekoodi = Attrib(j).TextString
            End Select
          End If
          If Attrib(j).TextString <> "" Then
            If (Not OnkoVenttiilikartoitus Or _
                (UCase$(Attrib(j).TagString) <> "VALVE_FUNCTION" And UCase$(Attrib(j).TagString) <> "VALVE_NO")) And _
                (Not OnkoPiirikartoitus Or UCase$(Attrib(j).TagString) <> "LOOPSYMB") Then
              If FieldExists(Tiedot, Attrib(j).TagString) Then
                If OnkoPiirikartoitus And StrComp(Attrib(j).TagString, "LOOPNO", vbTextCompare) = 0 Then
                  AsetaTekstikentta Tiedot, Attrib(j).TagString, Right$(Attrib(j).TextString, 5), oBlock.Handle
                Else
                  AsetaTekstikentta Tiedot, Attrib(j).TagString, Attrib(j).TextString, oBlock.Handle
                End If
              End If
            End If
          End If
        Next j
      End If
      If OnkoVenttiilikartoitus Then
        If Len(Trim$(Aluekoodi)) = 0 Then
          Err.Raise vbObjectError + 1105, "LueTiedotYhteinen", "Blokilta puuttuu AREACODE (handle " & oBlock.Handle & ")."
        End If
        If Len(Trim$(LoopNo)) = 0 Then
          Err.Raise vbObjectError + 1106, "LueTiedotYhteinen", "Blokilta puuttuu LOOPNO (handle " & oBlock.Handle & ")."
        End If
        AsetaTekstikentta Tiedot, "AREACODE", Aluekoodi, oBlock.Handle
        AsetaTekstikentta Tiedot, "VALVE_FUNCTION", Left$(LoopNo, 5), oBlock.Handle
        AsetaTekstikentta Tiedot, "VALVE_NO", Right$(LoopNo, 5), oBlock.Handle
      ElseIf OnkoPiirikartoitus Then
        If Len(Trim$(LoopNo)) = 0 Then
          Err.Raise vbObjectError + 1106, "LueTiedotYhteinen", "Blokilta puuttuu LOOPNO (handle " & oBlock.Handle & ")."
        End If
        AsetaTekstikentta Tiedot, "LOOPSYMB", Left$(LoopNo, 5), oBlock.Handle
        AsetaTekstikentta Tiedot, "LOOPNO", Right$(LoopNo, 5), oBlock.Handle
      End If
      Tiedot.Update
      Lukumaara = Lukumaara + 1
    Next i

    Joukko.Delete
    Set Joukko = Nothing
    VapautaPiirustus
  Loop

  If Kuvia = 0 Then
    Err.Raise vbObjectError + 1102, "LueTiedotYhteinen", "Hakukansiosta ei löytynyt DWG-piirustuksia."
  End If

  Application.SysCmd acSysCmdClearStatus
  If Lukumaara = 0 Then
    Tyoalue.Rollback
  Else
    Tyoalue.CommitTrans
    TiedotVahvistettu = True
  End If
  TapahtumaKaynnissa = False
  Tiedot.Close
  Set Tiedot = Nothing
  Set oDOC = Nothing
  Set oACAD = Nothing
  LueTiedotYhteinen = Lukumaara
  Exit Function

ErrorHandler:
  VirheNumero = Err.Number
  VirheKuvaus = Err.Description
  VirheLahde = Err.Source
  If Tied <> "" Then VirheKuvaus = VirheKuvaus & " (piirustus: " & Tied & ")"
  If TiedotVahvistettu Then VirheKuvaus = VirheKuvaus & vbCrLf & "Uudet välitaulun tiedot ehdittiin jo vahvistaa."
  On Error Resume Next
  If TapahtumaKaynnissa Then
    Err.Clear
    Tyoalue.Rollback
    If Err.Number <> 0 Then VirheKuvaus = VirheKuvaus & vbCrLf & "Välitaulun palautus epäonnistui: " & Err.Description
  End If
  If Not Tiedot Is Nothing Then
    Err.Clear
    Tiedot.Close
    If Err.Number <> 0 Then VirheKuvaus = VirheKuvaus & vbCrLf & "Tietuejoukon sulkeminen epäonnistui: " & Err.Description
  End If
  If Not Joukko Is Nothing Then
    Err.Clear
    Joukko.Delete
    If Err.Number <> 0 Then VirheKuvaus = VirheKuvaus & vbCrLf & "Valintajoukon poisto epäonnistui: " & Err.Description
  End If
  Set Joukko = Nothing
  Err.Clear
  VapautaPiirustus
  If Err.Number <> 0 Then VirheKuvaus = VirheKuvaus & vbCrLf & "Piirustuksen vapautus epäonnistui: " & Err.Description
  Application.SysCmd acSysCmdClearStatus
  Set oDOC = Nothing
  Set oACAD = Nothing
  On Error GoTo 0
  Err.Raise VirheNumero, VirheLahde, VirheKuvaus
End Function

'================================================================================
' Funktio: Yhdista
' Tarkoitus: Yhdistää käynnissä olevaan AutoCAD-instanssiin
' Palautusarvo: True jos yhteys onnistui; virhe välitetään kutsujalle
'================================================================================
Private Function Yhdista() As Boolean
  On Error GoTo ErrorHandler
  Set oACAD = GetObject(, "AutoCAD.Application")
  Yhdista = True
  Exit Function
ErrorHandler:
  Err.Raise vbObjectError + 1103, "Yhdista", "Käynnissä olevaan AutoCADiin ei saatu yhteyttä: " & Err.Description
End Function

'================================================================================
' Funktio: SeuraavaKuva
' Tarkoitus: Ottaa seuraavan DWG-piirustuksen käyttöön eräajossa
' Palautusarvo: True, jos piirustus on käytettävissä; False, kun tiedostot loppuvat
'================================================================================
Private Function SeuraavaKuva() As Boolean
  Dim Polku As String
  Dim Tiedosto As String
  Dim AvoinKuva As Object
  Dim Tiedostot As Object
  
  On Error GoTo ErrorHandler
  
  If oACAD Is Nothing Then
    If Not Yhdista() Then Err.Raise vbObjectError + 1103, "SeuraavaKuva", "AutoCAD-yhteyden muodostaminen epäonnistui."
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
  
  Set Tiedostot = CreateObject("Scripting.FileSystemObject")
  Tiedosto = Tiedostot.GetAbsolutePathName(Polku & Tied)
  For Each AvoinKuva In oACAD.Documents
    If StrComp(Tiedostot.GetAbsolutePathName(AvoinKuva.FullName), Tiedosto, vbTextCompare) = 0 Then
      Set oDOC = AvoinKuva
      KuvaAvattuHaussa = False
      SeuraavaKuva = True
      Exit Function
    End If
  Next AvoinKuva

  Set oDOC = oACAD.Documents.Open(Tiedosto)
  KuvaAvattuHaussa = True
  SeuraavaKuva = True
  Exit Function
  
ErrorHandler:
  Err.Raise Err.Number, "SeuraavaKuva", "Piirustuksen avaaminen epäonnistui: " & Err.Description
End Function

Private Sub VapautaPiirustus()
  If Not oDOC Is Nothing Then
    If KuvaAvattuHaussa Then
      If Not oDOC.Saved Then
        Err.Raise vbObjectError + 1111, "VapautaPiirustus", "Haun avaamassa piirustuksessa on tallentamattomia muutoksia; piirustus jätettiin auki: " & oDOC.Name
      End If
      oDOC.Close False
    End If
    Set oDOC = Nothing
  End If
  KuvaAvattuHaussa = False
End Sub

Private Sub PoistaVanhaValintajoukko()
  Dim VanhaJoukko As Object
  For Each VanhaJoukko In oDOC.SelectionSets
    If StrComp(VanhaJoukko.Name, "APUPICK", vbTextCompare) = 0 Then
      VanhaJoukko.Delete
      Exit Sub
    End If
  Next VanhaJoukko
End Sub

'================================================================================
' Aliohjelma: PoimiJoukko
' Tarkoitus: Luo suodatettu valintajoukko määritetyistä blokkityypeistä
' Parametrit:
'   Nimet - Haettavat blokkien nimet pilkuilla erotettuina
'================================================================================
Private Sub PoimiJoukko(Nimet As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Long
  Dim j As Long
  Dim Blokit As Variant
  Dim Poista() As Object  ' AutoCAD-olio myöhäisellä sidonnalla
  Dim L As Long
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
  
  PoistaVanhaValintajoukko
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Suodatus: poistetaan blokit, jotka eivät ole listalla
  ' Vertailussa poistetaan mahdollinen ulkoisen viitteen nimen etuliite.
  Dim sNimi As String
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      sNimi = UCase$(Joukko(j).EffectiveName)
      If InStr(sNimi, "|") > 0 Then sNimi = Mid$(sNimi, InStrRev(sNimi, "|") + 1)
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If sNimi = Blokit(i) Then
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
  Err.Raise Err.Number, "PoimiJoukko", "Virhe valintajoukon luomisessa: " & Err.Description
End Sub

'================================================================================
' Aliohjelma: PoimiJoukkoByAttribute
' Tarkoitus: Luo valintajoukko blokeista, joilla on vähintään yksi annetuista attribuuteista
' Parametrit:
'   AttribNames - Etsittävät attribuuttien tunnisteet pilkuilla erotettuina
'================================================================================
Private Sub PoimiJoukkoByAttribute(AttribNames As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Long
  Dim j As Long
  Dim k As Long
  Dim Attribs As Variant
  Dim AttribList As Variant
  Dim Poista() As Object  ' AutoCAD-olio myöhäisellä sidonnalla
  Dim L As Long
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
  
  PoistaVanhaValintajoukko
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Mukaan jäävät blokit, joilla on vähintään yksi haetuista attribuuteista.
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
  Err.Raise Err.Number, "PoimiJoukkoByAttribute", "Virhe attribuuttisuodatetun valintajoukon luomisessa: " & Err.Description
End Sub

'================================================================================
' Aliohjelma: PoimiJoukkoByBlockAndAttributeValue
' Tarkoitus: Luo valintajoukko joko hyväksyen tai pois sulkien annetut attribuuttiarvot
' Parametrit:
'   BlockName - Haettavan blokin nimi
'   AttribName - Tarkistettavan attribuutin tunniste
'   AttributeValues - Vertailtavat arvot pilkuilla erotettuina; kirjainkoolla ei ole merkitystä
'   SuljeArvot - True sulkee täsmäävät arvot pois; puuttuva attribuutti jää mukaan
'================================================================================
Private Sub PoimiJoukkoByBlockAndAttributeValue(BlockName As String, AttribName As String, AttributeValues As String, Optional SuljeArvot As Boolean = False)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Long
  Dim j As Long
  Dim Attribs As Variant
  Dim Poista() As Object  ' AutoCAD-olio myöhäisellä sidonnalla
  Dim L As Long
  Dim HasMatch As Boolean
  Dim NimiTasmaa As Boolean
  Dim AttribValue As String
  Dim Arvot As Variant
  Dim k As Long
  Dim BlockNimi As String
  
  On Error GoTo ErrorHandler
  
  BlockName = UCase$(Trim$(BlockName))
  AttribName = UCase$(Trim$(AttribName))
  Arvot = Split(AttributeValues, ",")
  
  ' Luodaan valintajoukko kaikista blokeista
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  
  PoistaVanhaValintajoukko
  
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  ' Attribuuttiarvot joko hyväksytään tai suljetaan pois hakutavan mukaan.
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      HasMatch = False
      
      BlockNimi = UCase$(Joukko(j).EffectiveName)
      If InStr(BlockNimi, "|") > 0 Then BlockNimi = Mid$(BlockNimi, InStrRev(BlockNimi, "|") + 1)
      NimiTasmaa = (BlockNimi = BlockName)
      If NimiTasmaa Then
        ' Tarkistetaan, onko blokilla attribuutteja
        If Joukko(j).HasAttributes Then
          Attribs = Joukko(j).GetAttributes
          
          ' Käydään läpi kaikki attribuutit
          For i = 0 To UBound(Attribs)
            If UCase$(Attribs(i).TagString) = AttribName Then
              AttribValue = Trim$(Attribs(i).TextString)
              For k = LBound(Arvot) To UBound(Arvot)
                If StrComp(AttribValue, Trim$(Arvot(k)), vbTextCompare) = 0 Then
                  HasMatch = True
                  Exit For
                End If
              Next k
              If HasMatch Then Exit For
            End If
          Next i
        End If
      End If
      
      ' Jos blokki ei täsmää kriteereihin, merkitään poistettavaksi
      If Not NimiTasmaa Or (HasMatch = SuljeArvot) Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  
  If L > 0 Then Joukko.RemoveItems Poista
  Exit Sub
  
ErrorHandler:
  Err.Raise Err.Number, "PoimiJoukkoByBlockAndAttributeValue", "Virhe suodatetun valintajoukon luomisessa: " & Err.Description
End Sub
