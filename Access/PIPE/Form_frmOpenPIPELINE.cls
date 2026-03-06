Option Compare Database
Option Explicit

'================================================================================
' Lomake: frmOpenPIPELINE
' Tarkoitus: Monisegmenttisten putkilinjojen segmentinvalitsin
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki lisätty
'
' Kuvaus:
'   Kun putkilinja ulottuu useisiin piirustuksiin, tällä lomakkeella valitaan segmentti:
'   - Näyttää putkilinjasegmenttien listan PIPELINEDATA-taulukosta
'   - Näyttää polun, piirustuksen, alueen, linjan, FROM ja TO kullekin segmentille
'   - Opens selected drawing and zooms to pipeline block
'   - Avautuu automaattisesti, jos vain 2 segmenttiä (yksi on otsikkorivi)
'
' Riippuvuudet:
'   - PIPELINEDATA-taulukko (tuotu putkilinjadata)
'   - AvaaKuvasta-funktio Koodit-moduulista
'
' OpenArgs-muoto:
'   - "AluetKoodi,LinjaNo" (pilkuilla erotettu)
'
' Kontrollit:
'   - Lista: Listaruutu putkilinjasegmenteille
'   - Command3: Avaa valittu -painike
'   - Command4: Sulje lomake -painike
'================================================================================

'--------------------------------------------------------------------------------
' Tapahtuma: Command3_Click
' Tarkoitus: Avaa valittu putkilinjasegmentti
'--------------------------------------------------------------------------------
Private Sub Command3_Click()
  AvaaValittu False
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Command4_Click
' Tarkoitus: Sulkee lomakkeen
'--------------------------------------------------------------------------------
Private Sub Command4_Click()
  DoCmd.Close
End Sub

'--------------------------------------------------------------------------------
' Tapahtuma: Form_Load
' Tarkoitus: Alustetaan lomake ja tÃ¤ytetÃ¤Ã¤n putkilinjasegmenttilista
'
' Toiminta:
'   1. Jäsennetään OpenArgs (aluekoodi ja linjanumero)
'   2. Rakennetaan SQL-kysely PIPELINEDATA-taulukkoon
'   3. Täytetään Lista-listaruutu tuloksilla
'   4. Jos vain 2 riviä (otsikko + 1 segmentti), avataan automaattisesti
'   5. Jos dataa ei löydy, näytetään virhe ja suljetaan
'
' Huom:
'   - Avattava OpenArgs-parametrilla muotoiltuna "Alue,LinjaNo"
'   - Valitsee ensimmäisen segmentin automaattisesti (rivi 1)
'   - Näyttää virheen, jos lomake avataan väärin
'--------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim SQL As String
Dim Alue As String
Dim Linja As String
Dim Parametrit As Variant

  If IsNull(Me.OpenArgs) Then
     MsgBox "Avaa tämä lomake putkilinjataulukon kautta," & vbCrLf & "painamalla Add-Ins valikosta 'Etsi kohde'", vbCritical, "Virhe"
     DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
     Exit Sub
  End If
  
  ' Jäsennetään OpenArgs-parametrit
  Parametrit = Split(Me.OpenArgs, ",")
  If UBound(Parametrit) <> 1 Then
    MsgBox "Virheelliset parametrit", vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
    Exit Sub
  End If
  
  Alue = Trim$(Parametrit(0))
  Linja = Trim$(Parametrit(1))
  
  ' Rakennetaan SQL-kysely
  SQL = "SELECT PATH, ImpFileID, Handles, DEP, FLOWSUBSTMAINGROUP AS FSMG, LINENO, FLOWSUBSTCODE AS FSC, DN, PN,  [FROM], [TO] FROM PIPELINEDATA "
  SQL = SQL & "WHERE DEP = '" & Alue & "' AND LINENO = '" & Linja & "'"
  Lista.RowSource = SQL
  
  ' Tarkistetaan tulokset - avataan automaattisesti, jos yksi segmentti
  If Lista.ListCount = 2 Then
    ' Vain otsikko + 1 datarivi, avataan automaattisesti
    Lista.Selected(1) = True
    AvaaValittu True
  ElseIf Lista.ListCount < 2 Then
    MsgBox "Putkilinjatietoja ei ole tuotu virtauskaavioista", vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
  Else
    ' Useita segmentejä, valitaan ensimmäinen
    Lista.Selected(1) = True
  End If
  Exit Sub

ErrorHandler:
    MsgBox "Virhe lomakkeen latauksessa: " & Err.Description, vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
End Sub
'--------------------------------------------------------------------------------
' Aliohjelma: AvaaValittu
' Tarkoitus: Avaa valittu putkilinjasegmentti AutoCADissa
'
' Parametrit:
'   Sulje - Boolean, True = suljetaan lomake avauksen jälkeen
'
' Toiminta:
'   1. Poimitaan valitun rivin tiedot Lista-listaruudusta
'   2. Haetaan polku, piirustus (ImpFileID), Handle ja tunnistetiedot
'   3. Kutsutaan AvaaKuvasta avaamaan piirustus ja zoomaamaan blokkiin
'   4. Suljetaan lomake tarvittaessa
'
' Huom:
'   - Lista.ListIndex on 0-pohjainen, lisätään 1 otsikon ohittamiseksi
'   - Sarakeindeksit: 0=PATH, 1=ImpFileID, 2=Handles, 3=DEP, 5=LINENO
'--------------------------------------------------------------------------------
Private Sub AvaaValittu(Sulje As Boolean)
On Error GoTo ErrorHandler
Dim Polku As String
Dim DWG As String
Dim Handle As String
Dim Info As String

    Polku = Lista.Column(0, Lista.ListIndex)
    DWG = Lista.Column(1, Lista.ListIndex)
    Handle = Lista.Column(2, Lista.ListIndex)
    Info = Lista.Column(3, Lista.ListIndex) & "-" & Lista.Column(5, Lista.ListIndex)
    
    AvaaKuvasta Polku, DWG, Handle, Info
    
    If Sulje Then
      DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Virhe putkilinjan avaamisessa: " & Err.Description, vbCritical, "Virhe"
End Sub
