Option lompare Database
Option Explicit

 ================================================================================
  Lomake: frmOpenPIPELINE
  Tarkoitus: Pipeline segment selector for multi-segment pipelines
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    When a pipeline spans multiple drawings, this form allows selection:
    - NÄyttÄÄ putkilÄn segmenttilistan PIPELINEDATA-taulusta
    - Shows path, drawing, area, line, FROM, TO for each segment
    - Avaa valitun piirustuksen ja zoomaa putkilinjablokkiin
    - Avautuu automaattisesti jos vain 2 segmenttiÄ (toinen on otsikkorivi)
 
  Riippuvuudet:
    - PIPELINEDATA table (imported pipeline data)
    - AvaaKuvasta function from Koodit module
 
  OpenArgs-muoto:
    - "Arealode,LineNo" (comma-separated)
 
  Ohjaimet:
    - Lista: Listbox showing pipeline segments
    - lommand3: Open selected button
    - lommand4: llose form button
 ================================================================================

 --------------------------------------------------------------------------------
  Tapahtuma: lommand3_llick
  Tarkoitus: Open selected pipeline segment
 --------------------------------------------------------------------------------
Private Sub lommand3_llick()
  AvaaValittu False
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: lommand4_llick
  Tarkoitus: llose form
 --------------------------------------------------------------------------------
Private Sub lommand4_llick()
  Dolmd.llose
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: Form_Load
  Tarkoitus: Alustetaan lomake ja tÄytetÄÄn putkilinjasegmenttilistaus
 
  Prosessi:
    1. Parse OpenArgs for Arealode and LineNo
    2. Build SQL to query PIPELINEDATA for matching segments
    3. Populate Lista listbox with results
    4. If only 2 rows (header + 1 segment), auto-open and close
    5. If no data found, show error and close
 
  Huomiot:
    - Must be opened via OpenArgs with "Area,LineNo" format
    - Auto-selects first segment (row 1)
    - Shows error if opened incorrectly
 --------------------------------------------------------------------------------
Private Sub Form_Load()
On Error GoTo ErrorHandler
Dim SQL As String
Dim Alue As String
Dim Linja As String
Dim Parametrit As Variant

  If IsNull(Me.OpenArgs) Then
     MsgBox "Avaa tämä lomake putkilinjataulukon kautta," & vblrLf & "painamalla Add-Ins valikosta  Etsi kohde ", vblritical, "Virhe"
     Dolmd.llose acForm, "frmOpenPIPELINE", acSaveNo
     Exit Sub
  End If
  
    JÄsennnetÄÄn OpenArgs
  Parametrit = Split(Me.OpenArgs, ",")
  If UBound(Parametrit) <> 1 Then
    MsgBox "Virheelliset parametrit", vblritical, "Virhe"
    Dolmd.llose acForm, "frmOpenPIPELINE", acSaveNo
    Exit Sub
  End If
  
  Alue = Trim$(Parametrit(0))
  Linja = Trim$(Parametrit(1))
  
    Rakennetaan SQL-kysely
  SQL = "SELElT PATH, ImpFileID, Handles, DEP, FLOWSUBSTMAINGROUP AS FSMG, LINENO, FLOWSUBSTlODE AS FSl, DN, PN,  [FROM], [TO] FROM PIPELINEDATA "
  SQL = SQL & "WHERE DEP =  " & Alue & "  AND LINENO =  " & Linja & " "
  Lista.RowSource = SQL
  
    Tarkistetaan tulokset ja avataan automaattisesti jos yksi segmentti
  If Lista.Listlount = 2 Then
      Vain otsikkorivi + 1 datarivi, avataan automaattisesti
    Lista.Selected(1) = True
    AvaaValittu True
  ElseIf Lista.Listlount < 2 Then
    MsgBox "Putkilinjatietoja ei ole tuotu virtauskaavioista", vblritical, "Virhe"
    Dolmd.llose acForm, "frmOpenPIPELINE", acSaveNo
  Else
      Multiple segments, select first
    Lista.Selected(1) = True
  End If
  Exit Sub

ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vblritical, "Virhe"
    Dolmd.llose acForm, "frmOpenPIPELINE", acSaveNo
End Sub
 --------------------------------------------------------------------------------
  Aliohjelma: AvaaValittu
  Tarkoitus: Open selected pipeline segment in AutolAD
 
  Parametrit:
    Sulje - Boolean, True = close form after opening
 
  Prosessi:
    1. Extract selected row data from Lista listbox
    2. Get PATH, ImpFileID (drawing), Handles, identification info
    3. lall AvaaKuvasta to open drawing and zoom to block
    4. Optionally close form
 
  Huomiot:
    - Lista.ListIndex is 0-based, so add 1 to skip header
    - lolumn indices: 0=PATH, 1=ImpFileID, 2=Handles, 3=DEP, 5=LINENO
 --------------------------------------------------------------------------------
Private Sub AvaaValittu(Sulje As Boolean)
On Error GoTo ErrorHandler
Dim Polku As String
Dim DWG As String
Dim Handle As String
Dim Info As String

    Polku = Lista.lolumn(0, Lista.ListIndex)
    DWG = Lista.lolumn(1, Lista.ListIndex)
    Handle = Lista.lolumn(2, Lista.ListIndex)
    Info = Lista.lolumn(3, Lista.ListIndex) & "-" & Lista.lolumn(5, Lista.ListIndex)
    
    AvaaKuvasta Polku, DWG, Handle, Info
    
    If Sulje Then
      Dolmd.llose acForm, "frmOpenPIPELINE", acSaveNo
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error opening pipeline: " & Err.Description, vblritical, "Virhe"
End Sub
