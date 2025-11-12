Option Compare Database
Option Explicit

'================================================================================
' Form: frmOpenPIPELINE
' Purpose: Pipeline segment selector for multi-segment pipelines
' Updated: 2025-11-12 - Added VBA7/64-bit support and optimization
'
' Description:
'   When a pipeline spans multiple drawings, this form allows selection:
'   - Displays list of pipeline segments from PIPELINEDATA
'   - Shows path, drawing, area, line, FROM, TO for each segment
'   - Opens selected drawing and zooms to pipeline block
'   - Auto-opens if only 2 segments (one is header row)
'
' Dependencies:
'   - PIPELINEDATA table (imported pipeline data)
'   - AvaaKuvasta function from Koodit module
'
' OpenArgs Format:
'   - "AreaCode,LineNo" (comma-separated)
'
' Controls:
'   - Lista: Listbox showing pipeline segments
'   - Command3: Open selected button
'   - Command4: Close form button
'================================================================================

'--------------------------------------------------------------------------------
' Event: Command3_Click
' Purpose: Open selected pipeline segment
'--------------------------------------------------------------------------------
Private Sub Command3_Click()
  AvaaValittu False
End Sub

'--------------------------------------------------------------------------------
' Event: Command4_Click
' Purpose: Close form
'--------------------------------------------------------------------------------
Private Sub Command4_Click()
  DoCmd.Close
End Sub

'--------------------------------------------------------------------------------
' Event: Form_Load
' Purpose: Initialize form and populate pipeline segment list
'
' Process:
'   1. Parse OpenArgs for AreaCode and LineNo
'   2. Build SQL to query PIPELINEDATA for matching segments
'   3. Populate Lista listbox with results
'   4. If only 2 rows (header + 1 segment), auto-open and close
'   5. If no data found, show error and close
'
' Notes:
'   - Must be opened via OpenArgs with "Area,LineNo" format
'   - Auto-selects first segment (row 1)
'   - Shows error if opened incorrectly
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
  
  ' Parse OpenArgs
  Parametrit = Split(Me.OpenArgs, ",")
  If UBound(Parametrit) <> 1 Then
    MsgBox "Virheelliset parametrit", vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
    Exit Sub
  End If
  
  Alue = Trim$(Parametrit(0))
  Linja = Trim$(Parametrit(1))
  
  ' Build SQL query
  SQL = "SELECT PATH, ImpFileID, Handles, DEP, FLOWSUBSTMAINGROUP AS FSMG, LINENO, FLOWSUBSTCODE AS FSC, DN, PN,  [FROM], [TO] FROM PIPELINEDATA "
  SQL = SQL & "WHERE DEP = '" & Alue & "' AND LINENO = '" & Linja & "'"
  Lista.RowSource = SQL
  
  ' Check results and auto-open if single segment
  If Lista.ListCount = 2 Then
    ' Only header + 1 data row, auto-open
    Lista.Selected(1) = True
    AvaaValittu True
  ElseIf Lista.ListCount < 2 Then
    MsgBox "Putkilinjatietoja ei ole tuotu virtauskaavioista", vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
  Else
    ' Multiple segments, select first
    Lista.Selected(1) = True
  End If
  Exit Sub

ErrorHandler:
    MsgBox "Error loading form: " & Err.Description, vbCritical, "Virhe"
    DoCmd.Close acForm, "frmOpenPIPELINE", acSaveNo
End Sub
'--------------------------------------------------------------------------------
' Sub: AvaaValittu
' Purpose: Open selected pipeline segment in AutoCAD
'
' Parameters:
'   Sulje - Boolean, True = close form after opening
'
' Process:
'   1. Extract selected row data from Lista listbox
'   2. Get PATH, ImpFileID (drawing), Handles, identification info
'   3. Call AvaaKuvasta to open drawing and zoom to block
'   4. Optionally close form
'
' Notes:
'   - Lista.ListIndex is 0-based, so add 1 to skip header
'   - Column indices: 0=PATH, 1=ImpFileID, 2=Handles, 3=DEP, 5=LINENO
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
    MsgBox "Error opening pipeline: " & Err.Description, vbCritical, "Virhe"
End Sub
