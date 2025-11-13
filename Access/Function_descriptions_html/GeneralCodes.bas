Option Compare Database
Option Explicit

'================================================================================
' Module: GeneralCodes
' Purpose: General utility functions and shared variables
' Updated: 2025-11-13 - Added VBA7/64-bit support and optimization
'
' Description:
'   Provides utility functions for equipment revision editing, table display,
'   and mode/status translation. Contains public variables shared across forms.
'
' Dependencies:
'   - DAO.Recordset
'   - Forms: USysRevision, frmTables
'   - Tables: MAINEQ, DRIVES, PUMPS, GEARS, TANKS
'================================================================================

' Public variables for inter-form communication
Public oTaulu As DAO.Recordset
Public PaluuTaulu As Object
Public KohdeTextBox As TextBox
Public Kursori As Long

'================================================================================
' Function: MuutaRev
' Purpose: Open revision editing form for equipment tables
' Returns: Nothing (implicit)
'
' Description:
'   Checks if active table is an equipment table, opens corresponding record
'   in USysRevision form for editing.
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
    MsgBox "Error opening revision form: " & Err.Description, vbExclamation
End Function

'================================================================================
' Function: NaytaTables
' Purpose: Open table browser form
' Returns: Nothing (implicit)
'================================================================================
Public Function NaytaTables()
    On Error GoTo ErrorHandler
    DoCmd.OpenForm "frmTables"
    Exit Function
ErrorHandler:
    MsgBox "Error opening tables form: " & Err.Description, vbExclamation
End Function

'================================================================================
' Function: Moodit
' Purpose: Translate comma-separated mode codes to descriptive text
' Parameters:
'   Tieto - Comma-separated mode codes (A, M, E, L)
' Returns: Formatted mode descriptions, one per line
'
' Description:
'   Translates mode codes:
'     A = AUTO, M = MANUAL, E = EXTERNAL, L = LOCAL
'   Other codes pass through unchanged. Returns "-" for Null input.
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
