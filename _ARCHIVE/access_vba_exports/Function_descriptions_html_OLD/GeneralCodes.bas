Option Compare Database
Option Explicit
Public oTaulu As Recordset
Public PaluuTaulu As Object
Public KohdeTextBox As TextBox
Public Kursori As Long
Public Function MuutaRev()
Dim Taul As String
Taul = Application.CurrentObjectName
If Taul = "MAINEQ" Or Taul = "DRIVES" Or Taul = "PUMPS" Or Taul = "GEARS" Or Taul = "TANKS" Then
  Set PaluuTaulu = Screen.ActiveDatasheet
  Set oTaulu = CurrentDb.OpenRecordset("SELECT * FROM " & Taul & " WHERE ID=" & Screen.ActiveDatasheet("ID").Value)
  DoCmd.OpenForm "USysRevision"
End If

End Function
Public Function NaytaTables()
    DoCmd.OpenForm "frmTables"
End Function
Function Moodit(Tieto As Variant) As Variant
Dim Tiedot As Variant
Dim i As Integer
Dim T As String
If IsNull(Tieto) Then
  Moodit = "-" 'Null
Else
  Tiedot = Split(Tieto, ",")
  For i = 0 To UBound(Tiedot)
    Select Case UCase(Trim(Tiedot(i)))
      Case "A"
        T = "AUTO"
      Case "M"
        T = "MANUAL"
      Case "E"
        T = "EXTERNAL"
      Case "L"
        T = "LOCAL"
      Case Else
        T = UCase(Trim(Tiedot(i)))
    End Select
    If i > 0 Then
      Moodit = Moodit & vbCrLf & T
    Else
      Moodit = T
    End If
      
  Next i
End If
End Function
