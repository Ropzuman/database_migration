VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Revisiointi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Sub ATaulu_Change()
  PaivitaIDRev
End Sub
Private Sub FiltteriB_Click()
  PaivitaTaulut
End Sub

Private Sub Form_Load()
Dim BuffSize As Long
Dim NBuffer As String

  BuffSize = 256
  NBuffer = Space$(BuffSize)
    
  FiltteriT.Value = "*"
  PaivitaTaulut
  TDate.Value = Date
    
  If api_GetUserName(NBuffer, BuffSize) Then
     TMaker.Value = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
  End If
  TComment.Value = "Updated"
End Sub
Private Sub PaivitaTaulut()
  If FiltteriT.Value = "" Then FiltteriT.Value = "*"
  ATaulu.RowSource = "SELECT Name From MSysObjects WHERE [Type] = 1 and Name Not Like 'MSys*' AND Name Like '" & FiltteriT.Value & "'"
  UTaulu.RowSource = "SELECT Name From MSysObjects WHERE [Type] = 1 and Name Not Like 'MSys*' AND Name Like '" & FiltteriT.Value & "'"
  ATaulu.Value = ATaulu.ItemData(0)
  UTaulu.Value = ATaulu.ItemData(0)
  PaivitaIDRev
End Sub
Private Sub PaivitaIDRev()
Dim i As Integer
Dim Taul As Recordset
Dim Revi As String
  ID.RowSource = ATaulu.Value
  Rev.RowSource = ATaulu.Value
For i = 0 To Rev.ListCount - 1
  If Left(LCase(Rev.ItemData(i)), 3) = "rev" Then
    Rev.Value = Rev.ItemData(i)
    Exit For
  End If
Next i
Set Taul = CurrentDB.OpenRecordset(ATaulu.Value)
For i = 0 To Taul.Fields.Count - 1
  If Taul.Fields(i).Type = 4 And Taul.Fields(i).Attributes = 49 Then
    ID.Value = Taul.Fields(i).Name
    Exit For
  End If
Next i
If ATaulu <> "" And Rev <> "" Then
  Set Taul = CurrentDB.OpenRecordset("SELECT " & Rev & " FROM " & ATaulu & " WHERE " & Rev & " Is Not Null")
  If Not Taul.EOF Then
    Revi = Left(Taul.Fields(0), InStr(Taul.Fields(0), " ") - 1)
    If Len(Revi) > 1 Then
      Revi = CStr(VAL(Revi) + 1)
      If Len(Revi) < 2 Then
        Revi = "0" & Revi
      End If
    Else
      Revi = Chr(Asc(Revi) + 1)
    End If
    TRev = Revi
  End If
End If
End Sub
Private Sub Sulje_Click()
 DoCmd.Close
End Sub
Private Sub Tee_Click()
Dim Vert As Recordset
Dim Muuttunut As Boolean
Dim i As Integer
If ATaulu = "" Or UTaulu = "" Or ID = "" Or Rev = "" Then
  MsgBox "Valitse ensin taulukot ja sarakkeet", vbCritical, "Tee merkinnät"
  Exit Sub
End If
If TRev = "" Or TDate = "" Or TMaker = "" Then
  MsgBox "Laita ensin revisiomerkinnät", vbCritical, "Tee merkinnät"
  Exit Sub
End If
Merk.Visible = True
Me.Repaint
Set Vert = CurrentDB.OpenRecordset("SELECT [" & ATaulu & "].*, [" & UTaulu & "].* FROM [" & ATaulu & "] RIGHT JOIN [" & UTaulu & "] ON [" & UTaulu & "]." & ID & " = [" & ATaulu & "]." & ID & ";")
Do While Not Vert.EOF
  Muuttunut = False
  For i = 0 To ID.ListCount - 1
    If ID.ItemData(i) <> ID.Value And ID.ItemData(i) <> Rev.Value Then
      If Nz(Vert.Fields(ATaulu & "." & ID.ItemData(i)).Value) <> Nz(Vert.Fields(UTaulu & "." & ID.ItemData(i)).Value) Then
        Muuttunut = True
        Exit For
      End If
    End If
  Next i
  If Muuttunut Then
    Vert.Edit
    If IsNull(Vert.Fields(UTaulu & "." & Rev).Value) Or Vert.Fields(UTaulu & "." & Rev).Value = "" Then
      Vert.Fields(UTaulu & "." & Rev).Value = TRev & " " & TDate & "/" & TMaker & "/" & TComment
    Else
      Vert.Fields(UTaulu & "." & Rev).Value = TRev & " " & TDate & "/" & TMaker & "/" & TComment & vbCrLf & Vert.Fields(UTaulu & "." & Rev).Value
    End If
    Vert.Update
  Else
    Debug.Print "EI MUUTTUNUT"
  End If
  Vert.MoveNext
Loop
Merk.Visible = False
MsgBox "Valmis", vbOKOnly, "Tee merkinnät"

End Sub
