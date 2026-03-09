Option Compare Database
Option Explicit

' Updated 2025-11-11: Added VBA7 conditional compilation for 64-bit compatibility
' This API retrieves the current Windows user name for revision tracking
#If VBA7 Then
    Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#End If

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

' PaivitaIDRev: Updates the ID and Revision combo boxes based on selected table
' Automatically detects AutoNumber fields and revision columns
' Updated 2025-11-11: Added DAO typing, improved comments
Private Sub PaivitaIDRev()
Dim i As Integer
Dim Taul As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
Dim Revi As String
  ID.RowSource = ATaulu.Value
  Rev.RowSource = ATaulu.Value
For i = 0 To Rev.ListCount - 1
  If Left$(LCase$(Rev.ItemData(i)), 3) = "rev" Then
    Rev.Value = Rev.ItemData(i)
    Exit For
  End If
Next i
' Avataan taulu kenttien tyyppien tarkistamista varten (suljetaan ennen uudelleenkäyttöä)
Set Taul = CurrentDb.OpenRecordset(ATaulu.Value)
For i = 0 To Taul.Fields.Count - 1
  If Taul.Fields(i).Type = 4 And Taul.Fields(i).Attributes = 49 Then
    ID.Value = Taul.Fields(i).Name
    Exit For
  End If
Next i
Taul.Close
Set Taul = Nothing
If ATaulu <> "" And Rev <> "" Then
  ' Avataan uusi kysely viimeisimmän revision selvittämiseksi
  Set Taul = CurrentDb.OpenRecordset("SELECT " & Rev & " FROM " & ATaulu & " WHERE " & Rev & " Is Not Null")
  If Not Taul.EOF Then
    Revi = Left$(Taul.Fields(0), InStr(Taul.Fields(0), " ") - 1)
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
  Taul.Close
  Set Taul = Nothing
End If
End Sub
Private Sub Sulje_Click()
 DoCmd.Close
End Sub

' Tee: Main revision marking function
' Compares two tables and marks changed records with revision information
' Updated 2025-11-11: Added DAO typing, improved comments
Private Sub Tee_Click()
Dim Vert As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
Dim Muuttunut As Boolean
Dim i As Integer

' Validate that all required selections are made
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

' Create a joined recordset comparing the two tables
Set Vert = CurrentDb.OpenRecordset("SELECT [" & ATaulu & "].*, [" & UTaulu & "].* FROM [" & ATaulu & "] RIGHT JOIN [" & UTaulu & "] ON [" & UTaulu & "]." & ID & " = [" & ATaulu & "]." & ID & ";")  ' Updated 2025-11-11: CurrentDB -> CurrentDb
Do While Not Vert.EOF
  Muuttunut = False
  ' Check each field (except ID and Rev) for changes
  For i = 0 To ID.ListCount - 1
    If ID.ItemData(i) <> ID.Value And ID.ItemData(i) <> Rev.Value Then
      ' Compare values between the two tables (Nz handles nulls)
      If Nz(Vert.Fields(ATaulu & "." & ID.ItemData(i)).Value) <> Nz(Vert.Fields(UTaulu & "." & ID.ItemData(i)).Value) Then
        Muuttunut = True
        Exit For
      End If
    End If
  Next i
  
  ' If record has changed, update revision information
  If Muuttunut Then
    Vert.Edit
    If IsNull(Vert.Fields(UTaulu & "." & Rev).Value) Or Vert.Fields(UTaulu & "." & Rev).Value = "" Then
      ' First revision entry
      Vert.Fields(UTaulu & "." & Rev).Value = TRev & " " & TDate & "/" & TMaker & "/" & TComment
    Else
      ' Append to existing revision history
      Vert.Fields(UTaulu & "." & Rev).Value = TRev & " " & TDate & "/" & TMaker & "/" & TComment & vbCrLf & Vert.Fields(UTaulu & "." & Rev).Value
    End If
    Vert.Update
  End If
  Vert.MoveNext
Loop
Merk.Visible = False
Vert.Close
Set Vert = Nothing
MsgBox "Valmis", vbOKOnly, "Tee merkinnät"
End Sub
