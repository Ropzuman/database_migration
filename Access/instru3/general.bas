Option Compare Database
Option Explicit
Public Sivunro As Integer
Public EdelArea As Integer
Public Sivuja As Integer

' Updated 2025-11-11: Added VBA7/64-bit support for GetOpenFileName API
#If VBA7 Then
    Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As LongPtr  ' Updated for 64-bit
        hInstance As LongPtr  ' Updated for 64-bit
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As LongPtr  ' Updated for 64-bit
        lpfnHook As LongPtr  ' Updated for 64-bit
        lpTemplateName As String
    End Type
#Else
    Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
    Public Type OPENFILENAME
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
    End Type
#End If
Public Function PilkkuPiste(Luku As Variant) As String
Dim Osoitin As Long
If Nz(Luku) = "" Then
PilkkuPiste = ""
Else

  Osoitin = InStr(Luku, ",")
  If Osoitin = 0 Then
    PilkkuPiste = Luku
  Else
    PilkkuPiste = Left(Luku, Osoitin - 1) & "." & Mid(Luku, Osoitin + 1)
  End If
End If
End Function
Public Function UdNoteToRev(UdNote As Variant) As Variant
Dim Paiva As String
Dim Os As Long
Dim VP As Date
Dim RevTaul As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
If IsNull(UdNote) Then
  UdNoteToRev = Null
Else
 Os = InStr(UdNote, ":")
 If Os > 0 Then
   Paiva = Mid(UdNote, Os + 1)
   Paiva = Left(Paiva, InStr(Paiva, "|") - 1)
   VP = DateValue(Paiva)
   Paiva = Month(VP) & "/" & Day(VP) & "/" & Year(VP)   'Esim. 2/1/2007
   Set RevTaul = CurrentDb.OpenRecordset("SELECT * FROM _Revisions WHERE (((BeforeDate) > #" & Paiva & "#)) ORDER BY BeforeDate ASC;")
   If RevTaul.RecordCount > 0 Then
     UdNoteToRev = RevTaul.Fields("Rev").Value
   End If
 End If
End If
End Function
Function EtsiLoop(Alue As String, Looppi As String) As String
Dim Taul As DAO.Recordset  ' Updated 2025-11-11: Added DAO prefix for early binding
Set Taul = CurrentDb.OpenRecordset("SELECT * From qrysolvalve WHERE AreaCode='" & Alue & "' AND LoopNo='" & Looppi & "'")
If Taul.EOF Then
  EtsiLoop = ""
Else
  EtsiLoop = "1"
End If
Taul.Close
Set Taul = Nothing
End Function
