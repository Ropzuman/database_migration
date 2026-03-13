Attribute VB_Name = "Start"
' --- 64-bitti korjattu: PtrSafe lisatty kaikkiin Declare-lauseisiin ---
' --------- [ CHOOSE FILE ] -----------------
Declare PtrSafe Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Declare PtrSafe Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As LongPtr
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
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
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
End Type
Public DB As New ADODB.Connection
Public TauluLoops As New ADODB.Recordset
Public TauluMotors As New ADODB.Recordset
Public TauluINSTTAG As New ADODB.Recordset
Public Tietokanta As String
Public StartOK As Boolean
Public AreaNo As String
Public PreviewText As String
Public PreviewFile As String
Public BLOCKPATH  As String
Public LinkkiFormi As Boolean
Sub StartInterlocking()
  StartOK = False
  StartForm.Show
  If StartOK Then MainForm.Show
End Sub
Sub StartLink()
  LinkForm.Show
End Sub
