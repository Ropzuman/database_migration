Option Compare Database 'Use database order for string comparisons
Option Explicit
'DOCUMENTS-alaformien välistä tiedonsiirtoa varten
Public Revisioteksti As String
Public Revisionumero As String
Public UusiRevisio As Boolean
Public RevDefaults As String
Public CurRecord As String
Public CurTyyppi As String
Public Common As String
Public CommonOts As String
Public CommonType As Integer
Public DocStatus As String
'UUDEN täytetyn dokumentin tietojen muistamista varten
Public DefDName1 As String
Public DefDName2 As String
Public DefDName3 As String
Public DefArea As String
Public DefDept As String
Public DefCode As String
Public DefSize As String
Public DefCDocNo As String
Public DefClass As String
Public DefPages As String
Public DefPath As String
Public DefSort As String
Public DefRev As String
Public DefRevText As String
'---------------------------------------
'Edellisen kirjoitetun revision muistamista varten
Public MRevRev As String
Public MRevDrawn As String
Public MRevChecked As String
Public MRevApproved As String
Public MRevDescription As String
'---------------------------------------

'---------- [Hakemiston valintaa varten API-funktiot ja muuttujat] ------------
' API declarations for 64-bit compatibility
' Folder browser dialog functions
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As LongPtr)
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

'***************************************************************************
'* REMOVED: Unused VBA71.dll API declarations                              *
'* Date: November 8, 2025                                                  *
'* Reason: GetCurrentVbaProject, GetFuncID, and GetAddr were declared      *
'*         but never called anywhere in the codebase. These are advanced   *
'*         VBA project introspection functions not needed for this app.    *
'* Impact: ~120 bytes less compiled size, cleaner code                     *
'***************************************************************************

' BrowseInfo structure for folder selection dialog
Private Type BrowseInfo
    hOwner As LongPtr         ' Window handle of parent form
    pIDLRoot As LongPtr       ' Root folder PIDL (NULL for Desktop)
    pszDisplayName As LongPtr ' Pointer to display name buffer
    lpszTitle As LongPtr      ' Pointer to dialog title string
    ulFlags As Long           ' Dialog behavior flags (BIF_*)
    lpfn As LongPtr           ' Callback function pointer
    lParam As LongPtr         ' Application-defined parameter
    iImage As Long            ' Image index (output only)
End Type
Public CDialogPath As String  ' Default path for folder browser dialog
'---------- [Hakemiston valintaa varten API-funktiot Loppu] ------------

' Network username API
Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
Public Function IsLoaded(ByVal strFormName As String) As Integer
 ' Palauttaa arvon "Tosi", jos m��ritetty lomake on avoinna
 ' lomake- tai taulukkon�kym�ss�.
    Const conObjStateClosed = 0
    Const conDesignView = 0
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
End Function

Public Function IsTableLoaded(TableName As String) As Integer
 IsTableLoaded = SysCmd(acSysCmdGetObjectState, acTable, TableName)
End Function

Public Function NetworkUserName() As String
    Dim lngStringLength As LongPtr ' Changed to LongPtr
    Dim sString As String * 255
    lngStringLength = Len(sString)
    sString = String$(lngStringLength, 0)
    If wu_GetUserName(sString, lngStringLength) Then
        NetworkUserName = Left$(sString, lngStringLength - 1)
    Else
        NetworkUserName = "Unknown"
    End If
End Function
Public Function ValitseHakem(Handle As LongPtr, Optional StartPath As String) As String
 Dim lpIDList As LongPtr ' Changed to LongPtr
 Dim lpSelPath As Long
 Dim ThePath As String
 Dim Otsikko As String
 Dim tBrowseInfo As BrowseInfo
 Const LMEM_FIXED = &H0
 Const LMEM_ZEROINIT = &H40
 Const LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
 Const BIF_RETURNONLYFSDIRS = 1
 Const BIF_DONTGOBELOWDOMAIN = 2
 Const MAX_PATH = 260

 If IsMissing(StartPath) Or StartPath = "" Then
  CDialogPath = "K:\PROJECTS"
 Else
  If Right$(StartPath, 2) = ":\" Then
   CDialogPath = StartPath
  ElseIf Right$(StartPath, 1) = "\" Then
   CDialogPath = Left$(StartPath, Len(StartPath) - 1)
  Else
   CDialogPath = StartPath
  End If
 End If
    
    Otsikko = "Choose path:"
    With tBrowseInfo
      .hOwner = Handle
      .pIDLRoot = 0
      .lpszTitle = lstrcat(Otsikko, "")
      .lpfn = DummyFunc(AddressOf BrowseCallbackProc)
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        ThePath = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, ThePath
        ThePath = Left$(ThePath, InStr(ThePath, vbNullChar) - 1)
    Else
      ThePath = ""
    End If
    CoTaskMemFree lpIDList
    If ThePath <> "" Then
      If Right$(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
    End If
    ValitseHakem = ThePath
End Function
Public Function DummyFunc(ByVal param As LongPtr) As LongPtr ' Changed to LongPtr
  DummyFunc = param
End Function
Public Function BrowseCallbackProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr ' Changed to LongPtr
  Const BFFM_INITIALIZED = 1
  Const BFFM_SETSELECTION = &H466
  Dim retval As LongPtr ' Changed to LongPtr

  Select Case uMsg
  Case BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELECTION, ByVal 1&, ByVal StrPtr(CDialogPath))
  Case Else
  End Select
  BrowseCallbackProc = 0
End Function
