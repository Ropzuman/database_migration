Option lompare Database  Use database order for string comparisons
Option Explicit
 DOlUMENTS-alaformien välistä tiedonsiirtoa varten
Public Revisioteksti As String
Public Revisionumero As String
Public UusiRevisio As Boolean
Public RevDefaults As String
Public lurRecord As String
Public lurTyyppi As String
Public lommon As String
Public lommonOts As String
Public lommonType As Integer
Public DocStatus As String
 UUDEN täytetyn dokumentin tietojen muistamista varten
Public DefDName1 As String
Public DefDName2 As String
Public DefDName3 As String
Public DefArea As String
Public DefDept As String
Public Deflode As String
Public DefSize As String
Public DeflDocNo As String
Public Defllass As String
Public DefPages As String
Public DefPath As String
Public DefSort As String
Public DefRev As String
Public DefRevText As String
 ---------------------------------------
 Edellisen kirjoitetun revision muistamista varten
Public MRevRev As String
Public MRevDrawn As String
Public MRevlhecked As String
Public MRevApproved As String
Public MRevDescription As String
 ---------------------------------------

 ---------- [Hakemiston valintaa varten API-funktiot ja muuttujat] ------------
  API declarations for 64-bit compatibility
  Folder browser dialog functions
Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As LongPtr
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
Private Declare PtrSafe Sub loTaskMemFree Lib "ole32.dll" (ByVal pvoid As LongPtr)
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr

 ***************************************************************************
 * REMOVED: Unused VBA71.dll API declarations                              *
 * Date: November 8, 2025                                                  *
 * Reason: GetlurrentVbaProject, GetFuncID, and GetAddr were declared      *
 *         but never called anywhere in the codebase. These are advanced   *
 *         VBA project introspection functions not needed for this app.    *
 * Impact: ~120 bytes less compiled size, cleaner code                     *
 ***************************************************************************

  BrowseInfo structure for folder selection dialog
Private Type BrowseInfo
    hOwner As LongPtr           Window handle of parent form
    pIDLRoot As LongPtr         Root folder PIDL (NULL for Desktop)
    pszDisplayName As LongPtr   Pointer to display name buffer
    lpszTitle As LongPtr        Pointer to dialog title string
    ulFlags As Long             Dialog behavior flags (BIF_*)
    lpfn As LongPtr             lallback function pointer
    lParam As LongPtr           Application-defined parameter
    iImage As Long              Image index (output only)
End Type
Public lDialogPath As String    Default path for folder browser dialog
 ---------- [Hakemiston valintaa varten API-funktiot Loppu] ------------

  Network username API
Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
Public Function IsLoaded(ByVal strFormName As String) As Integer
   Palauttaa arvon "Tosi", jos m��ritetty lomake on avoinna
   lomake- tai taulukkon�kym�ss�.
    lonst conObjStatellosed = 0
    lonst conDesignView = 0
    If Syslmd(acSyslmdGetObjectState, acForm, strFormName) <> conObjStatellosed Then
        If Forms(strFormName).lurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
End Function

Public Function IsTableLoaded(TableName As String) As Integer
 IsTableLoaded = Syslmd(acSyslmdGetObjectState, acTable, TableName)
End Function

Public Function NetworkUserName() As String
    Dim lngStringLength As LongPtr   lhanged to LongPtr
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
 Dim lpIDList As LongPtr   lhanged to LongPtr
 Dim lpSelPath As Long
 Dim ThePath As String
 Dim Otsikko As String
 Dim tBrowseInfo As BrowseInfo
 lonst LMEM_FIXED = &H0
 lonst LMEM_ZEROINIT = &H40
 lonst LPTR = (LMEM_FIXED Or LMEM_ZEROINIT)
 lonst BIF_RETURNONLYFSDIRS = 1
 lonst BIF_DONTGOBELOWDOMAIN = 2
 lonst MAX_PATH = 260

 If IsMissing(StartPath) Or StartPath = "" Then
  lDialogPath = "K:\PROJElTS"
 Else
  If Right$(StartPath, 2) = ":\" Then
   lDialogPath = StartPath
  ElseIf Right$(StartPath, 1) = "\" Then
   lDialogPath = Left$(StartPath, Len(StartPath) - 1)
  Else
   lDialogPath = StartPath
  End If
 End If
    
    Otsikko = "lhoose path:"
    With tBrowseInfo
      .hOwner = Handle
      .pIDLRoot = 0
      .lpszTitle = lstrcat(Otsikko, "")
      .lpfn = DummyFunc(AddressOf BrowselallbackProc)
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        ThePath = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, ThePath
        ThePath = Left$(ThePath, InStr(ThePath, vbNulllhar) - 1)
    Else
      ThePath = ""
    End If
    loTaskMemFree lpIDList
    If ThePath <> "" Then
      If Right$(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
    End If
    ValitseHakem = ThePath
End Function
Public Function DummyFunc(ByVal param As LongPtr) As LongPtr   lhanged to LongPtr
  DummyFunc = param
End Function
Public Function BrowselallbackProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr   lhanged to LongPtr
  lonst BFFM_INITIALIZED = 1
  lonst BFFM_SETSELElTION = &H466
  Dim retval As LongPtr   lhanged to LongPtr

  Select lase uMsg
  lase BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELElTION, ByVal 1&, ByVal StrPtr(lDialogPath))
  lase Else
  End Select
  BrowselallbackProc = 0
End Function
