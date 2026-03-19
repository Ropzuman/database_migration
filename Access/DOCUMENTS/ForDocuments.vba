Option Compare Database
Option Explicit
'==========================================================================
' MODUULI  : ForDocuments
' SOVELLUS : DOCUMENTS — Jaetut muuttujat ja apufunktiot
' KUVAUS   : Sisältää lomakkeiden välisen tiedonsiirron julkiset muuttujat,
'            Shell-kansion valintadialogin API-kutsut (BrowseForFolder),
'            verkkonimen haun (NetworkUserName) sekä apufunktiot
'            hakemistopoluille (ValitseHakem, BrowseCallbackProc).
'            Kaikki API-julistukset suojattu #If VBA7 -ehdolla:
'            64-bittinen haara käyttää PtrSafe + LongPtr,
'            32-bittinen haara käyttää tavallista Declare + Long.
'            nSize GetUserNameA:ssa on ByRef Long (LPDWORD) molemmissa haaroissa.
' PÄIVITETTY: 2026-03-03
'==========================================================================

' DOCUMENTS-alaformien välinen tiedonsiirto
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
' API declarations — guarded for both 32-bit (VBA6) and 64-bit (VBA7/PtrSafe)
#If VBA7 Then
    ' 64-bit: pointer-sized handles and return values use LongPtr
    Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As LongPtr
    Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As LongPtr, ByVal lpBuffer As String) As Long
    Private Declare PtrSafe Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As LongPtr
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As LongPtr)
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, lParam As Any) As LongPtr
    ' nSize is LPDWORD (pointer to a 32-bit DWORD) — ByRef Long is correct on both bitnesses
    Private Declare PtrSafe Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#Else
    ' 32-bit: no PtrSafe, pointer-sized values are plain Long
    Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
    Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
    Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
    Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As Long)
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
#End If

'***************************************************************************
'* REMOVED: Unused VBA71.dll API declarations                              *
'* Date: November 8, 2025                                                  *
'* Reason: GetCurrentVbaProject, GetFuncID, and GetAddr were declared      *
'*         but never called anywhere in the codebase. These are advanced   *
'*         VBA project introspection functions not needed for this app.    *
'* Impact: ~120 bytes less compiled size, cleaner code                     *
'***************************************************************************

' BrowseInfo structure — pointer fields must match the process bitness
#If VBA7 Then
Private Type BrowseInfo
    hOwner         As LongPtr  ' Window handle of parent form
    pIDLRoot       As LongPtr  ' Root folder PIDL (NULL for Desktop)
    pszDisplayName As LongPtr  ' Pointer to display name buffer
    lpszTitle      As LongPtr  ' Pointer to dialog title string
    ulFlags        As Long     ' Dialog behavior flags (BIF_*)
    lpfn           As LongPtr  ' Callback function pointer
    lParam         As LongPtr  ' Application-defined parameter
    iImage         As Long     ' Image index (output only)
End Type
#Else
Private Type BrowseInfo
    hOwner         As Long     ' Window handle of parent form
    pIDLRoot       As Long     ' Root folder PIDL (NULL for Desktop)
    pszDisplayName As Long     ' Pointer to display name buffer
    lpszTitle      As Long     ' Pointer to dialog title string
    ulFlags        As Long     ' Dialog behavior flags (BIF_*)
    lpfn           As Long     ' Callback function pointer
    lParam         As Long     ' Application-defined parameter
    iImage         As Long     ' Image index (output only)
End Type
#End If

Public CDialogPath As String  ' Default path for folder browser dialog
'---------- [Hakemiston valintaa varten API-funktiot Loppu] ------------
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
    Dim lngStringLength As Long  ' LPDWORD-yhteensopiva: 32-bittinen arvo, EI LongPtr
    Dim sString As String * 255
    lngStringLength = Len(sString)
    sString = String$(lngStringLength, 0)
    If wu_GetUserName(sString, lngStringLength) Then
        NetworkUserName = Left$(sString, lngStringLength - 1)
    Else
        NetworkUserName = "Unknown"
    End If
End Function
' ValitseHakem: Handle parameter is pointer-sized — LongPtr on 64-bit, Long on 32-bit.
' The entire function body is duplicated in each branch because #If cannot split a
' Function definition across its boundary — the header and End Function must be in the same block.
#If VBA7 Then
Public Function ValitseHakem(Handle As LongPtr, Optional ByVal StartPath As String = "") As String
  Dim dlg As Object
  Dim ThePath As String

  On Error GoTo ErrorHandler

  If StartPath = "" Then
    CDialogPath = "K:\PROJECTS"
  Else
    CDialogPath = CStr(StartPath)
  End If

  Set dlg = Application.FileDialog(4)
  With dlg
    .Title = "Choose path:"
    .AllowMultiSelect = False
    If CDialogPath <> "" Then .InitialFileName = CDialogPath
    If .Show <> -1 Then
      ValitseHakem = ""
      GoTo Cleanup
    End If
    ThePath = CStr(.SelectedItems(1))
  End With

  If ThePath <> "" Then
    If Right$(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
  End If
  ValitseHakem = ThePath

Cleanup:
  Set dlg = Nothing
  Exit Function

ErrorHandler:
  ValitseHakem = ""
  Resume Cleanup
End Function
#Else
Public Function ValitseHakem(Handle As Long, Optional ByVal StartPath As String = "") As String
  Dim dlg As Object
  Dim ThePath As String

  On Error GoTo ErrorHandler

  If StartPath = "" Then
    CDialogPath = "K:\PROJECTS"
  Else
    CDialogPath = CStr(StartPath)
  End If

  Set dlg = Application.FileDialog(4)
  With dlg
    .Title = "Choose path:"
    .AllowMultiSelect = False
    If CDialogPath <> "" Then .InitialFileName = CDialogPath
    If .Show <> -1 Then
      ValitseHakem = ""
      GoTo Cleanup
    End If
    ThePath = CStr(.SelectedItems(1))
  End With

  If ThePath <> "" Then
    If Right$(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
  End If
  ValitseHakem = ThePath

Cleanup:
  Set dlg = Nothing
  Exit Function

ErrorHandler:
  ValitseHakem = ""
  Resume Cleanup
End Function
#End If
#If VBA7 Then
Public Function DummyFunc(ByVal param As LongPtr) As LongPtr
  DummyFunc = param
End Function
Public Function BrowseCallbackProc(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal lParam As LongPtr, ByVal lpData As LongPtr) As LongPtr
  Const BFFM_INITIALIZED = 1
  Const BFFM_SETSELECTION = &H466
  Dim retval As LongPtr

  Select Case uMsg
  Case BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELECTION, ByVal 1&, ByVal StrPtr(CDialogPath))
  Case Else
  End Select
  BrowseCallbackProc = 0
End Function
#Else
Public Function DummyFunc(ByVal param As Long) As Long
  DummyFunc = param
End Function
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
  Const BFFM_INITIALIZED = 1
  Const BFFM_SETSELECTION = &H466
  Dim retval As Long

  Select Case uMsg
  Case BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELECTION, ByVal 1&, ByVal CDialogPath)
  Case Else
  End Select
  BrowseCallbackProc = 0
End Function
#End If

Public Sub AsetaCommonArvot(ByVal Teksti As String, ByVal Otsikko As String, ByVal Tyyppi As Integer)
    ' Keskitetään yhteisen tekstidialogin syöttöarvot yhteen paikkaan.
    Common = Teksti
    CommonOts = Otsikko
    CommonType = Tyyppi
End Sub

Public Function HaeCommonTeksti() As String
    HaeCommonTeksti = Common
End Function

Public Sub AsetaCommonTeksti(ByVal Teksti As String)
    Common = Teksti
End Sub

Public Function HaeRevisioTekstiJaettu() As String
    HaeRevisioTekstiJaettu = Revisioteksti
End Function

Public Sub AsetaRevisioTekstiJaettu(ByVal Teksti As String)
    Revisioteksti = Teksti
End Sub
