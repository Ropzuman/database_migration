Option Compare Database 'Use database order for string comparisons
Option Explicit
'DOCUMENTS-alaformien v�list� tiedosiirtoa varten
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
'UUDEN t�ytetyn dokumentin tietojen muistamista varten
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
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pvoid As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetCurrentVbaProject Lib "vba332.dll" Alias "EbGetExecutingProj" (hProject As Long) As Long
Private Declare Function GetFuncID Lib "vba332.dll" Alias "TipGetFunctionId" (ByVal hProject As Long, ByVal strFunctionName As String, ByRef strFunctionId As String) As Long
Private Declare Function GetAddr Lib "vba332.dll" Alias "TipGetLpfnOfFunctionId" (ByVal hProject As Long, ByVal strFunctionId As String, ByRef lpfn As Long) As Long
Private Type BrowseInfo
    hOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfn           As Long
    lParam         As Long
    iImage         As Long
End Type
Public CDialogPath As String
'---------- [Hakemiston valintaa varten API-funktiot Loppu] ------------
Declare Function wu_GetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
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
   Dim lngStringLength As Long
   Dim sString As String * 255
   lngStringLength = Len(sString)
   sString = String$(lngStringLength, 0)
   If wu_GetUserName(sString, lngStringLength) Then
       NetworkUserName = Left$(sString, lngStringLength - 1)
   Else
       NetworkUserName = "Unknown"
   End If
End Function
Public Function ValitseHakem(Handle As Long, Optional StartPath As String) As String
  Dim lpIDList As Long
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
    If Right(StartPath, 2) = ":\" Then
      CDialogPath = StartPath
    ElseIf Right(StartPath, 1) = "\" Then
      CDialogPath = Left(StartPath, Len(StartPath) - 1)
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
        ThePath = Left(ThePath, InStr(ThePath, vbNullChar) - 1)
    Else
      ThePath = ""
    End If
    CoTaskMemFree lpIDList
    If ThePath <> "" Then
      If Right(ThePath, 1) <> "\" Then ThePath = ThePath & "\"
    End If
    ValitseHakem = ThePath
End Function
Public Function DummyFunc(ByVal param As Long) As Long
  DummyFunc = param
End Function
Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
  Const BFFM_INITIALIZED = 1
  Const BFFM_SETSELECTION = &H466
  Dim retval As Long         'Return value

  Select Case uMsg
  Case BFFM_INITIALIZED
    retval = SendMessage(hWnd, BFFM_SETSELECTION, True, ByVal CDialogPath)
  Case Else
  End Select
  BrowseCallbackProc = 0
End Function
