Attribute VB_Name = "Koodit"
Option Compare Database
Option Explicit
Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Sub AvaaBlock()
Dim DWG As String
Dim Polku As String
Dim Handle As String
Dim Info As String
Dim i As Integer
Dim Doku As String
Dim Tyyppi As Integer
Dim Taulu As Recordset
Dim Alue As String
Dim Linja As String


Polku = Left(CurrentDb.Name, Len(CurrentDb.Name) - Len(Dir(CurrentDb.Name))) & "..\..\R\FlowSheets\"

If UCase(Application.CurrentObjectName) = "MANUALVALVES" Or UCase(Application.CurrentObjectName) = "PIPELINES" Then
  If UCase(Application.CurrentObjectName) = "MANUALVALVES" Then Tyyppi = 1
  If Tyyppi = 1 Then
    Set Taulu = CurrentDb.OpenRecordset("SELECT * FROM MANVALVEDATA WHERE AREACODE = '" & Screen.ActiveDatasheet("Area").Value & "' AND VAL_NO = '" & Screen.ActiveDatasheet("ValveNo").Value & "'")
    DWG = LCase(Taulu.Fields("ImpFileID"))
    Handle = Taulu.Fields("Handles")
    Info = Taulu.Fields("AREACODE") & "-" & Taulu.Fields("VAL_NO")
  Else
    Alue = Nz(Screen.ActiveDatasheet("Area").Value)
    Linja = Nz(Screen.ActiveDatasheet("LineNo").Value)
    DoCmd.OpenForm "frmOpenPIPELINE", acNormal, , , , , Alue & "," & Linja
    Exit Sub
  End If
  If Polku = "" Then
    MsgBox "Ei ole tietoa missä kuvassa kohde " & Info & " on !", vbCritical, "Etsi kohde"
    Exit Sub
  End If
  AvaaKuvasta Polku, DWG, Handle, Info
ElseIf UCase(Application.CurrentObjectName) = "PIPELINEDATA" Or UCase(Application.CurrentObjectName) = "MANVALVEDATA" Then
    Polku = Screen.ActiveDatasheet("PATH").Value
    If Right(Polku, 1) <> "\" Then Polku = Polku & "\"
    
    DWG = LCase(Screen.ActiveDatasheet("ImpFileID").Value)
    Handle = Screen.ActiveDatasheet("Handles").Value
    If UCase(Application.CurrentObjectName) = "PIPELINEDATA" Then
      Info = Screen.ActiveDatasheet("DEP").Value & "-" & Screen.ActiveDatasheet("LINENO").Value
    Else
      Info = Screen.ActiveDatasheet("AREACODE").Value & "-" & Screen.ActiveDatasheet("VALVE_NO").Value
    End If
    AvaaKuvasta Polku, DWG, Handle, Info
Else
  MsgBox "MANUALVALVES tai PIPELINES taulukon tulee avaoinna näytöllä!", vbCritical, "Etsi kohde"
End If
End Sub
Sub AvaaKuvasta(Polku As String, Nimi As String, Handle As String, Info As String)
Dim oACAD As AcadApplication
Dim Entity As AcadEntity
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim i As Integer
Dim OK As Boolean
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Etsi Kohde"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
  OK = False
  For i = 0 To oACAD.Documents.Count - 1
    If LCase(oACAD.Documents(i).Name) = Nimi Then 'Sama kuva
      oACAD.Documents(i).Activate
      OK = True
      Exit For
    End If
  Next i
  If Not OK Then
    On Error Resume Next
    oACAD.Documents.Open Polku & Nimi
    If Err = 0 Then
      OK = True
    Else
      MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Nimi, vbCritical, "Etsi kohde"
      Err.Clear
    End If
    On Error GoTo 0
  End If
  If OK Then
    If Handle = "" Then
      MsgBox "Kohteen  " & Info & " sijainti ei ole tiedossa, vain kuva avattiin.", vbCritical, "Etsi kohde"
    Else
      oACAD.ActiveDocument.ActiveSpace = acModelSpace
      On Error Resume Next
      Set Entity = oACAD.ActiveDocument.HandleToObject(Handle)
      If Err <> 0 Then
        MsgBox "Kuvasta ei löytynyt kohdetta tietokannan tiedoilla (Handle oli väärä)!", vbCritical, "Etsi kohde"
        Err.Clear
      Else
        Entity.GetBoundingBox MinPoint, MaxPoint
        oACAD.ActiveDocument.WindowState = acMax
        oACAD.ZoomWindow MinPoint, MaxPoint
        oACAD.ZoomScaled 0.3, acZoomScaledRelative
        Entity.Highlight True
      End If
      On Error GoTo 0
    End If
    AppActivate oACAD.Caption, True
    Set Entity = Nothing
    Set oACAD = Nothing
  End If
End Sub
Public Function NetworkUserName() As String
   Dim lngStringLength As Long
   Dim sString As String * 255
   lngStringLength = Len(sString)
   sString = String$(lngStringLength, 0)
'   If wu_GetUserName(sString, lngStringLength) Then
'       NetworkUserName = Left$(sString, lngStringLength - 1)
'   Else
'       NetworkUserName = "Unknown"
'   End If
End Function
Function SetStartup()
    Dim DB As Database
    Dim Taulu As Recordset
    Dim NWUserName As String
    Dim CName As String
    Dim BuffSize As Long
    Dim NBuffer As String
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Unknown"
    End If
       
    Set DB = CurrentDb
    Set Taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With Taulu
        .AddNew
        .Fields(0) = NWUserName     'Users Name In Network
        .Fields(1) = CurrentUser()  'Users Name In This Database
        .Fields(2) = CName          'Users Computer Name
        .Fields(3) = Now            'Time At the Moment
        .Update
    End With
    Set DB = Nothing
    Set Taulu = Nothing
End Function

Function POIMI(Tieto As Variant, osa As Integer) As Variant
Dim Osat As Variant
  If IsNull(Tieto) Or Tieto = "" Then
     POIMI = Null
  Else
    Osat = Split(Tieto, "-")
    POIMI = Osat(osa - 1)
  End If
End Function
