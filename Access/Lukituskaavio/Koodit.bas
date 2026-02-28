Option Compare Database
Option Explicit
Public oACAD As AcadApplication    'AutoCad objekti
Public oDOC As AcadDocument
Public BlockPath As String
Sub KillLinks()
Dim T As DAO.TableDef
Dim LinkCount As Long
On Error GoTo ErrorHandler

Debug.Print "KillLinks: Starting - Dropping linked tables"
LinkCount = 0

For Each T In CurrentDb.TableDefs
    If T.Connect <> "" And Left$(T.Name, 1) <> "~" Then
        Debug.Print "  Dropping table: " & T.Name
        CurrentDb.Execute ("DROP TABLE " & T.Name)
        LinkCount = LinkCount + 1
    End If
Next

Debug.Print "KillLinks: COMPLETED - Tables dropped: " & LinkCount
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in KillLinks: " & Err.Number & " - " & Err.Description
  Debug.Print "    Current table: " & T.Name
  MsgBox "Error: " & Err.Description, vbCritical
End Sub
Public Function AvaaBlock()
Dim DWG As String
Dim Polku As String
Dim Handle As String
Dim i As Integer
Dim Doku As String
Dim OK As Boolean
Dim Entity As AcadEntity
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim TAULUKKO As String
On Error GoTo ErrorHandler

Debug.Print "AvaaBlock: Starting - Open block in AutoCAD"
TAULUKKO = UCase$(Application.CurrentObjectName)
Debug.Print "  Table: " & TAULUKKO
  If TAULUKKO = "DATA9" Then
    Polku = Screen.ActiveDatasheet("PATH").VALUE
    DWG = IIf(IsNull(Screen.ActiveDatasheet("DWG").VALUE), "", Screen.ActiveDatasheet("DWG").VALUE)
    Handle = IIf(IsNull(Screen.ActiveDatasheet("HANDLE").VALUE), "", Screen.ActiveDatasheet("HANDLE").VALUE)
    Debug.Print "  Path: " & Polku
    Debug.Print "  DWG: " & DWG
    Debug.Print "  Handle: " & Handle
    
    If Polku = "" Then
      Debug.Print "  ERROR: No path information"
      MsgBox "Ei ole tietoa missä kuvassa kohde on !", vbCritical, "Etsi kohde"
      Exit Function
    End If
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
    If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
      Debug.Print "  ERROR: AutoCAD not running"
      MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Etsi Kohde"
      Set oACAD = Nothing
      Exit Function
    End If
    Debug.Print "  Connected to AutoCAD"
    On Error GoTo 0
    If InStr(DWG, ".dwg") Then
      Doku = DWG
    Else
      Doku = LCase$(DWG) & ".dwg"
    End If
    OK = False
    For i = 0 To oACAD.Documents.Count - 1
      If LCase$(oACAD.Documents(i).Name) = Doku Then 'Sama kuva
        Debug.Print "  Document already open, activating: " & Doku
        oACAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    If Not OK Then
      On Error Resume Next
      Debug.Print "  Opening document: " & Polku & Doku
      oACAD.Documents.Open Polku & Doku
      If Err = 0 Then
        OK = True
        Debug.Print "  Document opened successfully"
      Else
        Debug.Print "  ERROR: Failed to open document: " & Err.Description
        MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Doku, vbCritical, "Etsi kohde"
        Err.Clear
      End If
      On Error GoTo 0
    End If
    If OK Then
      If Handle = "" Then
        MsgBox "Kohteen sijainti ei ole tiedossa, vain kuva avattiin.", vbCritical, "Etsi kohde"
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
      AppActivate "AutoCAD", True
      Set Entity = Nothing
      Set oACAD = Nothing
    End If
  End If
End Function
