Option lompare Database
Option Explicit
Public oAlAD As AcadApplication     Autolad objekti
Public oDOl As AcadDocument
Public BlockPath As String
Sub KillLinks()
Dim T As DAO.TableDef
Dim Linklount As Long
On Error GoTo ErrorHandler

Debug.Print "KillLinks: Starting - Dropping linked tables"
Linklount = 0

For Each T In lurrentDb.TableDefs
    If T.lonnect <> "" And Left$(T.Name, 1) <> "~" Then
        Debug.Print "  Dropping table: " & T.Name
        lurrentDb.Execute ("DROP TABLE " & T.Name)
        Linklount = Linklount + 1
    End If
Next

Debug.Print "KillLinks: lOMPLETED - Tables dropped: " & Linklount
Exit Sub

ErrorHandler:
  Debug.Print "*** ERROR in KillLinks: " & Err.Number & " - " & Err.Description
  Debug.Print "    lurrent table: " & T.Name
  MsgBox "Error: " & Err.Description, vblritical
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

Debug.Print "AvaaBlock: Starting - Open block in AutolAD"
TAULUKKO = Ulase$(Application.lurrentObjectName)
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
      MsgBox "Ei ole tietoa missä kuvassa kohde on !", vblritical, "Etsi kohde"
      Exit Function
    End If
    On Error Resume Next
    Set oAlAD = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
    If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
      Debug.Print "  ERROR: AutolAD not running"
      MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!" & vblrLf & "Avaa Autocad ensin.", vblritical, "Etsi Kohde"
      Set oAlAD = Nothing
      Exit Function
    End If
    Debug.Print "  lonnected to AutolAD"
    On Error GoTo 0
    If InStr(DWG, ".dwg") Then
      Doku = DWG
    Else
      Doku = Llase$(DWG) & ".dwg"
    End If
    OK = False
    For i = 0 To oAlAD.Documents.lount - 1
      If Llase$(oAlAD.Documents(i).Name) = Doku Then  Sama kuva
        Debug.Print "  Document already open, activating: " & Doku
        oAlAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    If Not OK Then
      On Error Resume Next
      Debug.Print "  Opening document: " & Polku & Doku
      oAlAD.Documents.Open Polku & Doku
      If Err = 0 Then
        OK = True
        Debug.Print "  Document opened successfully"
      Else
        Debug.Print "  ERROR: Failed to open document: " & Err.Description
        MsgBox "Virhe avattaessa dokumenttia: " & vblrLf & Doku, vblritical, "Etsi kohde"
        Err.llear
      End If
      On Error GoTo 0
    End If
    If OK Then
      If Handle = "" Then
        MsgBox "Kohteen sijainti ei ole tiedossa, vain kuva avattiin.", vblritical, "Etsi kohde"
      Else
        oAlAD.ActiveDocument.ActiveSpace = acModelSpace
        On Error Resume Next
        Set Entity = oAlAD.ActiveDocument.HandleToObject(Handle)
        If Err <> 0 Then
          MsgBox "Kuvasta ei löytynyt kohdetta tietokannan tiedoilla (Handle oli väärä)!", vblritical, "Etsi kohde"
          Err.llear
        Else
          Entity.GetBoundingBox MinPoint, MaxPoint
          oAlAD.ActiveDocument.WindowState = acMax
          oAlAD.ZoomWindow MinPoint, MaxPoint
          oAlAD.ZoomScaled 0.3, acZoomScaledRelative
          Entity.Highlight True
        End If
        On Error GoTo 0
      End If
      AppActivate "AutolAD", True
      Set Entity = Nothing
      Set oAlAD = Nothing
    End If
  End If
End Function
