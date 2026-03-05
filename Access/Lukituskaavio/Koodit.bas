Option Compare Database
Option Explicit
' AutoCAD-objektit muutettu late binding -tyypiksi 64-bit yhteensopivuuden vuoksi
Public oACAD As Object    ' AcadApplication - muutettu late binding (64-bit)
Public oDOC As Object     ' AcadDocument - muutettu late binding (64-bit)
Public BlockPath As String
Sub KillLinks()
Dim T As DAO.TableDef
Dim LinkCount As Long

On Error GoTo ErrorHandler

LinkCount = 0

For Each T In CurrentDb.TableDefs
    If T.Connect <> "" And Left$(T.Name, 1) <> "~" Then
        CurrentDb.Execute ("DROP TABLE " & T.Name)
        LinkCount = LinkCount + 1
    End If
Next

Exit Sub

ErrorHandler:
  MsgBox "Error: " & Err.Description, vbCritical
End Sub
Public Function AvaaBlock()
Dim DWG As String
Dim Polku As String
Dim Handle As String
Dim i As Integer
Dim Doku As String
Dim OK As Boolean
Dim Entity As Object  ' AcadEntity - muutettu late binding (64-bit)
Dim MinPoint As Variant
Dim MaxPoint As Variant
Dim TAULUKKO As String
On Error GoTo ErrorHandler

TAULUKKO = UCase$(Application.CurrentObjectName)
  If TAULUKKO = "DATA9" Then
    Polku = Screen.ActiveDatasheet("PATH").VALUE
    DWG = IIf(IsNull(Screen.ActiveDatasheet("DWG").VALUE), "", Screen.ActiveDatasheet("DWG").VALUE)
    Handle = IIf(IsNull(Screen.ActiveDatasheet("HANDLE").VALUE), "", Screen.ActiveDatasheet("HANDLE").VALUE)
    
    If Polku = "" Then
      MsgBox "Ei ole tietoa missä kuvassa kohde on !", vbCritical, "Etsi kohde"
      Exit Function
    End If
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
    If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
      MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Etsi Kohde"
      Set oACAD = Nothing
      Exit Function
    End If
    On Error GoTo 0
    If InStr(DWG, ".dwg") Then
      Doku = DWG
    Else
      Doku = LCase$(DWG) & ".dwg"
    End If
    OK = False
    For i = 0 To oACAD.Documents.Count - 1
      If LCase$(oACAD.Documents(i).Name) = Doku Then 'Sama kuva
        oACAD.Documents(i).Activate
        OK = True
        Exit For
      End If
    Next i
    If Not OK Then
      On Error Resume Next
      oACAD.Documents.Open Polku & Doku
      If Err = 0 Then
        OK = True
      Else
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
  Exit Function

ErrorHandler:
  ' Vapautetaan resurssit AutoCAD-virheen sattuessa
  Set Entity = Nothing
  Set oACAD = Nothing
  MsgBox "Virhe: " & Err.Description, vbCritical, "AvaaBlock"
End Function
