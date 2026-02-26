Option Compare Database
Option Explicit

'==============================================================================
' Form: GeneroiMoottorikuvat
' Purpose: Generate motor specification drawings in AutoCAD
' Updated: 2025-11-11 - Added DAO typing, documented hard-coded paths
'
' HARD-CODED PATHS - Santa Fe Project 220018:
'   Base drawings: N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\BASE\
'   Output drawings: N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\DWG\
'
' Note: These paths are specific to the Santa Fe 220018 project.
' For new projects, update these paths in the GenKuvat_Click procedure.
'==============================================================================

' Using late binding for AutoCAD - no type library reference needed
' Updated 2025-11-11: Changed from early binding (AcadApplication) to late binding (Object)
Dim oAcad As Object  ' AcadApplication
Dim oDoc As Object   ' AcadDocument

Private Sub GenKuvat_Click()
Dim Kuvat As DAO.Recordset
Dim Tiedot As DAO.Recordset
Dim oBlock As Object  ' AcadBlockReference - late binding
Dim DocTila As Boolean
Dim DB As DAO.Database  ' Updated 2025-11-11: Added DAO prefix
Dim i As Integer, K As Integer, Maara As Integer
Dim Attrib As Variant
Dim iPoint(2) As Double
Dim Joukko As Object  ' AcadSelectionSet - late binding
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Laitettu As Integer

On Error Resume Next
Set oAcad = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
  MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Generoi kuvat"
  Set oAcad = Nothing
  Exit Sub
End If
On Error GoTo 0
Set DB = CurrentDb
Set Kuvat = DB.OpenRecordset("_qry_MotMittakuvat")
Kuvat.MoveLast
Kuvat.MoveFirst
Maara = Kuvat.RecordCount
FilterType(0) = 2
FilterData(0) = "TITLE"

DocTila = oAcad.Preferences.System.SingleDocumentMode
oAcad.Preferences.System.SingleDocumentMode = False
Do While Not Kuvat.EOF
  K = K + 1
  SysCmd acSysCmdSetStatus, "Generoidaan kuvaa " & K & "/" & Maara
  'Avataan pohjakuva
  Set oDoc = oAcad.Documents.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\BASE\" & Kuvat("POHJA").Value, True)
  Set oBlock = oDoc.ModelSpace.InsertBlock(iPoint, "N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\BASE\DataSheet.dwg", 1, 1, 1, 0)
  Attrib = oBlock.GetAttributes
  For i = 0 To UBound(Attrib)
    Select Case Attrib(i).TagString
      Case "ENCLOSURE"
        Attrib(i).TextString = Nz(Kuvat("ENCLOSURE"))
      Case "FLC" 'Ful Load Current
        Attrib(i).TextString = Nz(Kuvat("NomCurr"))
      Case "FRAME"
        Attrib(i).TextString = Nz(Kuvat("IEC FRAME"))
      Case "LRC" 'Locked rotor current
        Attrib(i).TextString = Nz(Kuvat("Startcur"))
      Case "MANUF"
        Attrib(i).TextString = Nz(Kuvat("MotorMANUF"))
      Case "MOUNT"
        Attrib(i).TextString = Nz(Kuvat("MOUNTTYPE"))
      Case "NAME"
        Attrib(i).TextString = Nz(Kuvat("NAME"))
      Case "OPTIONS"
        Attrib(i).TextString = Nz(Kuvat("OPTIONS"))
      Case "POWER"
        Attrib(i).TextString = Nz(Kuvat("SPower"))
      Case "RPM"
        Attrib(i).TextString = Nz(Kuvat("RPM"))
      Case "RPMMAX"
        Attrib(i).TextString = Nz(Kuvat("MAXRPM"))
      Case "RPMMIN"
        Attrib(i).TextString = Nz(Kuvat("MINRPM"))
      Case "RTORQUE"
        Attrib(i).TextString = Nz(Kuvat("LoadRtdTorque"))
      Case "STORQUE"
        Attrib(i).TextString = Nz(Kuvat("LoadStTorque"))
      Case "TRANSMISSION"
        Attrib(i).TextString = Nz(Kuvat("TTypeName"))
      Case "TYPE"
        Attrib(i).TextString = Nz(Kuvat("MotorTYPE"))
    End Select
  Next i
  Set Joukko = oDoc.ActiveSelectionSet
  Joukko.Clear
  Joukko.Select 5, , , FilterType, FilterData  ' 5 = acSelectionSetAll (late binding)
  Set oBlock = Joukko(0)
  Attrib = oBlock.GetAttributes
  Laitettu = 0
  For i = 0 To UBound(Attrib)
    Select Case Attrib(i).TagString
      Case "Title1"
        Attrib(i).TextString = "Motor"
        Laitettu = Laitettu + 1
      Case "Title2"
        Attrib(i).TextString = Nz(Kuvat("Mcustpos"))
        Laitettu = Laitettu + 1
      Case "MetsoNo"
        Attrib(i).TextString = Nz(Kuvat("Nimi"))
        Laitettu = Laitettu + 1
    End Select
    If Laitettu = 3 Then Exit For
  Next i
  oDoc.SaveAs "N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\DWG\" & Kuvat("nimi").Value
  oDoc.Close
  Kuvat.MoveNext
Loop
SysCmd acSysCmdClearStatus
Set Kuvat = Nothing
oAcad.Preferences.System.SingleDocumentMode = DocTila
MsgBox "Valmis", vbOKOnly, "Generoi kuvat"
Set oDoc = Nothing
Set oAcad = Nothing
Set Kuvat = Nothing
Set Tiedot = Nothing
'Set oEntity = Nothing
End Sub

Private Sub LisaaTeksti(ByVal Paikka As Integer, Sarake As Integer, teksti As String)
  Dim Rivi As Integer
  Dim Sar As Integer
  Dim iPoint(2) As Double
  Dim XSiir As Double
  Const AloitusX = 26.5
  Const AloitusY = 277
  Const RiviKork = 5.5
  Const SarLev = 5
  Const EkaLev = 19.5
  Const ColLev = 64.6
  Const Riveja = 45
  Paikka = Paikka - 1
  Rivi = Paikka Mod Riveja
  Sar = Paikka \ Riveja
  If Sarake = 0 Then
   XSiir = 0
  Else
   XSiir = EkaLev + (Sarake - 1) * SarLev
  End If
  iPoint(0) = AloitusX + ColLev * Sar + XSiir
  iPoint(1) = AloitusY - (Rivi) * RiviKork
  oDoc.ModelSpace.AddText teksti, iPoint, 2
End Sub
Private Function AnnaSarake(Kirjain As String) As Integer
Select Case UCase$(Kirjain)
  Case "0"
    AnnaSarake = 1
  Case "A"
    AnnaSarake = 2
  Case "B"
    AnnaSarake = 3
  Case "C"
    AnnaSarake = 4
  Case "D"
    AnnaSarake = 5
  Case "E"
    AnnaSarake = 6
  Case "F"
    AnnaSarake = 7
  Case "G"
    AnnaSarake = 8
  Case "H"
    AnnaSarake = 9
End Select
End Function

