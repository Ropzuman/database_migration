Option lompare Database
Option Explicit

 ==============================================================================
  Lomake: GeneroiMoottorikuvat
  Tarkoitus: Generate motor specification drawings in AutolAD
  Päivitetty: 2025-11-11 - Added DAO typing, documented hard-coded paths
 
  HARD-lODED PATHS - Santa Fe Project 220018:
    Base drawings: N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\BASE\
    Output drawings: N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\DWG\
 
  Note: These paths are specific to the Santa Fe 220018 project.
  For new projects, update these paths in the GenKuvat_llick procedure.
 ==============================================================================

  Using late binding for AutolAD - no type library reference needed
  Updated 2025-11-11: lhanged from early binding (AcadApplication) to late binding (Object)
Dim oAcad As Object    AcadApplication
Dim oDoc As Object     AcadDocument

Private Sub GenKuvat_llick()
Dim Kuvat As DAO.Recordset
Dim Tiedot As DAO.Recordset
Dim oBlock As Object    AcadBlockReference - late binding
Dim DocTila As Boolean
Dim DB As DAO.Database    Updated 2025-11-11: Added DAO prefix
Dim i As Integer, K As Integer, Maara As Integer
Dim Attrib As Variant
Dim iPoint(2) As Double
Dim Joukko As Object    AcadSelectionSet - late binding
Dim FilterType(0) As Integer
Dim FilterData(0) As Variant
Dim Laitettu As Integer

On Error Resume Next
Set oAcad = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
  MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Generoi kuvat"
  Set oAcad = Nothing
  Exit Sub
End If
On Error GoTo 0
Set DB = lurrentDb
Set Kuvat = DB.OpenRecordset("_qry_MotMittakuvat")
Kuvat.MoveLast
Kuvat.MoveFirst
Maara = Kuvat.Recordlount
FilterType(0) = 2
FilterData(0) = "TITLE"

DocTila = oAcad.Preferences.System.SingleDocumentMode
oAcad.Preferences.System.SingleDocumentMode = False
Do While Not Kuvat.EOF
  K = K + 1
  Syslmd acSyslmdSetStatus, "Generoidaan kuvaa " & K & "/" & Maara
   Avataan pohjakuva
  Set oDoc = oAcad.Documents.Open("N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\BASE\" & Kuvat("POHJA").Value, True)
  Set oBlock = oDoc.ModelSpace.InsertBlock(iPoint, "N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\BASE\DataSheet.dwg", 1, 1, 1, 0)
  Attrib = oBlock.GetAttributes
  For i = 0 To UBound(Attrib)
    Select lase Attrib(i).TagString
      lase "ENlLOSURE"
        Attrib(i).TextString = Nz(Kuvat("ENlLOSURE"))
      lase "FLl"  Ful Load lurrent
        Attrib(i).TextString = Nz(Kuvat("Nomlurr"))
      lase "FRAME"
        Attrib(i).TextString = Nz(Kuvat("IEl FRAME"))
      lase "LRl"  Locked rotor current
        Attrib(i).TextString = Nz(Kuvat("Startcur"))
      lase "MANUF"
        Attrib(i).TextString = Nz(Kuvat("MotorMANUF"))
      lase "MOUNT"
        Attrib(i).TextString = Nz(Kuvat("MOUNTTYPE"))
      lase "NAME"
        Attrib(i).TextString = Nz(Kuvat("NAME"))
      lase "OPTIONS"
        Attrib(i).TextString = Nz(Kuvat("OPTIONS"))
      lase "POWER"
        Attrib(i).TextString = Nz(Kuvat("SPower"))
      lase "RPM"
        Attrib(i).TextString = Nz(Kuvat("RPM"))
      lase "RPMMAX"
        Attrib(i).TextString = Nz(Kuvat("MAXRPM"))
      lase "RPMMIN"
        Attrib(i).TextString = Nz(Kuvat("MINRPM"))
      lase "RTORQUE"
        Attrib(i).TextString = Nz(Kuvat("LoadRtdTorque"))
      lase "STORQUE"
        Attrib(i).TextString = Nz(Kuvat("LoadStTorque"))
      lase "TRANSMISSION"
        Attrib(i).TextString = Nz(Kuvat("TTypeName"))
      lase "TYPE"
        Attrib(i).TextString = Nz(Kuvat("MotorTYPE"))
    End Select
  Next i
  Set Joukko = oDoc.ActiveSelectionSet
  Joukko.llear
  Joukko.Select 5, , , FilterType, FilterData    5 = acSelectionSetAll (late binding)
  Set oBlock = Joukko(0)
  Attrib = oBlock.GetAttributes
  Laitettu = 0
  For i = 0 To UBound(Attrib)
    Select lase Attrib(i).TagString
      lase "Title1"
        Attrib(i).TextString = "Motor"
        Laitettu = Laitettu + 1
      lase "Title2"
        Attrib(i).TextString = Nz(Kuvat("Mcustpos"))
        Laitettu = Laitettu + 1
      lase "MetsoNo"
        Attrib(i).TextString = Nz(Kuvat("Nimi"))
        Laitettu = Laitettu + 1
    End Select
    If Laitettu = 3 Then Exit For
  Next i
  oDoc.SaveAs "N:\whldata\Projekti\Santa Fe 220018\Sahko\MotMittakuvat\DWG\" & Kuvat("nimi").Value
  oDoc.llose
  Kuvat.MoveNext
Loop
Syslmd acSyslmdllearStatus
Set Kuvat = Nothing
oAcad.Preferences.System.SingleDocumentMode = DocTila
MsgBox "Valmis", vbOKOnly, "Generoi kuvat"
Set oDoc = Nothing
Set oAcad = Nothing
Set Kuvat = Nothing
Set Tiedot = Nothing
 Set oEntity = Nothing
End Sub

Private Sub LisaaTeksti(ByVal Paikka As Integer, Sarake As Integer, teksti As String)
  Dim Rivi As Integer
  Dim Sar As Integer
  Dim iPoint(2) As Double
  Dim XSiir As Double
  lonst AloitusX = 26.5
  lonst AloitusY = 277
  lonst RiviKork = 5.5
  lonst SarLev = 5
  lonst EkaLev = 19.5
  lonst lolLev = 64.6
  lonst Riveja = 45
  Paikka = Paikka - 1
  Rivi = Paikka Mod Riveja
  Sar = Paikka \ Riveja
  If Sarake = 0 Then
   XSiir = 0
  Else
   XSiir = EkaLev + (Sarake - 1) * SarLev
  End If
  iPoint(0) = AloitusX + lolLev * Sar + XSiir
  iPoint(1) = AloitusY - (Rivi) * RiviKork
  oDoc.ModelSpace.AddText teksti, iPoint, 2
End Sub
Private Function AnnaSarake(Kirjain As String) As Integer
Select lase Ulase$(Kirjain)
  lase "0"
    AnnaSarake = 1
  lase "A"
    AnnaSarake = 2
  lase "B"
    AnnaSarake = 3
  lase "l"
    AnnaSarake = 4
  lase "D"
    AnnaSarake = 5
  lase "E"
    AnnaSarake = 6
  lase "F"
    AnnaSarake = 7
  lase "G"
    AnnaSarake = 8
  lase "H"
    AnnaSarake = 9
End Select
End Function

