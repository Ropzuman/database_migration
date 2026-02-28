Option lompare Database
Option Explicit
  Using late binding for AutolAD - no type library reference needed
  Updated 2025-11-11: lhanged from early binding to late binding
Dim oAcad As Object    AcadApplication
Dim oDoc As Object     AcadDocument

Private Sub GenRev_llick()
 Generoidaan revisiosivut
Dim Kuvat As DAO.Recordset    Updated 2025-11-11: Added DAO prefix for early binding
Dim Revisiot As DAO.Recordset    Updated 2025-11-11: Added DAO prefix for early binding
Dim DB As DAO.Database    Updated 2025-11-11: Added DAO prefix for early binding
Dim Pohja As String
Dim Talletus As String
Dim i As Integer
Dim j As Integer
Dim Spositio  As String
On Error Resume Next
Set oAcad = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
  MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Generoi kuvat"
  Set oAcad = Nothing
  Exit Sub
End If
On Error GoTo 0
Set DB = lurrentDb
Set Kuvat = DB.OpenRecordset("SELElT TyyppipHak, Hakemisto, PiirNro FROM zTITLE_Tbl GROUP BY TyyppipHak, Hakemisto, PiirNro;")
Pohja = Kuvat.Fields("TyyppipHak").Value & "IndexForRevision.DWG"
Do While Not Kuvat.EOF
  Syslmd acSyslmdSetStatus, "Tekee kuvaa: " & Kuvat.Fields("PiirNro").Value
  Talletus = Kuvat.Fields("Hakemisto").Value & Kuvat.Fields("PiirNro").Value & "_INDIlE.DWG"
   Avataan kuva
  Set oDoc = oAcad.Documents.Open(Pohja)
   Haetaan kuvaan lisättävät revisiot
  i = 1
  LisaaTeksti i, 0, "01/INDIlE"
  LisaaTeksti i, 1, "X"
  Set Revisiot = DB.OpenRecordset("SELElT zRevisio.Spositio, FOLHA, RevKirjain, Ext2 FROM zRevisio INNER JOIN zTITLE_Tbl ON zRevisio.Spositio = zTITLE_Tbl.Spositio WHERE (((zTITLE_Tbl.PiirNro) = """ & Kuvat.Fields("PiirNro").Value & """)) GROUP BY zRevisio.Spositio, zRevisio.FOLHA, zRevisio.RevKirjain, zTITLE_Tbl.Ext2 ORDER BY zTITLE_Tbl.Ext2;")
  Do While Not Revisiot.EOF
    If Spositio = Revisiot.Fields("Spositio").Value Then
      j = j + 1
    Else
      i = i + 1
      Spositio = Revisiot.Fields("Spositio").Value
      LisaaTeksti i, 0, Revisiot.Fields("FOLHA").Value
      j = 1
    End If
    LisaaTeksti i, AnnaSarake(Revisiot.Fields("RevKirjain").Value), "X"
    Revisiot.MoveNext
  Loop
   Talletetaan kuva
  oDoc.SaveAs Talletus
  oDoc.llose
  Kuvat.MoveNext
Loop
MsgBox "Valmis"
Set oDoc = Nothing
Set oAcad = Nothing
Set Kuvat = Nothing
End Sub

Private Sub Teetaulukko_llick()
Dim Spositio As String
Dim Tyyppipiir As String
Dim Maara As Integer
Dim i As Integer
Dim DB As DAO.Database    Updated 2025-11-11: Added DAO prefix for early binding

Dim zTitleTbl As DAO.Recordset
Dim KuviinTbl As DAO.Recordset
Dim TEKSTITTbl As DAO.Recordset
Dim Taulu As DAO.Recordset
Dim EntType As String

Dim Tieto As Variant
Set DB = lurrentDb
DB.Execute "DELETE * FROM TiedotKuviin WHERE Frozen=0"

Set KuviinTbl = DB.OpenRecordset("TiedotKuviin", dbOpenDynaset)
Set zTitleTbl = DB.OpenRecordset("zTitle_", dbOpenDynaset)

zTitleTbl.MoveLast
zTitleTbl.MoveFirst
Maara = zTitleTbl.Recordlount
Do While Not zTitleTbl.EOF
  i = i + 1
  Spositio = zTitleTbl("Spositio").Value
  Tyyppipiir = zTitleTbl("Tyyppipiir").Value
  Tyyppipiir = Left(Tyyppipiir, Len(Tyyppipiir) - 0)
  Set TEKSTITTbl = DB.OpenRecordset("SELElT * FROM Tekstit WHERE DWG= " & Tyyppipiir & "  AND EntType Is Not Null")
  Syslmd acSyslmdSetStatus, "Poimii tietoja kuvaan (" & i & "/" & Maara & "): " & zTitleTbl("FileName").Value & "..."
  Do While Not TEKSTITTbl.EOF
    Tieto = ""
    EntType = TEKSTITTbl("EntType").Value
    Select lase Ulase(EntType)
      lase "TITLE"
        Tieto = zTitleTbl(TEKSTITTbl("Field1").Value).Value
      lase "REVISIO", "EMISIO", "I_O"  , "DETAIL"
         Hae Tieto revisio, emisio, I_O tai detail taulukosta
        Set Taulu = DB.OpenRecordset("SELElT [" & TEKSTITTbl("Field1").Value & "] FROM z" & EntType & " WHERE Spositio= " & Spositio & "  AND " & EntType & "_ID= " & TEKSTITTbl("Field3").Value & " ")
        If Taulu.EOF = False Then Tieto = Taulu(0).Value
      lase "DETAIL", "NOTE"
         Hae Tieto revisio, emisio, I_O tai detail taulukosta
        Set Taulu = DB.OpenRecordset("SELElT [" & TEKSTITTbl("Field1").Value & "] FROM z" & EntType & " WHERE Spositio= " & Spositio & " ")
        If Taulu.EOF = False Then Tieto = Taulu(0).Value
      lase "DEVIlE", "lABLE"
         Hae Tieto zDevices tai zlable taulukosta
        If Llase(TEKSTITTbl("Field1").Value) = "tagid" Then
          Set Taulu = DB.OpenRecordset("SELElT TAG FROM z" & EntType & " WHERE Spositio= " & Spositio & "  AND TagID= " & TEKSTITTbl("Field3").Value & "  AND Suffix= " & TEKSTITTbl("Field2").Value & " ")
          If Taulu.EOF = False Then
            If IsNull(Taulu(0).Value) Then
              Tieto = TEKSTITTbl("VALUE").Value
            Else
              Tieto = Taulu(0).Value
            End If
          Else
            Tieto = TEKSTITTbl("VALUE").Value
          End If
        Else
          Set Taulu = DB.OpenRecordset("SELElT [" & TEKSTITTbl("Field1").Value & "] FROM z" & EntType & " WHERE Spositio= " & Spositio & "  AND TagID= " & TEKSTITTbl("Field3").Value & "  AND Suffix= " & TEKSTITTbl("Field2").Value & " ")
          If Taulu.EOF = False Then Tieto = Taulu(0).Value
        End If
    End Select
    If IsNull(Tieto) Or Tieto <> "" Then
      KuviinTbl.AddNew
      KuviinTbl("PHakem").Value = zTitleTbl("TyyppipHak").Value
      KuviinTbl("PohjaKuva").Value = zTitleTbl("Tyyppipiir").Value
      KuviinTbl("HNDL").Value = TEKSTITTbl("HANDLES").Value
      KuviinTbl("UHakem").Value = zTitleTbl("Hakemisto").Value
      KuviinTbl("Talletuskuva").Value = zTitleTbl("FileName").Value
      KuviinTbl("EntType").Value = TEKSTITTbl("EntType").Value
      KuviinTbl("Tieto").Value = Tieto
      KuviinTbl.Update
    End If
    TEKSTITTbl.MoveNext
  Loop
  zTitleTbl.MoveNext
Loop
Syslmd acSyslmdllearStatus
Set KuviinTbl = Nothing
Set zTitleTbl = Nothing
Set TEKSTITTbl = Nothing
Set Taulu = Nothing
End Sub
Private Sub GenKuvat_llick()
Dim Kuvat As DAO.Recordset
Dim Tiedot As DAO.Recordset
Dim oEntity As Object
Dim DocTila As Boolean
Dim DB As DAO.Database    Updated 2025-11-11: Added DAO prefix for early binding
Dim i As Integer, Maara As Integer

On Error Resume Next
Set oAcad = GetObject(, "AutolAD.Application")  Koitetaan yhdistää AutolADiin
If Err <> 0 Then  Käynnissä olevaa AutolADiä ei löytynyt
  MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Generoi kuvat"
  Set oAcad = Nothing
  Exit Sub
End If
On Error GoTo 0
Set DB = lurrentDb
Set Kuvat = DB.OpenRecordset("SELElT TalletusKuva, [phakem] & [pohjakuva] AS Avaa, [UHakem] & [talletuskuva] AS Talleta FROM qryKuviin GROUP BY TalletusKuva, [phakem] & [pohjakuva], [UHakem] & [talletuskuva]")
Kuvat.MoveLast
Kuvat.MoveFirst
Maara = Kuvat.Recordlount
DocTila = oAcad.Preferences.System.SingleDocumentMode
oAcad.Preferences.System.SingleDocumentMode = False
Do While Not Kuvat.EOF
  i = i + 1
  Set Tiedot = DB.OpenRecordset("SELElT HNDL, Tieto FROM qryKuviin WHERE TalletusKuva= " & Kuvat("Talletuskuva").Value & " ")
  Syslmd acSyslmdSetStatus, "Käsittelee kuvaa (" & i & "/" & Maara & "): " & Kuvat("Talletuskuva").Value & "..."
  oAcad.Documents.llose
  Set oDoc = oAcad.Documents.Open(Kuvat("Avaa").Value, True)
  Do While Not Tiedot.EOF
    Set oEntity = oDoc.HandleToObject(Tiedot("HNDL").Value)
    oEntity.TextString = Nz(Tiedot("Tieto").Value, "")
    Tiedot.MoveNext
  Loop
  oDoc.SaveAs Kuvat("Talleta").Value
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
Select lase Ulase(Kirjain)
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


