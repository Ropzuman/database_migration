Option Compare Database
Option Explicit

'==============================================================================
' Lomake: GeneroiMoottorikuvat
' Tarkoitus: Generoi moottoreiden tietolehtikuvat AutoCADissa
' Päivitetty: 2025-11-11 - DAO-tyypit, myöhäinen sidonta AutoCADiin
'             2026-11-10 - Kovakoodatut projektipolut poistettu: käyttäjä valitsee
'                          MotMittakuvat-kansion, virheenkäsittely lisätty, käyttämättömät
'                          LisaaTeksti- ja AnnaSarake-rutiinit poistettu
'
' Kansiorakenne (käyttäjän valitsema kansio):
'   BASE\  - pohjakuvat ja DataSheet.dwg
'   DWG\   - valmiit kuvat
'==============================================================================

' Myöhäinen sidonta AutoCADiin - tyyppikirjastoviittausta ei tarvita
Private oAcad As Object  ' AcadApplication
Private oDoc As Object   ' AcadDocument

' Yhdistää käynnissä olevaan AutoCADiin; palauttaa False, jos sitä ei löydy
Private Function YhdistaAutoCAD() As Boolean
    On Error Resume Next
    Set oAcad = GetObject(, "AutoCAD.Application")
    If Err.Number <> 0 Or oAcad Is Nothing Then
        Err.Clear
        Set oAcad = Nothing
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Generoi kuvat"
    Else
        YhdistaAutoCAD = True
    End If
End Function

' Tarkistaa, onko piirustus jo auki AutoCADissa (ettei käyttäjän avointa kuvaa suljeta)
Private Function OnkoAvoinna(ByVal Polku As String) As Boolean
    Dim Dok As Object  ' Avoin piirustus

    For Each Dok In oAcad.Documents
        If StrComp(Dok.FullName, Polku, vbTextCompare) = 0 Then
            OnkoAvoinna = True
            Exit Function
        End If
    Next Dok
End Function

' Näyttää kansionvalintaikkunan; palauttaa polun kenoviivaan päättyen tai "" jos peruttiin
Private Function ValitseKansio(ByVal Otsikko As String) As String
    Dim fd As Object  ' FileDialog-objekti

    Set fd = Application.FileDialog(4)  ' msoFileDialogFolderPicker = 4
    With fd
        .Title = Otsikko
        .AllowMultiSelect = False
        If .Show = -1 Then
            ValitseKansio = .SelectedItems(1)
            If Right$(ValitseKansio, 1) <> "\" Then ValitseKansio = ValitseKansio & "\"
        End If
    End With
    Set fd = Nothing
End Function

' Palauttaa kentän arvon tekstinä (Null -> tyhjä merkkijono)
Private Function Teksti(ByVal rs As DAO.Recordset, ByVal Kentta As String) As String
    Teksti = CStr(Nz(rs.Fields(Kentta).Value, ""))
End Function

'------------------------------------------------------------------------------
' Tapahtuma: GenKuvat_Click
' Tarkoitus: Luo jokaiselle _qry_MotMittakuvat-rivin moottorille tietolehtikuva
'------------------------------------------------------------------------------
Private Sub GenKuvat_Click()
On Error GoTo ErrorHandler
    Dim Kuvat As DAO.Recordset  ' Generoitavat moottorit
    Dim oBlock As Object  ' AcadBlockReference (myöhäinen sidonta)
    Dim DocTila As Boolean  ' AutoCADin SingleDocumentMode ennen ajoa
    Dim TilaTallennettu As Boolean  ' Onko SingleDocumentMode tallennettu
    Dim DB As DAO.Database  ' Nykyinen tietokanta
    Dim i As Long, K As Long, Maara As Long  ' Laskurit
    Dim Attrib As Variant  ' Blokin attribuutit
    Dim iPoint(2) As Double  ' Lisäyspiste (0,0,0)
    Dim Joukko As Object  ' AcadSelectionSet (myöhäinen sidonta)
    Dim FilterType(0) As Integer  ' Valintasuodattimen DXF-koodi
    Dim FilterData(0) As Variant  ' Valintasuodattimen arvo
    Dim Laitettu As Integer  ' Täytettyjen TITLE-attribuuttien määrä
    Dim Juuri As String  ' Käyttäjän valitsema MotMittakuvat-kansio
    Dim Pohjahakem As String  ' Pohjakuvien hakemisto
    Dim Kohdehakem As String  ' Valmiiden kuvien hakemisto
    Dim Pohja As String  ' Avattava pohjakuva
    Dim virhenumero As Long  ' Virheen numero käsittelijässä
    Dim virhekuvaus As String  ' Virheen kuvaus käsittelijässä

    If Not YhdistaAutoCAD Then Exit Sub

    Juuri = ValitseKansio("Valitse MotMittakuvat-kansio (sisältää BASE- ja DWG-alikansiot)")
    If Juuri = "" Then GoTo Siivous
    Pohjahakem = Juuri & "BASE\"
    Kohdehakem = Juuri & "DWG\"
    If Len(Dir$(Pohjahakem, vbDirectory)) = 0 Or Len(Dir$(Kohdehakem, vbDirectory)) = 0 Then
        MsgBox "Valitusta kansiosta puuttuu BASE- tai DWG-alikansio.", vbExclamation, "Generoi kuvat"
        GoTo Siivous
    End If
    If Len(Dir$(Pohjahakem & "DataSheet.dwg")) = 0 Then
        MsgBox "Tiedostoa DataSheet.dwg ei löydy kansiosta " & Pohjahakem, vbExclamation, "Generoi kuvat"
        GoTo Siivous
    End If

    Set DB = CurrentDb
    Set Kuvat = DB.OpenRecordset("_qry_MotMittakuvat", dbOpenSnapshot)
    If Kuvat.EOF Then
        MsgBox "Generoitavia moottorikuvia ei löytynyt.", vbInformation, "Generoi kuvat"
        GoTo Siivous
    End If
    Kuvat.MoveLast
    Kuvat.MoveFirst
    Maara = Kuvat.RecordCount
    FilterType(0) = 2  ' DXF 2 = blokin nimi
    FilterData(0) = "TITLE"

    DocTila = oAcad.Preferences.System.SingleDocumentMode
    TilaTallennettu = True
    oAcad.Preferences.System.SingleDocumentMode = False

    Do While Not Kuvat.EOF
        K = K + 1
        SysCmd acSysCmdSetStatus, "Generoidaan kuvaa " & K & "/" & Maara
        Pohja = Pohjahakem & Teksti(Kuvat, "POHJA")
        If OnkoAvoinna(Pohja) Then Err.Raise vbObjectError + 1001, , "Pohjakuva on jo auki AutoCADissa, sulje se ensin: " & Pohja

        ' Avataan pohjakuva vain luku -tilassa ja lisätään siihen tietolehtiblokki
        Set oDoc = oAcad.Documents.Open(Pohja, True)
        Set oBlock = oDoc.ModelSpace.InsertBlock(iPoint, Pohjahakem & "DataSheet.dwg", 1, 1, 1, 0)
        Attrib = oBlock.GetAttributes
        For i = 0 To UBound(Attrib)
            Select Case Attrib(i).TagString
                Case "ENCLOSURE"
                    Attrib(i).TextString = Teksti(Kuvat, "ENCLOSURE")
                Case "FLC"  ' Täyskuormavirta
                    Attrib(i).TextString = Teksti(Kuvat, "NomCurr")
                Case "FRAME"
                    Attrib(i).TextString = Teksti(Kuvat, "IEC FRAME")
                Case "LRC"  ' Pyörivän roottorin lukitusvirta
                    Attrib(i).TextString = Teksti(Kuvat, "Startcur")
                Case "MANUF"
                    Attrib(i).TextString = Teksti(Kuvat, "MotorMANUF")
                Case "MOUNT"
                    Attrib(i).TextString = Teksti(Kuvat, "MOUNTTYPE")
                Case "NAME"
                    Attrib(i).TextString = Teksti(Kuvat, "NAME")
                Case "OPTIONS"
                    Attrib(i).TextString = Teksti(Kuvat, "OPTIONS")
                Case "POWER"
                    Attrib(i).TextString = Teksti(Kuvat, "SPower")
                Case "RPM"
                    Attrib(i).TextString = Teksti(Kuvat, "RPM")
                Case "RPMMAX"
                    Attrib(i).TextString = Teksti(Kuvat, "MAXRPM")
                Case "RPMMIN"
                    Attrib(i).TextString = Teksti(Kuvat, "MINRPM")
                Case "RTORQUE"
                    Attrib(i).TextString = Teksti(Kuvat, "LoadRtdTorque")
                Case "STORQUE"
                    Attrib(i).TextString = Teksti(Kuvat, "LoadStTorque")
                Case "TRANSMISSION"
                    Attrib(i).TextString = Teksti(Kuvat, "TTypeName")
                Case "TYPE"
                    Attrib(i).TextString = Teksti(Kuvat, "MotorTYPE")
            End Select
        Next i

        ' Haetaan pohjakuvan TITLE-blokki ja täytetään sen otsikkotiedot
        Set Joukko = oDoc.ActiveSelectionSet
        Joukko.Clear
        Joukko.Select 5, , , FilterType, FilterData  ' 5 = acSelectionSetAll
        If Joukko.Count = 0 Then Err.Raise vbObjectError + 1003, , "TITLE-blokkia ei löytynyt kuvasta " & Pohja
        Set oBlock = Joukko(0)
        Attrib = oBlock.GetAttributes
        Laitettu = 0
        For i = 0 To UBound(Attrib)
            Select Case Attrib(i).TagString
                Case "Title1"
                    Attrib(i).TextString = "Motor"
                    Laitettu = Laitettu + 1
                Case "Title2"
                    Attrib(i).TextString = Teksti(Kuvat, "Mcustpos")
                    Laitettu = Laitettu + 1
                Case "MetsoNo"
                    Attrib(i).TextString = Teksti(Kuvat, "Nimi")
                    Laitettu = Laitettu + 1
            End Select
            If Laitettu = 3 Then Exit For
        Next i

        oDoc.SaveAs Kohdehakem & Teksti(Kuvat, "nimi")
        oDoc.Close False
        Set oDoc = Nothing
        Kuvat.MoveNext
    Loop
    MsgBox "Valmis", vbOKOnly, "Generoi kuvat"

Siivous:
    On Error Resume Next
    SysCmd acSysCmdClearStatus
    If TilaTallennettu Then oAcad.Preferences.System.SingleDocumentMode = DocTila
    If Not Kuvat Is Nothing Then Kuvat.Close
    Set Kuvat = Nothing
    Set Joukko = Nothing
    Set oBlock = Nothing
    Set oDoc = Nothing
    Set oAcad = Nothing
    Set DB = Nothing
    If virhenumero <> 0 Then
        MsgBox "Virhe kuvien generoinnissa (" & virhenumero & "): " & virhekuvaus, vbCritical, "Generoi kuvat"
    End If
    Exit Sub

ErrorHandler:
    virhenumero = Err.Number
    virhekuvaus = Err.Description
    On Error Resume Next
    ' Suljetaan keskeneräinen piirustus tallentamatta
    If Not oDoc Is Nothing Then oDoc.Close False
    Resume Siivous
End Sub
