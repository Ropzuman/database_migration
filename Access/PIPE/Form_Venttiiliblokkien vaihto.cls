Option Compare Database
Option Explicit

'================================================================================
' Lomake: Venttiiliblokkien vaihto
' Tarkoitus: Korvaa käsiventtiiliblokit P&I-piirustuksissa Araucon standardiblokeilla
' Päivitetty: 2025-11-12 - VBA7/64-bit-tuki ja optimointi lisätty
'
' Kuvaus:
'   Korvaa käsiventtiiliblokit eräajona AutoCAD P&I -piirustuksissa:
'   - Lukee venttiilirecordit MANUALVALVES-taulukosta
'   - Avaa jokaisen venttiilitä sisältävän .DWG-tiedoston
'   - Korvaa vanhat venttiiliblokit uusilla standardiblokeilla
'   - Säilyttää attribuutit, tason ja kiertymän (huomioi peilauksen)
'   - Päivittää tietokantaan uuden Handle- ja blokin nimen
'
' Riippuvuudet:
'   - AutoCAD-sovellus (COM-automaatio)
'   - DAO.Recordset tietokantaoperaatioihin
'   - MANUALVALVES-taulukko, jossa DwgFileID, DwgHandle, Code -kentät
'   - Standardiventtiiliblokit hakemistossa L:\Projekti\ITATA\Process\Flowsheets\BLOCKS\
'
' Venttiilikoodi-kuvaus:
'   Koodi 2 → V11 (sulkuventtiili)
'   Koodi 3 → V05/V06 (palloventtiili, riippuu lähdeblokista)
'   Koodi 4 → V22/V23 (3-tieventtiili, riippuu lähdeblokista)
'   Koodi 5,6 → V09 (takaiskuventtiili)
'   Koodi 12 → V17 (säätöventtiili)
'   Koodi 1,8,11 → Ei korvata
'
' Piirustuspolut:
'   - POR1050970, POR1051314, POR1050965-69 → PINE MACHINE 1
'   - POR1051315, POR1050976-81 → EUCA MACHINE 2
'
' Kontrollit:
'   - Command0: Käynnistää venttiilienvaihdon eräajon
'================================================================================

'================================================================================
' Tapahtuma: Command0_Click
' Tarkoitus: Käynnistää venttiilienvaihdon eräajon
'================================================================================
Private Sub Command0_Click()
  Dim Taulukko As DAO.Recordset
  Dim DWG As String
  Dim Handle As String
  Dim Polku As String
  Dim oACAD As Object  ' AcadApplication - muutettu late binding (64-bit)
  Dim oEntity As Object  ' AcadEntity - muutettu late binding (64-bit)
  Dim oBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim UusiBlock As Object  ' AcadBlockReference - muutettu late binding (64-bit)
  Dim Attribuutit As Variant
  Dim UAttrib As Variant
  Dim Attrib As Object  ' AcadAttributeReference - muutettu late binding (64-bit)
  Dim OK As Boolean
  Dim i As Integer, j As Integer
  Dim BPolku As String
  Dim BNimi As String
  Dim Koodi As Integer
  Dim Rotation As Double
  
  On Error GoTo ErrorHandler

  BPolku = "L:\Projekti\ITATA\Process\Flowsheets\BLOCKS\"
  
  ' Yhdistetään AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application")
  If Err <> 0 Then
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Vaihda Blokit"
    Set oACAD = Nothing
    GoTo Loppu
  End If
  On Error GoTo ErrorHandler
  
  ' Käsitellään kaikki venttiilit
  Set Taulukko = CurrentDb.OpenRecordset("SELECT * FROM MANUALVALVES WHERE (DwgFileID Is Not Null AND DwgHandle Is Not Null AND Deleted = False) Order by DwgFileID")
  
  Do While Not Taulukko.EOF
    ' Avataan uusi piirustus, jos DWG-tiedosto muuttuu
    If UCase$(Nz(Taulukko.Fields("DwgFileID").Value)) <> DWG Then
      DWG = UCase$(Nz(Taulukko.Fields("DwgFileID").Value))
      OK = False
      Select Case DWG
        Case "POR1050970", "POR1051314", "POR1050965", "POR1050966", "POR1050967", "POR1050968", "POR1050969"
          Polku = "L:\Projekti\ITATA\Process\Flowsheets\PINE MACHINE 1\"
        Case "POR1051315", "POR1050976", "POR1050977", "POR1050978", "POR1050979", "POR1050980", "POR1050981"
          Polku = "L:\Projekti\ITATA\Process\Flowsheets\EUCA MACHINE 2\"
        Case Else
          Polku = ""
      End Select
      
      If Dir$(Polku & DWG & ".dwg") = "" Then
        OK = False
      Else
        oACAD.Documents.Open Polku & DWG & ".dwg"
        OK = True
      End If
    End If
    
    ' Käsitellään venttiili, jos tiedot ovat kelvolliset
    Handle = UCase$(Nz(Taulukko.Fields("DwgHandle").Value))
    If Handle <> "" And OK And Nz(Taulukko.Fields("Code").Value) <> "" Then
      Set oBlock = oACAD.ActiveDocument.HandleToObject(Handle)
      Attribuutit = oBlock.GetAttributes
      
      If UBound(Attribuutit) > 1 Then
        ' Tallennetaan venttiilikoodi uuteen blokin nimeen
        Koodi = CInt(Val(Nz(Taulukko.Fields("Code").Value)))
        Select Case Koodi
          Case 2
            BNimi = "V11"
          Case 3
            If oBlock.Name = "UP051" Then
              BNimi = "V06"
            Else
              BNimi = "V05"
            End If
          Case 4
            If oBlock.Name = "UP051" Then
              BNimi = "V23"
            Else
              BNimi = "V22"
            End If
          Case 5, 6 'Check valve
            BNimi = "V09"
          Case 12
            BNimi = "V17"
          Case 1, 8, 11
            BNimi = ""
          Case Else
            BNimi = ""
        End Select
        
        ' Korvataan blokki, jos kuvaus löytyi
        If BNimi <> "" Then
          ' Käsitellään kiertymä (huomioidaan peilaus)
          Rotation = oBlock.Rotation
          If oBlock.XScaleFactor <> 1 Then
            If Rotation > 3 Then
              Rotation = Rotation - 3.141592654
            Else
              Rotation = Rotation + 3.141592654
            End If
          End If
          
          ' Sijoitetaan uusi blokki
          Set UusiBlock = oACAD.ActiveDocument.ModelSpace.InsertBlock(oBlock.InsertionPoint, BPolku & BNimi & ".dwg", 1, 1, 1, Rotation)
          UusiBlock.Layer = oBlock.Layer
          
          ' Kopioidaan attribuutit
          UAttrib = UusiBlock.GetAttributes
          For i = LBound(Attribuutit) To UBound(Attribuutit)
            For j = LBound(UAttrib) To UBound(UAttrib)
              If UCase$(Attribuutit(i).TagString) = UCase$(UAttrib(j).TagString) Then
                UAttrib(j).TextString = Attribuutit(i).TextString
                UAttrib(j).Alignment = Attribuutit(i).Alignment
                UAttrib(j).Rotation = Attribuutit(i).Rotation
                UAttrib(j).InsertionPoint = Attribuutit(i).InsertionPoint
                Exit For
              End If
            Next j
          Next i
          
          ' Päivitetään tietokantaan uusi handle ja blokin nimi
          Taulukko.Edit
          Taulukko.Fields("DwgHandle").Value = UusiBlock.Handle
          Taulukko.Fields("DwgBlock").Value = BNimi
          Taulukko.Update
          
          ' Poistetaan vanha blokki
          oBlock.Delete
        End If
      End If
    End If
    
    Taulukko.MoveNext
  Loop
  
  MsgBox "Valmis", vbInformation
  GoTo Loppu
  
ErrorHandler:
  MsgBox "Virhe venttiilien vaihdossa: " & Err.Description, vbExclamation
  
Loppu:
  ' Cleanup
  If Not Taulukko Is Nothing Then
    Taulukko.Close
    Set Taulukko = Nothing
  End If
  Set UusiBlock = Nothing
  Set oBlock = Nothing
  Set oACAD = Nothing
End Sub
