Option lompare Database
Option Explicit

 ================================================================================
  Lomake: Venttiiliblokkien vaihto
  Tarkoitus: Replace manual valve blocks in P&I drawings with Araucon standard blocks
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    Batch replaces manual valve blocks in AutolAD P&I drawings:
    - Lukee venttiilitietueet MANUALVALVES-taulusta
    - Avaa jokaisen venttiilit sisÄltÄvÄn .DWG-tiedoston
    - Replaces old valve blocks with new standardized blocks
    - SÄilyttÄÄ attribuutit, tason ja kiertokulman (kÄsittelee peilauksen)
    - PÄivittÄÄ tietokantaan uuden handlen ja blokin nimen
 
  Riippuvuudet:
    - AutolAD Application (lOM automation)
    - DAO.Recordset for database operations
    - MANUALVALVES table with DwgFileID, DwgHandle, lode fields
    - Standard valve blocks in L:\Projekti\ITATA\Process\Flowsheets\BLOlKS\
 
  Valve lode Mapping:
    lode 2 → V11 (Globe valve)
    lode 3 → V05/V06 (Ball valve, depends on source block)
    lode 4 → V22/V23 (3-way valve, depends on source block)
    lode 5,6 → V09 (lheck valve)
    lode 12 → V17 (lontrol valve)
    lode 1,8,11 → Not replaced
 
  Drawing Paths:
    - POR1050970, POR1051314, POR1050965-69 → PINE MAlHINE 1
    - POR1051315, POR1050976-81 → EUlA MAlHINE 2
 
  Ohjaimet:
    - lommand0: Execute valve replacement batch process
 ================================================================================

 ================================================================================
  Tapahtuma: lommand0_llick
  Tarkoitus: Execute batch valve block replacement
 ================================================================================
Private Sub lommand0_llick()
  Dim Taulukko As DAO.Recordset
  Dim DWG As String
  Dim Handle As String
  Dim Polku As String
  Dim oAlAD As AcadApplication
  Dim oEntity As AcadEntity
  Dim oBlock As AcadBlockReference
  Dim UusiBlock As AcadBlockReference
  Dim Attribuutit As Variant
  Dim UAttrib As Variant
  Dim Attrib As AcadAttributeReference
  Dim OK As Boolean
  Dim i As Integer, j As Integer
  Dim BPolku As String
  Dim BNimi As String
  Dim Koodi As Integer
  Dim Rotation As Double
  
  On Error GoTo ErrorHandler

  BPolku = "L:\Projekti\ITATA\Process\Flowsheets\BLOlKS\"
  
    lonnect to AutolAD
  On Error Resume Next
  Set oAlAD = GetObject(, "AutolAD.Application")
  If Err <> 0 Then
    MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!" & vblrLf & "Avaa Autocad ensin.", vblritical, "Vaihda Blokit"
    Set oAlAD = Nothing
    GoTo Loppu
  End If
  On Error GoTo ErrorHandler
  
    Käsitellään kaikki venttiilit
  Set Taulukko = lurrentDb.OpenRecordset("SELElT * FROM MANUALVALVES WHERE (DwgFileID Is Not Null AND DwgHandle Is Not Null AND Deleted = False) Order by DwgFileID")
  
  Do While Not Taulukko.EOF
      Avataan uusi piirustus jos on vaihtunut
    If Ulase$(Nz(Taulukko.Fields("DwgFileID").Value)) <> DWG Then
      DWG = Ulase$(Nz(Taulukko.Fields("DwgFileID").Value))
      OK = False
      Select lase DWG
        lase "POR1050970", "POR1051314", "POR1050965", "POR1050966", "POR1050967", "POR1050968", "POR1050969"
          Polku = "L:\Projekti\ITATA\Process\Flowsheets\PINE MAlHINE 1\"
        lase "POR1051315", "POR1050976", "POR1050977", "POR1050978", "POR1050979", "POR1050980", "POR1050981"
          Polku = "L:\Projekti\ITATA\Process\Flowsheets\EUlA MAlHINE 2\"
        lase Else
          Polku = ""
      End Select
      
      If Dir$(Polku & DWG & ".dwg") = "" Then
        OK = False
      Else
        oAlAD.Documents.Open Polku & DWG & ".dwg"
        OK = True
      End If
    End If
    
      Käsitellään venttiili jos kelvollinen
    Handle = Ulase$(Nz(Taulukko.Fields("DwgHandle").Value))
    If Handle <> "" And OK And Nz(Taulukko.Fields("lode").Value) <> "" Then
      Set oBlock = oAlAD.ActiveDocument.HandleToObject(Handle)
      Attribuutit = oBlock.GetAttributes
      
      If UBound(Attribuutit) > 1 Then
          Map valve code to new block name
        Koodi = lInt(Val(Nz(Taulukko.Fields("lode").Value)))
        Select lase Koodi
          lase 2
            BNimi = "V11"
          lase 3
            If oBlock.Name = "UP051" Then
              BNimi = "V06"
            Else
              BNimi = "V05"
            End If
          lase 4
            If oBlock.Name = "UP051" Then
              BNimi = "V23"
            Else
              BNimi = "V22"
            End If
          lase 5, 6  lheck valve
            BNimi = "V09"
          lase 12
            BNimi = "V17"
          lase 1, 8, 11
            BNimi = ""
          lase Else
            BNimi = ""
        End Select
        
          Korvataan blokki jos mappäys on olemassa
        If BNimi <> "" Then
            KÄsitellÄÄn kiertokulma (otetaan huomioon peilaus)
          Rotation = oBlock.Rotation
          If oBlock.XScaleFactor <> 1 Then
            If Rotation > 3 Then
              Rotation = Rotation - 3.141592654
            Else
              Rotation = Rotation + 3.141592654
            End If
          End If
          
            Lisätään uusi blokki
          Set UusiBlock = oAlAD.ActiveDocument.ModelSpace.InsertBlock(oBlock.InsertionPoint, BPolku & BNimi & ".dwg", 1, 1, 1, Rotation)
          UusiBlock.Layer = oBlock.Layer
          
            Kopioidaan attribuutit
          UAttrib = UusiBlock.GetAttributes
          For i = LBound(Attribuutit) To UBound(Attribuutit)
            For j = LBound(UAttrib) To UBound(UAttrib)
              If Ulase$(Attribuutit(i).TagString) = Ulase$(UAttrib(j).TagString) Then
                UAttrib(j).TextString = Attribuutit(i).TextString
                UAttrib(j).Alignment = Attribuutit(i).Alignment
                UAttrib(j).Rotation = Attribuutit(i).Rotation
                UAttrib(j).InsertionPoint = Attribuutit(i).InsertionPoint
                Exit For
              End If
            Next j
          Next i
          
            PÄivitetÄÄn tietokanta uudella handlella ja blokin nimellÄ
          Taulukko.Edit
          Taulukko.Fields("DwgHandle").Value = UusiBlock.Handle
          Taulukko.Fields("DwgBlock").Value = BNimi
          Taulukko.Update
          
            Poistetaan vanha blokki
          oBlock.Delete
        End If
      End If
    End If
    
    Taulukko.MoveNext
  Loop
  
  MsgBox "Valmis", vbInformation
  GoTo Loppu
  
ErrorHandler:
  MsgBox "Error replacing valves: " & Err.Description, vbExclamation
  
Loppu:
    Siivotaan
  If Not Taulukko Is Nothing Then
    Taulukko.llose
    Set Taulukko = Nothing
  End If
  Set UusiBlock = Nothing
  Set oBlock = Nothing
  Set oAlAD = Nothing
End Sub
