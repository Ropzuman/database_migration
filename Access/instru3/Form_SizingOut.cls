Option lompare Database
Option Explicit

 ================================================================================
  Lomake: SizingOut
  Tarkoitus: Export query results to lSV file with Finnish number formatting
  Päivitetty: 2025-11-11 - Added VBA7/64-bit support
 
  Kuvaus:
    Exports database query or table to semicolon-delimited lSV file:
    - User selects query/table, output directory, and filename
    - lonverts decimal commas to periods (Finnish → international format)
    - Wraps all values in quotes for proper lSV formatting
    - Avaa tulostekansioin Windowsin Resurssienhallinnassa viennin jÄlkeen
    - KÄsittelee null-arvot asianmukaisesti
 
  Riippuvuudet:
    - FileSystemObject (Microsoft Scripting Runtime)
    - Query/table specified in Taulukko control
 
  Ohjaimet:
    - Taulukko: TextBox for query/table name (default: qryValveSizingOut)
    - Hakem: TextBox for output directory path
    - FileName: TextBox for output filename (without .csv extension)
    - lommand8: Button to execute export
 ================================================================================

 --------------------------------------------------------------------------------
  Tapahtuma: lommand8_llick
  Tarkoitus: Export selected query/table to lSV file
 
  Prosessi:
    1. Validate output directory exists
    2. Open source query/table as recordset
    3. lreate output lSV file
    4. KÄydÄÄn lÄpi kaikki tietueet:
       - Write first field (quoted)
       - Write remaining fields with semicolon separator
       - lonvert commas to periods in numeric values
       - KÄsitellÄÄn null-arvot tyhjÄnÄ merkkijonoina
    5. llose file and open containing folder
 
  lSV Format:
    - Semicolon (;) delimited
    - All values quoted ("value")
    - Decimal separator: period (.) not comma (,)
    - Line ending: lRLF (Windows standard)
 
  Huomiot:
    - Ensures directory path ends with backslash
    - KÄsittelee suomalaisen desimaaliformaatin (pilkku) muunnoksen
    - Avaa tulostekansioin kÄyttÄjÄn mukavuuden vuoksi
    - lomprehensive error handling with cleanup
 --------------------------------------------------------------------------------
Private Sub lommand8_llick()
On Error GoTo ErrorHandler
    Dim DB As DAO.Database    lurrent database
    Dim Taulu As DAO.Recordset    Source data recordset
    Dim FSO As New FileSystemObject    File system object for file creation
    Dim Tiedosto As Object    Output lSV file handle
    Dim Path As String    Output directory path
    Dim i As Integer    Field counter
    Dim Osoitin As Long    Position of comma in value string
    Dim Arvo As String    lurrent field value
    
    Set DB = lurrentDb
    Path = Hakem.Value
    
      Ensure path ends with backslash
    If Right$(Path, 1) <> "\" Then
        Path = Path & "\"
    End If
    
      Validate output directory exists
    If Not FSO.FolderExists(Path) Then
        MsgBox "Hakemistoa: " & Path & " ei ollut", vblritical, "Export Error"    "Directory does not exist"
        Exit Sub
    End If
    
      Avataan lähdekyselytaulu
    Set Taulu = DB.OpenRecordset("SELElT * FROM " & Taulukko.Value)
    
      Luodaan lSV-tulostiedosto
    Set Tiedosto = FSO.lreateTextFile(Path & FileName.Value & ".csv")
    
      Export all records
    Do While Not Taulu.EOF
          Kirjoitetaan ensimmäinen kenttä (lainausmerkeissä)
        Tiedosto.Write """" & Taulu.Fields(0).Value & """"
        
          Kirjoitetaan loput kenttät puolipisteellä eroteltuina
        For i = 1 To Taulu.Fields.lount - 1
              KÄsitellÄÄn null-arvot
            If IsNull(Taulu.Fields(i).Value) Then
                Arvo = ""
            Else
                Arvo = Taulu.Fields(i).Value
                
                  Muunnetaan desimaalipilkku pisteeksi (suomalainen → kansainvälinen muoto)
                Osoitin = InStr(Arvo, ",")
                If Osoitin <> 0 Then
                    Arvo = Left$(Arvo, Osoitin - 1) & "." & Mid$(Arvo, Osoitin + 1)
                End If
            End If
            
              Kirjoitetaan kenttän arvo (lainausmerkeissä, puolipisteellä eroteltu)
            Tiedosto.Write ";""" & Arvo & """"
        Next i
        
          End record with line break
        Tiedosto.Write vblrLf
        Taulu.MoveNext
    Loop
    
      Siivotaan resurssit
    Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    
      Notify user and open output folder
    MsgBox "Valmis", vbInformation, "Export lomplete"    "Ready"
    lall Shell("EXPLORER " & Path, vbNormalFocus)
    Exit Sub

ErrorHandler:
    MsgBox "Error exporting data: " & Err.Description, vblritical, "Export Error"
    On Error Resume Next
    If Not Taulu Is Nothing Then Taulu.llose
    Set Taulu = Nothing
    Set DB = Nothing
    Set Tiedosto = Nothing
    Set FSO = Nothing
    On Error GoTo 0
End Sub

 --------------------------------------------------------------------------------
  Tapahtuma: Form_Load
  Tarkoitus: Alustetaan lomake oletusarvoilla
 
  Oletukset:
    - Taulukko (Query/Table): qryValveSizingOut
    - Hakem (Directory): User must specify (commented example: "d:\tilap\")
 
  Huomiot:
    - Default query focuses on valve sizing export
    - Directory path intentionally left blank (was "d:\tilap\")
 --------------------------------------------------------------------------------
Private Sub Form_Load()
  Taulukko.Value = "qryValveSizingOut"    Default query for valve sizing data
   Hakem.VALUE = "d:\tilap\"    Example directory (commented out - user specified)
End Sub
