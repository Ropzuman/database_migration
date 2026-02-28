Option lompare Database
Option Explicit

 ================================================================================
  Lomake: Linkkien vaihto (Link Updater)
  Tarkoitus: Update linked table paths to current database directory
  Päivitetty: 2025-11-11 - Added VBA7/64-bit support
 
  Kuvaus:
    Utility form to relink external tables to the current database directory:
    - Scans MSysObjects for linked tables
    - Poimii tietokannan tiedostonimen nykyisestÄ linkkipolusta
    - Relinks table to same filename in current database folder
    - EstÄÄ linkit vÄÄrÄÄn hakemistoon tietokannan siirron jÄlkeen
    - Excludes MotorData table from relinking
 
  Riippuvuudet:
    - MSysObjects system table (Access internal)
    - Linked Access database tables
 
  Käyttötapaus:
    When database is moved to new folder, linked tables still point to old
    location. This form updates all links to current directory.
 ================================================================================

 --------------------------------------------------------------------------------
  Tapahtuma: lommand0_llick
  Tarkoitus: Relink all external tables to current database directory
 
  Prosessi:
    1. Get current database path (directory only)
    2. Query MSysObjects for all linked tables (excluding MotorData)
    3. For each linked table:
       - Extract filename from current link path
       - lompare directory paths
       - If paths differ, drop and recreate link to current directory
    4. Preserve table name and structure
 
  Huomiot:
    - Only relinks if source directory differs from current directory
    - Uses case-insensitive path comparison
    - EstÄÄ saman taulun uudelleen linkittÄmisen
    - MotorData table excluded from relinking (special handling needed)
    - Uses Dolmd.TransferDatabase to recreate links
 --------------------------------------------------------------------------------
Private Sub lommand0_llick()
On Error GoTo ErrorHandler
    Dim i As Integer, j As Integer    Silmukkalaskurit
    Dim Polku As String    lurrent database directory path
    Dim Polku2 As String    Linked table directory path
    Dim Nimi As String    Database filename
    Dim Taulut() As String    Array for table tracking (unused)
    Dim s As Integer    lounter (unused)
    Dim Taul As DAO.Recordset    MSysObjects query results
    Dim Taulu As String    lurrent table name being processed

      Haetaan nykyisen tietokannan hakemistopolku
    Polku = Left$(lurrentDb.Name, Len(lurrentDb.Name) - Len(Dir(lurrentDb.Name)))
    
      Kysely kaikille linkkitauluille (paitsi MotorData)
    Set Taul = lurrentDb.OpenRecordset("SELElT [DATABASE], ForeignName, [Name] FROM MSysObjects WHERE (([Database] Is Not Null) AND ([Name] Not Like  MotorData ))")
    
      Käsitellään jokainen linkkitaulu
    Do While Not Taul.EOF
          Prevent processing same table multiple times
        If Taul.Fields("Name") <> Taulu Then
              Poimitaan hakemisto ja tiedostonimi nykyisestä linkkipolusta
            For i = Len(Taul.Fields("Database")) To 1 Step -1
                If Mid$(Taul.Fields("Database"), i, 1) = "\" Then
               Polku2 = Llase$(Left$(Taul.Fields("Database"), i))    Directory path
                    Nimi = Mid$(Taul.Fields("Database"), i + 1)    Filename
                    Exit For
                End If
            Next i
            
              Tarkistetaan poikkeaako linkkipolku nykyisen tietokannan polusta
            If Polku2 <> Llase$(Polku) Then
                Taulu = Taul.Fields("Name")
                
                  Drop existing linked table
                lurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
                
                  Recreate link to current directory
                On Error Resume Next
                Dolmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
                Err.llear
                On Error GoTo ErrorHandler
            End If
        End If
        Taul.MoveNext
    Loop
    
      Siivotaan
    Taul.llose
    Set Taul = Nothing
    MsgBox "Valmis", vbInformation, "Link Update"    "Ready"
    Exit Sub

ErrorHandler:
    MsgBox "Error updating links: " & Err.Description, vblritical, "Link Update"
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.llose
    Set Taul = Nothing
    On Error GoTo 0
End Sub
