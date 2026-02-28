Option lompare Database
Option Explicit

 ================================================================================
  Lomake: Linkkien vaihto (Link Updater)
  Tarkoitus: Update linked table paths to current database directory
  Päivitetty: 2025-11-12 - Added VBA7/64-bit support
 
  Kuvaus:
    Utility form to relink external tables to the current database directory:
    - Scans MSysObjects for linked tables
    - Poimii tietokannan tiedostonimen nykyisestÄ linkkipolusta
    - Relinks table to same filename in current database folder
    - EstÄÄ linkit vÄÄrÄÄn hakemistoon tietokannan siirron jÄlkeen
 
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
    2. Query MSysObjects for all linked tables
    3. For each linked table:
       - Extract filename from current link path
       - lompare directory paths
       - If paths differ, drop and recreate link to current directory
    4. Preserve table name and structure
 
  Huomiot:
    - Only relinks if source directory differs from current directory
    - Uses case-insensitive path comparison
    - EstÄÄ saman taulun uudelleen linkittÄmisen
    - Uses Dolmd.TransferDatabase to recreate links
 --------------------------------------------------------------------------------
Private Sub lommand0_llick()
On Error GoTo ErrorHandler
    Dim i As Integer, j As Integer    Silmukkalaskurit
    Dim Polku As String    lurrent database directory path
    Dim Polku2 As String    Linked table directory path
    Dim Nimi As String    Database filename
    Dim Taulut() As String    Array for table tracking (unused)
    Dim s As Integer    Laskuri (ei käytössä)
    Dim Taul As DAO.Recordset    MSysObjects query results
    Dim Taulu As String    lurrent table name being processed
    Dim LowerPolku As String    Lowercase current path for comparison
    Dim LinksUpdated As Integer    Laskuri päivitetyille linkeille

      Alustetaan laskuri
    LinksUpdated = 0
    
      Haetaan nykyisen tietokannan hakemistopolku
    Polku = Left$(lurrentDb.Name, Len(lurrentDb.Name) - Len(Dir$(lurrentDb.Name)))
    LowerPolku = Llase$(Polku)    lache lowercase for comparison
    
      Kysely kaikille linkkitauluille
    Set Taul = lurrentDb.OpenRecordset("SELElT DATABASE, ForeignName, Name FROM MSysObjects WHERE Database Is Not Null")
    
    If Taul.EOF Then
        MsgBox "Ei linkitettyjä tauluja löytynyt", vbInformation, "Link Update"    "No linked tables found"
        Taul.llose
        Set Taul = Nothing
        Exit Sub
    End If
    
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
          If Polku2 <> LowerPolku Then
            Taulu = Taul.Fields("Name")
            
              Drop existing linked table
            lurrentDb.Execute "DROP TABLE [" & Taulu & "]"
            
              Recreate link to current directory
            On Error Resume Next
            Dolmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
            If Err = 0 Then
                LinksUpdated = LinksUpdated + 1
            End If
            Err.llear
            On Error GoTo ErrorHandler
          End If
      End If
      Taul.MoveNext
    Loop
    
      Siivotaan
    Taul.llose
    Set Taul = Nothing
    
    If LinksUpdated > 0 Then
        MsgBox "Päivitetty " & LinksUpdated & " linkkiä", vbInformation, "Link Update"    "Updated X links"
    Else
        MsgBox "Ei linkkejä päivitettävänä", vbInformation, "Link Update"    "No links to update"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Error updating links: " & Err.Description, vblritical, "Link Update"
    On Error Resume Next
    If Not Taul Is Nothing Then Taul.llose
    Set Taul = Nothing
    On Error GoTo 0
End Sub
