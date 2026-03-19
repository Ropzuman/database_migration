Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysAddToDistr
' SOVELLUS : DOCUMENTS — Lisää dokumentti distribuutioon
' KUVAUS   : Tarkistaa onko dokumentti jo lisätty valittuun distribuutioon
'            ja lisää sen, jos ei. Käyttää transaktioita atomaarisuuden
'            takaamiseksi.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Const TV_DOC_CURRENT_COUNTER As String = "DOC_CurrentCounter"
Private Const TV_DOC_DIST_REV As String = "DOC_AddToDistr_Rev"
Private Const TV_DOC_DIST_STATUS As String = "DOC_AddToDistr_Status"


'''
' Add Document to Distribution
' OPTIMIZED 2025-11-09: Added error handling, explicit DAO typing, transaction support, null-checking
'''
Private Sub Add_Click()
  Dim DB As DAO.Database
  Dim taulu As DAO.Recordset
  Dim DokumenttiID As Long
  Dim DokumentinRev As String
  Dim DokumentinStatus As String
  Dim inTransaction As Boolean
  
  On Error GoTo ErrorHandler
  inTransaction = False
  
  ' Esitarkistus: varmistetaan että distribuutio on valittu
  If IsNull(Me.Distribution.Value) Then
    MsgBox "Please select a distribution first.", vbExclamation, "Validation"
    Exit Sub
  End If

  DokumenttiID = LueTempVarPitka(TV_DOC_CURRENT_COUNTER)
  DokumentinRev = LueTempVarTeksti(TV_DOC_DIST_REV)
  DokumentinStatus = LueTempVarTeksti(TV_DOC_DIST_STATUS)
  If DokumenttiID = 0 Then
    MsgBox "Document context missing. Please close and reopen this dialog.", vbCritical, "Validation"
    Exit Sub
  End If
  
  Set DB = CurrentDb

  ' Suorituskykyisempi DCount-tarkistus SQL-injektioriskin minimoimiseksi
  If DCount("*", "SentDocuments", "TDoc=" & DokumenttiID & " AND TDistr=" & CLng(Me.Distribution.Value)) > 0 Then
    MsgBox "Document were already on selected distribution!", vbCritical, "Add document to distribution"
  Else
    DBEngine.BeginTrans
    inTransaction = True

    ' Avataan taulu pelkkää lisäystä varten
    Set taulu = DB.OpenRecordset("SentDocuments", dbOpenDynaset, dbAppendOnly)
    taulu.AddNew
    taulu.Fields(0) = DokumenttiID
    taulu.Fields(1) = Me.Distribution.Value
    taulu.Fields(2) = IIf(DokumentinRev = "", Null, DokumentinRev)
    If DokumentinStatus <> "" Then taulu.Fields(3) = DokumentinStatus
    taulu.Update

    DBEngine.CommitTrans
    inTransaction = False

    DoCmd.Close
  End If
  
Cleanup:
  If Not taulu Is Nothing Then
    taulu.Close
    Set taulu = Nothing
  End If
  ' EI DB.Close: CurrentDb-viitettä ei tule sulkea
  Set DB = Nothing
  Exit Sub
  
ErrorHandler:
  If inTransaction Then
    DBEngine.Rollback  ' Vapautetaan tietokanta lukot epäonnistuneen tallennuksen jälkeen
    inTransaction = False
  End If
  MsgBox "Error adding document to distribution: " & Err.Description, vbCritical, "Add Error"
  Resume Cleanup
End Sub

Private Sub Command14_Click()
  DoCmd.Close
End Sub

Private Sub Form_Load()
  ' Valitaan oletuksena ensimmäinen distribuutiovaihtoehto
  Me.Distribution.Value = Me.Distribution.ItemData(1)
End Sub

Private Function LueTempVarPitka(ByVal Nimi As String) As Long
  Dim Arvo As String
  ' Käytä paikallista lukijaa, se kutsuu pääformin versiota
  Arvo = LueTempVarTeksti(Nimi)
  If Arvo <> "" And IsNumeric(Arvo) Then
    LueTempVarPitka = CLng(Arvo)
  Else
    LueTempVarPitka = 0
  End If
End Function

Private Function LueTempVarTeksti(ByVal Nimi As String) As String
  On Error GoTo ErrorHandler
  LueTempVarTeksti = Nz(TempVars(Nimi).Value, "")
  Exit Function
ErrorHandler:
  LueTempVarTeksti = ""
End Function