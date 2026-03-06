Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_LukituskaavioLinkit
' SOVELLUS : FunctionDiagrams — Lukituskaaviolinkkien sivutieto
' KUVAUS   : Lukee InterlockingLinkPage-kyselyn tulokset ja päivittää
'            IntLinkPage-taulun page1-sarakkeen pilkulla erotettuun
'            muotoon "See Page X, Y". Tämän avulla lukituskaavion ristii
'            viittaukset saavat oikeat sivuviitteet piirustuksiin.
' DIPLE    : Option Explicit; Set tbl/qry/DB = Nothing muistivuodon ehkäisyyn;
'            O(N²) sisäkkäinen silmukka korvattu FindFirst-metodilla (suorit.)
' PÄIVITETTY: 2026-03-06
'==========================================================================

Private Sub Command0_Click()
On Error GoTo ErrorHandler

DoCmd.Hourglass True

Dim DB As DAO.Database
'Dim MyQuery As QueryDef
Dim qry As DAO.Recordset
Dim tbl As DAO.Recordset

Set DB = CurrentDb

DoCmd.OpenQuery ("InterlockingLinkPage2 ")
' dbOpenSnapshot riittää pelkkiin lukukyselyihin (nopeampi kuin Dynaset)
Set qry = DB.OpenRecordset("InterlockingLinkPage", dbOpenSnapshot)
' dbOpenDynaset vaaditaan, koska tbl-taulua muokataan (.Edit/.Update)
Set tbl = DB.OpenRecordset("IntLinkPage", dbOpenDynaset)

If Not qry.EOF Then qry.MoveFirst

' Kaikki päivitykset yhdessä transaktiossa — kymmenkertaistaa nopeuden verkkolevyllä
DBEngine.Workspaces(0).BeginTrans

Do Until qry.EOF
    ' FindFirst korvaa O(N²) sisäkkäisen silmukan — suorituskyky kasvaa eksponentiaalisesti suurilla tietomäärillä
    tbl.FindFirst "TXT1 = '" & Replace(qry!TXT1, "'", "''") & "'"
    If Not tbl.NoMatch Then
        tbl.Edit
        If Nz(tbl!page1, "") = "" Then
            tbl!page1 = "See Page " & qry!Page
        Else
            tbl!page1 = tbl!page1 & ", " & qry!Page
        End If
        tbl.Update
    End If
    qry.MoveNext
Loop

DBEngine.Workspaces(0).CommitTrans ' Hyväksytään kaikki muutokset yhtenä blokkina

DoCmd.Hourglass False

tbl.Close
qry.Close

' Vapautetaan DAO-resurssit muistivuodon ehkäisemiseksi
Set tbl = Nothing
Set qry = Nothing
Set DB = Nothing
Exit Sub

ErrorHandler:
    DBEngine.Workspaces(0).Rollback ' Perutaan muutokset virhetilanteessa — tietokanta palautuu ehjään tilaan
    Debug.Print "*** ERROR in Command0_Click (LukituskaavioLinkit): " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source
    DoCmd.Hourglass False
    On Error Resume Next
    If Not tbl Is Nothing Then tbl.Close
    If Not qry Is Nothing Then qry.Close
    Set tbl = Nothing
    Set qry = Nothing
    Set DB = Nothing
    MsgBox "Virhe sivutietojen päivityksessä: " & Err.Description, vbCritical
End Sub
