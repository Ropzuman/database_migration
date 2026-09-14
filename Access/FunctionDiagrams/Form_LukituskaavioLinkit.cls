Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_LukituskaavioLinkit
' SOVELLUS : FunctionDiagrams ó Lukituskaaviolinkkien sivutieto
' KUVAUS   : Lukee InterlockingLinkPage-kyselyn tulokset ja p‰ivitt‰‰
'            IntLinkPage-taulun page1-sarakkeen pilkulla erotettuun
'            muotoon "See Page X, Y". T‰m‰n avulla lukituskaavion ristii
'            viittaukset saavat oikeat sivuviitteet piirustuksiin.
' DIPLE    : Option Explicit; Set tbl/qry/DB = Nothing muistivuodon ehk‰isyyn;
'            O(N≤) sis‰kk‰inen silmukka korvattu FindFirst-metodilla (suorit.)
' PƒIVITETTY: 2026-03-06
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
' dbOpenSnapshot riitt‰‰ pelkkiin lukukyselyihin (nopeampi kuin Dynaset)
Set qry = DB.OpenRecordset("InterlockingLinkPage", dbOpenSnapshot)
' dbOpenDynaset vaaditaan, koska tbl-taulua muokataan (.Edit/.Update)
Set tbl = DB.OpenRecordset("IntLinkPage", dbOpenDynaset)

If Not qry.EOF Then qry.MoveFirst

' Kaikki p‰ivitykset yhdess‰ transaktiossa ó kymmenkertaistaa nopeuden verkkolevyll‰
DBEngine.Workspaces(0).BeginTrans

Do Until qry.EOF
    ' FindFirst korvaa O(N≤) sis‰kk‰isen silmukan ó suorituskyky kasvaa eksponentiaalisesti suurilla tietom‰‰rill‰
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

DBEngine.Workspaces(0).CommitTrans ' Hyv‰ksyt‰‰n kaikki muutokset yhten‰ blokkina

DoCmd.Hourglass False

tbl.Close
qry.Close

' Vapautetaan DAO-resurssit muistivuodon ehk‰isemiseksi
Set tbl = Nothing
Set qry = Nothing
Set DB = Nothing
Exit Sub

ErrorHandler:
    DBEngine.Workspaces(0).Rollback ' Perutaan muutokset virhetilanteessa ó tietokanta palautuu ehj‰‰n tilaan
    Debug.Print "*** ERROR in Command0_Click (LukituskaavioLinkit): " & Err.Number & " - " & Err.Description
    Debug.Print "    Source: " & Err.Source
    DoCmd.Hourglass False
    On Error Resume Next
    If Not tbl Is Nothing Then tbl.Close
    If Not qry Is Nothing Then qry.Close
    Set tbl = Nothing
    Set qry = Nothing
    Set DB = Nothing
    MsgBox "Virhe sivutietojen p‰ivityksess‰: " & Err.Description, vbCritical
End Sub
