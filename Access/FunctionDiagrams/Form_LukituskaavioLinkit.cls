Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_LukituskaavioLinkit
' SOVELLUS : FunctionDiagrams — Lukituskaaviolinkkien sivutieto
' KUVAUS   : Lukee InterlockingLinkPage-kyselyn tulokset ja päivittää
'            IntLinkPage-taulun page1-sarakkeen pilkulla erotettuun
'            muotoon "See Page X, Y". Tämän avulla lukituskaavion ristii
'            viittaukset saavat oikeat sivuviitteet piirustuksiin.
' DIPLE    : Option Explicit; Set tbl/qry/DB = Nothing muistivuodon ehkäisyyn.
' PÄIVITETTY: 2026-03-03
'==========================================================================

Private Sub Command0_Click()

DoCmd.Hourglass True

Dim DB As DAO.Database
'Dim MyQuery As QueryDef
Dim qry As DAO.Recordset
Dim tbl As DAO.Recordset

Set DB = CurrentDb

DoCmd.OpenQuery ("InterlockingLinkPage2 ")
Set qry = DB.OpenRecordset("InterlockingLinkPage")
Set tbl = DB.OpenRecordset("IntLinkPage")

qry.MoveFirst
Do Until qry.EOF

Do Until tbl.EOF
            If tbl!TXT1 = qry!TXT1 Then
                tbl.Edit
                    If tbl!page1 = "" Then
                    tbl!page1 = "See Page " & qry!Page
                    Else
                    tbl!page1 = tbl!page1 & ", " & qry!Page
                    End If
               tbl.Update
            End If
   tbl.MoveNext
Loop
tbl.MoveFirst
    qry.MoveNext
Loop

DoCmd.Hourglass False

tbl.Close
qry.Close

' Vapautetaan DAO-resurssit muistivuodon ehkäisemiseksi
Set tbl = Nothing
Set qry = Nothing
Set DB = Nothing

End Sub
