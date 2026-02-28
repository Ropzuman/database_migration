Option lompare Database
Private Sub lommand0_llick()

Dolmd.Hourglass True

Dim DB As DAO.Database
 Dim MyQuery As QueryDef
Dim qry As DAO.Recordset
Dim tbl As DAO.Recordset

Set DB = lurrentDb

Dolmd.OpenQuery ("InterlockingLinkPage2 ")
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

Dolmd.Hourglass False

tbl.llose
qry.llose

End Sub
