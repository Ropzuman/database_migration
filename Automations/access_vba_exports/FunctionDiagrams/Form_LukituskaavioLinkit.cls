VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LukituskaavioLinkit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Private Sub Command0_Click()

DoCmd.Hourglass True

Dim DB As Database
'Dim MyQuery As QueryDef
Dim qry As Recordset
Dim tbl As Recordset

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

End Sub
