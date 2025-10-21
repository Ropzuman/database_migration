VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FromTo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Command0_Click()
Dim DB As Database
Dim MyQuery As QueryDef
Dim qry As Recordset
Dim tbl As Recordset
Dim sSQL As String
Dim ssSQL As String
Dim qry2 As Recordset
Dim MyString As String

DoCmd.Hourglass True
Set DB = CurrentDb

DB.Execute "to_letter_update"

L = Chr(34)

Set qry = DB.OpenRecordset("FromInt")
Set qry2 = DB.OpenRecordset("ToInt")

Set tbl = DB.OpenRecordset("DATA9")

tbl.MoveFirst
Do Until tbl.EOF
    tbl.Edit
    tbl!to2 = ""
    tbl!to3 = ""
    tbl!to4 = ""
    tbl.Update
    tbl.MoveNext
Loop



edel = ""
qry.MoveFirst
Do Until qry.EOF

Do Until qry2.EOF
            If qry2!Text_4 = qry!Text_4 And qry2!out_4 = qry!out_4 Then
                qry2.Edit
             '   OnJo = InStr(tbl!Page, qry!Page)
              '  If OnJo = 0 Then
              If Len(qry2!to3) > 26 Then
              qry2!to4 = qry2!to4 & "," & Right(qry!DWG, 3)
              Else
                If Len(qry2!to2) > 26 Then
                    qry2!to3 = qry2!to3 & "," & Right(qry!DWG, 3)
                Else
                    qry2!to2 = qry2!to2 & "," & Right(qry!DWG, 3)
                End If
              End If
               qry2.Update
               
               
            End If
   qry2.MoveNext
Loop
qry2.MoveFirst
    qry.MoveNext
    
Loop

qry.MoveFirst
Do Until qry.EOF

Do Until qry2.EOF
            If qry2!Text_1 = qry!Text_1 And qry2!out_4 = qry!out_4 Then
                qry2.Edit
             '   OnJo = InStr(tbl!Page, qry!Page)
              '  If OnJo = 0 Then
            If Len(qry2!to3) > 26 Then
              qry2!to4 = qry2!to4 & "," & Right(qry!DWG, 3)
              Else
              
                If Len(qry2!to2) > 26 Then
                    qry2!to3 = qry2!to3 & "," & Right(qry!DWG, 3)
                Else
                    qry2!to2 = qry2!to2 & "," & Right(qry!DWG, 3)
                End If
              End If
               qry2.Update
               
               
            End If
   qry2.MoveNext
Loop
qry2.MoveFirst
    qry.MoveNext
    
Loop




tbl.MoveFirst
Do Until tbl.EOF
    tbl.Edit
        If Len(tbl!to2) > 2 Then
        tbl!to2 = Mid(tbl!to2, 2)
        End If
        If Len(tbl!to3) > 2 Then
        tbl!to3 = Mid(tbl!to3, 2)
        End If
        If Len(tbl!to4) > 2 Then
        tbl!to4 = Mid(tbl!to4, 2)
        End If
    tbl.Update
    tbl.MoveNext
Loop
       
qry2.MoveFirst
qry.MoveFirst 'lisätty jotta saataisiin myös eka

Do Until qry2.EOF

Do Until qry.EOF
            If qry!Text_4 = qry2!Text_4 And qry!out_4 = qry2!out_4 Then
                qry.Edit
                qry!Text_5 = qry2!Text_5
                qry.Update
            Else
            End If
               
            
   qry.MoveNext
Loop
qry.MoveFirst
    qry2.MoveNext

Loop


qry2.MoveFirst
Do Until qry2.EOF

Do Until qry.EOF
            If qry!Text_1 = qry2!Text_1 And qry!out_4 = qry2!out_4 Then
                qry.Edit
                qry!Text_5 = qry2!Text_5
                qry.Update
            Else
            End If
               
            
   qry.MoveNext
Loop
qry.MoveFirst
    qry2.MoveNext

Loop




DoCmd.Hourglass False

qry2.Close
qry.Close
DoCmd.OpenTable "DATA9"

End Sub

Private Sub Command12_Click()
'Vie tiedot kuviin
Dim EdTied As String
Dim Taul As Recordset
Dim oBlock As AcadBlockReference
Dim ATTRIB As Variant
Dim i As Integer
Set Taul = CurrentDb.OpenRecordset("SELECT * FROM DATA9 order by Path, DWG")
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
  Do While Not Taul.EOF
    If EdTied <> Taul.Fields("DWG") Then
      If Not oDOC Is Nothing Then
        oDOC.Close True
      End If
      Set oDOC = oACAD.Documents.Open(Taul.Fields("Path") & Taul.Fields("DWG") & ".dwg", False)
      EdTied = Taul.Fields("DWG")
    End If
    Set oBlock = oDOC.HandleToObject(Taul.Fields("HANDLE").VALUE)
    ATTRIB = oBlock.GetAttributes
    For i = 0 To UBound(ATTRIB)
      ATTRIB(i).TextString = Taul.Fields(ATTRIB(i).TagString)
    Next i
    Taul.MoveNext
  Loop
  oDOC.Save
  oDOC.Close
  Set oDOC = Nothing
  Set oACAD = Nothing
End Sub

Private Sub Command5_Click()
'Listataan tiedostot polussa
Dim FS As New FileSystemObject
Dim Tied As String
Dim i As Integer
If TPolku.VALUE <> "" Then
  If Right(TPolku.VALUE, 1) <> "\" Then
   TPolku.VALUE = TPolku.VALUE & "\"
  End If
Else
  MsgBox "Kirjoita ensin polku.", vbCritical, "Listaa tiedostot"
  Exit Sub
End If

If FS.FolderExists(TPolku.VALUE) Then
  For i = 1 To Lista.ListCount
    Lista.RemoveItem 1
  Next i
  Tied = Dir(TPolku & "*.dwg")
  Do While Tied <> ""
    Lista.AddItem Tied
    Tied = Dir()
  Loop
Else
  MsgBox "Virheellinen polku", vbCritical, "Listaa tiedostot"
End If
End Sub

Private Sub Command6_Click()
  NoudaTiedot True
 
End Sub
Private Sub NoudaTiedot(Listalta As Boolean)
Dim i As Integer
Dim Tied As String
'Yhdistetään ensin AutoCADiin
'Otetaan yhteys AutoCADiin
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
  If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
    MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Exit Sub
  End If
  On Error GoTo 0
  CurrentDb.Execute "DELETE * FROM DATA9"
  If Listalta Then
    For i = 1 To Lista.ItemsSelected.Count
        PoimiTiedot Lista.ItemData(Lista.ItemsSelected(i - 1))
    Next i
  Else
    Tied = Dir(TPolku & "*.dwg")
    Do While Tied <> ""
      PoimiTiedot Tied
      Tied = Dir()
    Loop
  End If
  Set oDOC = Nothing
  Set oACAD = Nothing
  DoCmd.OpenTable "DATA9"
End Sub
Private Sub PoimiTiedot(Tiedosto)
Dim FT(0) As Integer
Dim FD(0) As Variant
Dim Joukko As AcadSelectionSet
Dim i As Integer, j As Integer
Dim Taul As Recordset
Dim ATTRIB As Variant
Dim TNimi As String
  Set Taul = CurrentDb.OpenRecordset("DATA9")
  FT(0) = 2       'ObjecName
  FD(0) = FROMTOBLOCKS.VALUE
  Set oDOC = oACAD.Documents.Open(TPolku & Tiedosto)
  Set Joukko = oDOC.SelectionSets.Add("APUFROMTO")
  Joukko.Select acSelectionSetAll, , , FT, FD
  TNimi = Left(Tiedosto, Len(Tiedosto) - 4)
  For i = 0 To Joukko.Count - 1
    ATTRIB = Joukko(i).GetAttributes
    Taul.AddNew
    Taul.Fields("PATH") = TPolku
    Taul.Fields("DWG") = TNimi
    Taul.Fields("HANDLE") = Joukko(i).Handle
    Taul.Fields("BLOCK") = Joukko(i).Name
    
    For j = 0 To UBound(ATTRIB)
      Taul.Fields(ATTRIB(j).TagString) = ATTRIB(j).TextString
    Next j
    Taul.Update
  Next i
  oDOC.Close False
  
  Set oDOC = Nothing
End Sub

Private Sub Command7_Click()
  NoudaTiedot False
End Sub

Private Sub Form_Load()
  TPolku.VALUE = DLookup("[VALUE]", "SETTINGS", "SETTING = 'FROMTOPATH'")
  FROMTOBLOCKS.VALUE = DLookup("[VALUE]", "SETTINGS", "SETTING = 'FROMTOBLOCKS'")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  PaivitaPolku
End Sub


Private Sub FROMTOBLOCKS_AfterUpdate()
 CurrentDb.Execute "UPDATE SETTINGS SET SETTINGS.[VALUE] = '" & FROMTOBLOCKS.VALUE & "' WHERE (((SETTINGS.SETTING)='FROMTOBLOCKS'));"
End Sub

Private Sub TPolku_AfterUpdate()
  PaivitaPolku
End Sub
Private Sub PaivitaPolku()
 CurrentDb.Execute "UPDATE SETTINGS SET SETTINGS.[VALUE] = '" & TPolku.VALUE & "' WHERE (((SETTINGS.SETTING)='FROMTOPATH'));"
End Sub
