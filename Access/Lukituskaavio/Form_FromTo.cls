Option Compare Database
Option Explicit

Private Sub Command0_Click()
Dim DB As DAO.Database
Dim MyQuery As DAO.QueryDef
Dim qry As DAO.Recordset
Dim tbl As DAO.Recordset
Dim sSQL As String
Dim ssSQL As String
Dim qry2 As DAO.Recordset
Dim MyString As String
Dim L As String
Dim edel As String

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
              qry2!to4 = qry2!to4 & "," & Right$(qry!DWG, 3)
              Else
                If Len(qry2!to2) > 26 Then
                    qry2!to3 = qry2!to3 & "," & Right$(qry!DWG, 3)
                Else
                    qry2!to2 = qry2!to2 & "," & Right$(qry!DWG, 3)
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
              qry2!to4 = qry2!to4 & "," & Right$(qry!DWG, 3)
              Else
              
                If Len(qry2!to2) > 26 Then
                    qry2!to3 = qry2!to3 & "," & Right$(qry!DWG, 3)
                Else
                    qry2!to2 = qry2!to2 & "," & Right$(qry!DWG, 3)
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
        tbl!to2 = Mid$(tbl!to2, 2)
        End If
        If Len(tbl!to3) > 2 Then
        tbl!to3 = Mid$(tbl!to3, 2)
        End If
        If Len(tbl!to4) > 2 Then
        tbl!to4 = Mid$(tbl!to4, 2)
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
' Vie DATA9-taulun tiedot AutoCAD-piirustuksissa oleviin blokkeihin
Dim EdTied As String
Dim Taul As DAO.Recordset
Dim oBlock As Object  ' AcadBlockReference — myöhäinen sidonta 64-bit-yhteensopivuuden vuoksi
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
Dim FS As Object
Dim Tied As String
Dim i As Integer
Set FS = CreateObject("Scripting.FileSystemObject")
If Me.TPolku.Value <> "" Then
  If Right$(Me.TPolku.Value, 1) <> "\" Then
   Me.TPolku.Value = Me.TPolku.Value & "\"
  End If
Else
  MsgBox "Kirjoita ensin polku.", vbCritical, "Listaa tiedostot"
  Exit Sub
End If

If FS.FolderExists(Me.TPolku.Value) Then
  For i = 1 To Me.Lista.ListCount
    Me.Lista.RemoveItem 1
  Next i
  Tied = Dir(Me.TPolku & "*.dwg")
  Do While Tied <> ""
    Me.Lista.AddItem Tied
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
  Dim i As Long
  Dim Tied As String
  Dim DB As DAO.Database
  Dim TransaktioAuki As Boolean
  Dim TuontiValmis As Boolean
  Dim virheKuvaus As String

  On Error GoTo ErrorHandler
  Set oACAD = GetObject(, "AutoCAD.Application")
  If oACAD Is Nothing Then Err.Raise vbObjectError + 2200, "NoudaTiedot", "Käynnissä olevaa AutoCADia ei löytynyt."

  Set DB = CurrentDb
  DBEngine.Workspaces(0).BeginTrans
  TransaktioAuki = True
  DB.Execute "DELETE * FROM DATA9", dbFailOnError
  If Listalta Then
    For i = 1 To Me.Lista.ItemsSelected.Count
      PoimiTiedot DB, CStr(Me.Lista.ItemData(Me.Lista.ItemsSelected(i - 1)))
    Next i
  Else
    Tied = Dir(Me.TPolku & "*.dwg")
    Do While Tied <> ""
      PoimiTiedot DB, Tied
      Tied = Dir()
    Loop
  End If

  DBEngine.Workspaces(0).CommitTrans
  TransaktioAuki = False
  TuontiValmis = True
  Set oDOC = Nothing
  Set oACAD = Nothing
  Set DB = Nothing
  DoCmd.OpenTable "DATA9"
  Exit Sub

ErrorHandler:
  virheKuvaus = Err.Description
  On Error Resume Next
  If TransaktioAuki Then DBEngine.Workspaces(0).Rollback
  Set oDOC = Nothing
  Set oACAD = Nothing
  Set DB = Nothing
  On Error GoTo 0
  If TuontiValmis Then
    MsgBox "Tietojen tuonti tallennettiin, mutta DATA9-taulun avaaminen epäonnistui: " & virheKuvaus, vbExclamation
  Else
    MsgBox "Tietojen tuonti epäonnistui. DATA9-taulu palautettiin ennalleen: " & virheKuvaus, vbCritical
  End If
End Sub

Private Sub PoimiTiedot(ByVal DB As DAO.Database, ByVal Tiedosto As String)
  Dim FT(0) As Integer
  Dim FD(0) As Variant
  Dim TiedostonDOC As Object
  Dim Joukko As Object  ' AutoCAD-valintajoukko, myöhäinen sidonta
  Dim i As Long, j As Long
  Dim Taul As DAO.Recordset
  Dim ATTRIB As Variant
  Dim TNimi As String
  Dim virheNumero As Long
  Dim virheKuvaus As String

  On Error GoTo ErrorHandler
  Set Taul = DB.OpenRecordset("DATA9", dbOpenDynaset)
  FT(0) = 2       ' DXF-ryhmäkoodi: objektin nimi
  FD(0) = Me.FROMTOBLOCKS.Value
  Set TiedostonDOC = oACAD.Documents.Open(Me.TPolku & Tiedosto)
  Set Joukko = TiedostonDOC.SelectionSets.Add("APUFROMTO")
  Joukko.Select AC_SELECTION_SET_ALL, , , FT, FD
  TNimi = Left$(Tiedosto, Len(Tiedosto) - 4)
  For i = 0 To Joukko.Count - 1
    ATTRIB = Joukko(i).GetAttributes
    Taul.AddNew
    Taul.Fields("PATH") = Me.TPolku
    Taul.Fields("DWG") = TNimi
    Taul.Fields("HANDLE") = Joukko(i).Handle
    Taul.Fields("BLOCK") = Joukko(i).Name
    For j = 0 To UBound(ATTRIB)
      Taul.Fields(ATTRIB(j).TagString) = ATTRIB(j).TextString
    Next j
    Taul.Update
  Next i

  Taul.Close
  Set Taul = Nothing
  Joukko.Delete
  Set Joukko = Nothing
  TiedostonDOC.Close False
  Set TiedostonDOC = Nothing
  Exit Sub

ErrorHandler:
  virheNumero = Err.Number
  virheKuvaus = Err.Description
  On Error Resume Next
  If Not Taul Is Nothing Then Taul.Close
  Set Taul = Nothing
  If Not Joukko Is Nothing Then Joukko.Delete
  Set Joukko = Nothing
  If Not TiedostonDOC Is Nothing Then TiedostonDOC.Close False
  Set TiedostonDOC = Nothing
  On Error GoTo 0
  Err.Raise virheNumero, "PoimiTiedot", virheKuvaus
End Sub

Private Sub Command7_Click()
  NoudaTiedot False
End Sub

Private Sub Form_Load()
  Me.TPolku.Value = DLookup("[VALUE]", "SETTINGS", "SETTING = 'FROMTOPATH'")
  Me.FROMTOBLOCKS.Value = DLookup("[VALUE]", "SETTINGS", "SETTING = 'FROMTOBLOCKS'")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  PaivitaPolku
End Sub


Private Sub FROMTOBLOCKS_AfterUpdate()
 CurrentDb.Execute "UPDATE SETTINGS SET SETTINGS.[VALUE] = '" & Me.FROMTOBLOCKS.Value & "' WHERE (((SETTINGS.SETTING)='FROMTOBLOCKS'));"
End Sub

Private Sub TPolku_AfterUpdate()
  PaivitaPolku
End Sub
Private Sub PaivitaPolku()
 CurrentDb.Execute "UPDATE SETTINGS SET SETTINGS.[VALUE] = '" & Me.TPolku.Value & "' WHERE (((SETTINGS.SETTING)='FROMTOPATH'));"
End Sub
