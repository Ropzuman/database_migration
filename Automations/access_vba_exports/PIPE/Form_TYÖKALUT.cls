VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_TY÷KALUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Dim oACAD As AcadApplication
Dim oDOC As AcadDocument
Dim Tied As String
Dim Joukko As AcadSelectionSet

Private Sub BPipelineToManval_Click()
'Putkilinjatiedot k‰siventtiileihin
   If MsgBox("Onko haluttu P&I kaavio auki AutoCADiss‰?", vbYesNo, "Pipeline To Manualvalve") = vbYes Then
    DoCmd.OpenForm "USysPipeToOther", , , , , , "MANUALVALVE_DATA" 'MANVALVEBLOCK
  End If
End Sub

Private Sub Command11_Click()
'Lue putkilinjat
  Tied = ""
  LueTiedot "PIPELINEDATA", "PIPELINE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Putkilinjojen luku valmis"
End Sub

Private Sub Command12_Click()
  ' Lis‰‰ putkilinjat taulukkoon
  CurrentDb.Execute "PIPELINES_APPEND"
  MsgBox "Putkilinjojen lis‰ys valmis"
End Sub

Private Sub Command13_Click()
' Lis‰‰ putkilinjat taulukosta tietokantaan
  If MsgBox("Varoitus! T‰m‰ ajaa kaavioissa olevat tiedot kannassa olevien tietojen p‰‰lle." & vbCrLf & "Haluatko jatkaa?", vbYesNo, "P‰ivit‰ tiedot") = vbYes Then
    CurrentDb.Execute "DELETE * From UsysPIPEUPD"
    CurrentDb.Execute "PIPELINES_UPDATE_START"
    CurrentDb.Execute "PIPELINES_UPDATE"
    MsgBox "Putkilinjojen p‰ivitys valmis"
  End If
End Sub

Private Sub Command18_Click()
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADiss‰?", vbYesNo, "FlowPickNo") = vbYes Then
    DoCmd.OpenForm "USysFlowpickNo"
  End If
End Sub

Private Sub Command19_Click()
'FROM/TO
  If MsgBox("Onko haluttu P&I kaavio auki AutoCADiss‰?", vbYesNo, "Pipeline From/To") = vbYes Then
    DoCmd.OpenForm "USysPipeFromTo"
  End If
End Sub

Private Sub Command21_Click()
  'Lue k‰siventtiilit
  Tied = ""
  LueTiedot "MANVALVEDATA", "MANUALVALVE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "K‰siventtiilien luku valmis"
End Sub

Private Sub Command22_Click()
  'Lis‰‰ k‰siventtiilit taulukkoon
  CurrentDb.Execute "MANVALVE_APPEND"
  MsgBox "K‰siventtiilien lis‰ys valmis"
End Sub

Private Sub Command23_Click()
'Lis‰‰ k‰siventtiilit taulukosta tietokantaan
  If MsgBox("Varoitus! T‰m‰ ajaa kaavioissa olevat tiedot kannassa olevien tietojen p‰‰lle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia n‰ihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "P‰ivit‰ tiedot") = vbYes Then
    CurrentDb.Execute "MANVALVE_UPDATE"
    MsgBox "K‰siventtiilien p‰ivitys valmis"
  End If
End Sub

Private Sub Command31_Click()
'Lue instrumenttiventtiilit
  Tied = ""
  LueTiedot "INSTVALVEDATA", "INSTVALVE_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Instrumenttiventtiilien luku valmis"
End Sub

Private Sub Command32_Click()
  'Lis‰‰ instrumenttiventtiilit taulukkoon
  CurrentDb.Execute "INSTVALVE_APPEND"
  MsgBox "Instrumenttiventtiilien lis‰ys valmis"
End Sub

Private Sub Command33_Click()
  'Lis‰‰ instrumenttiventtiilit taulukosta tietokantaan
  If MsgBox("Varoitus! T‰m‰ ajaa kaavioissa olevat tiedot kannassa olevien tietojen p‰‰lle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia n‰ihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "P‰ivit‰ tiedot") = vbYes Then
    CurrentDb.Execute "INSTVALVE_UPDATE"
    MsgBox "Instrumenttiventtiilien p‰ivitys valmis"
  End If
End Sub

Private Sub Command40_Click()
'Putkilinjatiedot piireihin
   If MsgBox("Onko haluttu P&I kaavio auki AutoCADiss‰?", vbYesNo, "Pipeline To Manualvalve") = vbYes Then
    DoCmd.OpenForm "USysPipeToOther", , , , , , "UP079D,UP008C,UP085D,UP08C,"
  End If
End Sub

Private Sub Command41_Click()
  'Lue kentt‰instrumenttit
  Tied = ""
  LueTiedot "FIELDINSTDATA", "FIELDINST_DATA"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Kentt‰instrumenttien luku valmis"
End Sub

Private Sub Command42_Click()
  'Lis‰‰ kentt‰instrumentit taulukkoon
  CurrentDb.Execute "FIELDINST_APPEND"
  MsgBox "Kentt‰instrumenttien lis‰ys valmis"
End Sub

Private Sub Command43_Click()
  'Lis‰‰ kentt‰instrumentit taulukosta tietokantaan
  If MsgBox("Varoitus! T‰m‰ ajaa kaavioissa olevat tiedot kannassa olevien tietojen p‰‰lle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia n‰ihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "P‰ivit‰ tiedot") = vbYes Then
    CurrentDb.Execute "FIELDINST_UPDATE"
    MsgBox "Kentt‰instrumenttien p‰ivitys valmis"
  End If
End Sub

Private Sub Command51_Click()
  'Lue ohjauspiirit
  Tied = ""
  LueTiedot "INSTLOOPDATA", "UP079D"
  Application.SysCmd acSysCmdClearStatus
  MsgBox "Instrumenttipiirien luku valmis"
End Sub

Private Sub Command52_Click()
  'Lis‰‰ ohjauspiirit taulukkoon
  CurrentDb.Execute "INSTLOOP_APPEND"
  MsgBox "Instrumenttipiirien lis‰ys valmis"
End Sub

Private Sub Command53_Click()
  'Lis‰‰ ohjauspiirit taulukosta tietokantaan
  If MsgBox("Varoitus! T‰m‰ ajaa kaavioissa olevat tiedot kannassa olevien tietojen p‰‰lle." & vbCrLf _
         & "Uudet code ja pipeline tiedot otetaan kaaviosta tuoduista tiedoista." & vbCrLf _
         & "Jos olet tehnyt muutoksia n‰ihin tietoihin kannassa, ne ylikirjoitetaan kaavion tiedoilla." & vbCrLf _
         & "Haluatko jatkaa?", vbYesNo, "P‰ivit‰ tiedot") = vbYes Then
    CurrentDb.Execute "INSTLOOP_UPDATE"
    MsgBox "Instrumenttipiirien p‰ivitys valmis"
  End If
End Sub

Private Sub LueTiedot(Taulukko As String, Blokki As String)
Dim i As Integer, j As Integer
Dim oBlock As AcadBlockReference
Dim Attrib As Variant
Dim Tiedot As Recordset
  CurrentDb.Execute "DELETE * FROM " & Taulukko
  Do
    If SeuraavaKuva = False Then
      Exit Do
    Else
      Set Tiedot = CurrentDb.OpenRecordset(Taulukko)
      PoimiJoukko Blokki
      For i = 0 To Joukko.Count - 1
         Application.SysCmd acSysCmdSetStatus, "Haetaan tietoja " & i + 1 & "/" & Joukko.Count
         Set oBlock = Joukko(i)
         Attrib = oBlock.GetAttributes
         Tiedot.AddNew
         Tiedot.Fields("PATH") = oDOC.Path
         Tiedot.Fields("ImpFileID") = oDOC.Name
         Tiedot.Fields("BLOCK") = oBlock.EffectiveName
         Tiedot.Fields("Handles") = oBlock.Handle
         For j = 0 To UBound(Attrib)
           If Attrib(j).TextString <> "" Then
             Tiedot.Fields("[" & Attrib(j).TagString & "]") = Attrib(j).TextString
           End If
         Next j
         Tiedot.Update
      Next i
      Joukko.Delete
      Set Joukko = Nothing
    End If
  Loop
  Set oACAD = Nothing
  Set Taul = Nothing
  Set oDOC = Nothing

End Sub
Private Function Yhdista() As Boolean
  On Error Resume Next
  Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdist‰‰ AutoCADiin
  If Err <> 0 Then 'K‰ynniss‰ olevaa AutoCADi‰ ei lˆytynyt
    MsgBox "K‰ynniss‰ olevaa AutoCADi‰ ei lˆytynyt!", vbCritical, "Virhe!"
    Set oACAD = Nothing
    Yhdista = False
    Exit Function
  Else
    Yhdista = True
  End If
  On Error GoTo 0
End Function
Private Function SeuraavaKuva() As Boolean
Dim Polku As String
   If oACAD Is Nothing Then
     If Yhdista = False Then
       SeuraavaKuva = False
       Exit Function
     End If
  End If
  Polku = Left(CurrentDb.Name, InStrRev(CurrentDb.Name, "\"))
  Polku = Polku & LPolku.Caption
    
  If Tied = "" Then
    Tied = Dir(Polku & "*.dwg")
  Else
    Tied = Dir
  End If
  If Tied = "" Then
    SeuraavaKuva = False
    Exit Function
  End If
  If Not oDOC Is Nothing Then
    oDOC.Close False
  End If
  Set oDOC = oACAD.Documents.Open(Polku & Tied)
  SeuraavaKuva = True
End Function
Private Sub PoimiJoukko(Nimet As String)
  Dim FilterType(0) As Integer
  Dim FilterData(0) As Variant
  Dim i As Integer
  Dim Blokit As Variant
  Dim Poista() As AcadEntity
  Dim L As Integer
  Dim EiPoisteta As Boolean
  
  
  Blokit = Split(Nimet, ",")
  For i = 0 To UBound(Blokit)
    Blokit(i) = Trim(Blokit(i))
  Next i
  
  FilterType(0) = 0
  FilterData(0) = "INSERT"
  For i = 0 To oDOC.SelectionSets.Count - 1
    If oDOC.SelectionSets(i).Name = "APUPICK" Then
      oDOC.SelectionSets(i).Delete
      Exit For
    End If
  Next i
  Set Joukko = oDOC.SelectionSets.Add("APUPICK")
  Joukko.Select acSelectionSetAll, , , FilterType, FilterData
  
  L = 0
  For j = 0 To Joukko.Count - 1
    If Joukko(j).EntityName = "AcDbBlockReference" Then
      EiPoisteta = False
      For i = 0 To UBound(Blokit)
        If UCase(Joukko(j).EffectiveName) = Blokit(i) Then
          EiPoisteta = True
          Exit For
        End If
      Next i
      If EiPoisteta = False Then
        ReDim Preserve Poista(L)
        Set Poista(L) = Joukko(j)
        L = L + 1
      End If
    End If
  Next j
  If L > 0 Then Joukko.RemoveItems Poista
  
End Sub
