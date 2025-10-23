VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Venttiiliblokkien vaihto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub Command0_Click()
'Vaihtaa PI-Kaavioista venttiilit Araucon malleihin.
Dim Taulukko As Recordset
Dim DWG As String
Dim Handle As String
Dim Polku As String
Dim oACAD As AcadApplication
Dim oEntity As AcadEntity
Dim oBlock As AcadBlockReference
Dim UusiBlock As AcadBlockReference
Dim Attribuutit As Variant
Dim UAttrib As Variant
Dim OK As Boolean
Dim i As Integer, j As Integer
Dim BPolku As String
Dim BNimi As String
Dim Koodi As Integer
Dim Rotation As Double

BPolku = "L:\Projekti\ITATA\Process\Flowsheets\BLOCKS\"
On Error Resume Next
Set oACAD = GetObject(, "AutoCAD.Application") 'Koitetaan yhdistää AutoCADiin
If Err <> 0 Then 'Käynnissä olevaa AutoCADiä ei löytynyt
  MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!" & vbCrLf & "Avaa Autocad ensin.", vbCritical, "Vaihda Blokit"
  Set oACAD = Nothing
  GoTo Loppu
End If
On Error GoTo 0

Set Taulukko = CurrentDb.OpenRecordset("SELECT * FROM MANUALVALVES WHERE (DwgFileID Is Not Null AND DwgHandle Is Not Null AND Deleted = False) Order by DwgFileID")

Do While Not Taulukko.EOF
  If UCase(Nz(Taulukko.Fields("DwgFileID").Value)) <> DWG Then
    DWG = UCase(Nz(Taulukko.Fields("DwgFileID").Value))
    OK = False
    Select Case DWG
      Case "POR1050970", "POR1051314", "POR1050965", "POR1050966", "POR1050967", "POR1050968", "POR1050969"
        Polku = "L:\Projekti\ITATA\Process\Flowsheets\PINE MACHINE 1\"
      Case "POR1051315", "POR1050976", "POR1050977", "POR1050978", "POR1050979", "POR1050980", "POR1050981"
        Polku = "L:\Projekti\ITATA\Process\Flowsheets\EUCA MACHINE 2\"
      Case Else
        Polku = ""
    End Select
    If Dir(Polku & DWG & ".dwg") = "" Then
      OK = False
    Else
      oACAD.Documents.Open Polku & DWG & ".dwg"
      OK = True
    End If
  End If
  Handle = UCase(Nz(Taulukko.Fields("DwgHandle").Value))
  If Handle <> "" And OK And Nz(Taulukko.Fields("Code").Value) <> "" Then
    Set oBlock = oACAD.ActiveDocument.HandleToObject(Handle)
    Attribuutit = oBlock.GetAttributes
    If UBound(Attribuutit) > 1 Then
      Koodi = CInt(Val(Nz(Taulukko.Fields("Code").Value)))
      Select Case Koodi
        Case 2
          BNimi = "V11"
        Case 3
          If oBlock.Name = "UP051" Then
            BNimi = "V06"
          Else
            BNimi = "V05"
          End If
        Case 4
          If oBlock.Name = "UP051" Then
            BNimi = "V23"
          Else
            BNimi = "V22"
          End If
        Case 5, 6 'Check valve
          BNimi = "V09"
        Case 12
          BNimi = "V17"
        Case 1, 8, 11
          BNimi = ""
        Case Else
          BNimi = ""
      End Select
      If BNimi <> "" Then
        Rotation = oBlock.Rotation
        If oBlock.XScaleFactor <> 1 Then
          If Rotation > 3 Then
            Rotation = Rotation - 3.141592654
          Else
            Rotation = Rotation + 3.141592654
          End If
        End If
        Set UusiBlock = oACAD.ActiveDocument.ModelSpace.InsertBlock(oBlock.InsertionPoint, BPolku & BNimi & ".dwg", 1, 1, 1, Rotation)
        UusiBlock.Layer = oBlock.Layer
        UAttrib = UusiBlock.GetAttributes
        For i = LBound(Attribuutit) To UBound(Attribuutit)
          For j = LBound(UAttrib) To UBound(UAttrib)
            If UCase(Attribuutit(i).TagString) = UCase(UAttrib(j).TagString) Then
              UAttrib(j).TextString = Attribuutit(i).TextString
              UAttrib(j).Alignment = Attribuutit(i).Alignment
'              UAttrib(j).TextAlignmentPoint = Attribuutit(i).TextAlignmentPoint
              UAttrib(j).Rotation = Attribuutit(i).Rotation
              UAttrib(j).InsertionPoint = Attribuutit(i).InsertionPoint
              Dim Attrib As AcadAttributeReference
 '             attrib.Backward
              Exit For
            End If
          Next j
        Next i
        Taulukko.Edit
        Taulukko.Fields("DwgHandle").Value = UusiBlock.Handle
        Taulukko.Fields("DwgBlock").Value = BNimi
        Taulukko.Update
        oBlock.Delete
      End If
    End If
  End If
  Taulukko.MoveNext
Loop
MsgBox "Valmis"
Loppu:
Set UusiBlock = Nothing
Set oBlock = Nothing
Set oACAD = Nothing
Set Taulukko = Nothing
End Sub
