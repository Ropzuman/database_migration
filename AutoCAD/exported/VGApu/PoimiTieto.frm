VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PoimiTieto 
   Caption         =   "Pick MotorNumber"
   ClientHeight    =   480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1665
   OleObjectBlob   =   "PoimiTieto.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PoimiTieto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Vaihda_Click()
Dim Valinta As Object
Dim IPiste As Variant
Dim Attribuutit As Variant
Dim i As Integer, j As Integer
Dim Motor As String
  
  On Error Resume Next
  
  Me.Hide
  'Pyydet‰‰n k‰ytt‰j‰‰ valitsemaan Moottoriblokki
  Do
    ActiveDocument.Utility.GetEntity Valinta, IPiste, "Select Motor..."
    If Err = 0 Then 'Valinta osui objektiin tai ei painettu Esci‰
      If Valinta.ObjectName = "AcDbBlockReference" Then  'Valittiin Blokki
         If UCase(Valinta.Name) = "MOTOR" Then
           Attribuutit = Valinta.GetAttributes
           Motor = Attribuutit(0).TextString
           ActiveDocument.Utility.Prompt vbCrLf & "Selected Motor: " & Motor & vbCrLf
         End If
      End If
        'Pyydet‰‰n k‰ytt‰j‰‰ valitsemaan piiri
        ActiveDocument.Utility.GetEntity Valinta, IPiste, "Select Instrument..."
        If Err = 0 Then 'Valinta osui objektiin tai ei painettu Esci‰
          If Valinta.ObjectName = "AcDbBlockReference" Then  'Valittiin Blokki
             If UCase(Valinta.Name) = "UP079" Then
               Attribuutit = Valinta.GetAttributes
               For i = 0 To UBound(Attribuutit)
                 If Attribuutit(i).TagString = "SDPOS" Then
                   Attribuutit(i).TextString = Motor
                   Exit For
                 End If
               Next i
             End If
          End If
        Else
          Err.Clear
        End If
    Else
      Exit Do
    End If
  Loop
  
  Err.Clear
  On Error GoTo 0
  Set Valinta = Nothing
  Me.Show False
End Sub
