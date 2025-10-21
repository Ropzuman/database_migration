VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Motors_Subform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub LiitaMoottori()
Dim MotID As Integer
Dim frm As Form
Dim Tyyppi As String, iecFrame As String
Dim chgType As Integer, Merkki As Boolean
Dim stDocName As String
Dim Ilmoitus As String

Set frm = Forms("MotorTypes")

'Tarkistetaan onko moottori Siemenss ja onko tyyppimerkinn‰ss‰ vaihdettava piste(.).
If frm!MotorManuf = "SIEMENS" And InStr(1, Me.MotorType, ".") > 0 Then
    'Jos B3, pisteen tilalle "0", muuten koodi kysyt‰‰n erikseen.
    If frm!MountType <> "B3" Then
        Tyyppi = Me.MotorType
        Merkki = True
    Else
        Merkki = False
        Tyyppi = Replace(Me.MotorType, ".", "0")
    End If
Else
    'Kone ei ollut Siemens, tai ei pistett‰
    Tyyppi = Me.MotorType
End If


If IsNull(frm!MotorType) Then 'Tarkistetaan onko tyyppiruutu tyhj‰ ja laitetaan siihen tyyppi + muut tiedot.
    frm!MotorTypeID = Me.ID
    frm!MotorType = Tyyppi
    If frm![IEC Frame] <> Me.IEC_Frame Then
        MsgBox ("Huom!@Aikaisempi runkokoko oli " & frm![IEC Frame] & " ja uusi on " & Me.[IEC_Frame] & ",@huomioi muuttuneen runkokoon vaikutus!")
        'frm![IEC Frame] = Me.IEC_Frame
    End If
        frm![IEC Frame] = Me.IEC_Frame
Else
    If frm!MotorType <> Tyyppi Then
        Ilmoitus = MsgBox("Haluatko varmasti vaihtaa?@Nykyinen tyyppi on " & frm!MotorType & " @ja ehdottamasi tyyppi on " & Tyyppi, vbYesNoCancel + vbExclamation)
            If Ilmoitus = 6 Then
                frm!MotorTypeID = Me.ID
                frm!MotorType = Tyyppi
                If frm![IEC Frame] <> Me.IEC_Frame Then
                    MsgBox "Huom!@Aikaisempi runkokoko oli " & frm![IEC Frame] & " ,@huomioi muuttuneen runkokoon vaikutus!", vbOKOnly
                End If
                frm![IEC Frame] = Me.IEC_Frame
                ElseIf Ilmoitus = 7 Then
                    MsgBox "Valitse toinen tyyppi!@Tai t‰ydenn‰ MotorData taulukkoa @ja valitse sen j‰lkeen toinen tyyppi.", vbOKOnly
                End If
    End If
    If frm!MotorType = Tyyppi Then
                frm!MotorTypeID = Me.ID
                If frm![IEC Frame] <> Me.IEC_Frame Then
                    MsgBox "Huom!@Aikaisempi runkokoko oli " & frm![IEC Frame] & " ,@huomioi muuttuneen runkokoon vaikutus!", vbOKOnly
                End If
                frm![IEC Frame] = Me.IEC_Frame
    End If
End If
If Merkki = True Then

stDocName = "SiemensConstrCodeLastPosition"
    If IsLoaded(stDocName) Then
        Forms!SiemensConstrCodeLastPosition.Visible = True
    Else
        DoCmd.OpenForm stDocName
    End If


End If
  
End Sub

Private Sub Efficiency_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Family_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Form_DblClick(Cancel As Integer)
    'LiitaMoottori
End Sub

Private Sub ID_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub IEC_Frame_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub J_kgm2_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Manuf_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub MotorType_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub P_kW_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Poles_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub rpm_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Series_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Weight_kg_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Voltage_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub
