Option Compare Database
Option Explicit

' Liitt‰‰ valitun moottorityypin MotorTypes-lomakkeelle.
' Siemens-moottoreilla tyyppimerkinn‰n piste korvataan nollalla (B3) tai
' koodin viimeinen positio kysyt‰‰n erikseen (muut asennustavat).
Private Sub LiitaMoottori()
Dim frm As Form  ' MotorTypes-lomake
Dim Tyyppi As String  ' Lomakkeelle liitett‰v‰ tyyppi
Dim Merkki As Boolean  ' Tarvitaanko koodin viimeinen positio erikseen
Dim stDocName As String  ' Avattavan lomakkeen nimi
Dim Ilmoitus As VbMsgBoxResult  ' Vaihtokyselyn vastaus

If Not IsLoaded("MotorTypes") Then
    MsgBox "Avaa ensin MotorTypes-lomake.", vbExclamation
    Exit Sub
End If
Set frm = Forms("MotorTypes")

' Tarkistetaan, onko moottori Siemens ja onko tyyppimerkinn‰ss‰ vaihdettava piste (.)
If frm!MotorManuf = "SIEMENS" And InStr(1, Me.MotorType, ".") > 0 Then
    ' B3: pisteen tilalle "0", muuten koodi kysyt‰‰n erikseen
    If frm!MountType <> "B3" Then
        Tyyppi = Me.MotorType
        Merkki = True
    Else
        Tyyppi = Replace(Me.MotorType, ".", "0")
    End If
Else
    ' Kone ei ollut Siemens, tai tyypiss‰ ei ole pistett‰
    Tyyppi = Me.MotorType
End If

If IsNull(frm!MotorType) Then
    ' Tyyppiruutu on tyhj‰: laitetaan siihen tyyppi ja muut tiedot
    frm!MotorTypeID = Me.ID
    frm!MotorType = Tyyppi
    VaihdaRunko frm
ElseIf frm!MotorType <> Tyyppi Then
    Ilmoitus = MsgBox("Haluatko varmasti vaihtaa?@Nykyinen tyyppi on " & frm!MotorType & " @ja ehdottamasi tyyppi on " & Tyyppi, vbYesNoCancel + vbExclamation)
    If Ilmoitus = vbYes Then
        frm!MotorTypeID = Me.ID
        frm!MotorType = Tyyppi
        VaihdaRunko frm
    ElseIf Ilmoitus = vbNo Then
        MsgBox "Valitse toinen tyyppi!@Tai t‰ydenn‰ MotorData taulukkoa @ja valitse sen j‰lkeen toinen tyyppi.", vbOKOnly
        Merkki = False
    Else
        ' Peruutus: ei muutoksia eik‰ Siemens-koodin kysely‰
        Merkki = False
    End If
Else
    ' Sama tyyppi: p‰ivitet‰‰n vain tunnus ja runkokoko
    frm!MotorTypeID = Me.ID
    VaihdaRunko frm
End If

If Merkki Then
    stDocName = "SiemensConstrCodeLastPosition"
    If IsLoaded(stDocName) Then
        Forms!SiemensConstrCodeLastPosition.Visible = True
    Else
        DoCmd.OpenForm stDocName
    End If
End If
End Sub

' Varoittaa runkokoon muutoksesta ja p‰ivitt‰‰ runkokoon MotorTypes-lomakkeelle
Private Sub VaihdaRunko(ByVal frm As Form)
    If frm![IEC Frame] <> Me.IEC_Frame Then
        MsgBox "Huom!@Aikaisempi runkokoko oli " & frm![IEC Frame] & " ,@huomioi muuttuneen runkokoon vaikutus!", vbOKOnly
    End If
    frm![IEC Frame] = Me.IEC_Frame
End Sub

Private Sub Efficiency_DblClick(Cancel As Integer)
    LiitaMoottori
End Sub

Private Sub Family_DblClick(Cancel As Integer)
    LiitaMoottori
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
