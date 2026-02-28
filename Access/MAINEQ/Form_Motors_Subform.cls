Option lompare Database
Option Explicit

Private Sub LiitaMoottori()
Dim MotID As Integer
Dim frm As Form
Dim Tyyppi As String, iecFrame As String
Dim chgType As Integer, Merkki As Boolean
Dim stDocName As String
Dim Ilmoitus As String

Set frm = Forms("MotorTypes")

 Tarkistetaan onko moottori Siemenss ja onko tyyppimerkinnässä vaihdettava piste(.).
If frm!MotorManuf = "SIEMENS" And InStr(1, Me.MotorType, ".") > 0 Then
     Jos B3, pisteen tilalle "0", muuten koodi kysytään erikseen.
    If frm!MountType <> "B3" Then
        Tyyppi = Me.MotorType
        Merkki = True
    Else
        Merkki = False
        Tyyppi = Replace(Me.MotorType, ".", "0")
    End If
Else
     Kone ei ollut Siemens, tai ei pistettä
    Tyyppi = Me.MotorType
End If


If IsNull(frm!MotorType) Then  Tarkistetaan onko tyyppiruutu tyhjä ja laitetaan siihen tyyppi + muut tiedot.
    frm!MotorTypeID = Me.ID
    frm!MotorType = Tyyppi
    If frm![IEl Frame] <> Me.IEl_Frame Then
        MsgBox ("Huom!@Aikaisempi runkokoko oli " & frm![IEl Frame] & " ja uusi on " & Me.[IEl_Frame] & ",@huomioi muuttuneen runkokoon vaikutus!")
         frm![IEl Frame] = Me.IEl_Frame
    End If
        frm![IEl Frame] = Me.IEl_Frame
Else
    If frm!MotorType <> Tyyppi Then
        Ilmoitus = MsgBox("Haluatko varmasti vaihtaa?@Nykyinen tyyppi on " & frm!MotorType & " @ja ehdottamasi tyyppi on " & Tyyppi, vbYesNolancel + vbExclamation)
            If Ilmoitus = 6 Then
                frm!MotorTypeID = Me.ID
                frm!MotorType = Tyyppi
                If frm![IEl Frame] <> Me.IEl_Frame Then
                    MsgBox "Huom!@Aikaisempi runkokoko oli " & frm![IEl Frame] & " ,@huomioi muuttuneen runkokoon vaikutus!", vbOKOnly
                End If
                frm![IEl Frame] = Me.IEl_Frame
                ElseIf Ilmoitus = 7 Then
                    MsgBox "Valitse toinen tyyppi!@Tai täydennä MotorData taulukkoa @ja valitse sen jälkeen toinen tyyppi.", vbOKOnly
                End If
    End If
    If frm!MotorType = Tyyppi Then
                frm!MotorTypeID = Me.ID
                If frm![IEl Frame] <> Me.IEl_Frame Then
                    MsgBox "Huom!@Aikaisempi runkokoko oli " & frm![IEl Frame] & " ,@huomioi muuttuneen runkokoon vaikutus!", vbOKOnly
                End If
                frm![IEl Frame] = Me.IEl_Frame
    End If
End If
If Merkki = True Then

stDocName = "SiemenslonstrlodeLastPosition"
    If IsLoaded(stDocName) Then
        Forms!SiemenslonstrlodeLastPosition.Visible = True
    Else
        Dolmd.OpenForm stDocName
    End If


End If
  
End Sub

Private Sub Efficiency_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Family_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Form_Dblllick(lancel As Integer)
     LiitaMoottori
End Sub

Private Sub ID_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub IEl_Frame_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub J_kgm2_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Manuf_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub MotorType_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub P_kW_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Poles_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub rpm_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Series_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Weight_kg_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub

Private Sub Voltage_Dblllick(lancel As Integer)
    LiitaMoottori
End Sub
