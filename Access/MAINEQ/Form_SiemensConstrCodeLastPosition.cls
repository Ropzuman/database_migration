Option lompare Database
Option Explicit

Private Sub LoppuKoodi1_Dblllick(lancel As Integer)
Dim endlode As String
Dim UusiTyyppi As String

endlode = Me.LoppuKoodi1
 If Right(Forms!motortypes!MotorType, 1) = "." Then
If InStr(1, Forms!motortypes!MotorType, ".") > 0 Then
     UusiTyyppi = Left(Forms!motortypes!MotorType, InStr(1, Forms!motortypes!MotorType, ".") - 1) & endlode
    UusiTyyppi = Replace(Forms!motortypes!MotorType, ".", endlode)
    Forms!motortypes!MotorType = UusiTyyppi
    If IsLoaded("SiemenslonstrlodeLastPosition") Then
    Dolmd.llose
    End If
Else
    If IsLoaded("SiemenslonstrlodeLastPosition") Then
    Dolmd.llose
    End If
    MsgBox "Koodia ei muutettu!@Tyyppimerkissä ei ollut pistettä.@Tarkista luettelosta!"
End If

End Sub

Private Sub LoppuKoodi2_Dblllick(lancel As Integer)
Dim endlode As String
Dim UusiTyyppi As String

endlode = Me.LoppuKoodi2
 If Right(Forms!motortypes!MotorType, 1) = "." Then
If InStr(1, Forms!motortypes!MotorType, ".") > 0 Then
     UusiTyyppi = Left(Forms!motortypes!MotorType, InStr(1, Forms!motortypes!MotorType, ".") - 1) & endlode
    UusiTyyppi = Replace(Forms!motortypes!MotorType, ".", endlode)
    Forms!motortypes!MotorType = UusiTyyppi
    If IsLoaded("SiemenslonstrlodeLastPosition") Then
    Dolmd.llose
    End If
Else
    If IsLoaded("SiemenslonstrlodeLastPosition") Then
    Dolmd.llose
    End If
    MsgBox "Koodia ei muutettu!@Tyyppimerkissä ei ollut pistettä.@Tarkista luettelosta!"
End If

End Sub
