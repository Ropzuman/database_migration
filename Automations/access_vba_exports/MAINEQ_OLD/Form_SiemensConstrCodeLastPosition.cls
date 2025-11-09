Option Compare Database
Option Explicit

Private Sub LoppuKoodi1_DblClick(Cancel As Integer)
Dim endCode As String
Dim UusiTyyppi As String

endCode = Me.LoppuKoodi1
'If Right(Forms!motortypes!MotorType, 1) = "." Then
If InStr(1, Forms!motortypes!MotorType, ".") > 0 Then
    'UusiTyyppi = Left(Forms!motortypes!MotorType, InStr(1, Forms!motortypes!MotorType, ".") - 1) & endCode
    UusiTyyppi = Replace(Forms!motortypes!MotorType, ".", endCode)
    Forms!motortypes!MotorType = UusiTyyppi
    If IsLoaded("SiemensConstrCodeLastPosition") Then
    DoCmd.Close
    End If
Else
    If IsLoaded("SiemensConstrCodeLastPosition") Then
    DoCmd.Close
    End If
    MsgBox "Koodia ei muutettu!@Tyyppimerkiss� ei ollut pistett�.@Tarkista luettelosta!"
End If

End Sub

Private Sub LoppuKoodi2_DblClick(Cancel As Integer)
Dim endCode As String
Dim UusiTyyppi As String

endCode = Me.LoppuKoodi2
'If Right(Forms!motortypes!MotorType, 1) = "." Then
If InStr(1, Forms!motortypes!MotorType, ".") > 0 Then
    'UusiTyyppi = Left(Forms!motortypes!MotorType, InStr(1, Forms!motortypes!MotorType, ".") - 1) & endCode
    UusiTyyppi = Replace(Forms!motortypes!MotorType, ".", endCode)
    Forms!motortypes!MotorType = UusiTyyppi
    If IsLoaded("SiemensConstrCodeLastPosition") Then
    DoCmd.Close
    End If
Else
    If IsLoaded("SiemensConstrCodeLastPosition") Then
    DoCmd.Close
    End If
    MsgBox "Koodia ei muutettu!@Tyyppimerkiss� ei ollut pistett�.@Tarkista luettelosta!"
End If

End Sub
