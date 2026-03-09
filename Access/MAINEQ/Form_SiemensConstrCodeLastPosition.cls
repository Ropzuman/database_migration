Option Compare Database
Option Explicit

Private Sub LoppuKoodi1_DblClick(Cancel As Integer)
    Dim endCode As String
    Dim UusiTyyppi As String
    Dim origType As String

    endCode = IIf(IsNull(Me.LoppuKoodi1), "", Me.LoppuKoodi1)
    origType = IIf(IsNull(Forms!motortypes!MotorType), "", Forms!motortypes!MotorType)

    If InStr(1, origType, ".") > 0 Then
        ' Korvataan VAIN ENSIMMÄINEN piste (Count:=1) — estää usean pisteen vahingollisen korvauksen
        UusiTyyppi = Replace(origType, ".", endCode, 1, 1)
        Forms!motortypes!MotorType = UusiTyyppi
        If IsLoaded("SiemensConstrCodeLastPosition") Then
            DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        End If
    Else
        If IsLoaded("SiemensConstrCodeLastPosition") Then
            DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        End If
        MsgBox "Koodia ei muutettu!" & vbCrLf & "Tyyppimerkissä ei ollut pistettä." & vbCrLf & "Tarkista luettelosta!", vbExclamation, "Huomio"
    End If
End Sub

Private Sub LoppuKoodi2_DblClick(Cancel As Integer)
    Dim endCode As String
    Dim UusiTyyppi As String
    Dim origType As String

    endCode = IIf(IsNull(Me.LoppuKoodi2), "", Me.LoppuKoodi2)
    origType = IIf(IsNull(Forms!motortypes!MotorType), "", Forms!motortypes!MotorType)

    If InStr(1, origType, ".") > 0 Then
        ' Korvataan VAIN ENSIMMÄINEN piste (Count:=1) — estää usean pisteen vahingollisen korvauksen
        UusiTyyppi = Replace(origType, ".", endCode, 1, 1)
        Forms!motortypes!MotorType = UusiTyyppi
        If IsLoaded("SiemensConstrCodeLastPosition") Then
            DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        End If
    Else
        If IsLoaded("SiemensConstrCodeLastPosition") Then
            DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        End If
        MsgBox "Koodia ei muutettu!" & vbCrLf & "Tyyppimerkissä ei ollut pistettä." & vbCrLf & "Tarkista luettelosta!", vbExclamation, "Huomio"
    End If
End Sub
