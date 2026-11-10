Option Compare Database
Option Explicit

' Korvaa MotorTypes-lomakkeen tyyppimerkinn‰n ensimm‰isen pisteen valitulla
' p‰‰tekoodilla ja sulkee t‰m‰n lomakkeen.
Private Sub KorvaaPiste(ByVal endCode As String)
    Dim UusiTyyppi As String  ' Uusi tyyppimerkint‰
    Dim origType As String  ' Nykyinen tyyppimerkint‰

    If Not IsLoaded("MotorTypes") Then
        MsgBox "MotorTypes-lomake ei ole auki.", vbExclamation, "Huomio"
        Exit Sub
    End If
    origType = Nz(Forms!motortypes!MotorType, "")

    If InStr(1, origType, ".") > 0 Then
        ' Korvataan vain ensimm‰inen piste (Count:=1): est‰‰ usean pisteen vahingollisen korvauksen
        UusiTyyppi = Replace(origType, ".", endCode, 1, 1)
        Forms!motortypes!MotorType = UusiTyyppi
        DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
    Else
        DoCmd.Close acForm, "SiemensConstrCodeLastPosition"
        MsgBox "Koodia ei muutettu!" & vbCrLf & "Tyyppimerkinn‰ss‰ ei ollut pistett‰." & vbCrLf & "Tarkista luettelosta!", vbExclamation, "Huomio"
    End If
End Sub

Private Sub LoppuKoodi1_DblClick(Cancel As Integer)
    KorvaaPiste CStr(Nz(Me.LoppuKoodi1, ""))
End Sub

Private Sub LoppuKoodi2_DblClick(Cancel As Integer)
    KorvaaPiste CStr(Nz(Me.LoppuKoodi2, ""))
End Sub
