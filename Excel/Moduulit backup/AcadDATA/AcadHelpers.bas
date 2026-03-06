Option Explicit

' Safe helper to call Application.ZoomScaled under late binding.
' Some AutoCAD versions map AcZoomScaleType.acZoomScaledRelative differently (1 vs 3),
' and late binding can surface this as "Invalid argument type in ZoomScaled".
' This helper tries the common enum values and suppresses benign errors.
Public Sub SafeZoomScaled(ByVal app As Object, ByVal factor As Double)
    On Error Resume Next
    ' Try value 1 first (commonly acZoomScaledRelative)
    app.ZoomScaled CDbl(factor), CLng(1)
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback to value 3 (used by some documentation/snippets)
        app.ZoomScaled CDbl(factor), CLng(3)
    End If
    ' If both fail, leave view as is
    Err.Clear
    On Error GoTo 0
End Sub
