Option Explicit

' Updated 2025-10-26: 64-bit compatibility, improved error handling, performance optimizations
' Worksheet event handler for double-click: Locates and zooms to block in AutoCAD drawing
' Changes: Integer → Long (64-bit), early binding → late binding (compatibility)

' ============================================================================
' AutoCAD Constants - Required for Late Binding
' ============================================================================
' When using late binding (Object instead of AcadApplication, AcadEntity, etc.),
' the AutoCAD Type Library is not referenced, so built-in constants are not available.
' These must be manually defined with their numeric values.
' Source: Autodesk AutoCAD ActiveX/VBA Reference Documentation
' ============================================================================

Private Const acModelSpace As Long = 1              ' Model space (vs paper space)
Private Const acMax As Long = 3                     ' Maximize window


Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Excel.Range, Cancel As Boolean)
    Dim oACAD As Object ' AcadApplication (late binding for compatibility)
    Dim Entity As Object ' AcadEntity
    Dim Avataan As Boolean
    Dim OK As Boolean
    Dim Doku As String
    Dim Tiedosto As String
    Dim MinPoint As Variant
    Dim MaxPoint As Variant
    Dim i As Long ' Changed from Integer to Long for 64-bit compatibility
    
    On Error GoTo ErrHandler
    
    ' Validate that row has data
    If Cells(Target.Row, 1).Value = "" Then
        MsgBox "Ei kuvaa valitulla rivillä", vbInformation, "Etsi blokki"
        Cancel = True
        Exit Sub
    End If
    
    ' Connect to running AutoCAD instance
    On Error Resume Next
    Set oACAD = GetObject(, "AutoCAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutoCADiä ei löytynyt!", vbCritical, "Etsi blokki"
        Cancel = True
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
    ' Get document name from cell
    Doku = LCase(Cells(Target.Row, 2).Value) & ".dwg"
    
    ' Check if correct document is open
    If oACAD.Preferences.System.SingleDocumentMode Then
        ' SDI mode - only one document can be open
        If LCase(oACAD.ActiveDocument.Name) <> Doku Then
            If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKCancel, "Etsi blokki") = vbOK Then
                Avataan = True
            End If
        Else
            OK = True
        End If
    Else
        ' MDI mode - multiple documents can be open
        For i = 0 To oACAD.Documents.Count - 1
            If LCase(oACAD.Documents(i).Name) = Doku Then
                oACAD.Documents(i).Activate 
                OK = True
                Exit For
            End If
        Next i
        
        If Not OK Then
            If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKCancel, "Etsi blokki") = vbOK Then
                Avataan = True
            End If
        End If
    End If
    
    ' Open document if requested
    If Avataan Then
        Tiedosto = Cells(Target.Row, 1).Value & "\" & Doku
        
        On Error Resume Next
        oACAD.Documents.Open Tiedosto
        
        If Err.Number = 0 Then
            OK = True
        Else
            MsgBox "Virhe avattaessa dokumenttia: " & vbCrLf & Tiedosto, vbCritical, "Etsi blokki"
            OK = False
        End If
        On Error GoTo ErrHandler
    End If
    
    ' Zoom to entity if document is open
    If OK Then
        oACAD.ActiveDocument.ActiveSpace = acModelSpace
        Set Entity = oACAD.ActiveDocument.HandleToObject(Cells(Target.Row, 4).Value)
        
        ' Get bounding box and zoom to entity
        Entity.GetBoundingBox MinPoint, MaxPoint
    oACAD.ActiveDocument.WindowState = acMax
    oACAD.ZoomWindow MinPoint, MaxPoint
    ' Robust late-binding zoom-out: try enum value 1, then 3
    SafeZoomScaled oACAD, 0.5
        
        ' Activate AutoCAD window
        On Error Resume Next
        AppActivate oACAD.Caption, True
        On Error GoTo ErrHandler
    End If
    
Cleanup:
    ' Release COM objects
    Set Entity = Nothing
    Set oACAD = Nothing
    Cancel = True
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vbCrLf & Err.Description, vbCritical, "Etsi blokki"
    Resume Cleanup
End Sub