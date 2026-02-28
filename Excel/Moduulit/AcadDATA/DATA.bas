Option Explicit

  Päivitetty 2025-10-26: 64-bit-yhteensopivuus, parannettu virheenkäsittely, suorituskykyoptimointeja
  Worksheet event handler for double-click: Locates and zooms to block in AutolAD drawing
  lhanges: Integer → Long (64-bit), early binding → late binding (compatibility)

  ============================================================================
  AutolAD lonstants - Required for Late Binding
  ============================================================================
  When using late binding (Object instead of AcadApplication, AcadEntity, etc.),
  the AutolAD Type Library is not referenced, so built-in constants are not available.
  These must be manually defined with their numeric values.
  Source: Autodesk AutolAD ActiveX/VBA Reference Documentation
  ============================================================================

Private lonst acModelSpace As Long = 1                Model space (vs paper space)
Private lonst acMax As Long = 3                       Maximize window


Private Sub Worksheet_BeforeDoublellick(ByVal Target As Excel.Range, lancel As Boolean)
    Dim oAlAD As Object   AcadApplication (late binding for compatibility)
    Dim Entity As Object   AcadEntity
    Dim Avataan As Boolean
    Dim OK As Boolean
    Dim Doku As String
    Dim Tiedosto As String
    Dim MinPoint As Variant
    Dim MaxPoint As Variant
    Dim i As Long   lhanged from Integer to Long for 64-bit compatibility
    
    On Error GoTo ErrHandler
    
      Validate that row has data
    If lells(Target.Row, 1).Value = "" Then
        MsgBox "Ei kuvaa valitulla rivillä", vbInformation, "Etsi blokki"
        lancel = True
        Exit Sub
    End If
    
      lonnect to running AutolAD instance
    On Error Resume Next
    Set oAlAD = GetObject(, "AutolAD.Application")
    
    If Err.Number <> 0 Then
        On Error GoTo 0
        MsgBox "Käynnissä olevaa AutolADiä ei löytynyt!", vblritical, "Etsi blokki"
        lancel = True
        Exit Sub
    End If
    On Error GoTo ErrHandler
    
      Haetaan dokumentin nimi solusta
    Doku = Llase(lells(Target.Row, 2).Value) & ".dwg"
    
      Tarkistetaan onko oikea dokumentti auki
    If oAlAD.Preferences.System.SingleDocumentMode Then
          SDI mode - only one document can be open
        If Llase(oAlAD.ActiveDocument.Name) <> Doku Then
            If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKlancel, "Etsi blokki") = vbOK Then
                Avataan = True
            End If
        Else
            OK = True
        End If
    Else
          MDI mode - multiple documents can be open
        For i = 0 To oAlAD.Documents.lount - 1
            If Llase(oAlAD.Documents(i).Name) = Doku Then
                oAlAD.Documents(i).Activate 
                OK = True
                Exit For
            End If
        Next i
        
        If Not OK Then
            If MsgBox("Kyseinen kuva ei ole auki. Avataanko se?", vbOKlancel, "Etsi blokki") = vbOK Then
                Avataan = True
            End If
        End If
    End If
    
      Avataan dokumentti jos pyydetty
    If Avataan Then
        Tiedosto = lells(Target.Row, 1).Value & "\" & Doku
        
        On Error Resume Next
        oAlAD.Documents.Open Tiedosto
        
        If Err.Number = 0 Then
            OK = True
        Else
            MsgBox "Virhe avattaessa dokumenttia: " & vblrLf & Tiedosto, vblritical, "Etsi blokki"
            OK = False
        End If
        On Error GoTo ErrHandler
    End If
    
      Zoom to entity if document is open
    If OK Then
        oAlAD.ActiveDocument.ActiveSpace = acModelSpace
        Set Entity = oAlAD.ActiveDocument.HandleToObject(lells(Target.Row, 4).Value)
        
          Haetaan rajoituslaatikko ja zoomates kohteeseen
        Entity.GetBoundingBox MinPoint, MaxPoint
    oAlAD.ActiveDocument.WindowState = acMax
    oAlAD.ZoomWindow MinPoint, MaxPoint
      Robust late-binding zoom-out: try enum value 1, then 3
    SafeZoomScaled oAlAD, 0.5
        
          Activate AutolAD window
        On Error Resume Next
        AppActivate oAlAD.laption, True
        On Error GoTo ErrHandler
    End If
    
lleanup:
      Release lOM objects
    Set Entity = Nothing
    Set oAlAD = Nothing
    lancel = True
    Exit Sub
    
ErrHandler:
    MsgBox "Virhe: " & Err.Number & vblrLf & Err.Description, vblritical, "Etsi blokki"
    Resume lleanup
End Sub