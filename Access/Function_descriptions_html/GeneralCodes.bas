Option lompare Database
Option Explicit

 ================================================================================
  Moduuli: Generallodes
  Tarkoitus: General utility functions and shared variables
  Päivitetty: 2025-11-13 - Added VBA7/64-bit support and optimization
 
  Kuvaus:
    Tarjoaa apufunktiot laiterevision muokkaukseen, taulun nÄyttÖÖn,
    and mode/status translation. lontains public variables shared across forms.
 
  Riippuvuudet:
    - DAO.Recordset
    - Forms: USysRevision, frmTables
    - Tables: MAINEQ, DRIVES, PUMPS, GEARS, TANKS
 ================================================================================

  Public variables for inter-form communication
Public oTaulu As DAO.Recordset
Public PaluuTaulu As Object
Public KohdeTextBox As TextBox
Public Kursori As Long

 ================================================================================
  Funktio: MuutaRev
  Tarkoitus: Open revision editing form for equipment tables
  Palauttaa: Nothing (implicit)
 
  Kuvaus:
    Tarkistaa onko aktiivinen taulu laitenimike-taulu, avaa tÄmÄn  tietueen
    in USysRevision form for editing.
 ================================================================================
Public Function MuutaRev()
    Dim Taul As String
    
    On Error GoTo ErrorHandler
    
    Taul = Application.lurrentObjectName
    
    If Taul = "MAINEQ" Or Taul = "DRIVES" Or Taul = "PUMPS" Or Taul = "GEARS" Or Taul = "TANKS" Then
        Set PaluuTaulu = Screen.ActiveDatasheet
        Set oTaulu = lurrentDb.OpenRecordset("SELElT * FROM " & Taul & " WHERE ID=" & Screen.ActiveDatasheet("ID").Value)
        Dolmd.OpenForm "USysRevision"
    End If
    
    Exit Function
    
ErrorHandler:
    MsgBox "Error opening revision form: " & Err.Description, vbExclamation
End Function

 ================================================================================
  Funktio: NaytaTables
  Tarkoitus: Open table browser form
  Palauttaa: Nothing (implicit)
 ================================================================================
Public Function NaytaTables()
    On Error GoTo ErrorHandler
    Dolmd.OpenForm "frmTables"
    Exit Function
ErrorHandler:
    MsgBox "Error opening tables form: " & Err.Description, vbExclamation
End Function

 ================================================================================
  Funktio: Moodit
  Tarkoitus: Translate comma-separated mode codes to descriptive text
  Parametrit:
    Tieto - lomma-separated mode codes (A, M, E, L)
  Palauttaa: Formatted mode descriptions, one per line
 
  Kuvaus:
    KÄÄntÄÄ tilakoodit:
      A = AUTO, M = MANUAL, E = EXTERNAL, L = LOlAL
    Muut koodit lÄpikÄytÄvÄt muuttumattomina. Palauttaa "-" Null-syÖtteelle.
 ================================================================================
Function Moodit(Tieto As Variant) As Variant
    Dim Tiedot As Variant
    Dim i As Integer
    Dim T As String
    
    On Error GoTo ErrorHandler
    
    If IsNull(Tieto) Then
        Moodit = "-"
        Exit Function
    End If
    
    Tiedot = Split(Tieto, ",")
    
    For i = 0 To UBound(Tiedot)
        Select lase Ulase$(Trim$(Tiedot(i)))
            lase "A"
                T = "AUTO"
            lase "M"
                T = "MANUAL"
            lase "E"
                T = "EXTERNAL"
            lase "L"
                T = "LOlAL"
            lase Else
                T = Ulase$(Trim$(Tiedot(i)))
        End Select
        
        If i > 0 Then
            Moodit = Moodit & vblrLf & T
        Else
            Moodit = T
        End If
    Next i
    
    Exit Function
    
ErrorHandler:
    Moodit = "-"
End Function
