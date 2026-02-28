Option Compare Database
Option Explicit

'================================================================================
' Module: KAANNOS (Translation)
' Purpose: Translate equipment/loop references in text to descriptive names
' Updated: 2025-11-13 - Added VBA7/64-bit support and optimization
'
' Description:
'   Processes text containing equipment references in braces {xx-xx-xx description}
'   and translates them to actual equipment names from MAINEQ or Loops tables.
'   Validates references and marks deleted or missing items with error tags.
'
' Dependencies:
'   - Tables: MAINEQ, Loops
'   - DLookup function
'
' Format Examples:
'   Input:  "{60-20-01 Motor description}"
'   Output: "60-20-01 Actual Motor Name" (from MAINEQ table)
'   Input:  "{10-TIC-001 Loop description}"
'   Output: "10-TIC-001 Actual Loop Description" (from Loops table)
'================================================================================

'================================================================================
' Function: Kaanna
' Purpose: Translate equipment/loop references to actual names
' Parameters:
'   Tieto - Text containing references in format {Area-Type-Seq Description}
' Returns: Translated text with actual equipment names, or error markers
'
' Description:
'   Parses text for {POS description} patterns where:
'   - POS format: Area-Type-Seq (e.g., 60-20-01 or 10-TIC-001)
'   - If Area = "60": Looks up in MAINEQ (motors/equipment)
'   - Otherwise: Looks up in Loops table (process loops)
'   
'   Error handling:
'   - [ERR: Not found]: Equipment doesn't exist in database
'   - [DELETED!]: Equipment marked as deleted
'   - [ERR: No translation]: Equipment exists but has no description
'================================================================================
Function Kaanna(Tieto As Variant) As Variant
    Dim OS As Long, OS2 As Long, OS3 As Long, OS4 As Long
    Dim tPOS As String
    Dim Osat As Variant
    Dim Nimitys As Variant
    Dim Poistettu As Variant
    Dim Virheet As Long
    
    On Error GoTo ErrorHandler
    
    If IsNull(Tieto) Then
        Kaanna = Null
        Exit Function
    End If
    
    OS = InStr(Tieto, "{")
    If OS = 0 Then
        ' No references to translate
        Kaanna = Tieto
        Exit Function
    End If
    
    ' Initialize output with text before first reference
    Kaanna = Left$(Tieto, OS)
    
    Do While OS > 0
        ' Find position markers: { POS } structure
        OS2 = InStr(OS + 1, Tieto, " ")    ' Space after position
        OS3 = InStr(OS + 1, Tieto, "}")    ' Closing brace
        OS4 = InStr(OS3 + 1, Tieto, "{")   ' Next opening brace
        
        ' Extract position code (e.g., "60-20-01" or "10-TIC-001")
        tPOS = Mid$(Tieto, OS + 1, OS2 - OS - 1)
        Osat = Split(tPOS, "-")
        
        ' Determine table based on area code
        If Osat(0) = "60" Then
            ' Motor/equipment from MAINEQ
            Nimitys = DLookup("[EqNameSW20]", "MAINEQ", "[Department] = '" & Osat(1) & "' AND [EqSeq] = '" & Osat(2) & "'")
            Poistettu = DLookup("[Deleted]", "MAINEQ", "[Department] = '" & Osat(1) & "' AND [EqSeq] = '" & Osat(2) & "'")
        Else
            ' Process loop from Loops table
            Nimitys = DLookup("[Descr26_P]", "Loops", "[AreaCode] = '" & Osat(0) & "' AND [LoopSymb] = '" & Osat(1) & "' AND [LoopNo] = '" & Osat(2) & "'")
            Poistettu = DLookup("[DELETED]", "Loops", "[AreaCode] = '" & Osat(0) & "' AND [LoopSymb] = '" & Osat(1) & "' AND [LoopNo] = '" & Osat(2) & "'")
        End If
        
        ' Check for errors and mark accordingly
        If IsNull(Poistettu) Then
            Nimitys = "[ERR: Not found] " & Mid$(Tieto, OS2 + 1, OS3 - OS2 - 1)
            Virheet = Virheet + 1
        ElseIf Poistettu Then
            Nimitys = "[DELETED!] " & Mid$(Tieto, OS2 + 1, OS3 - OS2 - 1)
            Virheet = Virheet + 1
        ElseIf IsNull(Nimitys) Then
            Nimitys = "[ERR: No translation] " & Mid$(Tieto, OS2 + 1, OS3 - OS2 - 1)
            Virheet = Virheet + 1
        End If
        
        ' Build output: position + translation
        Kaanna = Kaanna & tPOS & " " & Nimitys
        
        ' Add text between this reference and next (or end)
        If OS4 <> 0 Then
            Kaanna = Kaanna & Mid$(Tieto, OS3, OS4 - OS3)
        Else
            Kaanna = Kaanna & Mid$(Tieto, OS3)
        End If
        
        ' Move to next reference
        OS = InStr(OS + 1, Tieto, "{")
    Loop
    
    ' Log errors to Immediate window for debugging
    If Virheet > 0 Then
        Debug.Print "Kaanna: " & Virheet & " error(s) in translation"
    End If
    
    Exit Function
    
ErrorHandler:
    Kaanna = "[ERR: " & Err.Description & "] " & Tieto
    Debug.Print "Kaanna error: " & Err.Description
End Function
