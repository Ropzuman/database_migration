Option Compare Database
Function Kaanna(Tieto As Variant) As Variant
Dim OS As Long, OS2 As Long, OS3 As Long, OS4 As Long
Dim tArea As String, tSymb As String, tLoop As String
Dim tPOS As String
Dim Osat As Variant
Dim Nimitys As Variant
Dim Poistettu As Variant
Dim Virheet As Long

If IsNull(Tieto) Then
  Kaanna = Null
Else
  OS = InStr(Tieto, "{")
  If OS Then
    Kaanna = Left(Tieto, OS)
    Do While OS > 0
      OS2 = InStr(OS + 1, Tieto, " ")
      OS3 = InStr(OS + 1, Tieto, "}")
      OS4 = InStr(OS3 + 1, Tieto, "{")
      tPOS = Mid(Tieto, OS + 1, OS2 - OS - 1)
      Osat = Split(tPOS, "-")
      If Osat(0) = "60" Then 'Moottori
        Nimitys = DLookup("[EqNameSW20]", "MAINEQ", "[Department] = '" & Osat(1) & "' AND [EqSeq] = '" & Osat(2) & "'")
        Poistettu = DLookup("[Deleted]", "MAINEQ", "[Department] = '" & Osat(1) & "' AND [EqSeq] = '" & Osat(2) & "'")
      Else
        Nimitys = DLookup("[Descr26_P]", "Loops", "[AreaCode] = '" & Osat(0) & "' AND [LoopSymb] = '" & Osat(1) & "' AND [LoopNo] = '" & Osat(2) & "'")
        Poistettu = DLookup("[DELETED]", "Loops", "[AreaCode] = '" & Osat(0) & "' AND [LoopSymb] = '" & Osat(1) & "' AND [LoopNo] = '" & Osat(2) & "'")
      End If
      If IsNull(Poistettu) Then
        Nimitys = "[ERR: Not found] " & Mid(Tieto, OS2 + 1, OS3 - OS2 - 1)
        Virheet = Virheet + 1
      ElseIf Poistettu Then
        Nimitys = "[DELETED!] " & Mid(Tieto, OS2 + 1, OS3 - OS2 - 1)
        Virheet = Virheet + 1
      ElseIf IsNull(Nimitys) Then
        Nimitys = "[ERR: No translation] " & Mid(Tieto, OS2 + 1, OS3 - OS2 - 1)
        Virheet = Virheet + 1
      End If
      Kaanna = Kaanna & tPOS & " " & Nimitys
      If OS4 <> 0 Then
        Kaanna = Kaanna & Mid(Tieto, OS3, OS4 - OS3)
      Else
        Kaanna = Kaanna & Mid(Tieto, OS3)
      End If
      
      
      
      
      OS = InStr(OS + 1, Tieto, "{")
    Loop
  Else
    Kaanna = Tieto
  End If
  If Virheet > 0 Then
    Debug.Print Virheet
  End If
End If
End Function
