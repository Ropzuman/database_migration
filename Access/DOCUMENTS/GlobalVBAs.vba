Option Explicit
'---------------------------------------------
' 2001 VG Codes for checking current user name
'---------------------------------------------
Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
Private Declare PtrSafe Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As LongPtr) As Long
Function SetStartup()
    Dim DB As DAO.Database ' Changed to DAO.Database for clarity and best practice
    Dim taulu As DAO.Recordset ' Changed to DAO.Recordset
    Dim NWUserName As String
    Dim CName As String
    Dim BuffSize As LongPtr ' Changed to LongPtr
    Dim NBuffer As String
    
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    
    If api_GetUserName(NBuffer, BuffSize) Then
      NWUserName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      NWUserName = "Unknown"
    End If
    BuffSize = 256
    NBuffer = Space$(BuffSize)
    If api_GetComputerName(NBuffer, BuffSize) Then
      CName = Left$(NBuffer, InStr(NBuffer, Chr(0)) - 1)
    Else
      CName = "Unknown"
    End If
        
    Set DB = CurrentDb
    Set taulu = DB.OpenRecordset("UsysUsers", dbOpenTable)
    With taulu
      .AddNew
      .Fields(0) = NWUserName    'Users Name In Network
      .Fields(1) = CurrentUser()  'Users Name In This Database
      .Fields(2) = CName          'Users Computer Name
      .Fields(3) = Now            'Time At the Moment
      .Update
    End With
    Set DB = Nothing
    Set taulu = Nothing
End Function
Public Function Yhdista(T1 As String, T2 As String, T3 As String) As String
'This combines the Document name fields into one column.
'This cannot be done directly in a query because Access interprets the line break (vbCrLf) as a field name.
Dim apu As String
apu = IIf(T1 = "", "0", "1") & IIf(T2 = "", "0", "1") & IIf(T3 = "", "0", "1")
Select Case apu
  Case "000"
    Yhdista = ""
  Case "001"
    Yhdista = T3
  Case "010"
    Yhdista = T2
  Case "011"
    Yhdista = T2 & vbCrLf & T3
  Case "100"
    Yhdista = T1
  Case "101"
    Yhdista = T1 & vbCrLf & T3
  Case "110"
    Yhdista = T1 & vbCrLf & T2
  Case "111"
    Yhdista = T1 & vbCrLf & T2 & vbCrLf & T3
End Select
End Function
'***************************************************************************
'* REMOVED: Custom Replace() function                                       *
'* Date: November 8, 2025                                                  *
'* Reason: Shadowed VBA's built-in Replace() function with identical       *
'*         functionality. VBA built-in is faster (compiled C vs VBA loop). *
'* Impact: No code changes needed - built-in has same signature.           *
'* Used in: Form_USysRevText.cls (2 locations)                             *
'***************************************************************************
Public Function aReplace(Source As String) As String
'***************************************************************************
'* This function replaces all unsuitable characters in the given string    *
'***************************************************************************
Dim Tmp As String
Dim Lahde As String
Dim Merkki As String
Dim i As Long
Lahde = Source
Tmp = ""
    For i = 1 To Len(Source)
      Merkki = Mid(Lahde, i, 1)
      Select Case Merkki
        Case "/", "\", "?", "*", ":", ",", ";", "."
          Merkki = "-"
        Case Else
      End Select
      Tmp = Tmp & Merkki
    Next i
    aReplace = Tmp
End Function
'---------------------------------------------------------
' Functions for parsing revisions
' The following functions are given a revision notation as input and they return the desired part of the input.
' HaeTekija: Returns the creator (i.e., the very first author)
' HaeRevisioija: Returns the latest author (if there is only one line of data, both the creator and reviser are the same)
' HaeRevisio: Returns the notation of the latest revision
' HaeViimPaiva: Returns the date of the latest noted revision
' HaePaiva: Returns the date of the first revision
'
' - VG/22.3.2002
' - Updated: November 8, 2025 - Removed unused variables for code clarity
'---------------------------------------------------------
Function HaeTekija(Revisio As Variant) As String
'''
' Extracts the original author name from a multi-line revision string.
' Parses backward to find the first (oldest) revision entry.
' @param Revisio: Revision string with format "Rev Date/Author/Checker/..." separated by vbCrLf
' @return Author name from the first revision, or empty string if Null
'''
Dim i As Long
  If IsNull(Revisio) Then
    HaeTekija = ""
  Else
    i = 2
    'Look for the first revision (parse from end to find oldest entry)
    If InStr(Revisio, vbCrLf) Then
      Do
        i = i + 1
      Loop Until InStr(Right(Revisio, i), vbCrLf) = 1 Or i = Len(Revisio)
      Revisio = Mid(Revisio, Len(Revisio) - i + 3)
    End If
    Revisio = Mid(Revisio, InStr(Revisio, "/") + 1)
    HaeTekija = Left(Revisio, InStr(Revisio, "/") - 1)
  End If
End Function
Function HaeRevisioija(Revisio As String) As String
Dim Teksti As String
  Teksti = Revisio
  If InStr(Teksti, vbCrLf) Then 'If the input contains a line break
    Teksti = Mid(Teksti, InStr(Teksti, "/") + 1)
    HaeRevisioija = Left(Teksti, InStr(Teksti, "/") - 1)
  Else 'Since the input has only one revision, a reviser is not needed
    HaeRevisioija = ""
  End If
End Function
Function HaeRevisioijaPvm(Revisio As String) As String
Dim Teksti As String
Dim Tekija As String
Dim Pvm As String
  Teksti = Revisio
  If InStr(Teksti, vbCrLf) Then 'If the input contains a line break
    Pvm = Mid(Teksti, InStr(Teksti, " ") + 1)
    Pvm = Left(Pvm, InStr(Pvm, "/") - 1)
    Teksti = Mid(Teksti, InStr(Teksti, "/") + 1)
    Tekija = Left(Teksti, InStr(Teksti, "/") - 1)
    HaeRevisioijaPvm = Tekija & ": " & Pvm
  Else 'Since the input has only one revision, a reviser is not needed
    HaeRevisioijaPvm = ""
  End If
End Function
Public Function EkaRevRivi(Revisio As String) As String
  If InStr(Revisio, vbCrLf) Then
    EkaRevRivi = Left(Revisio, InStr(Revisio, vbCrLf) - 1)
  Else
    EkaRevRivi = Revisio
  End If
End Function
Public Function HaeRevisio(Revisio As Variant) As String
'''
' Extracts the revision mark (e.g., "A", "B", "0") from revision string.
' @param Revisio: Revision string with format "Rev Date/Author/..."
' @return Revision mark before the first space, or empty string if Null
'''
  If IsNull(Revisio) Then
    HaeRevisio = ""
  Else
    HaeRevisio = Left(Revisio, InStr(Revisio, " ") - 1)
  End If
End Function

Function HaeViimPaiva(Revisio As String) As String
'''
' Extracts the date from the first (oldest) revision entry.
' Parses backward through multi-line revision string to find original date.
' @param Revisio: Revision string with format "Rev Date/Author/..." separated by vbCrLf
' @return Date string from the first revision
'''
Dim i As Long
Dim Teksti As String
  Teksti = Revisio
  i = 2
  'Look for the first revision (parse from end to find oldest entry)
  If InStr(Teksti, vbCrLf) Then 'If the input contains a line break
    Do
      i = i + 1
    Loop Until InStr(Right(Teksti, i), vbCrLf) = 1 Or i = Len(Teksti)
    Teksti = Mid(Teksti, Len(Teksti) - i + 3)
  End If
  Teksti = Mid(Teksti, InStr(Teksti, " ") + 1)
  HaeViimPaiva = Left(Teksti, InStr(Teksti, "/") - 1)
End Function
Function HaePaiva(Revisio As String) As String
Dim Teksti As String
  Teksti = Revisio
  Teksti = Mid(Teksti, InStr(Teksti, " ") + 1)
  HaePaiva = Left(Teksti, InStr(Teksti, "/") - 1)
End Function
