Attribute VB_Name = "GlobalVBAs"
Option Explicit
'---------------------------------------------
' 2001 VG Codes for checking current user name
'---------------------------------------------
Private Declare Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function api_GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Function SetStartup()
    Dim DB As Database
    Dim taulu As Recordset
    Dim NWUserName As String
    Dim CName As String
    Dim BuffSize As Long
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
        .Fields(0) = NWUserName     'Users Name In Network
        .Fields(1) = CurrentUser()  'Users Name In This Database
        .Fields(2) = CName          'Users Computer Name
        .Fields(3) = Now            'Time At the Moment
        .Update
    End With
    Set DB = Nothing
    Set taulu = Nothing
End Function
Public Function Yhdista(T1 As String, T2 As String, T3 As String) As String
'Tämä yhdistää Dokumentin nimikentät yhdeksi sarakkeeksi.
'Suoraan kyselyssä tämä ei onnistu, koska Access tulkitsee kyselyyn asetetun rivinvaihdon kenttänimeksi  (vbCrLf)
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
Public Function Replace(Src As String, Etsi As String, Uusi As String) As String
'***************************************************************************
'* Tämä Funktio korvaa annetusta merkkijonosta kaikki vaihdettavat         *
'* merkit (Etsi) vaihdettavalla merkillä (Uusi) ja                         *
'* palauttaa merkkijonon, jossa korvaukset on tehty.                       *
'* Esim. Replace("Matti;Maija;Liisa", ";", ", ") = "Matti, Maija, Liisa"   *
'*      Replace("Matti Maija Liisa", " ", "_") = "Matti_Maija_Liisa"       *
'***************************************************************************
Dim Pos As Long
Dim Pointer As Long
Dim Tmp As String
Dim Pituus As Integer
Dim Pituus2 As Integer
   Replace = Src
   Pointer = 1
   Pituus = Len(Etsi)
   Pituus2 = Len(Uusi)
   Do
      Pos = InStr(Pointer, Replace, Etsi)
      If Pos = 0 Then Exit Do
      Tmp = Left(Replace, Pos - 1) 'Muuttujan alku
      Replace = Tmp & Uusi & Mid(Replace, Pos + Pituus) 'Alku + Uusi + Loppu
      Pointer = Pos + Pituus2
   Loop
End Function
Public Function aReplace(Source As String) As String
'***************************************************************************
'* Tämä Funktio korvaa annetusta merkkijonosta kaikki sopimattomat merkit  *
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
' Funktiot Revisioiden parsimista varten
' Seuraaville funktiolle annetaan revisiomerkintä syötteenä ja ne palauttavat syötteestä halutun osan
' HaeTekija: Palauttaa tekijän (siis ihka ensimmäisen authorin)
' HaeRevisioija: Palauttaa viimeisimmäin authorin (jos on vain yksi rivi tietoa, sekä tekijä että revisioija ovat samat)
' HaeRevisio: Palauttaa viimeisen revision merkinnän
' HaeViimPaiva: Palauttaa viimeisimmän merkitys revision päivämäärän
' HaePaiva: Palauttaa ensimmäisen revision päivämäärään
'
' - VG/22.3.2002
'---------------------------------------------------------
Function HaeTekija(Revisio As Variant) As String
Dim i As Integer
Dim Pituus As Long
  If IsNull(Revisio) Then
    HaeTekija = ""
  Else
    i = 2
    Pituus = Len(Revisio)
    'Etsitään ensimmäine revisio
    If InStr(Revisio, vbCrLf) Then
      Do
        i = i + 1
      Loop Until InStr(Right(Revisio, i), vbCrLf) = 1 Or i = Pituus
      Revisio = Mid(Revisio, Pituus - i + 3)
    End If
    Revisio = Mid(Revisio, InStr(Revisio, "/") + 1)
    HaeTekija = Left(Revisio, InStr(Revisio, "/") - 1)
  End If
End Function
Function HaeRevisioija(Revisio As String) As String
Dim Teksti As String
  Teksti = Revisio
  If InStr(Teksti, vbCrLf) Then 'Jos syötteestä löytyy rivinvaihto
    Teksti = Mid(Teksti, InStr(Teksti, "/") + 1)
    HaeRevisioija = Left(Teksti, InStr(Teksti, "/") - 1)
  Else 'Koska syötteessä on vain yksi revisio, revisioijaa ei tarvita
    HaeRevisioija = ""
  End If
End Function
Function HaeRevisioijaPvm(Revisio As String) As String
Dim Teksti As String
Dim Tekija As String
Dim Pvm As String
  Teksti = Revisio
  If InStr(Teksti, vbCrLf) Then 'Jos syötteestä löytyy rivinvaihto
    Pvm = Mid(Teksti, InStr(Teksti, " ") + 1)
    Pvm = Left(Pvm, InStr(Pvm, "/") - 1)
    Teksti = Mid(Teksti, InStr(Teksti, "/") + 1)
    Tekija = Left(Teksti, InStr(Teksti, "/") - 1)
    HaeRevisioijaPvm = Tekija & ": " & Pvm
  Else 'Koska syötteessä on vain yksi revisio, revisioijaa ei tarvita
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
  If IsNull(Revisio) Then
    HaeRevisio = ""
  Else
    HaeRevisio = Left(Revisio, InStr(Revisio, " ") - 1)
  End If
End Function
Function HaeViimPaiva(Revisio As String) As String
Dim i As Integer
Dim Pituus As Long
Dim Teksti As String
  Teksti = Revisio
  i = 2
  Pituus = Len(Teksti)
  'Etsitään ensimmäine revisio
  If InStr(Teksti, vbCrLf) Then 'Jos syötteestä löytyy rivinvaihto
    Do
      i = i + 1
    Loop Until InStr(Right(Teksti, i), vbCrLf) = 1 Or i = Pituus
    Teksti = Mid(Teksti, Pituus - i + 3)
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
