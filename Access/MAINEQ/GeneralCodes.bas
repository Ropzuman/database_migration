Attribute VB_Name = "GeneralCodes"
Option Compare Database
Option Explicit
Public Revisioteksti As String
'---------------------------------------
'Edellisen kirjoitetun revision muistamista varten
Public MRevRev As String
Public MRevDrawn As String
Public MRevChecked As String
Public MRevApproved As String
Public MRevDescription As String
'---------------------------------------
Public TRevHist As String
Public TRevDesc As String
Function IsLoaded(ByVal strFormName As String) As Integer
 ' Palauttaa arvon "Tosi", jos määritetty lomake on avoinna
 ' lomake- tai taulukkonäkymässä.
    
    Const conObjStateClosed = 0
    Const conDesignView = 0
    
    If SysCmd(acSysCmdGetObjectState, acForm, strFormName) <> conObjStateClosed Then
        If Forms(strFormName).CurrentView <> conDesignView Then
            IsLoaded = True
        End If
    End If
    
End Function
Function HaeViimPaiva(Revisio As String) As String
Dim i As Integer
Dim Pituus As Long
Dim teksti As String
  teksti = Revisio
  i = 2
  Pituus = Len(teksti)
  'Etsitään ensimmäine revisio
  If InStr(teksti, vbCrLf) Then 'Jos syötteestä löytyy rivinvaihto
    Do
      i = i + 1
    Loop Until InStr(Right(teksti, i), vbCrLf) = 1 Or i = Pituus
    teksti = Mid(teksti, Pituus - i + 3)
  End If
  teksti = Mid(teksti, InStr(teksti, " ") + 1)
  HaeViimPaiva = Left(teksti, InStr(teksti, "/") - 1)
End Function

Public Function Replace(ByVal Source As String, Replaced As String, Replacement As String) As String
'***************************************************************************
'* Tämä Funktio korvaa annetusta merkkijonosta kaikki vaihdettavat         *
'* merkit (Replaced) vaihdettavalla merkillä (Replacement) ja              *
'* palauttaa merkkijonon, jossa korvaukset on tehty.                       *
'* Esim. Replace("Matti;Maija;Liisa", ";", ", ") = "Matti, Maija, Liisa"   *
'*      Replace("Matti Maija Liisa", " ", "_") = "Matti_Maija_Liisa"       *
'***************************************************************************
Dim pos As Long
Dim pointer As Long
Dim Tmp As String
Replace = Source
   pointer = 1
   Do
      pos = InStr(pointer, Replace, Replaced)
      If pos = 0 Then Exit Do
      Tmp = Mid(Replace, 1, pos - 1) 'Muuttujan alku talteen
      Replace = Tmp & Replacement & Mid(Replace, pos + Len(Replaced)) 'Alku + vaihdettu + loppu
      pointer = pos + Len(Replaced)
   Loop
End Function
Function Optiot(ByVal Drives_ID As Integer) As String
Dim DB As DAO.Database
Dim OptTaulu As DAO.Recordset
Dim teksti As String
Set DB = CurrentDB

Set OptTaulu = DB.OpenRecordset("SELECT Optio FROM MotorsOptions WHERE DrivesID = " & Drives_ID & ";")

teksti = ""
If Not (OptTaulu.EOF And OptTaulu.BOF) Then
    OptTaulu.MoveFirst
    teksti = "+"
    Do
        teksti = teksti & OptTaulu(0) & " +"
        OptTaulu.MoveNext
    Loop Until OptTaulu.EOF
    teksti = Left(teksti, Len(teksti) - 2)
End If
Optiot = teksti
End Function
Function Positiot(ByVal LaiteNr As String) As String
Dim DB As DAO.Database
Dim ElemTaulu As DAO.Recordset
Dim Teksti1 As String
Dim sqtxt As String

Set DB = CurrentDB
'sqtxt = "SELECT qryMotTilausSIEMENS1
sqtxt = "SELECT MAINEQ.ProjectElement, [maineq]![department] & '-' & [maineq]![eqtype] " _
    & "& '-' & [maineq]![eqseq] & ' / ' & [Drives].[suffix] AS Custpos FROM MAINEQ INNER JOIN DRIVES ON " _
    & "(MAINEQ.EqClass = DRIVES.EqClass) AND (MAINEQ.EqType = DRIVES.EqType) AND (MAINEQ.Eqseq = DRIVES.EqSeq) " _
    & "AND (MAINEQ.Department = DRIVES.Department) WHERE MAINEQ.ProjectElement= '" & LaiteNr & "';"
    
Set ElemTaulu = DB.OpenRecordset(sqtxt)


Teksti1 = ""
If Not (ElemTaulu.EOF And ElemTaulu.BOF) Then
    ElemTaulu.MoveFirst
    Teksti1 = "Pos: "
    Do
        Teksti1 = Teksti1 & ElemTaulu!Custpos & " and "
        ElemTaulu.MoveNext
    Loop Until ElemTaulu.EOF
    Teksti1 = Left(Teksti1, Len(Teksti1) - 5)
End If
Positiot = Teksti1
End Function
Function Vaihekulma(Cosfii)
Vaihekulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)
End Function
Function MotKaapUh(Cosfii As Single, Resist As Double, React As Double, Virta As Single, Voltage As Integer, Pituus As Integer)
Dim Kulma As Double
On Error GoTo MotKaapUhErr
Kulma = Atn(-Cosfii / Sqr(-Cosfii * Cosfii + 1)) + 2 * Atn(1)

MotKaapUh = Sqr(3) * Virta * ((Resist * Pituus * Cosfii) + (React * Pituus * Sin(Kulma)))
MotKaapUh = (MotKaapUh / Voltage) * 100
MotKaapUh = Format(MotKaapUh, "# ##0.0#") & " %"
Exit_Function:
    Exit Function

MotKaapUhErr:
    MotKaapUh = "00"
    Resume Exit_Function
End Function
Function LisaaNo(Tieto As Variant, Lisays As Integer) As String
Dim Pit As Integer
Dim No As Integer
Dim i As Integer
  If IsNull(Tieto) Then
    LisaaNo = ""
  Else
    Pit = Len(Tieto)
    No = VAL(Tieto)
    No = No + Lisays
    LisaaNo = CStr(No)
    For i = 0 To Len(LisaaNo) - Pit
    LisaaNo = "0" & LisaaNo
    Next i
  End If
End Function

'Esim.
'Kenttä: LisaaNo(Kentannimi;100)
