Option Compare Database
Option Explicit

' Haetaan Windows-k‰ytt‰j‰nimi revisiomerkint‰‰ varten.
' W-versio ja StrPtr-osoitin, jotta skandinaaviset merkit s‰ilyv‰t oikein 64-bittisess‰ Officessa.
Private Declare PtrSafe Function api_GetUserName Lib "advapi32.dll" Alias "GetUserNameW" _
    (ByVal lpBuffer As LongPtr, ByRef nSize As Long) As Long

Private Sub ATaulu_Change()
  PaivitaIDRev
End Sub
Private Sub FiltteriB_Click()
  PaivitaTaulut
End Sub

Private Sub Form_Load()
Dim BuffSize As Long  ' Puskurin koko merkkein‰
Dim NBuffer As String  ' K‰ytt‰j‰nimen puskuri

  BuffSize = 256
  NBuffer = Space$(BuffSize)
    
  FiltteriT.Value = "*"
  PaivitaTaulut
  TDate.Value = Date
    
  If api_GetUserName(StrPtr(NBuffer), BuffSize) Then
     TMaker.Value = Left$(NBuffer, BuffSize - 1)  ' P‰‰tt‰v‰ nollamerkki j‰tet‰‰n pois
  End If
  TComment.Value = "Updated"
End Sub

' P‰ivitt‰‰ taulujen valintalistat suodattimen mukaan
Private Sub PaivitaTaulut()
Dim Suodatin As String  ' Suodatin, jossa heittomerkit on tuplattu SQL:‰‰ varten

  If Nz(FiltteriT.Value, "") = "" Then FiltteriT.Value = "*"
  Suodatin = Replace(CStr(FiltteriT.Value), "'", "''")
  ATaulu.RowSource = "SELECT Name From MSysObjects WHERE [Type] = 1 and Name Not Like 'MSys*' AND Name Like '" & Suodatin & "'"
  UTaulu.RowSource = ATaulu.RowSource
  ' Tyhj‰ lista: ei valita mit‰‰n, jotta ItemData(0) ei kaadu
  If ATaulu.ListCount = 0 Then
    ATaulu.Value = Null
    UTaulu.Value = Null
    ID.RowSource = ""
    Rev.RowSource = ""
    Exit Sub
  End If
  ATaulu.Value = ATaulu.ItemData(0)
  UTaulu.Value = ATaulu.ItemData(0)
  PaivitaIDRev
End Sub

' PaivitaIDRev: P‰ivitt‰‰ ID- ja Rev-valintalistat valitun taulun mukaan.
' Tunnistaa automaattinumerokent‰n (ID) ja revisiosarakkeen (Rev) sek‰ ehdottaa
' seuraavaa revisiotunnusta taulun viimeisimm‰n merkinn‰n perusteella.
Private Sub PaivitaIDRev()
Dim i As Integer  ' Silmukkalaskuri
Dim Taul As DAO.Recordset  ' Taulun kentt‰tietojen / revisioiden luku
Dim Revi As String  ' Viimeisin revisiotunnus
Dim Teksti As String  ' Viimeisin revisiomerkint‰ kokonaisuudessaan
Dim Valco As Long  ' Ensimm‰isen v‰lilyˆnnin sijainti

  If Nz(ATaulu.Value, "") = "" Then Exit Sub
  ID.RowSource = ATaulu.Value
  Rev.RowSource = ATaulu.Value
  For i = 0 To Rev.ListCount - 1
    If Left$(LCase$(Rev.ItemData(i)), 3) = "rev" Then
      Rev.Value = Rev.ItemData(i)
      Exit For
    End If
  Next i
  ' Avataan taulu kenttien tyyppien tarkistamista varten
  Set Taul = CurrentDb.OpenRecordset("SELECT * FROM [" & ATaulu.Value & "] WHERE False", dbOpenSnapshot)
  For i = 0 To Taul.Fields.Count - 1
    If Taul.Fields(i).Type = 4 And Taul.Fields(i).Attributes = 49 Then
      ID.Value = Taul.Fields(i).Name
      Exit For
    End If
  Next i
  Taul.Close
  Set Taul = Nothing
  If Nz(Rev.Value, "") <> "" Then
    ' Haetaan viimeisin revisiomerkint‰ seuraavan tunnuksen selvitt‰miseksi
    Set Taul = CurrentDb.OpenRecordset("SELECT [" & Rev.Value & "] FROM [" & ATaulu.Value & "] WHERE [" & Rev.Value & "] Is Not Null", dbOpenSnapshot)
    If Not Taul.EOF Then
      Teksti = CStr(Taul.Fields(0).Value)
      Valco = InStr(Teksti, " ")
      If Valco > 1 Then
        Revi = Left$(Teksti, Valco - 1)
      Else
        Revi = Teksti
      End If
      If Len(Revi) > 1 Then
        ' Numeerinen revisio: kasvatetaan ja t‰ytet‰‰n kaksinumeroiseksi
        Revi = CStr(Val(Revi) + 1)
        If Len(Revi) < 2 Then
          Revi = "0" & Revi
        End If
      ElseIf Len(Revi) = 1 Then
        ' Kirjainrevisio: seuraava kirjain
        Revi = Chr$(Asc(Revi) + 1)
      End If
      TRev = Revi
    End If
    Taul.Close
    Set Taul = Nothing
  End If
End Sub
Private Sub Sulje_Click()
 DoCmd.Close
End Sub

' Tee: Revisiomerkintˆjen p‰‰funktio.
' Vertaa kahta taulua ja kirjaa revisiotiedon muuttuneille riveille.
' Kaikki muutokset tehd‰‰n yhdess‰ transaktiossa: virhe peruu koko merkinn‰n.
Private Sub Tee_Click()
On Error GoTo Virhe
Dim Vert As DAO.Recordset  ' Taulujen yhdistetty vertailutulos
Dim Tyotila As DAO.Workspace  ' Tyˆtila transaktiota varten
Dim TransAktiivinen As Boolean  ' Onko transaktio k‰ynniss‰
Dim Muuttunut As Boolean  ' Onko rivin tieto muuttunut
Dim i As Integer  ' Silmukkalaskuri
Dim Merkinta As String  ' Uusi revisiomerkint‰
Dim VanhaArvo As Variant  ' Rivin nykyinen revisiohistoria
Dim virhenumero As Long  ' Virheen numero k‰sittelij‰ss‰
Dim virhekuvaus As String  ' Virheen kuvaus k‰sittelij‰ss‰

  ' Tarkistetaan, ett‰ kaikki tarvittavat valinnat on tehty
  If Nz(ATaulu.Value, "") = "" Or Nz(UTaulu.Value, "") = "" Or Nz(ID.Value, "") = "" Or Nz(Rev.Value, "") = "" Then
    MsgBox "Valitse ensin taulukot ja sarakkeet", vbCritical, "Tee merkinn‰t"
    Exit Sub
  End If
  If Nz(TRev.Value, "") = "" Or Nz(TDate.Value, "") = "" Or Nz(TMaker.Value, "") = "" Then
    MsgBox "Laita ensin revisiomerkinn‰t", vbCritical, "Tee merkinn‰t"
    Exit Sub
  End If

  Merk.Visible = True
  Me.Repaint
  Merkinta = TRev & " " & TDate & "/" & TMaker & "/" & Nz(TComment.Value, "")

  Set Tyotila = DBEngine.Workspaces(0)
  Tyotila.BeginTrans
  TransAktiivinen = True

  ' Yhdistetty vertailutulos: uuden taulun rivit ja niit‰ vastaavat rivit vanhasta taulusta
  Set Vert = CurrentDb.OpenRecordset("SELECT [" & ATaulu & "].*, [" & UTaulu & "].* FROM [" & ATaulu & "] RIGHT JOIN [" & UTaulu & "] ON [" & UTaulu & "].[" & ID & "] = [" & ATaulu & "].[" & ID & "];", dbOpenDynaset)
  Do While Not Vert.EOF
    Muuttunut = False
    ' Verrataan kaikkia kentti‰ paitsi ID- ja Rev-kentt‰‰
    For i = 0 To ID.ListCount - 1
      If ID.ItemData(i) <> ID.Value And ID.ItemData(i) <> Rev.Value Then
        If Nz(Vert.Fields(ATaulu & "." & ID.ItemData(i)).Value, "") <> Nz(Vert.Fields(UTaulu & "." & ID.ItemData(i)).Value, "") Then
          Muuttunut = True
          Exit For
        End If
      End If
    Next i

    ' Muuttuneelle riville kirjataan revisiotieto (uusin ensin)
    If Muuttunut Then
      VanhaArvo = Vert.Fields(UTaulu & "." & Rev).Value
      Vert.Edit
      If IsNull(VanhaArvo) Then
        Vert.Fields(UTaulu & "." & Rev).Value = Merkinta
      ElseIf CStr(VanhaArvo) = "" Then
        Vert.Fields(UTaulu & "." & Rev).Value = Merkinta
      Else
        Vert.Fields(UTaulu & "." & Rev).Value = Merkinta & vbCrLf & CStr(VanhaArvo)
      End If
      Vert.Update
    End If
    Vert.MoveNext
  Loop
  Tyotila.CommitTrans
  TransAktiivinen = False

Siivous:
  On Error Resume Next
  Merk.Visible = False
  If Not Vert Is Nothing Then Vert.Close
  Set Vert = Nothing
  Set Tyotila = Nothing
  If virhenumero <> 0 Then
    MsgBox "Virhe revisiomerkinnˆiss‰, muutokset peruttiin (" & virhenumero & "): " & virhekuvaus, vbCritical, "Tee merkinn‰t"
  Else
    MsgBox "Valmis", vbOKOnly, "Tee merkinn‰t"
  End If
  Exit Sub

Virhe:
  virhenumero = Err.Number
  virhekuvaus = Err.Description
  On Error Resume Next
  If TransAktiivinen Then Tyotila.Rollback
  Resume Siivous
End Sub
