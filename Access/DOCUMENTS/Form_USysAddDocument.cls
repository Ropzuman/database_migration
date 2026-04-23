Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_USysAddDocument
' SOVELLUS : DOCUMENTS — Uuden dokumentin lisäys
' KUVAUS   : Lisää uuden dokumentin rekisteriin. Käyttää ed. dokumentin
'            tietoja oletusarvoina nopeuttaakseen sarjalisausta.
'            Tarkistaa asiakkaan numeron varauksen ennen tallennusta.
' PÄIVITETTY: 2026-03-03
'==========================================================================
Private Const TV_DOC_NEW_COUNTER As String = "DOC_NewRecordCounter"
Private Const TV_DOC_REV_TEXT As String = "DOC_Rev_Text"
Private Const TV_DOC_DEF_DNAME1 As String = "DOC_Def_DName1"
Private Const TV_DOC_DEF_DNAME2 As String = "DOC_Def_DName2"
Private Const TV_DOC_DEF_DNAME3 As String = "DOC_Def_DName3"
Private Const TV_DOC_DEF_AREA As String = "DOC_Def_Area"
Private Const TV_DOC_DEF_SORT As String = "DOC_Def_Sort"
Private Const TV_DOC_DEF_PAGES As String = "DOC_Def_Pages"
Private Const TV_DOC_DEF_PATH As String = "DOC_Def_Path"
Private Const TV_DOC_DEF_REV As String = "DOC_Def_Rev"

Public SelDept As String
Public Varaaja As String

' Oletusarvot luetaan ja kirjoitetaan TempVarsiin, jolloin lomake ei riipu Def*-globaaleista.
Private Sub ChoosePath_Click()
Dim Hakemisto As String
  Hakemisto = ValitseHakemPaikallinen(IIf(IsNull(Me.TPath.Value), "", Me.TPath.Value))
  If Hakemisto <> "" Then
    Me.TPath.Value = Hakemisto
  End If
End Sub
Private Sub Form_Load()
Dim taulu As DAO.Recordset
Dim OletusNimi1 As String
Dim OletusNimi2 As String
Dim OletusNimi3 As String
Dim OletusSivut As String
Dim OletusPolku As String
Dim OletusRev As String
  Varaaja = GetNetworkUserName()
' Käytetään viimeksi lisätyn dokumentin tietoja uuden pohjaksi
  OletusNimi1 = LueTempVarTeksti(TV_DOC_DEF_DNAME1)
  OletusNimi2 = LueTempVarTeksti(TV_DOC_DEF_DNAME2)
  OletusNimi3 = LueTempVarTeksti(TV_DOC_DEF_DNAME3)
  OletusSivut = LueTempVarTeksti(TV_DOC_DEF_PAGES)
  OletusPolku = LueTempVarTeksti(TV_DOC_DEF_PATH)
  OletusRev = LueTempVarTeksti(TV_DOC_DEF_REV)
  Set taulu = CurrentDb.OpenRecordset("SELECT Path FROM ProjInfo")
  Me.TDocName1.Value = IIf(OletusNimi1 <> "", OletusNimi1, Null)
  Me.TDocName2.Value = IIf(OletusNimi2 <> "", OletusNimi2, Null)
  Me.TDocName3.Value = IIf(OletusNimi3 <> "", OletusNimi3, Null)
  Me.TPages.Value = IIf(OletusSivut <> "", OletusSivut, 1)
  Me.TPath.Value = IIf(OletusPolku <> "", OletusPolku, taulu("Path"))
  Me.TRev.Value = IIf(OletusRev <> "", OletusRev, Null)
  Set taulu = Nothing
'--------------------------------------------------
End Sub
Private Sub Peruuta_Click()
  DoCmd.Close
End Sub
Private Sub TalletaNappi_Click()
    'Formin tietojen talletus ja sulkeminen
    If IsNull(Me.TDocName1.Value) And IsNull(Me.TDocName2.Value) Then
      MsgBox "You can't leave Document Name empty!" & vbCrLf _
           & "If you don't know document name now, put something like:" & vbCrLf _
           & """Reserved for subcontractor"", and write name down afterwards.", vbCritical
      Exit Sub
    ElseIf IsNull(Me.TClientDocNo.Value) And IsNull(Me.TMetsoDocNo.Value) And IsNull(Me.TExtDocID.Value) Then
      MsgBox "Document has to have some kind of ID!" & vbCrLf _
           & "If document is not going to client put at least Extra Document ID", vbCritical
      Exit Sub
    End If
    'Kaikki kunnossa, voidaan tallettaa tiedot ja poistua formista
    Dim DB As DAO.Database
    Dim taulu As DAO.Recordset
    Dim SQL As String
    Set DB = CurrentDb
      If IsNull(Me.TClientDocNo.Value) = False And Me.TClientDocNo.Value <> "" Then
        SQL = "SELECT * FROM UsysClientDocNos WHERE Number='AV-2090-210-DD-" & Me.TNumber.Value & "'"
        Set taulu = DB.OpenRecordset(SQL, , dbDenyWrite)
        If taulu.EOF = False Then
          If IsNull(taulu.Fields("Reserved")) Then
            taulu.Edit
            taulu.Fields("Number") = Me.TClientDocNo.Value
            taulu.Fields("Reserved") = Varaaja & "@" & Now
            taulu.Update
            taulu.Close
          Else
            MsgBox "Someone reserved number " & Me.TClientDocNo.Value & " at the same time!" & vbCrLf & "Please choose another number.", vbCritical, "Error!"
            taulu.Close
            DB.Close
            Set taulu = Nothing
            Set DB = Nothing
            Exit Sub
          End If
        End If
      End If
      Set taulu = DB.OpenRecordset("DOCUMENTS", dbOpenTable)
      With taulu
        .AddNew
        .Fields("DocName1") = Me.TDocName1.Value
        .Fields("DocName2") = Me.TDocName2.Value
        .Fields("DocName3") = Me.TDocName3.Value
        .Fields("Area") = Me.TArea.Value
        .Fields("Pick2") = Me.TPick2.Value
        .Fields("Rev") = Me.TRev.Value
        .Fields("ClientNo") = Me.TClientDocNo.Value
        .Fields("MetsoDocNo") = Me.TMetsoDocNo.Value
        .Fields("ExtDocID") = Me.TExtDocID.Value
        .Fields("Pages") = Me.TPages.Value
        .Fields("WorkPath") = Me.TPath.Value
        .Fields("Note") = Me.TNote.Value
        AsetaTempVar TV_DOC_NEW_COUNTER, CStr(.Fields("Counter"))
        .Update
      End With
      taulu.Close
      DB.Close
      Set taulu = Nothing
      Set DB = Nothing
' Tallennetaan viimeksi lisätyn dokumentin tiedot seuraavan pohjaksi
  AsetaTempVar TV_DOC_DEF_DNAME1, Nz(Me.TDocName1.Value, "")
  AsetaTempVar TV_DOC_DEF_DNAME2, Nz(Me.TDocName2.Value, "")
  AsetaTempVar TV_DOC_DEF_DNAME3, Nz(Me.TDocName3.Value, "")
  AsetaTempVar TV_DOC_DEF_AREA, Nz(Me.TArea.Value, "")
  AsetaTempVar TV_DOC_DEF_SORT, Nz(Me.TPick2.Value, "")
  AsetaTempVar TV_DOC_DEF_PAGES, Nz(Me.TPages.Value, "")
  AsetaTempVar TV_DOC_DEF_PATH, Nz(Me.TPath.Value, "")
  AsetaTempVar TV_DOC_DEF_REV, Nz(Me.TRev.Value, "")
'--------------------------------------------------
      DoCmd.Close
End Sub

Private Sub AsetaTempVar(ByVal Nimi As String, ByVal Arvo As String)
  On Error Resume Next
  TempVars.Remove Nimi
  On Error GoTo 0
  TempVars.Add Nimi, Arvo
End Sub

Private Function ValitseHakemPaikallinen(Optional ByVal AloitusPolku As String = "") As String
  Dim Dlg As Object
  Dim Polku As String

  On Error GoTo ErrorHandler

  Set Dlg = Application.FileDialog(4)
  With Dlg
    .Title = "Choose path:"
    .AllowMultiSelect = False
    If AloitusPolku <> "" Then .InitialFileName = AloitusPolku
    If .Show <> -1 Then
      ValitseHakemPaikallinen = ""
      GoTo Cleanup
    End If
    Polku = CStr(.SelectedItems(1))
  End With

  If Polku <> "" Then
    If Right$(Polku, 1) <> "\" Then Polku = Polku & "\"
  End If
  ValitseHakemPaikallinen = Polku

Cleanup:
  Set Dlg = Nothing
  Exit Function

ErrorHandler:
  ValitseHakemPaikallinen = ""
  Resume Cleanup
End Function

Private Function LueTempVarTeksti(ByVal Nimi As String) As String
  On Error GoTo ErrorHandler
  LueTempVarTeksti = Nz(TempVars(Nimi).Value, "")
  Exit Function
ErrorHandler:
  LueTempVarTeksti = ""
End Function

Private Function GetNetworkUserName() As String
  On Error GoTo FallbackUnknown
  ' Luetaan käyttäjänimi suoraan ympäristömuuttujasta, jotta lomake ei riipu ulkoisesta moduulifunktiosta.
  GetNetworkUserName = Trim$(CStr(Environ$("USERNAME")))
  If GetNetworkUserName = "" Then GoTo FallbackUnknown
  Exit Function
FallbackUnknown:
  GetNetworkUserName = "Unknown"
End Function

Private Sub TDiscipline_Change()
  MuutaNumero
End Sub

Private Sub TNumber_Change()
  MuutaNumero
End Sub
Private Sub TPaperSize_Change()
  MuutaNumero
End Sub
Private Sub TRev_Click()
  AvaaRevisio
End Sub
Private Sub MuutaNumero()
  If IsNull(Me.TPaperSize.Value) Or IsNull(Me.TDiscipline.Value) Or IsNull(Me.TNumber.Value) Then
  Else
    Me.TClientDocNo.Value = Me.TPaperSize.Value & "-2090-210-" & Me.TDiscipline.Value & "-" & Me.TNumber.Value
  End If
End Sub
Private Sub AvaaRevisio()
  ' Tallennetaan nykyinen revisio TempVarsiin, jotta USysRevText voi lukea ja kirjoittaa sen
  AsetaTempVar TV_DOC_REV_TEXT, IIf(IsNull(Me.TRev.Value), "", CStr(Me.TRev.Value))
  DoCmd.OpenForm "USysRevText", , , , , acDialog
  ' Luetaan päivitetty revisio takaisin TempVars:sta
  Dim UusiRev As String
  UusiRev = LueTempVarTeksti(TV_DOC_REV_TEXT)
  If UusiRev <> "" Then
    Me.TRev.Value = UusiRev
  End If
End Sub

