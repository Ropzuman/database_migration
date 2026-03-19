Option Compare Database
Option Explicit
'==========================================================================

Private Const TV_DOC_COMMON_TEXT As String = "DOC_Common_Text"
Private Const TV_DOC_COMMON_TITLE As String = "DOC_Common_Title"
Private Const TV_DOC_COMMON_TYPE As String = "DOC_Common_Type"
' LOMAKE   : Form_USysShowCommon
' SOVELLUS : DOCUMENTS — Yleiskäyttöinen tekstin näyttö/muokkaus
' KUVAUS   : Yleiskäyttöinen teksti-ikkuna lomakkeiden väliseen tiedonsiirtoon.
'            CommonType 0 = näytä ainoastaan (Locked), 1 = muokkaa.
'            SetFocus-kutsu on kosmeettinen — kaatumattomuuden vuoksi
'            suojattu On Error Resume Next -lohkolla.
' PÄIVITETTY: 2026-03-03
'==========================================================================

Private Sub BCancel_Click()
  DoCmd.Close
End Sub
Private Sub BOK_Click()
  Select Case HaeCommonTyyppi()
  Case 0
  Case 1
    AsetaTempVarYhteinen TV_DOC_COMMON_TEXT, Nz(Me.Teksti.Value, "")
  End Select
  DoCmd.Close
End Sub

Private Sub Form_Load()
  ' Jos OpenArgs on annettu, näytetään teksti suoraan ilman globaaleja välimuuttujia.
  If Not IsNull(Me.OpenArgs) Then
    Me.Otsikko.Caption = "Detail Address:"
    Me.Caption = "Detail Address:"
    Me.Teksti.Value = CStr(Me.OpenArgs)
    Me.BCancel.Visible = False
    Me.BOK.Caption = "Close"
    Me.Teksti.Locked = True
    Exit Sub
  End If

  Me.Otsikko.Caption = HaeCommonOtsikko()
  Me.Caption = HaeCommonOtsikko()
  Me.Teksti.Value = HaeCommonTeksti()
  Select Case HaeCommonTyyppi()
  Case 0
    Me.BCancel.Visible = False
    Me.BOK.Caption = "Close"
    Me.Teksti.Locked = True
  Case 1
    Me.BCancel.Visible = True
    Me.BOK.Caption = "OK"
    Me.Teksti.Locked = False
    ' SetFocus on kosmeettinen toiminto — suojataan kaatumiselta
    On Error Resume Next
    Me.Teksti.SetFocus
    On Error GoTo 0
  End Select
End Sub

Private Function HaeCommonTeksti() As String
  HaeCommonTeksti = LueTempVarTekstiYhteinen(TV_DOC_COMMON_TEXT)
End Function

Private Function HaeCommonOtsikko() As String
  HaeCommonOtsikko = LueTempVarTekstiYhteinen(TV_DOC_COMMON_TITLE)
End Function

Private Function HaeCommonTyyppi() As Integer
  Dim TyyppiTeksti As String

  TyyppiTeksti = LueTempVarTekstiYhteinen(TV_DOC_COMMON_TYPE)
  If TyyppiTeksti = "" Then
    HaeCommonTyyppi = 0
  Else
    HaeCommonTyyppi = CInt(Val(TyyppiTeksti))
  End If
End Function

Private Sub AsetaTempVarYhteinen(ByVal Nimi As String, ByVal Arvo As String)
  On Error Resume Next
  TempVars.Remove Nimi
  On Error GoTo 0
  TempVars.Add Nimi, Arvo
End Sub

Private Function LueTempVarTekstiYhteinen(ByVal Nimi As String) As String
  On Error GoTo ErrorHandler
  LueTempVarTekstiYhteinen = Nz(TempVars(Nimi).Value, "")
  Exit Function

ErrorHandler:
  LueTempVarTekstiYhteinen = ""
End Function