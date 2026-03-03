Option Compare Database
Option Explicit
'==========================================================================
' LOMAKE   : Form_Sub_RECIPES
' SOVELLUS : FunctionDiagrams — Reseptinvalinta-alilomake
' KUVAUS   : Funktiokaavion RECIPES-alilomake. Mikä tahansa kenttä
'            (CODE, ID, INFO, PROCESS) kaksoisnapsautettuna kutsuu
'            ReseptiKayttoon()-proseduuria. Se tarkistaa käyttäjältä
'            vahvistuksen ja lisää Control-tauluun vain ne blokit,
'            joita piirille ei vielä ole määritelty — olemassa olevia
'            blokkeja ei ylikirjoiteta.
' DIPLE    : Option Explicit; On Error GoTo + Set DB = Nothing lisätty.
' PÄIVITETTY: 2026-03-03
'==========================================================================

Private Sub CODE_DblClick(Cancel As Integer)
  ReseptiKayttoon
End Sub
Private Sub ID_DblClick(Cancel As Integer)
  ReseptiKayttoon
End Sub

Private Sub INFO_DblClick(Cancel As Integer)
  ReseptiKayttoon
End Sub

Private Sub PROCESS_DblClick(Cancel As Integer)
  ReseptiKayttoon
End Sub
Private Sub ReseptiKayttoon()
' Ottaa valitun reseptin käyttöön tälle piirille.
' Lisää Control-tauluun vain uudet blokit (jo olemassa oleviin ei kosketa).
On Error GoTo ErrorHandler
Dim TAreaCOde As String
Dim TLoopNo As String
Dim DB As DAO.Database
Dim sSQL As String

TAreaCOde = Me.Parent.AreaCode
TLoopNo = Me.Parent.LoopNo
Set DB = CurrentDb
If MsgBox("Otetaanko resepti (" & Me.PROCESS & " " & Me.CODE & ") käyttöön tälle piirille " & TAreaCOde & "-" & TLoopNo & "?", vbOKCancel, "Ota resepti käyttöön") = vbOK Then
  ' Kirjataan valittu RecipeID piirilomakkeeseen
  Me.Parent.RecipeID = Me.ID
  Me.Parent.Refresh
  
  ' Lisätään Control-tauluun vain blokit, joita piirille ei vielä ole määritelty
  sSQL = "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) "
  sSQL = sSQL & "SELECT Loops_With_Recipes.RecipeID, Loops_With_Recipes.AreaCode, Loops_With_Recipes.LoopNo, Loops_With_Recipes.FBlock, Loops_With_Recipes.LoopSymb, Loops_With_Recipes.Suffix, Loops_With_Recipes.BlockPlace, Loops_With_Recipes.Description2, Loops_With_Recipes.Note, Loops_With_Recipes.Sivu, Loops_With_Recipes.Sivuja "
  sSQL = sSQL & "FROM Loops_With_Recipes LEFT JOIN Control ON (Loops_With_Recipes.FBlock = Control.FBlock) AND (Loops_With_Recipes.Suffix = Control.Suffix) AND (Loops_With_Recipes.LoopSymb = Control.LoopSymb) AND (Loops_With_Recipes.LoopNo = Control.LoopNo) AND (Loops_With_Recipes.AreaCode = Control.AreaCode) AND (Loops_With_Recipes.RecipeID = Control.RecipeID) "
  sSQL = sSQL & "WHERE (((Loops_With_Recipes.AreaCode)='" & TAreaCOde & "') AND ((Loops_With_Recipes.LoopNo)='" & TLoopNo & "') AND ((Control.AreaCode) Is Null));"
  DB.Execute sSQL
  Me.Parent.Refresh
End If

' Vapautetaan tietokantaresurssi
Set DB = Nothing
Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ReseptiKayttoon: " & Err.Number & " - " & Err.Description
    Debug.Print "    AreaCode: " & TAreaCOde & ", LoopNo: " & TLoopNo
    Debug.Print "    Source: " & Err.Source & ", Line: " & Erl
    Set DB = Nothing
    MsgBox "Virhe reseptin käyttöönotossa: " & Err.Description, vbCritical
End Sub
