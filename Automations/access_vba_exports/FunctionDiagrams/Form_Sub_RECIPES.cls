VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_Sub_RECIPES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

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
'Tämä scripti ottaa valitun reseptin käyttöön piirille
Dim TAreaCOde As String
Dim TLoopNo As String
Dim DB As Database
Dim sSQL As String

TAreaCOde = Me.Parent.AreaCode
TLoopNo = Me.Parent.LoopNo
Set DB = CurrentDb
If MsgBox("Otetaanko resepti (" & Me.PROCESS & " " & Me.CODE & ") käyttöön tälle piirille " & TAreaCOde & "-" & TLoopNo & "?", vbOKCancel, "Ota resepti käyttöön") = vbOK Then
  'Syötetään ensin käytetty RecipeID käsiteltävän piirin RecipeID sarakkeeseen (moottorin tai loopin).
  Me.Parent.RecipeID = Me.ID
  Me.Parent.Refresh
  
  'Lisätään rivit Control taulukkoon. Tämä kysely lisää control taulukkoon vain blokit, joita piirillä ei vielä ole.
  'Jos piirillä on jo blokkeja, niitä ei tuhota. Käyttäjän pitää poistaa tarpeettomat/väärät blokit etukäteen.
  sSQL = "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) "
  sSQL = sSQL & "SELECT Loops_With_Recipes.RecipeID, Loops_With_Recipes.AreaCode, Loops_With_Recipes.LoopNo, Loops_With_Recipes.FBlock, Loops_With_Recipes.LoopSymb, Loops_With_Recipes.Suffix, Loops_With_Recipes.BlockPlace, Loops_With_Recipes.Description2, Loops_With_Recipes.Note, Loops_With_Recipes.Sivu, Loops_With_Recipes.Sivuja "
  sSQL = sSQL & "FROM Loops_With_Recipes LEFT JOIN Control ON (Loops_With_Recipes.FBlock = Control.FBlock) AND (Loops_With_Recipes.Suffix = Control.Suffix) AND (Loops_With_Recipes.LoopSymb = Control.LoopSymb) AND (Loops_With_Recipes.LoopNo = Control.LoopNo) AND (Loops_With_Recipes.AreaCode = Control.AreaCode) AND (Loops_With_Recipes.RecipeID = Control.RecipeID) "
  sSQL = sSQL & "WHERE (((Loops_With_Recipes.AreaCode)='" & TAreaCOde & "') AND ((Loops_With_Recipes.LoopNo)='" & TLoopNo & "') AND ((Control.AreaCode) Is Null));"
  DB.Execute sSQL
  Me.Parent.Refresh

End If


End Sub
