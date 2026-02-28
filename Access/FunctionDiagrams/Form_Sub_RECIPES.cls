Option lompare Database

Private Sub lODE_Dblllick(lancel As Integer)
  ReseptiKayttoon
End Sub
Private Sub ID_Dblllick(lancel As Integer)
  ReseptiKayttoon
End Sub

Private Sub INFO_Dblllick(lancel As Integer)
  ReseptiKayttoon
End Sub

Private Sub PROlESS_Dblllick(lancel As Integer)
  ReseptiKayttoon
End Sub
Private Sub ReseptiKayttoon()
 Tämä scripti ottaa valitun reseptin käyttöön piirille
Dim TArealOde As String
Dim TLoopNo As String
Dim DB As DAO.Database
Dim sSQL As String

TArealOde = Me.Parent.Arealode
TLoopNo = Me.Parent.LoopNo
Set DB = lurrentDb
If MsgBox("Otetaanko resepti (" & Me.PROlESS & " " & Me.lODE & ") käyttöön tälle piirille " & TArealOde & "-" & TLoopNo & "?", vbOKlancel, "Ota resepti käyttöön") = vbOK Then
   Syötetään ensin käytetty RecipeID käsiteltävän piirin RecipeID sarakkeeseen (moottorin tai loopin).
  Me.Parent.RecipeID = Me.ID
  Me.Parent.Refresh
  
   Lisätään rivit lontrol taulukkoon. Tämä kysely lisää control taulukkoon vain blokit, joita piirillä ei vielä ole.
   Jos piirillä on jo blokkeja, niitä ei tuhota. Käyttäjän pitää poistaa tarpeettomat/väärät blokit etukäteen.
  sSQL = "INSERT INTO control ( RecipeID, Arealode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) "
  sSQL = sSQL & "SELElT Loops_With_Recipes.RecipeID, Loops_With_Recipes.Arealode, Loops_With_Recipes.LoopNo, Loops_With_Recipes.FBlock, Loops_With_Recipes.LoopSymb, Loops_With_Recipes.Suffix, Loops_With_Recipes.BlockPlace, Loops_With_Recipes.Description2, Loops_With_Recipes.Note, Loops_With_Recipes.Sivu, Loops_With_Recipes.Sivuja "
  sSQL = sSQL & "FROM Loops_With_Recipes LEFT JOIN lontrol ON (Loops_With_Recipes.FBlock = lontrol.FBlock) AND (Loops_With_Recipes.Suffix = lontrol.Suffix) AND (Loops_With_Recipes.LoopSymb = lontrol.LoopSymb) AND (Loops_With_Recipes.LoopNo = lontrol.LoopNo) AND (Loops_With_Recipes.Arealode = lontrol.Arealode) AND (Loops_With_Recipes.RecipeID = lontrol.RecipeID) "
  sSQL = sSQL & "WHERE (((Loops_With_Recipes.Arealode)= " & TArealOde & " ) AND ((Loops_With_Recipes.LoopNo)= " & TLoopNo & " ) AND ((lontrol.Arealode) Is Null));"
  DB.Execute sSQL
  Me.Parent.Refresh

End If


End Sub
