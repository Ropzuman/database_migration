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
'T�m� scripti ottaa valitun reseptin k�ytt��n piirille
Dim TAreaCOde As String
Dim TLoopNo As String
Dim DB As Database
Dim sSQL As String

TAreaCOde = Me.Parent.AreaCode
TLoopNo = Me.Parent.LoopNo
Set DB = CurrentDb
If MsgBox("Otetaanko resepti (" & Me.PROCESS & " " & Me.CODE & ") k�ytt��n t�lle piirille " & TAreaCOde & "-" & TLoopNo & "?", vbOKCancel, "Ota resepti k�ytt��n") = vbOK Then
  'Sy�tet��n ensin k�ytetty RecipeID k�sitelt�v�n piirin RecipeID sarakkeeseen (moottorin tai loopin).
  Me.Parent.RecipeID = Me.ID
  Me.Parent.Refresh
  
  'Lis�t��n rivit Control taulukkoon. T�m� kysely lis�� control taulukkoon vain blokit, joita piirill� ei viel� ole.
  'Jos piirill� on jo blokkeja, niit� ei tuhota. K�ytt�j�n pit�� poistaa tarpeettomat/v��r�t blokit etuk�teen.
  sSQL = "INSERT INTO control ( RecipeID, AreaCode, LoopNo, FBlock, LoopSymb, Suffix, BlockPlace, Description2, [Note], Sivu, Sivuja ) "
  sSQL = sSQL & "SELECT Loops_With_Recipes.RecipeID, Loops_With_Recipes.AreaCode, Loops_With_Recipes.LoopNo, Loops_With_Recipes.FBlock, Loops_With_Recipes.LoopSymb, Loops_With_Recipes.Suffix, Loops_With_Recipes.BlockPlace, Loops_With_Recipes.Description2, Loops_With_Recipes.Note, Loops_With_Recipes.Sivu, Loops_With_Recipes.Sivuja "
  sSQL = sSQL & "FROM Loops_With_Recipes LEFT JOIN Control ON (Loops_With_Recipes.FBlock = Control.FBlock) AND (Loops_With_Recipes.Suffix = Control.Suffix) AND (Loops_With_Recipes.LoopSymb = Control.LoopSymb) AND (Loops_With_Recipes.LoopNo = Control.LoopNo) AND (Loops_With_Recipes.AreaCode = Control.AreaCode) AND (Loops_With_Recipes.RecipeID = Control.RecipeID) "
  sSQL = sSQL & "WHERE (((Loops_With_Recipes.AreaCode)='" & TAreaCOde & "') AND ((Loops_With_Recipes.LoopNo)='" & TLoopNo & "') AND ((Control.AreaCode) Is Null));"
  DB.Execute sSQL
  Me.Parent.Refresh

End If


End Sub
