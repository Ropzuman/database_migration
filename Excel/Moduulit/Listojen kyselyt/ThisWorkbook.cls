VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'''
' ThisWorkbook - Kytkentälista Excel-makrojärjestelmän työkirjatapahtumakäsittelijä.
' Workbook_Open nollaa ajokohtaiset globaalitilat avautuessa, jotta edellisen
' istunnon CheckOK=True-tila ei jää virheellisesti voimaan.
'''

Private Sub Workbook_Open()
  ' Nollataan globaali tila jotta edellisen istunnon CheckOK ei jää voimaan
  AlustaTila
  Debug.Print Format(Now, "hh:mm:ss") & " [Workbook_Open] Tila nollattu AlustaTila-kutsulla"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
  ' Nollataan tila myös sulkeutuessa (varmuuden vuoksi jos työkirja avataan uudelleen)
  AlustaTila
  Debug.Print Format(Now, "hh:mm:ss") & " [Workbook_BeforeClose] Tila nollattu"
End Sub
