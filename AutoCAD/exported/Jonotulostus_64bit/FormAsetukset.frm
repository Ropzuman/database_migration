VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAsetukset 
   Caption         =   "GhostScript asetukset"
   ClientHeight    =   3144
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5625
   OleObjectBlob   =   "FormAsetukset.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormAsetukset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Tallenna_asetukset()
    SaveSetting "Jonotulostus", "Asetukset", _
        "GSPath", GSPath.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "GSBinary", gsbinary.Value
    SaveSetting "Jonotulostus", "Asetukset", _
        "pc3Path", pc3Path.Value
End Sub


Private Sub GS_Exit_Click()
    Tallenna_asetukset
    Unload FormAsetukset
End Sub


Private Sub gsbinary_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim objFile As FileDialogs
    Dim strFilter As String
    Dim strDir As String
    Dim strFileName As String
    Dim tmpFilename As String
    
    Set fso = New FileSystemObject
    
    strFilter = "Application (*.exe)" & Chr(124) & "*.EXE"
    
    Set objFile = New FileDialogs
    objFile.OwnerHwnd = ThisDrawing.HWND

    objFile.Title = "Valitse GhostScript bin‰‰ri"
    objFile.StartInDir = GSPath.Value & "\bin"
    objFile.Filter = strFilter
    'return a valid filename
    strFileName = objFile.ShowOpen
    If Not strFileName = vbNullString Then
        'use this space to perform operation
        tmpFilename = fso.GetFileName(strFileName)
        gsbinary.Value = Left(tmpFilename, Len(tmpFilename) - 4)
        GSPath.Value = Left(strFileName, InStr(1, strFileName, "\bin") - 1)
    End If
    Set objFile = Nothing
    
    GS_Exit.SetFocus
    
End Sub

Private Sub UserForm_Initialize()
    GSPath.Text = GetSetting("Jonotulostus", "Asetukset", "GSPath", "C:\Data\Tools\gs\9.21")
    gsbinary.Value = GetSetting("Jonotulostus", "Asetukset", "GSBinary", "gswin64c")
    pc3Path.Value = GetSetting("Jonotulostus", "Asetukset", "pc3Path", "C:\Data\Tools")
End Sub

Private Sub pc3Path_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    pc3Muisti = pc3Path.Value
    pc3Uusi = Formi.BrowseForDirectory
    If pc3Uusi = "" Then
        pc3Path.Value = pc3Muisti
    Else
        pc3Path.Value = pc3Uusi
    End If
End Sub

Private Sub GSPath_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    GSPath.Value = Formi.BrowseForDirectory
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Tallenna_asetukset
End Sub
