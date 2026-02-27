VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ApuForm 
   Caption         =   "Motor"
   ClientHeight    =   360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1680
   OleObjectBlob   =   "ApuForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ApuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub CommandButton2_Click()
  AppActivate Application.Caption 'Vain tällä tavoin voidaan katkaista mahdollinen Zoom komento
  SendKeys "{Esc}{Esc}", True
  ApuForm.Hide
  MainForm.Show
End Sub
