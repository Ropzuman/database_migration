Public CheckOK As Boolean
Public PHStart As Long
Public PHEnd As Long
Public PFStart As Long
Public PFEnd As Long
Public DocStart As Long
Public DocEnd As Long
Public Sarakkeita As Long
Public RMAX As Long
Public POSheet As String
Public HideLINKING As Boolean
Public AddFooter As Boolean
Public DIContract As String
Public DIMill As String
Public DIDepartName As String
Public DICustomer As String
Public DIProject As String
Public DIProjNo As String
Public DIProjName As String
Public DIMunit As String
Public DIManager As String
Public DIDocNo As String
Public DIMetsoDocNo As String
Public DIDocName As String
Public DIDocName1 As String
Public DIDocName2 As String
Public DIDocName3 As String
Public DIPath As String
Public DIFile As String
Public DIDate As String
Public DIRev As String
Public DIRevArr() As String
Public DIRevID As String
Public DIRevDate As String
Public DIStatus As String

'''
' Module1.vba - Main logic for Kytkentälista Excel macro system.
' Handles data fetching from Access, checkouts and printout generation.
'''

' Performance/UX state (to minimize screen flashing and speed up macros.)
Private prevScreenUpdating As Boolean
Private prevCalculation As XlCalculation
Private prevEnableEvents As Boolean
Private prevDisplayAlerts As Boolean
Private prevDisplayStatusBar As Boolean

'''
' BeginFastMode: Temporarily disables Excel UI updates, events, and sets calculation to manual
' to speed up macro execution and prevent screen flicker.
'''
Private Sub BeginFastMode()
  prevScreenUpdating = Application.ScreenUpdating
  prevCalculation = Application.Calculation
  prevEnableEvents = Application.EnableEvents
  prevDisplayAlerts = Application.DisplayAlerts
  prevDisplayStatusBar = Application.DisplayStatusBar
  Application.ScreenUpdating = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  On Error Resume Next
  Application.DisplayStatusBar = False
  Application.AskToUpdateLinks = False
  On Error GoTo 0
End Sub