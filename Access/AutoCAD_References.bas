Attribute VB_Name = "AutoCAD_References"
Option Compare Database
Option Explicit

' Poistaa Access VBA -projektista AutoCADin version sidotut tyyppikirjastoviittaukset.
' Access-kannan lähdekoodi käyttää AutoCADia myöhäisellä sidonnalla, joten viittausta
' ei tarvita käännökseen eikä sitä pidä säilyttää tietokannan References-listassa.
Public Sub PoistaAutoCADReferences()
    Dim projekti As Object
    Dim viite As Object
    Dim poistettavat As Collection
    Dim poistettu As Long

    On Error GoTo Virhe

    Set projekti = Application.VBE.ActiveVBProject
    Set poistettavat = New Collection

    For Each viite In projekti.References
        If OnAutoCADReference(viite) Then
            poistettavat.Add viite
        End If
    Next viite

    For Each viite In poistettavat
        projekti.References.Remove viite
        poistettu = poistettu + 1
    Next viite

    MsgBox "Poistettu AutoCAD References -viitteitä: " & poistettu & vbCrLf & _
           "Suorita nyt Debug -> Compile VBAProject.", vbInformation, "AutoCAD-viitteet"

Poistu:
    Set poistettavat = Nothing
    Set projekti = Nothing
    Exit Sub

Virhe:
    MsgBox "AutoCAD References -viitteitä ei voitu käsitellä: " & Err.Description & vbCrLf & _
           "Tarkista, että Trust Centerissä on sallittu VBA-projektimallin käyttö.", _
           vbExclamation, "AutoCAD-viitteet"
    Resume Poistu
End Sub

Private Function OnAutoCADReference(ByVal viite As Object) As Boolean
    Dim nimi As String
    Dim kuvaus As String
    Dim polku As String
    Dim rikkinäinen As Boolean

    On Error Resume Next
    nimi = UCase$(CStr(viite.Name))
    kuvaus = UCase$(CStr(viite.Description))
    polku = UCase$(CStr(viite.FullPath))
    rikkinäinen = CBool(viite.IsBroken)
    On Error GoTo 0

    OnAutoCADReference = rikkinäinen _
        Or InStr(1, nimi & " " & kuvaus & " " & polku, "AUTOCAD", vbTextCompare) > 0 _
        Or InStr(1, nimi & " " & kuvaus & " " & polku, "ACDB", vbTextCompare) > 0 _
        Or InStr(1, nimi & " " & kuvaus & " " & polku, "AXDB", vbTextCompare) > 0
End Function