Sub ExportAllModules()
    Dim vbc As VBIDE.VBComponent
    Dim sPath As String

    ' Määritä polku, johon moduulit viedään.
    ' Tämän kansion pitää olla olemassa, tai koodi antaa virheen.
    sPath = "C:\Data\Opinnäytetyö\VBA\Access\"
    
    ' Luo FileSystemObject, jolla varmistetaan kansion olemassaolo
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Tarkista, onko kansio olemassa. Jos ei, luo se.
    If Not fso.FolderExists(sPath) Then
        fso.CreateFolder (sPath)
    End If
    
    ' Tässä silmukka käy läpi kaikki aktiivisen työkirjan VBA-komponentit
    For Each vbc In ActiveWorkbook.VBProject.VBComponents
        ' Vie moduuli riippuen sen tyypistä
        Select Case vbc.Type
            Case vbext_ct_StdModule
                ' Standardimoduuli (.bas)
                vbc.Export sPath & vbc.Name & ".bas"
            Case vbext_ct_MSForm
                ' Käyttöliittymä (UserForm) (.frm)
                vbc.Export sPath & vbc.Name & ".frm"
            Case vbext_ct_ClassModule
                ' Luokkamoduuli (.cls)
                vbc.Export sPath & vbc.Name & ".cls"
            ' Tyyppi vbext_ct_Document jätetään tarkoituksella huomioimatta.
            ' Se viittaa työkirjan (esim. Sheet1 tai ThisWorkbook) koodeihin.
        End Select
    Next vbc

    ' Ilmoita käyttäjälle onnistuneesta viennistä
    MsgBox "Kaikki moduulit vietiin onnistuneesti kansioon: " & sPath, vbInformation, "Vienti valmis"
End Sub
