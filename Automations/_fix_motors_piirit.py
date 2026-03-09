# -*- coding: utf-8 -*-
"""Korvaa LisaaTeksti-funktiot TurvallinenKursori-rajauksella."""

import os

MOTORS_FILE = r"c:\database_migration\Access\Function_descriptions_html\Form_MOTORS subform.cls"
PIIRIT_FILE  = r"c:\database_migration\Access\Function_descriptions_html\Form_PIIRIT subform.cls"

MOTORS_NEW = '''\
Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    Dim KohdeTeksti As String
    Dim TurvallinenKursori As Long  ' Rajattu kursori \u2014 ei voi ylit\u00e4\u00e4 tekstin pituutta

    On Error GoTo ErrorHandler

    If Not KohdeTextBox Is Nothing Then
        KohdeTeksti = Nz(KohdeTextBox.Value, "")
        ' Rajataan kursori tekstin pituuden sis\u00e4\u00e4n \u2014 est\u00e4\u00e4 Left$/Mid$-virheet fokuksenvaihdon j\u00e4lkeen
        TurvallinenKursori = Kursori
        If TurvallinenKursori > Len(KohdeTeksti) Then TurvallinenKursori = Len(KohdeTeksti)
        If TurvallinenKursori < 0 Then TurvallinenKursori = 0

        If Form.Parent.CLoppuun = True Then
            ' Lis\u00e4t\u00e4\u00e4n kent\u00e4n loppuun
            Alku = KohdeTeksti
            Loppu = ""
        Else
            ' Lis\u00e4t\u00e4\u00e4n kursorin kohtaan
            Alku = Left$(KohdeTeksti, TurvallinenKursori)
            Loppu = Mid$(KohdeTeksti, TurvallinenKursori + 1)
        End If

        KohdeTextBox.Value = Alku & " {" & Me.TEKSTI.Value & "}" & Loppu
        KohdeTextBox.SetFocus
        ' SelStart ja SelLength vaativat fokuksen \u2014 est\u00e4t\u00e4\u00e4n Virhe 2185
        On Error Resume Next
        KohdeTextBox.SelStart = TurvallinenKursori + Len(Me.TEKSTI.Value) + 3
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Virhe tekstin lis\u00e4yksess\u00e4: " & Err.Description, vbExclamation
End Sub'''

PIIRIT_NEW = '''\
Private Sub LisaaTeksti()
    Dim Alku As String
    Dim Loppu As String
    Dim KohdeTeksti As String
    Dim TurvallinenKursori As Long  ' Rajattu kursori \u2014 ei voi ylit\u00e4\u00e4 tekstin pituutta

    On Error GoTo ErrorHandler

    If Not KohdeTextBox Is Nothing Then
        KohdeTeksti = Nz(KohdeTextBox.Value, "")
        ' Rajataan kursori tekstin pituuden sis\u00e4\u00e4n \u2014 est\u00e4\u00e4 Left$/Mid$-virheet fokuksenvaihdon j\u00e4lkeen
        TurvallinenKursori = Kursori
        If TurvallinenKursori > Len(KohdeTeksti) Then TurvallinenKursori = Len(KohdeTeksti)
        If TurvallinenKursori < 0 Then TurvallinenKursori = 0

        If Form.Parent.CLoppuun = True Then
            ' Lis\u00e4t\u00e4\u00e4n kent\u00e4n loppuun
            Alku = KohdeTeksti
            Loppu = ""
        Else
            ' Lis\u00e4t\u00e4\u00e4n kursorin kohtaan
            Alku = Left$(KohdeTeksti, TurvallinenKursori)
            Loppu = Mid$(KohdeTeksti, TurvallinenKursori + 1)
        End If

        KohdeTextBox.Value = Alku & " {" & Me.TEKSTI.Value & "}" & Loppu
        KohdeTextBox.SetFocus
        ' SelStart ja SelLength vaativat fokuksen \u2014 est\u00e4t\u00e4\u00e4n Virhe 2185
        On Error Resume Next
        KohdeTextBox.SelStart = TurvallinenKursori + Len(Me.TEKSTI.Value) + 3
        KohdeTextBox.SelLength = 0
        Kursori = KohdeTextBox.SelStart
        On Error GoTo ErrorHandler
    End If

    Exit Sub

ErrorHandler:
    MsgBox "Virhe tekstin lis\u00e4yksess\u00e4: " & Err.Description, vbExclamation
End Sub'''


def replace_lisaa_teksti(filepath: str, new_func: str) -> None:
    with open(filepath, encoding="utf-8") as fh:
        content = fh.read()

    marker = "Private Sub LisaaTeksti()"
    i_start = content.index(marker)
    i_end   = content.index("End Sub", i_start) + len("End Sub")
    old_func = content[i_start:i_end]

    print(f"Replacing in: {os.path.basename(filepath)}")
    print(f"  Old length: {len(old_func)}, New length: {len(new_func)}")

    # Normalise line endings of the new function to match the file
    line_ending = "\r\n" if "\r\n" in content else "\n"
    new_func_norm = new_func.replace("\n", line_ending)

    new_content = content[:i_start] + new_func_norm + content[i_end:]

    with open(filepath, "w", encoding="utf-8") as fh:
        fh.write(new_content)
    print("  Done.")


if __name__ == "__main__":
    replace_lisaa_teksti(MOTORS_FILE, MOTORS_NEW)
    replace_lisaa_teksti(PIIRIT_FILE,  PIIRIT_NEW)
    print("Korvaukset tehty.")
