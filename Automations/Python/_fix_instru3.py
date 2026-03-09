# -*- coding: utf-8 -*-
"""
Refaktorointi: Access/instru3 - code review -korjaukset 2026-03-06

Muutokset:
  1. Form_DBUsers.cls    - Command Injection (net send -> msg, sanitointi)
                         - dbCurrent.Close CurrentDb-viittauksella poistettu
  2. Form_SizingOut.cls  - RFC 4180 CSV-sanitointi + Formula Injection -esto
  3. Form_Linkkien vaihto.cls - DROP TABLE -> RefreshLink (turvallisempi)
  4. general.bas         - OPENFILENAME API poistettu -> HaeTiedostoNimi()
  5. Form_CopyLoops.cls  - Form_Unload: RecordSource "" + DoEvents ennen DELETE
                         - ValitseKanta_Click: GetOpenFileName -> Application.FileDialog
"""

import os
import re

BASE = r"c:\database_migration\Access\instru3"

def read_utf8(path):
    with open(path, encoding="utf-8") as f:
        return f.read()

def write_utf8(path, content):
    with open(path, encoding="utf-8", newline="") as f:
        pass  # just to check it exists
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def replace_between(content, old, new, label):
    if old not in content:
        print(f"  [VIRHE] '{label}' ei löydy tiedostosta!")
        return content
    result = content.replace(old, new, 1)
    print(f"  [OK] {label}")
    return result

def replace_by_index(content, start_marker, end_marker, new_text, label):
    """Korvaa osuus start_marker:sta end_marker:iin (sisältäen molemmat)."""
    i_start = content.find(start_marker)
    if i_start < 0:
        print(f"  [VIRHE] Alkumerkkiä '{start_marker[:40]}...' ei löydy! ({label})")
        return content
    i_end = content.find(end_marker, i_start)
    if i_end < 0:
        print(f"  [VIRHE] Loppumerkkiä '{end_marker[:40]}...' ei löydy! ({label})")
        return content
    i_end += len(end_marker)
    result = content[:i_start] + new_text + content[i_end:]
    print(f"  [OK] {label}")
    return result

# ============================================================
# 1. Form_DBUsers.cls
# ============================================================
print("\n=== Form_DBUsers.cls ===")
path = os.path.join(BASE, "Form_DBUsers.cls")
c = read_utf8(path)

# 1a. Command27_Click - Command Injection + net send -> msg
OLD_CMD27 = '''Private Sub Command27_Click()
Dim Viesti As String
  Viesti = INPUTBOX("Send message to user " & Me.NetworkName.Value & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & "Give message:", "Send message")
  If Viesti <> "" Then
    Call Shell("net send " & Me.NetworkName.Value & " \\""\\" & Viesti & "\\""\\"")
  End If
End Sub'''

# Use index-based approach for Command27_Click
NEW_CMD27 = '''Private Sub Command27_Click()
    Dim Viesti As String
    Dim PuhdistettuViesti As String
    Dim PuhdistettuKohde  As String

    Viesti = InputBox("Syöt\u00e4 viesti k\u00e4ytt\u00e4j\u00e4lle " & Me.NetworkName.Value & ":" & vbCrLf & vbCrLf, "L\u00e4het\u00e4 viesti")

    If Trim(Viesti) <> "" Then
        ' Estet\u00e4\u00e4n Command Injection \u2014 poistetaan lainausmerkit ja Shell-erikoismerkit
        PuhdistettuViesti = Replace(Viesti, """", "'")
        PuhdistettuViesti = Replace(PuhdistettuViesti, "&", "ja")
        PuhdistettuViesti = Replace(PuhdistettuViesti, "|", "")
        PuhdistettuKohde  = Replace(CStr(Nz(Me.NetworkName.Value, "")), " ", "")

        ' K\u00e4ytet\u00e4\u00e4n modernia MSG.EXE-komentoa \u2014 NET SEND poistettu Windows Vistasta l\u00e4htien
        Call Shell("msg " & PuhdistettuKohde & " """ & PuhdistettuViesti & """", vbHide)
    End If
End Sub'''

c = replace_by_index(
    c,
    "Private Sub Command27_Click()",
    "End Sub",
    NEW_CMD27,
    "Command27_Click - Command Injection + msg-komento"
)

# 1b. dbCurrent.Close -> Set dbCurrent = Nothing
c = replace_between(
    c,
    "   Set dbCurrent = CurrentDb\n   SPath = dbCurrent.Name\n   dbCurrent.Close",
    "   Set dbCurrent = CurrentDb\n   SPath = dbCurrent.Name\n   ' CurrentDb-viittausta EI suljeta .Close-kutsulla \u2014 vain Set Nothing\n   Set dbCurrent = Nothing",
    "dbCurrent.Close -> Set Nothing"
)

# Paivitetty-header
c = replace_between(
    c,
    "' P\u00e4ivitetty: 2025-11-11",
    "'             2026-03-03 - Me.-etuliitteet ja kommentit suomeksi",
    "Paivitetty-header MOTORS"
)

# Rakenna uusi header
OLD_HDR_DBU = "' P\u00e4ivitetty: 2025-11-11"
# This is actually done above but let me check which header is in this file
# No header in Form_DBUsers.cls based on what I read - it jumps straight to UserRec type

with open(path, "w", encoding="utf-8") as f:
    f.write(c)
print("  Tallennettu.")

# ============================================================
# 2. Form_SizingOut.cls
# ============================================================
print("\n=== Form_SizingOut.cls ===")
path = os.path.join(BASE, "Form_SizingOut.cls")
c = read_utf8(path)

OLD_CSV = "            ' Kirjoitetaan kentt\u00e4arvo lainausmerkeiss\u00e4, puolipisteeroteltu\n            Tiedosto.Write \";\"\"\" & Arvo & \"\"\"\""
NEW_CSV = """            ' RFC 4180: tuplaa lainausmerkit \u2014 est\u00e4\u00e4 CSV-rakenteen rikkoutumisen
            Arvo = Replace(Arvo, \"\"\"\", \"\"\"\"\"\"\"\"\")
            ' Formula Injection -esto: lis\u00e4t\u00e4\u00e4n heittomerkki vaarallisten =-@-alkuisten arvojen eteen
            If Arvo <> \"\" Then
                If Left$(Arvo, 1) = \"=\" Or Left$(Arvo, 1) = \"@\" Or _
                   Left$(Arvo, 1) = \"+\" Or Left$(Arvo, 1) = \"-\" Then
                    Arvo = \"'\" & Arvo
                End If
            End If
            ' Kirjoitetaan kentt\u00e4arvo lainausmerkeiss\u00e4, puolipisteeroteltu
            Tiedosto.Write \";\"\"\" & Arvo & \"\"\"\"  """

c = replace_between(c, OLD_CSV, NEW_CSV, "CSV RFC 4180 + Formula Injection")

with open(path, "w", encoding="utf-8") as f:
    f.write(c)
print("  Tallennettu.")

# ============================================================
# 3. Form_Linkkien vaihto.cls
# ============================================================
print("\n=== Form_Linkkien vaihto.cls ===")
path = os.path.join(BASE, "Form_Linkkien vaihto.cls")
c = read_utf8(path)

# 3a. Add tdf declaration
c = replace_between(
    c,
    "    Dim Taul As DAO.Recordset  ' MSysObjects query results\n    Dim Taulu As String  ' Current table name being processed",
    "    Dim Taul As DAO.Recordset  ' MSysObjects query results\n    Dim Taulu As String  ' Current table name being processed\n    Dim tdf As DAO.TableDef  ' Linkitetyn taulun m\u00e4\u00e4ritys p\u00e4ivityst\u00e4 varten",
    "Dim tdf As DAO.TableDef"
)

# 3b. Replace DROP TABLE + DoCmd.TransferDatabase
OLD_DROP = """                ' Drop existing linked table
                CurrentDb.Execute "DROP TABLE [" & Taul.Fields("Name") & "]"
                
                ' Recreate link to current directory
                On Error Resume Next
                DoCmd.TransferDatabase acLink, "Microsoft Access", Polku & Nimi, acTable, Taulu, Taulu, False
                Err.Clear
                On Error GoTo ErrorHandler"""

NEW_DROP = """                ' P\u00e4ivitet\u00e4\u00e4n linkki turvallisesti \u2014 RefreshLink ei riko taulua eik\u00e4 sen relaatioita
                Set tdf = CurrentDb.TableDefs(Taulu)
                tdf.Connect = ";DATABASE=" & Polku & Nimi
                On Error Resume Next
                tdf.RefreshLink
                If Err.Number <> 0 Then
                    MsgBox "Virhe p\u00e4ivitett\u00e4ess\u00e4 taulua " & tdf.Name & ": " & Err.Description, vbExclamation, "Linkin p\u00e4ivitysvirhe"
                End If
                Err.Clear
                On Error GoTo ErrorHandler
                Set tdf = Nothing"""

c = replace_between(c, OLD_DROP, NEW_DROP, "DROP TABLE -> RefreshLink")

# 3c. Add Set tdf = Nothing to ErrorHandler
c = replace_between(
    c,
    "ErrorHandler:\n    MsgBox \"Error updating links: \" & Err.Description, vbCritical, \"Link Update\"\n    On Error Resume Next\n    If Not Taul Is Nothing Then Taul.Close\n    Set Taul = Nothing\n    On Error GoTo 0",
    "ErrorHandler:\n    MsgBox \"Error updating links: \" & Err.Description, vbCritical, \"Link Update\"\n    On Error Resume Next\n    If Not Taul Is Nothing Then Taul.Close\n    Set Taul = Nothing\n    Set tdf = Nothing\n    On Error GoTo 0",
    "ErrorHandler: Set tdf = Nothing"
)

with open(path, "w", encoding="utf-8") as f:
    f.write(c)
print("  Tallennettu.")

# ============================================================
# 4. general.bas
# ============================================================
print("\n=== general.bas ===")
path = os.path.join(BASE, "general.bas")
c = read_utf8(path)

# Header update
c = replace_between(
    c,
    "'             2026-03-03 - Kommentit suomeksi",
    "'             2026-03-03 - Kommentit suomeksi\n'             2026-03-06 - GetOpenFileName-API poistettu \u2014 korvattu HaeTiedostoNimi (Application.FileDialog)",
    "Header-paivitys"
)

# Replace the #If VBA7 Then ... #End If block with HaeTiedostoNimi
# Find start and end markers
START_API = "'--------------------------------------------------------------------------------\n' Windows Common Dialog API -m\u00e4\u00e4rittely\n' P\u00e4ivitetty 2025-11-11: VBA7/64-bit-tuki lis\u00e4tty GetOpenFileName-APIlle\n'--------------------------------------------------------------------------------"
END_API = "#End If"

NEW_API = """'--------------------------------------------------------------------------------
' Funktio: HaeTiedostoNimi
' Tarkoitus: Avaa Office-natiivi tiedostovalintaikkuna
'
' Palauttaa:
'   Merkkijono \u2014 valitun tiedoston t\u00e4ydellinen polku, tai \"\" jos peruttu
'
' Huomiot:
'   - Korvaa vanhan ja 64-bittisess\u00e4 Officessa ep\u00e4vakaan GetOpenFileName-API-rakenteen
'   - Application.FileDialog toimii luotettavasti kaikissa Office-versioissa (32/64-bit)
'--------------------------------------------------------------------------------
Public Function HaeTiedostoNimi() As String
    Dim fd As Object
    ' msoFileDialogFilePicker = 3
    Set fd = Application.FileDialog(3)
    With fd
        .Title = "Valitse tiedosto"
        .AllowMultiSelect = False
        If .Show = -1 Then
            HaeTiedostoNimi = .SelectedItems(1)
        Else
            HaeTiedostoNimi = ""
        End If
    End With
    Set fd = Nothing
End Function"""

c = replace_by_index(c, START_API, END_API, NEW_API, "OPENFILENAME API -> HaeTiedostoNimi")

with open(path, "w", encoding="utf-8") as f:
    f.write(c)
print("  Tallennettu.")

# ============================================================
# 5. Form_CopyLoops.cls
# ============================================================
print("\n=== Form_CopyLoops.cls ===")
path = os.path.join(BASE, "Form_CopyLoops.cls")
c = read_utf8(path)

# 5a. Form_Unload: Add RecordSource clear + DoEvents
OLD_UNLOAD = """Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  CurrentDb.TableDefs.Delete "LOOPLINK"  ' Poistetaan v\u00e4liaikainen LOOPLINK-taulu
  Err.Clear
  On Error GoTo 0
  If Not LahdeDB Is Nothing Then
    LahdeDB.Close
    Set LahdeDB = Nothing
  End If
End Sub"""

NEW_UNLOAD = """Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ' Vapautetaan alilomakkeen lukitukset ennen taulun poistoa \u2014 est\u00e4\u00e4 Access-kaatumisen
    Me.Loopit.Form.RecordSource = ""
    DoEvents  ' Annetaan Accessille hetki vapauttaa tiedostolukot
    CurrentDb.TableDefs.Delete "LOOPLINK"  ' Poistetaan v\u00e4liaikainen LOOPLINK-taulu
    Err.Clear
    If Not LahdeDB Is Nothing Then
        LahdeDB.Close
        Set LahdeDB = Nothing
    End If
    On Error GoTo 0
End Sub"""

c = replace_between(c, OLD_UNLOAD, NEW_UNLOAD, "Form_Unload: RecordSource + DoEvents")

# 5b. ValitseKanta_Click: Replace GetOpenFileName with Application.FileDialog
# Find the start and replace the entire sub
NEW_VALITSE = """Private Sub ValitseKanta_Click()
On Error GoTo ErrorHandler
    ' Avataan Office-natiivi tiedostovalintaikkuna tietokannan valitsemiseksi
    Dim fd As Object                   ' FileDialog-objekti
    Dim ValittuTiedosto As String      ' K\u00e4ytt\u00e4j\u00e4n valitsema tiedostopolku
    Dim AHakem As String               ' Alkuhakemisto dialogille
    Dim Linkki As DAO.TableDef         ' V\u00e4liaikaisen linkitetyn taulun m\u00e4\u00e4ritys

    Debug.Print "ValitseKanta: Opening database selection dialog"

    ' M\u00e4\u00e4ritet\u00e4\u00e4n alkuhakemisto aiemmin valitusta tai nykykannasta
    If InStr(Nz(Me.Kanta.Value), "\\\\") Then
        AHakem = Me.Kanta.Value
    Else
        AHakem = Application.CurrentDb.Name
    End If
    AHakem = Left$(AHakem, InStrRev(AHakem, "\\\\"))

    ' Application.FileDialog korvaa haurauden OPENFILENAME-API-rakenteen (kaataa 64-bit Accessin)
    Set fd = Application.FileDialog(3)  ' msoFileDialogFilePicker = 3
    With fd
        .Title = "Choose Instru Database"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Microsoft Access Databases", "*.accdb"
        .InitialFileName = AHakem
        If .Show = -1 Then
            ValittuTiedosto = .SelectedItems(1)
        End If
    End With
    Set fd = Nothing

    If ValittuTiedosto <> "" Then
        Me.Kanta.Value = ValittuTiedosto
        Debug.Print "  Valittu tietokanta: " & ValittuTiedosto

        ' Avataan l\u00e4hdetietokantayhteys
        Set LahdeDB = OpenDatabase(ValittuTiedosto)
        Debug.Print "  Tietokanta avattu onnistuneesti"

        ' Poistetaan mahdollinen olemassa oleva LOOPLINK-taulu
        On Error Resume Next
        CurrentDb.TableDefs.Delete "LOOPLINK"
        Err.Clear
        On Error GoTo ErrorHandler

        ' Luodaan linkitetty taulu l\u00e4hteen Loops-tauluun
        Set Linkki = CurrentDb.CreateTableDef("LOOPLINK")
        Linkki.Connect = ";DATABASE=" & ValittuTiedosto
        Linkki.SourceTableName = "Loops"
        CurrentDb.TableDefs.Append Linkki
        Set Linkki = Nothing

        ' N\u00e4ytet\u00e4\u00e4n saatavilla olevat loopit alilomakkeessa
        Me.Loopit.Form.RecordSource = "SELECT * FROM LOOPLINK"
    End If
    Me.Refresh
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR in ValitseKanta: " & Err.Number & " - " & Err.Description
    MsgBox "Virhe tietokannan valinnassa: " & Err.Description, vbCritical, "Valitse tietokanta"
    On Error Resume Next
    Set fd = Nothing
    Set Linkki = Nothing
    On Error GoTo 0
End Sub"""

# Fix the backslash escaping for actual content (Python string needed \\, VBA needs \)
NEW_VALITSE = NEW_VALITSE.replace("\\\\", "\\")

# Find ValitseKanta_Click start and end
i_start = c.find("Private Sub ValitseKanta_Click()")
i_end = c.find("End Sub", i_start) + len("End Sub")
old_valitse = c[i_start:i_end]
print(f"  ValitseKanta_Click: len={len(old_valitse)}")
c = c[:i_start] + NEW_VALITSE + c[i_end:]
print("  [OK] ValitseKanta_Click: GetOpenFileName -> Application.FileDialog")

with open(path, "w", encoding="utf-8") as f:
    f.write(c)
print("  Tallennettu.")

print("\n=== Kaikki korjaukset tehty ===")
