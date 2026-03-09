# -*- coding: utf-8 -*-
"""Siivous: poistaa kaksoiskopiot jotka syntyivät kaksoisajosta."""

import os

BASE = r"c:\database_migration\Access\instru3"

def read_utf8(path):
    with open(path, encoding="utf-8") as f:
        return f.read()

def write_utf8(path, content):
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

# 1. general.bas: poista kaksoiskopio Päivitetty-rivistä
print("=== general.bas ===")
path = os.path.join(BASE, "general.bas")
c = read_utf8(path)

DUPE = ("'             2026-03-06 - GetOpenFileName-API poistettu \u2014 korvattu HaeTiedostoNimi (Application.FileDialog)\n"
        "'             2026-03-06 - GetOpenFileName-API poistettu \u2014 korvattu HaeTiedostoNimi (Application.FileDialog)\n")
SINGLE = "'             2026-03-06 - GetOpenFileName-API poistettu \u2014 korvattu HaeTiedostoNimi (Application.FileDialog)\n"

if DUPE in c:
    c = c.replace(DUPE, SINGLE, 1)
    print("  [OK] Poistettu kaksoiskopio Paivitetty-rivistä")
else:
    print("  [VIRHE] Kaksoiskopioita ei löydy - tarkista käsin")

# Päivitä Kuvaus: poista comdlg32-viittaus
OLD_DEP = ("'   - Tiedoston avausdialogi (Windows Common Dialog)\n"
           "'\n"
           "' Riippuvuudet:\n"
           "'   - comdlg32.dll (Common Dialog API)\n")
NEW_DEP = ("'   - Tiedoston avausdialogi (Office FileDialog)\n"
           "'\n"
           "' Riippuvuudet:\n")
if OLD_DEP in c:
    c = c.replace(OLD_DEP, NEW_DEP, 1)
    print("  [OK] Poistettu comdlg32.dll-viittaus Riippuvuudet-osiosta")
else:
    print("  [OHITUS] comdlg32-riviä ei löydy - jo poistettu tai erilainen")

write_utf8(path, c)
print("  Tallennettu.")

# 2. Form_Linkkien vaihto.cls: poista kaksoiskopio Dim tdf
print("\n=== Form_Linkkien vaihto.cls ===")
path = os.path.join(BASE, "Form_Linkkien vaihto.cls")
c = read_utf8(path)

DIM_DUPE = ("    Dim tdf As DAO.TableDef  ' Linkitetyn taulun m\u00e4\u00e4ritys p\u00e4ivityst\u00e4 varten\n"
            "    Dim tdf As DAO.TableDef  ' Linkitetyn taulun m\u00e4\u00e4ritys p\u00e4ivityst\u00e4 varten\n")
DIM_SINGLE = "    Dim tdf As DAO.TableDef  ' Linkitetyn taulun m\u00e4\u00e4ritys p\u00e4ivityst\u00e4 varten\n"

if DIM_DUPE in c:
    c = c.replace(DIM_DUPE, DIM_SINGLE, 1)
    print("  [OK] Poistettu kaksoiskopio Dim tdf -rivistä")
else:
    print("  [VIRHE] Kaksoiskopioita ei löydy")

write_utf8(path, c)
print("  Tallennettu.")

print("\n=== Siivous valmis ===")
