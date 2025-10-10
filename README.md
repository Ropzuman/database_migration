# kanta3 VBA Project

Short description
-----------------

This repository contains Visual Basic for Applications (VBA) source code used in the "Kanta3" project. The primary source files are stored under the `Access/` directory; the main project module is `kanta3.vba`. The repository also contains other legacy or related modules (`kanta.vba`, `kanta1.vba`, `kanta2.vba`) and an `Excel/` folder for any related spreadsheets.

Repository contents
-------------------

- `Access/` — VBA modules and project files. Primary entry: `kanta3.vba`.
- `Excel/` — supporting Excel files used with the VBA code.
- `CHANGELOG_64bit_and_perf.md` — notes about 64-bit compatibility and performance changes.
- `README.md` — this file.

Usage
-----

1. Open Microsoft Access (or Excel if the code is used there).
2. Import or open the VBA modules from the `Access/` directory.
3. Review and run macros from the VBA editor (Alt+F11) after enabling macros and setting appropriate references.

Notes and suggestions
---------------------

- The code appears to be intended for Access/Excel environments. Ensure your Office version matches any API calls (32-bit vs 64-bit). See `CHANGELOG_64bit_and_perf.md` for migration notes.
- Consider adding a project layout (e.g., `src/`, `tests/`) if you plan to expand or port the code to other runtimes.
- Adding a LICENSE and a short CONTRIBUTING.md will help collaborators.

Contact / author
----------------

If you need help migrating, testing, or documenting the code, open an issue or contact the repository owner.
