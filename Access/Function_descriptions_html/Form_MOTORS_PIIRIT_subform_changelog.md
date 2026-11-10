# Muutosloki

- **Tiedostot:** `Form_MOTORS subform.cls`, `Form_PIIRIT subform.cls`
- **P‰iv‰m‰‰r‰:** 2026-11-10

## Kriittiset muutokset

- Ei toiminnallisia muutoksia.

## Siivous ja optimointi

- `Form.Parent.CLoppuun` ? `Me.Parent.CLoppuun` (selke‰mpi viittaus p‰‰loomakkeeseen).
- Kommenttien oikoluku: "yll‰pit‰‰", "kiedotaan", "estet‰‰n" ja "{aaltosulkeissa}" (viittaukset kirjoitetaan aaltosulkeisiin, ei hakasulkeisiin).
- J‰tetty ennalleen: molemmissa tiedostoissa on l‰hes identtinen `LisaaTeksti`. Sen voisi siirt‰‰ yhteiseen moduuliin erillisess‰ muutoksessa.
