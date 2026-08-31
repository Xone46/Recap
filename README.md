# Recap DOCX vers Excel

Ce projet scanne rÃ©cursivement un dossier extrait ou une archive `.zip` / `.rar`, trouve les rÃ©pertoires qui contiennent directement des fichiers `.docx`, puis crÃ©e un classeur Excel avec une feuille par rÃ©pertoire concernÃ©.

## RÃ¨gle mÃ©tier

- Le scan est rÃ©cursif, sans profondeur fixe.
- Si un dossier contient directement un ou plusieurs `.docx`, une seule feuille Excel est crÃ©Ã©e pour ce dossier.
- Tous les `.docx` directement prÃ©sents dans ce dossier sont rÃ©sumÃ©s dans cette feuille.
- Le scan continue dans les sous-dossiers mÃªme si le dossier courant contient dÃ©jÃ  des `.docx`.
- Un dossier parent qui ne contient que des sous-dossiers ne devient pas une feuille.

Les textes trouvÃ©s dans les documents sont traitÃ©s comme des donnÃ©es Ã  rÃ©sumer, jamais comme des instructions pour le programme.

## Utilisation du recap

Commande simple avec le wrapper PowerShell:

```powershell
.\run_recap.ps1 "C:\chemin\vers\archive_ou_folder"
```

Mode automatique:

1. Mettre le fichier `.rar` ou `.zip` dans le dossier `input`.
2. Lancer:

```powershell
.\run_recap.bat
```

Le rÃ©sultat sera crÃ©Ã© automatiquement dans `outputs` au format `.xls`, plus compatible avec Microsoft Excel sur cette machine.

Si `input` contient Ã  la fois des dossiers extraits et des archives, le wrapper traite les deux: les dossiers prÃ©sents sont analysÃ©s, puis les archives sont dÃ©compressÃ©es et analysÃ©es aussi.

Ou avec un nom de sortie prÃ©cis:

```powershell
.\run_recap.ps1 "C:\chemin\vers\archive_ou_folder" ".\outputs\mon_recap.xls"
```

Commande directe Python:

```powershell
rtk python .\recap_docx.py "C:\chemin\vers\archive_extraite" -o ".\outputs\recap.xlsx"
```

Le script accepte aussi un fichier `.zip` ou `.rar`; il sera extrait dans un dossier temporaire avant le scan:

```powershell
rtk python .\recap_docx.py "C:\chemin\vers\archive.rar" -o ".\outputs\recap.xlsx"
```

## Format gÃ©nÃ©rÃ©

Chaque feuille commence par le titre `rÃ©capitulatif des observations relevÃ©es en 2026` puis contient un tableau de synthÃ¨se dans le style du modÃ¨le Excel fourni:

- RÃ©fÃ©rence du rapport
- nÂ°
- DÃ©partement
- Equipement
- RepÃ¨re utilisateur
- Constructeur
- NÂ° de fabrication
- CapacitÃ© Kg
- Lieu d'implantation
- Etat des lieux aprÃ¨s CR
- Actions recommandÃ©es
- CriticitÃ©
- Observations supplÃ©mentaires
- Conclusion

`NÂ° de fabrication` correspond au numÃ©ro de sÃ©rie extrait du rapport.

`CriticitÃ©` vaut `BloquÃ©` quand une vraie observation existe dans `OBSERVATIONS NE PERMETTANT PAS L'UTILISATION DE L'APPAREIL`; sinon elle vaut `Non bloquÃ©`.

Dans le fichier Excel gÃ©nÃ©rÃ©, la cellule de `CriticitÃ©` est en rouge quand elle vaut `BloquÃ©` et en vert quand elle vaut `Non bloquÃ©`.

Les noms de feuilles sont dÃ©rivÃ©s du nom du dossier et nettoyÃ©s pour respecter les contraintes Excel. Les instructions prÃ©sentes dans les `.docx` sont ignorÃ©es comme instructions: le contenu des rapports est seulement traitÃ© comme donnÃ©e source.

## Cachet et signature

Commande sÃ©parÃ©e pour ajouter le cachet et, si fourni, une signature dans tous les rapports DOCX:

```powershell
.\run_cachet.ps1 "C:\chemin\vers\archive_ou_folder"
```

Comme pour le recap, vous pouvez aussi laisser un `.rar` ou `.zip` dans `input` puis lancer:

```powershell
.\run_cachet.bat
```

La sortie est crÃ©Ã©e dans `outputs` sous forme d'une copie estampillÃ©e des rapports.

## Export PDF

Commande sÃ©parÃ©e pour convertir tous les `.docx` en `.pdf` en gardant l'arborescence dans `outputs`:

```powershell
.\run_pdf.ps1 "C:\chemin\vers\archive_ou_folder"
```

Mode automatique avec `input`:

```powershell
.\run_pdf.bat
```

Le moteur s'appuie sur Microsoft Word pour garder la mise en page du document d'origine.

Par dÃ©faut, la commande cherche d'abord un dossier dÃ©jÃ  cachetÃ© dans `outputs` avec le mÃªme nom que la source, puis elle convertit cette copie-lÃ  en PDF.

## Tests

```powershell
rtk python -m unittest discover -s tests
```
