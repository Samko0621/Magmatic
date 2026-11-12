# Magmatic — capitalisation documentaire

Projet d'alternance (BUT GMP, Savoye — BE mécanique). Phase 2 : historique des affaires
Magmatic et guide des règles de conception.

## Arborescence

| Dossier | Contenu |
|---|---|
| `livrables/` | Guides (PDF / DOCX), historique des affaires et récap clients (XLSX) |
| `sources/` | Source HTML du guide, scripts de génération, illustrations découpées |
| `images/` | 77 captures de travail triées en 10 sous-dossiers (index dans `images/README.md`) |
| `donnees_entree/` | Classeur de lancement EHP01718LEVM001 |

## Génération

```powershell
powershell -ExecutionPolicy Bypass -File sources\build_guide.ps1   # figures + PDF (Edge headless)
powershell -ExecutionPolicy Bypass -File sources\build_docx.ps1    # DOCX (Word COM)
```

La version actuelle du guide (V3) est tenue sur Canva ; voir `CLAUDE.md`.

## Conventions

Niveaux de confiance utilisés dans le guide : **Confirmé** / **Déduit** / **À compléter**.
Une déduction n'est jamais présentée comme confirmée.
