# Charte éthique et durabilité numérique — Numérisation des titres de transport aux CFF

Travail réalisé dans le cadre du cours **Éthique et durabilité** (HEIG-VD).

**Groupe :** Sofian Ethenoz · Arthur Menétrey · Mirco Profico
**Sujet :** la numérisation des titres de transport aux CFF et ses enjeux éthiques pour les personnes vulnérables.

## Contenu du dépôt

| Fichier | Description |
|---|---|
| [`charte_ethique_CFF.pdf`](charte_ethique_CFF.pdf) | **Charte éthique** (document autonome à rendre) — 3 pages : préambule, valeurs, 6 axes d'engagements. |
| [`rapport_ethique_CFF.pdf`](rapport_ethique_CFF.pdf) | **Rapport écrit final** (à rendre) — présentation, explication des choix, conclusion, sources. |
| [`presentation_charte_CFF.pptx`](presentation_charte_CFF.pptx) | **Présentation** (6 diapositives) avec le script de chaque voix dans les notes du présentateur — support à enregistrer. |
| [`script_video.md`](script_video.md) | **Script de la vidéo** réparti en 3 voix, avec minutage (≈ 2 min 45 – 2 min 57). |
| `charte_ethique_CFF.html` | Source de la charte (HTML + CSS, sert à générer le PDF de la charte). |
| `question_charte.md` | Réponses sourcées aux questions guidées ayant servi de base à la charte. |
| `rapport_ethique_CFF.md` | Source du rapport (Markdown). |
| `rapport.html` | Version HTML mise en page (sert à générer le PDF du rapport). |
| `build_pptx.py` | Script de génération de la présentation. |

## Régénérer les fichiers

### PDF de la charte et du rapport (HTML → PDF)

Les PDF sont produits à partir des fichiers HTML via un navigateur Chromium en mode « headless »
(option `--print-to-pdf`). Le CSS gère déjà le format A4 et les sauts de page.

**Windows (PowerShell) — avec Chrome ou Edge :**

```powershell
# Adapter le chemin du navigateur si nécessaire (Chrome ou msedge.exe)
$chrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

# PDF de la charte
& $chrome --headless=new --disable-gpu --no-pdf-header-footer `
  --print-to-pdf="charte_ethique_CFF.pdf" "charte_ethique_CFF.html"

# PDF du rapport
& $chrome --headless=new --disable-gpu --no-pdf-header-footer `
  --print-to-pdf="rapport_ethique_CFF.pdf" "rapport.html"
```

**macOS / Linux :**

```bash
chromium --headless --no-pdf-header-footer \
  --print-to-pdf=charte_ethique_CFF.pdf charte_ethique_CFF.html
chromium --headless --no-pdf-header-footer \
  --print-to-pdf=rapport_ethique_CFF.pdf rapport.html
```

### Présentation PPTX (nécessite python-pptx)

```bash
python3 -m venv .venv && .venv/bin/pip install python-pptx
.venv/bin/python build_pptx.py
```

## Rendu

Délai : **vendredi 12 juin, 12 h 00** (Cyberlearn) — formats `.pdf` (rapport) et `.pptx`/`.mp4` (présentation avec voix enregistrées).
