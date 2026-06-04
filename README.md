# Charte éthique et durabilité numérique — Numérisation des titres de transport aux CFF

Travail réalisé dans le cadre du cours **Éthique et durabilité** (HEIG-VD).

**Groupe :** Sofian Ethenoz · Arthur Menétrey · Mirco Profico
**Sujet :** la numérisation des titres de transport aux CFF et ses enjeux éthiques pour les personnes vulnérables.

## Contenu du dépôt

| Fichier | Description |
|---|---|
| [`rapport_ethique_CFF.pdf`](rapport_ethique_CFF.pdf) | **Rapport écrit final** (à rendre) — présentation, charte en 13 engagements, explication des choix, conclusion, sources. |
| [`presentation_charte_CFF.pptx`](presentation_charte_CFF.pptx) | **Présentation** (6 diapositives) avec le script de chaque voix dans les notes du présentateur — support à enregistrer. |
| [`script_video.md`](script_video.md) | **Script de la vidéo** réparti en 3 voix, avec minutage (≈ 2 min 45 – 2 min 57). |
| `rapport_ethique_CFF.md` | Source du rapport (Markdown). |
| `rapport.html` | Version HTML mise en page (sert à générer le PDF). |
| `build_pptx.py` | Script de génération de la présentation. |

## Régénérer les fichiers

```bash
# PDF du rapport
chromium --headless --no-pdf-header-footer \
  --print-to-pdf=rapport_ethique_CFF.pdf rapport.html

# Présentation PPTX (nécessite python-pptx)
python3 -m venv .venv && .venv/bin/pip install python-pptx
.venv/bin/python build_pptx.py
```

## Rendu

Délai : **vendredi 12 juin, 12 h 00** (Cyberlearn) — formats `.pdf` (rapport) et `.pptx`/`.mp4` (présentation avec voix enregistrées).
