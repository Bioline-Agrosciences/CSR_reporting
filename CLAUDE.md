# CLAUDE.md

Reporting RSE mensuel de Bioline Agrosciences (6 BU + Safety). Le README
décrit le fonctionnement ; ce fichier ne garde que ce qui sert à travailler
sur le code.

## Où tourne quoi

- **Local** : seulement `scripts/generate_data_entry_file.py` (+
  `config.py`), qui écrit les fichiers de saisie dans SharePoint.
- **Fabric** (workspace BM_F_D - SAP-B1, lakehouse LH_CSR_Reporting) :
  `scripts/nb_consolidate_csr_data.py` (consolidation, sommes) puis
  `scripts/nb_consolidate_sap_and_csr_data.py` (ventes SAP + 7 ratios).
- `archive/` : scripts de migration et ancienne consolidation locale, gelés.
  Ne pas les modifier ni s'en servir comme référence de la logique actuelle
  (c'est le notebook qui fait foi). Historique des décisions :
  `archive/DECISIONS.md`.

## Copie manuelle vers Fabric

Les notebooks Fabric ne sont PAS synchronisés avec le dépôt : l'utilisatrice
recopie les cellules à la main depuis les `nb_*.py`, découpés en blocs
`# Cell N — ...`.

- Après toute modification d'un `nb_*.py`, lister précisément les cellules
  à recopier (numéro, quoi ajouter/supprimer), en distinguant les
  changements de code des changements de commentaires seuls.
- Ce qui se tape dans Fabric mais n'est pas du Python (`%pip ...`) est mis
  en commentaire dans le `.py`.

## Règles à respecter

- **Uniquement des sommes dans la consolidation**, les ratios dans le
  second notebook (un ratio ne s'additionne pas entre BU).
- **Clé (BU, Year, Month, ID)** partout.
- **Facteurs de conversion/émission** : dans `input_data/CSR_parameters.xlsx`
  (par BU et par année), jamais codés en dur.
- **`holidays` figé à 0.105** dans Fabric, identique à `uv.lock` : les deux
  se mettent à jour ensemble, sinon Saf.4.1 diverge.
- **Aucune donnée réelle dans git** : tout est sur SharePoint « CSR
  referents » / `General/` ; le dépôt est cloné hors OneDrive.
- **Documentation** : les en-têtes de fichiers et le README décrivent
  uniquement l'état actuel (quoi, où, comment) pour quelqu'un qui découvre
  le code. Le pourquoi et l'historique vont dans `archive/DECISIONS.md`
  (ajouter une entrée datée à chaque nouvelle décision).

## Commandes

```
uv sync
uv run scripts/generate_data_entry_file.py [BU ...]
uv run pytest                  # tests actifs (CI)
uv run pytest archive/tests    # tests des scripts archivés, à la main
```

## Git / GitHub

- Dépôt : github.com/Bioline-Agrosciences/CSR_reporting, PR vers `main`.
- Le GitHub CLI (`gh`) n'est pas installé : fournir le lien de comparaison
  et un titre/description à coller au lieu d'ouvrir la PR.
- Messages de commit en anglais ; commentaires des notebooks en français ;
  README et docstrings des scripts en anglais.

## Langue

L'utilisatrice écrit en français : répondre en français.
