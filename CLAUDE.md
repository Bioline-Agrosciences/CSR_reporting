# CLAUDE.md

Reporting RSE mensuel de Bioline Agrosciences (6 BU + Safety). Le README
décrit le fonctionnement ; ce fichier ne garde que ce qui sert à travailler
sur le code.

## Où tourne quoi

- **Local** : seulement `scripts/generate_data_entry_file.py` (+
  `config.py`), qui écrit les fichiers de saisie dans SharePoint.
- **Fabric** (workspace BM_F_D - SAP-B1, lakehouse LH_CSR_Reporting) :
  `scripts/nb_consolidate_csr_data.ipynb` (consolidation, sommes) puis
  `scripts/nb_consolidate_data_for_CSR_report.ipynb` (tables du tableau de
  bord : CSR_gold_reporting avec les ventes depuis le modèle sémantique
  commercial, CSR_gold_tracking ; aucun calcul).
- **Power BI** : les ratios, calculés à la volée dans le tableau de bord.
- `archive/` : scripts de migration et ancienne consolidation locale, gelés.
  Ne pas les modifier ni s'en servir comme référence de la logique actuelle
  (c'est le notebook qui fait foi). Historique des décisions :
  `archive/DECISIONS.md`.

## Notebooks Fabric

Fabric n'est pas connecté au dépôt. Les notebooks sont versionnés en
`.ipynb` exportés depuis Fabric, **toujours sans sorties** (elles
contiennent des données réelles : ventes, valeurs RSE).

- Fabric -> dépôt : l'utilisatrice exporte le notebook, puis
  `import-fabric-notebook <fichier.ipynb> scripts/` (commande du paquet
  bioline_utils, installée via `uv tool install -e`). Un test échoue si un
  notebook a encore des sorties.
- Dépôt -> Fabric : après une modification d'un `.ipynb` ici, lui dire de
  réimporter le notebook dans Fabric, ou lister les cellules à modifier
  (en les désignant par leur contenu, elles ne sont pas numérotées), en
  distinguant code et commentaires seuls.
- Modifier un `.ipynb` par script Python (json), en gardant les métadonnées
  du notebook (lakehouse par défaut) et les ids de cellules existants.

## Règles à respecter

- **Uniquement des sommes dans Fabric**, les ratios dans le tableau de bord
  Power BI (un ratio ne s'additionne pas entre BU).
- **Clé (BU, Year, Month, ID)** partout.
- **Facteurs de conversion/émission** : dans `input_data/CSR_parameters.xlsx`
  (par BU et par année), jamais codés en dur.
- **`holidays` figé à 0.105** dans l'Environment Fabric du notebook de
  consolidation, identique à `uv.lock` : les deux se mettent à jour
  ensemble, sinon Saf.4.1 diverge. Pas de `%pip` dans les notebooks : il est
  bloqué quand le pipeline les lance.
- **Pipeline quotidien à 10:30** (après la mise à jour des ventes vers 10h) :
  df_CSR_dimensions → nb_consolidate_csr_data →
  nb_consolidate_data_for_CSR_report → actualisation de ms_csr_reporting.
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
