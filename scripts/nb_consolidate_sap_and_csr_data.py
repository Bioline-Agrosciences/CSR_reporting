# Fabric notebook: nb_consolidate_sap_and_csr_data
# Workspace: BM_F_D - SAP-B1 / Lakehouse: LH_CSR_Reporting
#
# Version pandas : le notebook tourne toujours sur Spark (il faut donc bien
# avoir attaché LH_CSR_Reporting comme lakehouse par défaut), mais toute la
# logique entre la lecture et l'écriture est en pandas — mêmes réflexes que
# csr_calc_engine.py. Conversion Spark -> pandas juste après la lecture
# (.toPandas()), et pandas -> Spark juste avant l'écriture finale.
#
# IMPORTANT — quels ratios ont vraiment besoin d'être recalculés ici
# (mis à jour 23/09/2026 — le périmètre a changé, relire même si tu connais
# la version précédente de ce commentaire) :
# Décision d'Aurélie : scripts/csr_calc_engine.py ne calcule plus AUCUN
# ratio/division, seulement des sommes — un ratio par BU ne peut pas être
# sommé/moyenné pour obtenir un total groupe valide (contrairement à une
# somme), donc CSR_indicators_report ne contient plus AUCUN des 7 ratios
# suivants (avant le 23/09, seuls Wat.2/Ene.11/Was.4 en étaient absents,
# faute d'Env.1 ; Ene.10/Was.3/Saf.6/Saf.7 arrivaient déjà calculés). Les 7
# doivent maintenant être recalculés ici, à partir des composantes brutes
# déjà dans CSR_indicators_report (+ Env.1 pour les 3 qui en dépendent) :
#   Wat.2  = Wat.1 / Env.1                          (a besoin d'Env.1/SAP)
#   Ene.10 = (Ene.2 + Ene.3 + Ene.4) / Ene.9
#   Ene.11 = Ene.9 / Env.1                          (a besoin d'Env.1/SAP)
#   Was.3  = Was.2 / Was.1
#   Was.4  = Was.1 / Env.1                          (a besoin d'Env.1/SAP)
#   Saf.6  = (Saf.2 + Saf.3) / Saf.4.2 * 1 000 000
#   Saf.7  = Saf.5 / Saf.4.2 * 1 000
# (formules copiées telles quelles depuis l'ancien scripts/csr_calc_engine.py,
# avant qu'il ne soit restreint aux sommes — Saf.6/Saf.7 n'ont plus besoin du
# fallback Env.1×Saf.4.1×8 : Saf.4.2 est désormais toujours renseigné, réel,
# via Working_Hours_2026.xlsx côté pipeline Python.)
#
# IMPORTANT — l'année (22/09/2026) : Raw_data_CSR.xlsx / Consolidated_results_CSR.xlsx
# côté pipeline Python portent désormais une colonne "Year" (BU+Année+Mois+ID,
# plus seulement BU+Mois), pour ne pas écraser silencieusement une année sur
# l'autre. CSR_indicators_report doit donc, elle aussi, porter cette colonne
# "Year" une fois le dataflow Fabric rafraîchi — sans quoi le pivot ci-dessous
# fusionnerait par erreur 2026 et 2027 sur la même ligne BU/Mois. Si la
# colonne "Year" n'apparaît pas encore dans CSR_indicators_report après un
# refresh du dataflow, il faut d'abord aller la détecter dans Power Query
# (colonne source ajoutée après la dernière génération des requêtes).
#
# CSR_completion_tracking est volontairement laissé de côté (suivi de
# complétion, pas un indicateur de reporting) — à joindre séparément côté
# Power BI si besoin d'un visuel de statut.

# --------------------------------------------------------------------------
# Cell 1 — charger les 3 sources, tout de suite converti en pandas
# --------------------------------------------------------------------------
df_indicators = spark.read.table("CSR_indicators_report").toPandas()
df_tracking = spark.read.table("CSR_completion_tracking").toPandas()  # non utilisé plus bas, gardé pour référence
df_sales = spark.read.table("f_Sales").toPandas()

# --------------------------------------------------------------------------
# Cell 2 — isoler une ligne propre par (BU, Year, Month) depuis f_Sales
# --------------------------------------------------------------------------
# TODO — A ADAPTER une fois le schéma réel confirmé.
# Lance `df_sales.head()` et `df_sales.dtypes` (en pandas, plus besoin de
# printSchema/display) dans une cellule à part, et envoie-moi le résultat :
# je remplirai précisément les noms ci-dessous. Pour l'instant on suppose
# une colonne type BU, une colonne type année, une colonne type mois/période,
# une colonne numérique de ventes — mêmes codes BU
# (BAF/BFR/BIB/BUK/BUS/Viridaxis) et mêmes noms de mois (January...December)
# que dans CSR_indicators_report.

df_sales_monthly = (
    df_sales
    .groupby(["BU", "Year", "Month"], as_index=False)["Value"]  # <-- ADAPT: vrais noms de colonnes
    .sum()
    .rename(columns={"Value": "Env.1"})                          # <-- ADAPT: vrai nom de la colonne de ventes
)

# --------------------------------------------------------------------------
# Cell 3 — pivot des indicateurs en large, jointure des vraies ventes,
# recalcul des 7 ratios (par BU/Année/Mois, jamais mélangé entre années,
# jamais sommé/moyenné entre BU — voir le commentaire en tête de fichier)
# --------------------------------------------------------------------------
df_wide = (
    df_indicators
    .pivot_table(index=["BU", "Year", "Month"], columns="ID", values="Value", aggfunc="first")
    .reset_index()
)
df_wide.columns.name = None

# CSR_indicators_report ne contient plus Env.1 du tout (exclu côté pipeline
# Python, voir extract_reference_and_data.py) — on l'ajoute ici pour la
# première fois, depuis le vrai chiffre SAP, année par année. errors="ignore"
# reste défensif si jamais un refresh plus ancien du dataflow en portait
# encore un résidu.
df_wide = (
    df_wide.drop(columns=["Env.1"], errors="ignore")
    .merge(df_sales_monthly, on=["BU", "Year", "Month"], how="left")
)

df_wide["Wat.2"] = df_wide["Wat.1"] / df_wide["Env.1"]
df_wide["Ene.10"] = (df_wide["Ene.2"] + df_wide["Ene.3"] + df_wide["Ene.4"]) / df_wide["Ene.9"]
df_wide["Ene.11"] = df_wide["Ene.9"] / df_wide["Env.1"]
df_wide["Was.3"] = df_wide["Was.2"] / df_wide["Was.1"]
df_wide["Was.4"] = df_wide["Was.1"] / df_wide["Env.1"]
df_wide["Saf.6"] = (df_wide["Saf.2"] + df_wide["Saf.3"]) / df_wide["Saf.4.2"] * 1_000_000
df_wide["Saf.7"] = df_wide["Saf.5"] / df_wide["Saf.4.2"] * 1_000

# --------------------------------------------------------------------------
# Cell 4 — repasser en format long (une ligne par BU/Year/Month/ID/Value,
# même forme que Consolidated_results_CSR.xlsx / Power BI)
# --------------------------------------------------------------------------
id_cols = ["BU", "Year", "Month"]
value_cols = [c for c in df_wide.columns if c not in id_cols]

df_gold_pd = (
    df_wide
    .melt(id_vars=id_cols, value_vars=value_cols, var_name="ID", value_name="Value")
    .dropna(subset=["Value"])
)

# --------------------------------------------------------------------------
# Cell 5 — reconvertir en Spark et écrire la table Delta "gold"
# --------------------------------------------------------------------------
df_gold = spark.createDataFrame(df_gold_pd)
df_gold.write.format("delta").mode("overwrite").saveAsTable("CSR_gold_reporting")
