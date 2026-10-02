# Fabric notebook: nb_consolidate_sap_and_csr_data
# Workspace: BM_F_D - SAP-B1 / Lakehouse: LH_CSR_Reporting
#
# À lancer juste après nb_consolidate_csr_data, dans le même pipeline.
#
# CSR_indicators_report ne contient que des sommes. Ce notebook y joint les
# ventes SAP (Env.1, depuis f_Sales) et calcule les 7 ratios, par BU, année
# et mois, à partir des composantes brutes :
#   Wat.2  = Wat.1 / Env.1
#   Ene.10 = (Ene.2 + Ene.3 + Ene.4) / Ene.9
#   Ene.11 = Ene.9 / Env.1
#   Was.3  = Was.2 / Was.1
#   Was.4  = Was.1 / Env.1
#   Saf.6  = (Saf.2 + Saf.3) / Saf.4.2 * 1 000 000
#   Saf.7  = Saf.5 / Saf.4.2 * 1 000
# Un ratio ne s'additionne ni ne se moyenne entre BU : pour un total groupe,
# le recalculer à partir des sommes (mesure DAX dans Power BI).
#
# Écrit la table Delta CSR_gold_reporting (format long : BU, Year, Month,
# ID, Value).
#
# Lecture en Spark, logique en pandas, écriture en Spark. Lakehouse par
# défaut à attacher : LH_CSR_Reporting.

# --------------------------------------------------------------------------
# Cell 1 — charger les sources, tout de suite converti en pandas
# --------------------------------------------------------------------------
df_indicators = spark.read.table("CSR_indicators_report").toPandas()
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
# calcul des 7 ratios (par BU/Année/Mois — voir l'en-tête)
# --------------------------------------------------------------------------
df_wide = (
    df_indicators
    .pivot_table(index=["BU", "Year", "Month"], columns="ID", values="Value", aggfunc="first")
    .reset_index()
)
df_wide.columns.name = None

# CSR_indicators_report ne contient pas Env.1 : il vient uniquement de SAP.
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
# même forme que CSR_indicators_report)
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
