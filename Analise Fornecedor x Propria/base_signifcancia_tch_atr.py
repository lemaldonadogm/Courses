import pandas as pd
import scipy.stats as stats
import matplotlib.pyplot as plt
import seaborn as sns

# =====================================================
# 1. IMPORTAR BASE
# =====================================================
df = pd.read_excel("base_analises_prop_forn.xlsx")

# =====================================================
# 2. LIMPEZA (Base_Limpa)
# =====================================================
cols = [
    "dt_ano_safra",
    "ds_prop_forn",
    "vl_area_bloco",
    "vl_pd_real",
    "atr_cana",
    "vl_tch_bloco"
]

df = df[cols].copy()

for c in ["vl_area_bloco", "vl_pd_real", "atr_cana", "vl_tch_bloco"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")

df = df.dropna()
df = df[(df["vl_area_bloco"] > 0) & (df["vl_pd_real"] > 0)]
df["ds_prop_forn"] = df["ds_prop_forn"].str.lower()

# =====================================================
# 3. FUNÇÃO GENÉRICA PARA WELCH + BOXPLOT
# =====================================================
def analise_welch(
    df,
    variavel,
    nome_variavel,
    arquivo_excel,
    arquivo_figura
):

    resultados = []

    for safra, df_safra in df.groupby("dt_ano_safra"):

        propria = df_safra[df_safra["ds_prop_forn"] == "própria"][variavel]
        fornecedor = df_safra[df_safra["ds_prop_forn"] == "fornecedor"][variavel]

        if len(propria) > 30 and len(fornecedor) > 30:
            teste = stats.ttest_ind(
                propria,
                fornecedor,
                equal_var=False
            )

            resultados.append({
                "Safra": safra,
                f"{nome_variavel}_Medio_Propria": propria.mean(),
                f"{nome_variavel}_Medio_Fornecedor": fornecedor.mean(),
                "Diferença": fornecedor.mean() - propria.mean(),
                "p_value": teste.pvalue
            })

    tabela = pd.DataFrame(resultados).sort_values("Safra")

    # =========================
    # Exportar tabela
    # =========================
    tabela.to_excel(arquivo_excel, index=False)

    # =========================
    # Boxplot
    # =========================
    plt.figure(figsize=(14, 6))
    sns.set(style="whitegrid")

    ax = sns.boxplot(
        data=df,
        x="dt_ano_safra",
        y=variavel,
