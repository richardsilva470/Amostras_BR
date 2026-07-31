import io
import re
import unicodedata
import numpy as np
from scipy.stats import norm

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Desenho Amostral Eleitoral", page_icon="📊", layout="wide")

# =========================================================================== #
# Regras de negócio (fixas — mesma metodologia usada na planilha do Acre)
# =========================================================================== #

FUNDAMENTAL = {"ANALFABETO", "LÊ E ESCREVE", "ENSINO FUNDAMENTAL INCOMPLETO", "ENSINO FUNDAMENTAL COMPLETO"}
MEDIO = {"ENSINO MÉDIO INCOMPLETO", "ENSINO MÉDIO COMPLETO"}
SUPERIOR = {"SUPERIOR INCOMPLETO", "SUPERIOR COMPLETO"}

# padrões (regex) usados para reconhecer as colunas de faixas de renda da
# Tabela 10300 do IBGE/SIDRA, independentemente das letras de nota de rodapé
# que o IBGE anexa ao final do nome da coluna (ex.: "...salário mínimo e")
PADROES_RENDA = {
    "D/E": [r"at[ée]\s*1/4", r"1/4\s*a\s*1/2", r"1/2\s*a\s*1\s*sal", r"1\s*a\s*2\s*sal", r"sem\s*rendimento"],
    "C": [r"2\s*a\s*3\s*sal", r"3\s*a\s*5\s*sal"],
    "B": [r"5\s*a\s*10\s*sal", r"10\s*a\s*15\s*sal", r"15\s*a\s*20\s*sal"],
    "A": [r"mais\s*de\s*20\s*sal"],
}

# classificação alternativa de renda, em 3 faixas de salário mínimo (SM)
PADROES_RENDA_SM = {
    "ATÉ 2 SM": [r"at[ée]\s*1/4", r"1/4\s*a\s*1/2", r"1/2\s*a\s*1\s*sal", r"1\s*a\s*2\s*sal", r"sem\s*rendimento"],
    "2 A 5 SM": [r"2\s*a\s*3\s*sal", r"3\s*a\s*5\s*sal"],
    "MAIS DE 5 SM": [r"5\s*a\s*10\s*sal", r"10\s*a\s*15\s*sal", r"15\s*a\s*20\s*sal", r"mais\s*de\s*20\s*sal"],
}

FAIXAS_IDADE_ORDEM = ["16 a 24 anos", "25 a 34 anos", "35 a 44 anos", "45 a 59 anos", "60 anos ou mais"]


# =========================================================================== #
# Funções auxiliares
# =========================================================================== #

def normalizar_municipio(nome):
    """Remove sufixo de UF (ex. '(AC)'), acentos e caixa para permitir o cruzamento
    entre bases que grafam o nome do município de formas ligeiramente diferentes."""
    if pd.isna(nome):
        return nome
    nome = str(nome)
    nome = re.sub(r"\s*\([A-Za-zÀ-ÿ]{2}\)\s*$", "", nome)
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("utf-8")
    return re.sub(r"\s+", " ", nome).strip().upper()


def bucket_faixa_etaria(valor):
    """Converte valores como '17 anos', '21 a 24 anos', '100 anos ou mais' nas
    5 faixas etárias usadas no plano amostral."""
    if pd.isna(valor):
        return None
    m = re.search(r"(\d+)", str(valor))
    if not m:
        return None
    idade = int(m.group(1))
    if idade < 16:
        return None
    if idade <= 24:
        return "16 a 24 anos"
    if idade <= 34:
        return "25 a 34 anos"
    if idade <= 44:
        return "35 a 44 anos"
    if idade <= 59:
        return "45 a 59 anos"
    return "60 anos ou mais"


def bucket_escolaridade(valor):
    v = str(valor).strip().upper()
    if v in FUNDAMENTAL:
        return "fundamental"
    if v in MEDIO:
        return "médio"
    if v in SUPERIOR:
        return "superior"
    return None


def bucket_cor_raca(valor):
    v = str(valor).strip()
    if v in ("Amarela", "Indígena"):
        return "Amarelo + outros"
    return v


def _eh_numerico_ou_traco(v):
    if pd.isna(v):
        return True
    s = str(v).strip()
    if s in ("-", "..", "...", "X", ""):
        return True
    try:
        float(s.replace(".", "").replace(",", "."))
        return True
    except ValueError:
        return False


def classificar_colunas_renda(colunas, padroes_renda=PADROES_RENDA):
    """Agrupa as colunas de faixa de renda da tabela 10300 do IBGE segundo o
    dicionário de padrões informado (por padrão, Classe A/B/C/D-E), reconhecendo
    o texto da coluna por regex (robusto a diferentes letras de nota de rodapé
    entre exportações)."""
    grupos = {classe: [] for classe in padroes_renda}
    for col in colunas:
        col_norm = str(col).strip().lower()
        for classe, padroes in padroes_renda.items():
            if any(re.search(p, col_norm) for p in padroes):
                grupos[classe].append(col)
                break
    return grupos


def numerico(serie):
    return pd.to_numeric(serie, errors="coerce").fillna(0)


# --------------------------------------------------------------------------- #
# Leitura de arquivos
# --------------------------------------------------------------------------- #

def carregar_csv(uploaded_file):
    """Lê CSV detectando separador e codificação automaticamente."""
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    uploaded_file.seek(0)

    text = None
    for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw_bytes.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        st.error("Não foi possível ler o conteúdo do arquivo com uma codificação suportada.")
        st.stop()

    import csv
    try:
        delimiter = csv.Sniffer().sniff(text[:4096]).delimiter
        df = pd.read_csv(io.StringIO(text), sep=delimiter)
    except csv.Error:
        df = pd.read_csv(io.StringIO(text), sep=None, engine="python")
    return df


def carregar_tabela_ibge(uploaded_file, sheet_name=None):
    """Lê um export do SIDRA/IBGE (xlsx) detectando automaticamente em qual
    linha começa o cabeçalho de verdade — esses arquivos sempre trazem
    algumas linhas de título/metadado antes da tabela propriamente dita."""
    xl = pd.ExcelFile(uploaded_file)
    if sheet_name is None:
        sheet_name = "Tabela" if "Tabela" in xl.sheet_names else xl.sheet_names[0]
    raw = xl.parse(sheet_name, header=None)

    primeira_linha_dado = None
    for i in range(len(raw)):
        col0 = raw.iat[i, 0]
        if pd.isna(col0):
            continue
        s = str(col0).strip()
        if len(s) == 0 or len(s) > 60:
            continue
        if s.lower() in ("total", "município", "municípios") or s.isdigit():
            continue
        outros = raw.iloc[i, 1:]
        if len(outros) == 0 or outros.notna().sum() == 0:
            continue
        frac_numerico = sum(_eh_numerico_ou_traco(v) for v in outros) / len(outros)
        if frac_numerico >= 0.6:
            primeira_linha_dado = i
            break

    if primeira_linha_dado is None or primeira_linha_dado == 0:
        return xl.parse(sheet_name)

    linha_cabecalho = primeira_linha_dado - 1
    bloco_cabecalho = raw.iloc[0:linha_cabecalho + 1]
    colunas = []
    for j in range(raw.shape[1]):
        val = bloco_cabecalho.iloc[-1, j]
        if pd.isna(val):
            for k in range(linha_cabecalho, -1, -1):
                v2 = bloco_cabecalho.iloc[k, j]
                if pd.notna(v2):
                    val = v2
                    break
        colunas.append(str(val).strip() if pd.notna(val) else f"col_{j}")
    colunas[0] = "Município"

    dados = raw.iloc[primeira_linha_dado:].reset_index(drop=True)
    dados.columns = colunas
    return dados


def carregar_mapa_regioes(uploaded_file):
    """Lê o arquivo simples de mapeamento Município -> Microrregião (2 colunas)."""
    if uploaded_file.name.lower().endswith(".csv"):
        df = carregar_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    df = df.iloc[:, :2].copy()
    df.columns = ["Município", "Microrregião"]
    df["muni_norm"] = df["Município"].apply(normalizar_municipio)
    return df


def preparar_municipio(df, mapa, dropar_col_regiao_existente=True):
    """Cruza um dataframe com o mapa de microrregiões pelo nome do município."""
    df = df.copy()
    if dropar_col_regiao_existente:
        cols_regiao = [c for c in df.columns if "microrreg" in str(c).lower() or "mesorreg" in str(c).lower()]
        df = df.drop(columns=cols_regiao, errors="ignore")
    df["muni_norm"] = df["Município"].apply(normalizar_municipio)
    antes = len(df)
    df = df.merge(mapa[["muni_norm", "Microrregião"]], on="muni_norm", how="left")
    sem_match = df[df["Microrregião"].isna()]["Município"].dropna().unique().tolist()
    return df, sem_match


# =========================================================================== #
# Sidebar — upload dos 5 arquivos padrão (mesmo formato para qualquer estado)
# =========================================================================== #

st.sidebar.title("Menu")
st.sidebar.header("1. Mapa de Microrregiões")
st.sidebar.caption("2 colunas: Município e Microrregião/Mesorregião do estado.")
arq_mapa = st.sidebar.file_uploader("Município → Microrregião", type=["csv", "xlsx", "xls"], key="mapa")

st.sidebar.header("2. Eleitorado (TSE)")
arq_tse = st.sidebar.file_uploader("Base de eleitores (CSV)", type=["csv"], key="tse")

st.sidebar.header("3. PEA / NPEA (IBGE)")
arq_pea = st.sidebar.file_uploader("Tabela 10300 - Força de trabalho", type=["xlsx", "xls"], key="pea")

st.sidebar.header("4. Religião (IBGE)")
arq_rel = st.sidebar.file_uploader("Tabela 3458 - Religião", type=["xlsx", "xls"], key="rel")

st.sidebar.header("5. Renda / Salário (IBGE)")
arq_sal = st.sidebar.file_uploader("Tabela 10300 - Classes de rendimento", type=["xlsx", "xls"], key="sal")

st.sidebar.header("Desenho amostral")
n_amostra = st.sidebar.number_input("Tamanho da amostra", min_value=1, value=1000, step=1)
confianca = st.sidebar.number_input("Nível de confiança", min_value=0.0, max_value=1.0, value=0.95, step=0.01)
margem_erro = np.sqrt(0.25 / n_amostra) * -100 * norm.ppf((1 - confianca) / 2)


st.title("📊 Desenho Amostral Eleitoral — qualquer estado")
st.caption(
    "Carregue os 5 arquivos no menu à esquerda. O app cruza as bases pelo município, "
    "calcula a distribuição real da população/eleitorado por microrregião nas variáveis de "
    "ponderação (idade, gênero, escolaridade, classe social, PEA, cor/raça e religião) e monta "
    "a tabela final de amostra — do mesmo jeito que foi feito para o Acre."
)

arquivos_ok = all([arq_mapa, arq_tse, arq_pea, arq_rel, arq_sal])

if not arquivos_ok:
    st.info("Carregue os 5 arquivos no menu lateral para gerar a análise.")
    st.stop()

avisos_municipios_sem_match = {}

# --------------------------------------------------------------------------- #
# 1. Mapa de microrregiões
# --------------------------------------------------------------------------- #
mapa = carregar_mapa_regioes(arq_mapa)
microrregioes = sorted(mapa["Microrregião"].dropna().unique().tolist())

# --------------------------------------------------------------------------- #
# 2. Eleitorado (TSE)
# --------------------------------------------------------------------------- #
tse = carregar_csv(arq_tse)
colunas_esperadas_tse = {"Cor / Raça", "Faixa etária", "Gênero", "Grau de instrução", "Município", "Quantidade de eleitores"}
faltando = colunas_esperadas_tse - set(tse.columns)
if faltando:
    st.error(f"O arquivo de eleitorado não tem as colunas esperadas: {', '.join(faltando)}")
    st.stop()

tse["Quantidade de eleitores"] = numerico(tse["Quantidade de eleitores"])
tse, sem_match = preparar_municipio(tse, mapa)
if sem_match:
    avisos_municipios_sem_match["Eleitorado (TSE)"] = sem_match
tse = tse.dropna(subset=["Microrregião"])

total_eleitores = tse["Quantidade de eleitores"].sum()

# participação de cada microrregião no total do eleitorado
tab_total = tse.groupby("Microrregião")["Quantidade de eleitores"].sum() / total_eleitores

# faixa etária
tse["faixa_bucket"] = tse["Faixa etária"].apply(bucket_faixa_etaria)
tab_idade = (
    tse.dropna(subset=["faixa_bucket"])
    .groupby(["Microrregião", "faixa_bucket"])["Quantidade de eleitores"].sum()
    .unstack(fill_value=0) / total_eleitores
)
tab_idade = tab_idade.reindex(columns=FAIXAS_IDADE_ORDEM, fill_value=0)

# gênero
tse["genero_bucket"] = tse["Gênero"].str.upper().map({"MASCULINO": "Homens", "FEMININO": "Mulheres"}).fillna(tse["Gênero"])
tab_genero = (
    tse.groupby(["Microrregião", "genero_bucket"])["Quantidade de eleitores"].sum().unstack(fill_value=0) / total_eleitores
)

# escolaridade
tse["esc_bucket"] = tse["Grau de instrução"].apply(bucket_escolaridade)
tab_esc = (
    tse.dropna(subset=["esc_bucket"])
    .groupby(["Microrregião", "esc_bucket"])["Quantidade de eleitores"].sum()
    .unstack(fill_value=0) / total_eleitores
)
tab_esc = tab_esc.reindex(columns=["fundamental", "médio", "superior"], fill_value=0)

# cor/raça (exclui "não informado" e renormaliza)
tse_cr = tse[~tse["Cor / Raça"].astype(str).str.upper().str.contains("NAO INFORMADO|NÃO INFORMADO", na=False)].copy()
tse_cr["cor_bucket"] = tse_cr["Cor / Raça"].apply(bucket_cor_raca)
total_cr = tse_cr["Quantidade de eleitores"].sum()
tab_cor = (
    tse_cr.groupby(["Microrregião", "cor_bucket"])["Quantidade de eleitores"].sum().unstack(fill_value=0) / total_cr
)
tab_cor = tab_cor.reindex(columns=["Branca", "Preta", "Parda", "Amarelo + outros"], fill_value=0)

# --------------------------------------------------------------------------- #
# 3. PEA / NPEA
# --------------------------------------------------------------------------- #
pea = carregar_tabela_ibge(arq_pea)
if "Força de trabalho" not in pea.columns or "Fora da força de trabalho" not in pea.columns:
    st.error("O arquivo de PEA/NPEA não tem as colunas 'Força de trabalho' e 'Fora da força de trabalho'.")
    st.stop()
pea, sem_match = preparar_municipio(pea, mapa)
if sem_match:
    avisos_municipios_sem_match["PEA/NPEA"] = sem_match
pea = pea.dropna(subset=["Microrregião"])
pea["Força de trabalho"] = numerico(pea["Força de trabalho"])
pea["Fora da força de trabalho"] = numerico(pea["Fora da força de trabalho"])
total_pea = pea["Força de trabalho"].sum() + pea["Fora da força de trabalho"].sum()
tab_pea = pea.groupby("Microrregião")[["Força de trabalho", "Fora da força de trabalho"]].sum() / total_pea
tab_pea = tab_pea.rename(columns={"Força de trabalho": "PEA", "Fora da força de trabalho": "NPEA"})

# --------------------------------------------------------------------------- #
# 4. Religião
# --------------------------------------------------------------------------- #
rel = carregar_tabela_ibge(arq_rel)
rel, sem_match = preparar_municipio(rel, mapa)
if sem_match:
    avisos_municipios_sem_match["Religião"] = sem_match
rel = rel.dropna(subset=["Microrregião"])
colunas_religiao = [c for c in rel.columns if c not in ("Município", "Total", "muni_norm", "Microrregião")]
for c in colunas_religiao:
    rel[c] = numerico(rel[c])
total_rel = rel[colunas_religiao].sum().sum()
tab_rel = rel.groupby("Microrregião")[colunas_religiao].sum() / total_rel

# --------------------------------------------------------------------------- #
# 5. Renda / Classe social
# --------------------------------------------------------------------------- #
sal = carregar_tabela_ibge(arq_sal)
sal, sem_match = preparar_municipio(sal, mapa)
if sem_match:
    avisos_municipios_sem_match["Renda/Salário"] = sem_match
sal = sal.dropna(subset=["Microrregião"])
colunas_renda = [c for c in sal.columns if c not in ("Município", "Total", "muni_norm", "Microrregião")]
grupos_renda = classificar_colunas_renda(colunas_renda)
faltando_grupos = [k for k, v in grupos_renda.items() if not v]
if faltando_grupos:
    st.warning(
        f"Não reconheci automaticamente as colunas de renda para: {', '.join(faltando_grupos)}. "
        "Confira se o arquivo segue o padrão da Tabela 10300 do IBGE."
    )
for classe, cols in grupos_renda.items():
    for c in cols:
        sal[c] = numerico(sal[c])
    sal[f"Classe {classe}"] = sal[cols].sum(axis=1) if cols else 0

colunas_classe = [f"Classe {c}" for c in ["A", "B", "C", "D/E"]]
total_sal = sal[colunas_classe].sum().sum()
tab_classe = sal.groupby("Microrregião")[colunas_classe].sum() / total_sal

# classificação alternativa: RENDA em 3 faixas de salário mínimo
grupos_renda_sm = classificar_colunas_renda(colunas_renda, PADROES_RENDA_SM)
faltando_grupos_sm = [k for k, v in grupos_renda_sm.items() if not v]
if faltando_grupos_sm:
    st.warning(
        f"Não reconheci automaticamente as colunas de renda (faixas de SM) para: {', '.join(faltando_grupos_sm)}. "
        "Confira se o arquivo segue o padrão da Tabela 10300 do IBGE."
    )
for faixa, cols in grupos_renda_sm.items():
    for c in cols:
        sal[c] = numerico(sal[c])
    sal[faixa] = sal[cols].sum(axis=1) if cols else 0

colunas_renda_sm = ["ATÉ 2 SM", "2 A 5 SM", "MAIS DE 5 SM"]
total_sal_sm = sal[colunas_renda_sm].sum().sum()
tab_renda_sm = sal.groupby("Microrregião")[colunas_renda_sm].sum() / total_sal_sm

# --------------------------------------------------------------------------- #
# Monta a tabela final (equivalente à aba "Amostra")
# --------------------------------------------------------------------------- #

tabela_final = pd.DataFrame(index=microrregioes)
tabela_final["Total"] = tab_total.reindex(microrregioes)
for col in tab_idade.columns:
    tabela_final[col] = tab_idade[col].reindex(microrregioes)
for col in ["Homens", "Mulheres"]:
    tabela_final[col] = tab_genero[col].reindex(microrregioes) if col in tab_genero.columns else 0
for col in ["fundamental", "médio", "superior"]:
    tabela_final[col] = tab_esc[col].reindex(microrregioes)
for col in colunas_classe:
    tabela_final[col] = tab_classe[col].reindex(microrregioes)
for col in colunas_renda_sm:
    tabela_final[col] = tab_renda_sm[col].reindex(microrregioes)
for col in ["PEA", "NPEA"]:
    tabela_final[col] = tab_pea[col].reindex(microrregioes)
for col in ["Branca", "Preta", "Parda", "Amarelo + outros"]:
    tabela_final[col] = tab_cor[col].reindex(microrregioes)
for col in colunas_religiao:
    tabela_final[col] = tab_rel[col].reindex(microrregioes)

tabela_final = tabela_final.fillna(0)
linha_total = tabela_final.sum(numeric_only=True)
linha_total.name = "Total Geral"
tabela_final = pd.concat([tabela_final, linha_total.to_frame().T])

# alocação da amostra por microrregião (maior resto, para fechar em n_amostra)
alvo = (tabela_final.loc[microrregioes, "Total"] * n_amostra)
alocacao = alvo.round().astype(int)
diferenca = int(n_amostra) - alocacao.sum()
if diferenca != 0:
    restos = (alvo - alvo.round()).sort_values(ascending=(diferenca < 0))
    for muni in restos.index[:abs(diferenca)]:
        alocacao[muni] += 1 if diferenca > 0 else -1

# =========================================================================== #
# Exibição
# =========================================================================== #

st.subheader("Desenho amostral")
c1, c2, c3 = st.columns(3)
c1.metric("Amostra", f"{n_amostra}")
c2.metric("Nível de confiança", f"{confianca:.0%}")
c3.metric("Margem de erro", f"{margem_erro:.1f}%")

if avisos_municipios_sem_match:
    with st.expander("⚠️ Municípios que não foram encontrados no mapa de microrregiões"):
        for origem, lista in avisos_municipios_sem_match.items():
            st.write(f"**{origem}:** {', '.join(lista)}")
        st.caption(
            "Esses municípios ficaram de fora dos cálculos. Confira se o nome no mapa de "
            "microrregiões está grafado da mesma forma (acentos e UF entre parênteses são "
            "ignorados automaticamente)."
        )

st.subheader("Distribuição da amostra por microrregião")
tab_alocacao = alocacao.to_frame(name="Nº de entrevistas")
tab_alocacao["% da amostra"] = (tabela_final.loc[microrregioes, "Total"] * 100).round(1)
st.dataframe(tab_alocacao, use_container_width=True)

st.subheader("Tabela consolidada de ponderação (variáveis x microrregião)")
st.dataframe(tabela_final.style.format("{:.2%}"), use_container_width=True)

# --------------------------------------------------------------------------- #
# Download em Excel
# --------------------------------------------------------------------------- #

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    mapa.drop(columns=["muni_norm"]).to_excel(writer, sheet_name="Mesorregiões", index=False)
    tabela_final.to_excel(writer, sheet_name="Amostra")
    tab_alocacao.to_excel(writer, sheet_name="Alocação da amostra")
    tse.drop(columns=["muni_norm", "faixa_bucket", "genero_bucket", "esc_bucket"], errors="ignore").to_excel(
        writer, sheet_name="Dados TSE", index=False
    )
    pea.drop(columns=["muni_norm"], errors="ignore").to_excel(writer, sheet_name="PEA-NPEA", index=False)
    rel.drop(columns=["muni_norm"], errors="ignore").to_excel(writer, sheet_name="Religião", index=False)
    sal.drop(columns=["muni_norm"], errors="ignore").to_excel(writer, sheet_name="Renda", index=False)
buffer.seek(0)

st.download_button(
    "⬇️ Baixar planilha completa (.xlsx)",
    data=buffer,
    file_name="desenho_amostral.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)