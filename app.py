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

# padrões (regex) que reconhecem as colunas de renda da Tabela 10300 do IBGE/SIDRA
# (robustos às letras de nota de rodapé) e as agrupam em 3 faixas de salário mínimo (SM)
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


def classificar_colunas_renda(colunas, padroes_renda=PADROES_RENDA_SM):
    """Agrupa as colunas de faixa de renda da tabela 10300 do IBGE segundo o
    dicionário de padrões informado (faixas de salário mínimo), reconhecendo
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


def _eh_rotulo_generico(v):
    """Rótulos de cabeçalho do SIDRA que não identificam a coluna (Total, ano)."""
    if pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "total" or bool(re.fullmatch(r"\d{4}(\.0)?", s))


def carregar_tabela_ibge(uploaded_file, sheet_name=None):
    """Lê um export do SIDRA/IBGE (xlsx) no formato padrão de download:
    linhas de título, cabeçalho em várias linhas (ex.: faixa de renda / ano /
    "Total"), dados por município "Nome (UF)" e rodapé "Fonte: IBGE ...".

    - detecta onde começam os dados;
    - escolhe, para cada coluna, o rótulo que realmente a identifica (a linha de
      cabeçalho com os nomes das categorias, e não "Total" ou o ano);
    - remove o rodapé e colunas só de texto (ex.: "Classes de rendimento" = Total)."""
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
        if frac_numerico >= 0.5:
            primeira_linha_dado = i
            break

    if primeira_linha_dado is None or primeira_linha_dado == 0:
        return xl.parse(sheet_name)

    cab = raw.iloc[:primeira_linha_dado, 1:]

    # linha de rótulos = a mais baixa do cabeçalho com 2+ nomes de categoria
    linha_rotulos = None
    for k in range(len(cab) - 1, -1, -1):
        especificos = [v for v in cab.iloc[k] if not _eh_rotulo_generico(v)]
        if len(set(map(str, especificos))) >= 2:
            linha_rotulos = k
            break

    colunas = ["Município"]
    for j in range(cab.shape[1]):
        val = cab.iat[linha_rotulos, j] if linha_rotulos is not None else np.nan
        if pd.isna(val):  # sem rótulo na linha principal: usa o mais baixo disponível
            for k in range(len(cab) - 1, -1, -1):
                if pd.notna(cab.iat[k, j]):
                    val = cab.iat[k, j]
                    break
        colunas.append(str(val).strip() if pd.notna(val) else f"col_{j + 1}")

    # nomes repetidos ganham sufixo para não sobrescrever colunas
    vistos = {}
    for i, c in enumerate(colunas):
        if c in vistos:
            vistos[c] += 1
            colunas[i] = f"{c} ({vistos[c]})"
        else:
            vistos[c] = 0

    dados = raw.iloc[primeira_linha_dado:].reset_index(drop=True)
    dados.columns = colunas

    # remove rodapé (Fonte, Notas...) e linhas sem município
    mun = dados["Município"].astype(str).str.strip()
    dados = dados[dados["Município"].notna() & ~mun.str.lower().str.match(r"^(fonte|nota|legenda|\(?\d+\))")]

    # remove colunas que não são numéricas (ex.: coluna com o texto "Total")
    manter = ["Município"]
    for c in dados.columns[1:]:
        vals = dados[c].dropna()
        if len(vals) and all(_eh_numerico_ou_traco(v) for v in vals):
            manter.append(c)
    return dados[manter].reset_index(drop=True)


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

    # nomes abreviados num arquivo e completos no outro (ex.: "São Luiz" no mapa/TSE
    # e "São Luiz do Anauá" no IBGE): casa pelo início do nome, se for único
    nomes_mapa = set(mapa["muni_norm"].dropna())
    def _casar(n):
        if pd.isna(n) or n in nomes_mapa:
            return n
        cands = [m for m in nomes_mapa if n.startswith(m + " ") or m.startswith(n + " ")]
        return cands[0] if len(cands) == 1 else n
    df["muni_norm"] = df["muni_norm"].apply(_casar)

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
arq_sal = st.sidebar.file_uploader("Tabela 10300 - Faixas de rendimento", type=["xlsx", "xls"], key="sal")

st.sidebar.header("Desenho amostral")
n_amostra = st.sidebar.number_input("Tamanho da amostra", min_value=1, value=1000, step=1)
confianca = st.sidebar.number_input("Nível de confiança", min_value=0.0, max_value=1.0, value=0.95, step=0.01)
margem_erro = np.sqrt(0.25 / n_amostra) * -100 * norm.ppf((1 - confianca) / 2)

st.sidebar.header("Ajuste dos subgrupos")
ajuste_automatico = st.sidebar.checkbox(
    "Ajustar automaticamente subgrupos que não batem com o Total da linha", value=True
)
tol_pp = st.sidebar.number_input(
    "Tolerância (pontos percentuais)", min_value=0.0, max_value=5.0, value=0.01, step=0.01, format="%.3f",
    help="Linhas cuja soma do subgrupo difere do Total por mais que isso são ajustadas.",
)
tol_ajuste = tol_pp / 100


st.title("📊 Desenho Amostral Eleitoral — qualquer estado")
st.caption(
    "Carregue os 5 arquivos no menu à esquerda. O app cruza as bases pelo município, "
    "calcula a distribuição real da população/eleitorado por microrregião nas variáveis de "
    "ponderação (idade, gênero, escolaridade, renda em faixas de SM, PEA, cor/raça e religião) e monta "
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
# 5. Renda (faixas de salário mínimo)
# --------------------------------------------------------------------------- #
sal = carregar_tabela_ibge(arq_sal)
sal, sem_match = preparar_municipio(sal, mapa)
if sem_match:
    avisos_municipios_sem_match["Renda/Salário"] = sem_match
sal = sal.dropna(subset=["Microrregião"])
colunas_renda = [c for c in sal.columns if c not in ("Município", "Total", "muni_norm", "Microrregião")]

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
# Monta a tabela "bruta" (exatamente como vem do IBGE/TSE, antes do ajuste)
# --------------------------------------------------------------------------- #

tabela_bruta = pd.DataFrame(index=microrregioes)
tabela_bruta["Total"] = tab_total.reindex(microrregioes)
for col in tab_idade.columns:
    tabela_bruta[col] = tab_idade[col].reindex(microrregioes)
for col in ["Homens", "Mulheres"]:
    tabela_bruta[col] = tab_genero[col].reindex(microrregioes) if col in tab_genero.columns else 0
for col in ["fundamental", "médio", "superior"]:
    tabela_bruta[col] = tab_esc[col].reindex(microrregioes)
for col in colunas_renda_sm:
    tabela_bruta[col] = tab_renda_sm[col].reindex(microrregioes)
for col in ["PEA", "NPEA"]:
    tabela_bruta[col] = tab_pea[col].reindex(microrregioes)
for col in ["Branca", "Preta", "Parda", "Amarelo + outros"]:
    tabela_bruta[col] = tab_cor[col].reindex(microrregioes)
for col in colunas_religiao:
    tabela_bruta[col] = tab_rel[col].reindex(microrregioes)
tabela_bruta = tabela_bruta.fillna(0)

# grupos de subcolunas que, somadas, devem bater com o "Total" de cada linha
GRUPOS = {
    "Faixa etária": FAIXAS_IDADE_ORDEM,
    "Gênero": ["Homens", "Mulheres"],
    "Escolaridade": ["fundamental", "médio", "superior"],
    "Renda (faixas de SM)": colunas_renda_sm,
    "PEA / NPEA": ["PEA", "NPEA"],
    "Cor / Raça": ["Branca", "Preta", "Parda", "Amarelo + outros"],
    "Religião": colunas_religiao,
}


def ajustar_linhas(bloco, total_linha, tol):
    """Para cada linha cuja soma do subgrupo difere do Total da linha, aplica
    valor_ajustado = valor / soma(subgrupo da linha) * Total da linha.
    Ex.: PEA = 22,31% / (22,31% + 17,33%) * 41,22%."""
    soma = bloco.sum(axis=1)
    fora = (soma - total_linha).abs() > tol
    fator = (total_linha / soma.replace(0, np.nan)).fillna(0)
    ajustado = bloco.copy()
    ajustado.loc[fora] = bloco.loc[fora].mul(fator[fora], axis=0)
    return ajustado, fora


CASAS = 2                      # casas decimais em % (0,01 p.p.)
UNIDADE = 10 ** -(CASAS + 2)   # 0,01 p.p. expresso em proporção (0.0001)


def arredondar_maior_resto(serie, total_unidades):
    """Arredonda uma série (em proporção) para múltiplos de UNIDADE de modo que
    a soma dê exatamente total_unidades * UNIDADE (método do maior resto)."""
    u = serie.to_numpy(dtype=float) / UNIDADE
    base = np.floor(u + 1e-9)
    falta = int(round(total_unidades - base.sum()))
    ordem = np.argsort(-(u - base))
    if falta > 0:
        base[ordem[:falta]] += 1
    elif falta < 0:
        base[ordem[::-1][:abs(falta)]] -= 1
    return pd.Series(base * UNIDADE, index=serie.index)


def raking(bloco, total_linha, alvo_coluna, max_iter=200000, tol=1e-13):
    """Ajuste iterativo proporcional (IPF): normaliza linhas e colunas
    alternadamente até que TODA linha some o Total da linha e TODA coluna some
    o alvo informado (sem reescalar os alvos)."""
    m = bloco.to_numpy(dtype=float).copy()
    r = total_linha.to_numpy(dtype=float)
    c = alvo_coluna.to_numpy(dtype=float)
    # células zeradas travam o ajuste; uma semente mínima deixa sempre encaixar
    m = np.where((m <= 0) & (r[:, None] > 0) & (c[None, :] > 0), 1e-9, m)
    m = np.where((r[:, None] == 0) | (c[None, :] == 0), 0, m)
    for _ in range(max_iter):
        rs = m.sum(axis=1)
        m *= np.divide(r, rs, out=np.zeros_like(r), where=rs > 0)[:, None]
        cs = m.sum(axis=0)
        m *= np.divide(c, cs, out=np.zeros_like(c), where=cs > 0)[None, :]
        if np.abs(m.sum(axis=1) - r).max() < tol and np.abs(m.sum(axis=0) - c).max() < tol:
            break
    return pd.DataFrame(m, index=bloco.index, columns=bloco.columns)


def arredondamento_controlado(bloco, total_linha, alvo_coluna):
    """Arredonda cada célula para 0,01 p.p. garantindo que a soma de cada linha
    seja EXATAMENTE o Total da linha e a de cada coluna EXATAMENTE o alvo
    digitado. Resolve um problema de transporte (cada célula sobe ou desce 1
    unidade), preferindo arredondar para cima as células de maior resto."""
    from scipy.optimize import linprog

    u = bloco.to_numpy(dtype=float) / UNIDADE
    R = np.rint(total_linha.to_numpy(dtype=float) / UNIDADE).astype(int)
    C = np.rint(alvo_coluna.to_numpy(dtype=float) / UNIDADE).astype(int)
    base = np.floor(u + 1e-9)
    resto = u - base
    nl, nc = base.shape
    dr = R - base.sum(axis=1)
    dc = C - base.sum(axis=0)

    A = np.zeros((nl + nc, nl * nc))
    for i in range(nl):
        A[i, i * nc:(i + 1) * nc] = 1
    for j in range(nc):
        A[nl + j, j::nc] = 1
    b = np.concatenate([dr, dc])

    res = linprog(-resto.ravel(), A_eq=A, b_eq=b, bounds=[(0, 1)] * (nl * nc), method="highs")
    if not res.success:  # caso extremo: permite mover mais de 1 unidade por célula
        res = linprog(np.zeros(nl * nc), A_eq=A, b_eq=b,
                      bounds=[(-b_, None) for b_ in base.ravel()], method="highs")
    if not res.success:
        return None
    final = base + np.rint(res.x).reshape(nl, nc)
    return pd.DataFrame(final * UNIDADE, index=bloco.index, columns=bloco.columns)


def _resetar(chaves):
    for k in chaves:
        st.session_state.pop(k, None)


# alocação da amostra por microrregião (maior resto, para fechar em n_amostra)
alvo = (tabela_bruta.loc[microrregioes, "Total"] * n_amostra)
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

# --------------------------------------------------------------------------- #
# Ajuste dos subgrupos (automático + valores digitados pelo usuário)
# --------------------------------------------------------------------------- #

st.subheader("🔧 Ajuste dos subgrupos")
st.caption(
    "Quando a soma de um subgrupo (ex.: PEA + NPEA) não bate com o Total da linha, cada valor é "
    "corrigido por **valor / soma do subgrupo × Total da linha**. Você pode digitar os valores "
    "corretos em qualquer célula (em %, ou até em números absolutos — só a proporção entre as "
    "colunas da linha importa) e o app reajusta a linha para bater com o Total."
)

# Total de cada linha com 2 casas (em %) somando exatamente 100,00%
total_linha = arredondar_maior_resto(tabela_bruta["Total"], round(1 / UNIDADE))
tabela_final = tabela_bruta.copy()
tabela_final["Total"] = total_linha
log_ajustes = []

# grupos em que o total de cada coluna é SEMPRE informado pelo usuário
# (linha e coluna precisam bater ao mesmo tempo)
GRUPOS_COM_ALVO_FIXO = {"Renda (faixas de SM)"}

st.markdown("#### 🎯 Total desejado das colunas de Renda (faixas de SM)")
st.caption(
    "Digite o percentual estadual correto de cada faixa. O app ajusta a tabela para que cada linha "
    "feche no Total da mesorregião **e** cada coluna feche no valor que você informar."
)
area_alvos_sm = st.container()

for nome_grupo, cols in GRUPOS.items():
    cols = [c for c in cols if c in tabela_bruta.columns]
    if not cols:
        continue
    bloco_orig = tabela_bruta[cols]
    n_fora_orig = int(((bloco_orig.sum(axis=1) - total_linha).abs() > tol_ajuste).sum())

    slug = re.sub(r"\W+", "_", normalizar_municipio(nome_grupo).lower())
    k_ed, k_rake, k_alvo = f"ed_{slug}", f"rake_{slug}", f"alvo_{slug}"
    status = f"⚠️ {n_fora_orig} linha(s) não batem com o Total" if n_fora_orig else "✅ todas as linhas batem"

    with st.expander(f"{nome_grupo} — {status}"):
        base_editor = (bloco_orig * 100).round(4)
        editado = st.data_editor(
            base_editor,
            key=k_ed,
            use_container_width=True,
            column_config={c: st.column_config.NumberColumn(c, format="%.2f", min_value=0.0) for c in cols},
        )
        editado = editado.reindex(index=bloco_orig.index, columns=cols).astype(float).fillna(0).clip(lower=0)
        linhas_editadas = ~np.isclose(editado.to_numpy(), base_editor.to_numpy()).all(axis=1)
        linhas_editadas = pd.Series(linhas_editadas, index=bloco_orig.index)

        alvo_fixo = nome_grupo in GRUPOS_COM_ALVO_FIXO
        chaves_alvo_fixo = [f"{k_alvo}_{i}" for i in range(len(cols))]
        if alvo_fixo:
            usar_rake = True
        else:
            usar_rake = st.checkbox(
                "Também quero definir o total de cada coluna (o app ajusta linhas e colunas ao mesmo tempo — raking)",
                key=k_rake,
            )
        bloco_ed = editado / 100

        if alvo_fixo:
            soma_cols = bloco_orig.sum()
            alvo_padrao = (
                arredondar_maior_resto(soma_cols / soma_cols.sum(), round(1 / UNIDADE)) * 100
                if soma_cols.sum() > 0 else soma_cols * 0
            )
            with area_alvos_sm:
                campos = st.columns(len(cols))
                valores = []
                for i, (campo, c) in enumerate(zip(campos, cols)):
                    valores.append(campo.number_input(
                        f"{c} (%)", min_value=0.0, max_value=100.0,
                        value=float(round(alvo_padrao[c], CASAS)), step=0.01, format="%.2f",
                        key=chaves_alvo_fixo[i],
                        help=f"Valor calculado do IBGE: {alvo_padrao[c]:.2f}%",
                    ))
                # valores exatamente como digitados (em 0,01 p.p.), sem reescalar
                alvo_coluna = pd.Series(np.round(valores, CASAS) / 100, index=cols, dtype=float)
                soma_un = int(np.rint(alvo_coluna.sum() / UNIDADE))
                total_un = int(np.rint(total_linha.sum() / UNIDADE))

            if soma_un != total_un:
                diff = (total_un - soma_un) / 100
                with area_alvos_sm:
                    st.error(
                        f"Os valores digitados somam **{soma_un / 100:.2f}%**, mas as mesorregiões somam "
                        f"**{total_un / 100:.2f}%**. Como cada linha tem que fechar no seu Total, as colunas "
                        f"também precisam somar {total_un / 100:.2f}% — "
                        + (f"faltam **{diff:.2f} p.p.**" if diff > 0 else f"sobram **{-diff:.2f} p.p.**")
                        + ". Corrija um dos campos; até lá a renda fica só com o ajuste por linha."
                    )
                ajustado, fora = ajustar_linhas(bloco_ed, total_linha, tol_ajuste)
            else:
                rakeado = raking(bloco_ed, total_linha, alvo_coluna)
                exato = arredondamento_controlado(rakeado, total_linha, alvo_coluna)
                ajustado = exato if exato is not None else rakeado
                fora = (ajustado - bloco_ed).abs().max(axis=1) > 1e-9

                soma_l = np.rint(ajustado.sum(axis=1) / UNIDADE)
                soma_c = np.rint(ajustado.sum() / UNIDADE)
                ok = (
                    exato is not None
                    and (soma_l == np.rint(total_linha / UNIDADE)).all()
                    and (soma_c == np.rint(alvo_coluna / UNIDADE)).all()
                )
                with area_alvos_sm:
                    if ok:
                        st.success(
                            "✅ Encaixado: toda linha fecha exatamente no Total da mesorregião e toda coluna "
                            "fecha exatamente no valor digitado ("
                            + " · ".join(f"{c}: {ajustado[c].sum() * 100:.2f}%" for c in cols)
                            + "). Os valores estão arredondados em 2 casas, então a soma que você vê bate."
                        )
                    else:
                        st.error("Não consegui encaixar exatamente. Confira se há mesorregião com Total zerado.")
        elif usar_rake:
            soma_cols = bloco_orig.sum()
            alvo_padrao = (soma_cols / soma_cols.sum() * 100).round(2) if soma_cols.sum() > 0 else soma_cols
            alvo_df = st.data_editor(
                alvo_padrao.to_frame().T.rename(index={0: "Total da coluna (%)"}),
                key=k_alvo,
                use_container_width=True,
                column_config={c: st.column_config.NumberColumn(c, format="%.2f", min_value=0.0) for c in cols},
            )
            alvo_coluna = alvo_df.iloc[0].reindex(cols).astype(float).fillna(0)
            st.caption(
                f"Soma dos alvos: {alvo_coluna.sum():.2f}% (se não fechar em 100%, os alvos são "
                "normalizados proporcionalmente)."
            )
            if not ajuste_automatico:
                st.caption("O raking é aplicado mesmo com o ajuste automático desligado, pois você ativou esta opção.")
            ajustado = raking(bloco_ed, total_linha, alvo_coluna)
            fora = (ajustado - bloco_ed).abs().max(axis=1) > 1e-9  # linhas que o raking alterou
        elif ajuste_automatico:
            ajustado, fora = ajustar_linhas(bloco_ed, total_linha, tol_ajuste)
        else:
            ajustado, fora = bloco_ed.copy(), pd.Series(False, index=bloco_ed.index)

        sem_base = (bloco_ed.sum(axis=1) == 0) & (total_linha > 0)
        if sem_base.any():
            st.warning(
                "Estas linhas têm o subgrupo todo zerado e não podem ser ajustadas — digite os valores: "
                + ", ".join(sem_base[sem_base].index)
            )

        for regiao in bloco_ed.index:
            if fora[regiao] or linhas_editadas[regiao]:
                antes = bloco_ed.loc[regiao].sum()
                log_ajustes.append({
                    "Grupo": nome_grupo,
                    "Mesorregião": regiao,
                    "Editado manualmente": "sim" if linhas_editadas[regiao] else "não",
                    "Método": "raking" if usar_rake else ("proporcional" if ajuste_automatico else "sem ajuste"),
                    "Soma antes (%)": round(antes * 100, 4),
                    "Total da linha (%)": round(total_linha[regiao] * 100, 4),
                    "Diferença (p.p.)": round((antes - total_linha[regiao]) * 100, 4),
                    "Soma depois (%)": round(ajustado.loc[regiao].sum() * 100, 4),
                })

        previa = ajustado.copy()
        previa.insert(0, "Total da linha", total_linha)
        previa["Soma do subgrupo"] = ajustado.sum(axis=1)
        previa.loc["Total Geral"] = previa.sum()
        st.markdown("**Resultado após o ajuste:**")
        st.dataframe(previa.style.format("{:.2%}"), use_container_width=True)

        st.button(
            "↩️ Restaurar valores originais deste grupo",
            key=f"reset_{slug}",
            on_click=_resetar,
            args=([k_ed, k_rake, k_alvo] + chaves_alvo_fixo,),
        )

    tabela_final[cols] = ajustado

# linha de Total Geral
def _com_total(df):
    linha = df.sum(numeric_only=True)
    linha.name = "Total Geral"
    return pd.concat([df, linha.to_frame().T])

tabela_bruta = _com_total(tabela_bruta)
tabela_final = _com_total(tabela_final)
tab_log = pd.DataFrame(log_ajustes)

if tab_log.empty:
    st.success("Nenhuma linha precisou de ajuste.")
else:
    with st.expander(f"📋 Registro de ajustes ({len(tab_log)} linha(s) alteradas)"):
        st.dataframe(tab_log, use_container_width=True, hide_index=True)

st.subheader("Distribuição da amostra por microrregião")
tab_alocacao = alocacao.to_frame(name="Nº de entrevistas")
tab_alocacao["% da amostra"] = (tabela_final.loc[microrregioes, "Total"] * 100).round(1)
st.dataframe(tab_alocacao, use_container_width=True)

st.subheader("Tabela consolidada de ponderação (variáveis x microrregião) — ajustada")
st.dataframe(tabela_final.style.format("{:.2%}"), use_container_width=True)

with st.expander("Ver tabela original (sem ajuste, como veio do IBGE/TSE)"):
    st.dataframe(tabela_bruta.style.format("{:.2%}"), use_container_width=True)

# --------------------------------------------------------------------------- #
# Download em Excel
# --------------------------------------------------------------------------- #

buffer = io.BytesIO()
with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
    mapa.drop(columns=["muni_norm"]).to_excel(writer, sheet_name="Mesorregiões", index=False)
    tabela_final.to_excel(writer, sheet_name="Amostra")
    tabela_bruta.to_excel(writer, sheet_name="Amostra (original)")
    if not tab_log.empty:
        tab_log.to_excel(writer, sheet_name="Ajustes", index=False)
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
