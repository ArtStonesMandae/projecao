import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Projeção de Banho (Ouro/Ródio)", layout="wide")

st.title("Projeção de Metais para Banho")

st.markdown(
    """
Fluxo de uso:

1. **RETORNO DE BANHO**  
   - Peças já enviadas para o banho, mas **ainda não voltaram**.  
   - Colunas usadas: `Produto`, `Categoria`, `A Produzir`.

2. **PRODUÇÃO**  
   - Peças que **já voltaram do banho** e estão em produção.  
   - Mesmo layout: `Produto`, `Categoria`, `A Produzir`.

3. **PROJEÇÃO (WM10)**  
   - Relatório de previsão do WM10 (exportado em `.xls`).  
   - Colunas usadas: `Referência`, `Produto`,
     coluna que começa com **"Previsão de Venda"** e, se existir,
     coluna que começa com **"Estoque Atual"**.

O app cruza tudo por **Referência + Tipo de Banho** e calcula:

- Quantidade projetada
- Estoque atual
- Quantidade em produção
- Quantidade em retorno de banho
- Quantidade já coberta (estoque + produção + retorno)
- Quantidade a enviar com **30% de margem**
"""
)

# ------------------------------------------------------------------ #
# Funções auxiliares
# ------------------------------------------------------------------ #


def infer_banho(produto: str) -> str:
    """Tenta inferir o tipo de banho a partir do texto do produto."""
    s = str(produto).upper()
    if "OURO" in s:
        return "Ouro"
    if "RÓDIO" in s or "RODIO" in s:
        return "Ródio"
    if "PRATA" in s:
        return "Prata"
    return "Desconhecido"


def carregar_xls_html(uploaded_file) -> pd.DataFrame:
    """
    Relatórios do WM10 vêm como HTML disfarçado de .xls.
    Usa read_html e promove a primeira linha a cabeçalho.
    """
    data = uploaded_file.read()

    if not data.lstrip().startswith(b"<"):
        st.error(
            "O arquivo .xls não parece ser um relatório HTML do WM10. "
            "Tente exportar novamente ou converter para .xlsx/.csv."
        )
        return None

    try:
        tables = pd.read_html(io.BytesIO(data))
    except Exception as e:
        st.error(f"Não foi possível ler o relatório WM10 (.xls): {e}")
        return None

    if not tables:
        st.error("Nenhuma tabela encontrada no relatório WM10.")
        return None

    df = tables[0]

    # Primeira linha é o cabeçalho
    df2 = df.iloc[1:].copy()
    df2.columns = df.iloc[0]

    return df2


def carregar_planilha(uploaded_file) -> pd.DataFrame:
    """Lê CSV/XLSX normalmente e .XLS (WM10) via HTML."""
    if uploaded_file is None:
        return None

    nome = uploaded_file.name.lower()

    try:
        if nome.endswith(".xls"):
            return carregar_xls_html(uploaded_file)
        elif nome.endswith(".xlsx"):
            return pd.read_excel(uploaded_file, engine="openpyxl")
        elif nome.endswith(".csv"):
            return pd.read_csv(uploaded_file, sep=None, engine="python")
        else:
            st.error(f"Formato não suportado: {uploaded_file.name}")
            return None
    except Exception as e:
        st.error(f"Erro ao ler {uploaded_file.name}: {e}")
        return None


def preparar_retorno_ou_producao(df: pd.DataFrame, nome_qtd: str) -> pd.DataFrame:
    """
    Prepara base de RETORNO ou PRODUÇÃO:

    - Produto: "FO040 - Nome da peça"
    - Categoria: tipo de banho (Ouro, Ródio, etc.)
    - A Produzir: quantidade
    """
    col_obrigatorias = {"Produto", "Categoria", "A Produzir"}
    faltando = col_obrigatorias.difference(df.columns)
    if faltando:
        st.error(
            f"A planilha não contém as colunas obrigatórias {faltando}. "
            "Confirme se exportou o relatório correto."
        )
        return None

    base = (
        df.assign(
            referencia=lambda d: d["Produto"]
            .astype(str)
            .str.split(" - ")
            .str[0]
            .str.strip(),
            banho=lambda d: d["Categoria"].astype(str).str.strip(),
            qtd=lambda d: pd.to_numeric(d["A Produzir"], errors="coerce").fillna(0),
        )[["referencia", "banho", "qtd"]]
        .groupby(["referencia", "banho"], as_index=False)["qtd"]
        .sum()
        .rename(columns={"qtd": nome_qtd})
    )

    return base


def preparar_projecao(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara base de PROJEÇÃO a partir do relatório WM10:

    Esperado na planilha já lida:
    - Coluna 'Referência'
    - Coluna 'Produto'
    - Coluna cujo nome começa com 'Previsão de Venda'
    - Opcional: coluna cujo nome começa com 'Estoque Atual'
    """
    if df is None:
        return None

    cols = df.columns

    if "Referência" not in cols or "Produto" not in cols:
        st.error(
            "A planilha do WM10 precisa ter as colunas 'Referência' e 'Produto'. "
            "Verifique o layout do relatório."
        )
        return None

    # Localiza a coluna de previsão de venda (texto muda por causa da data)
    forecast_col = None
    for c in cols:
        if isinstance(c, str) and c.startswith("Previsão de Venda"):
            forecast_col = c
            break

    if not forecast_col:
        st.error(
            "Não foi encontrada nenhuma coluna que comece com 'Previsão de Venda' "
            "na planilha do WM10."
        )
        return None

    # Localiza a coluna de estoque atual (se existir)
    estoque_col = None
    for c in cols:
        if isinstance(c, str) and c.startswith("Estoque Atual"):
            estoque_col = c
            break

    df2 = df.copy()

    # Remove linhas de totais/rodapé
    df2 = df2[df2["Referência"].notna()]
    df2 = df2[df2["Referência"] != "Referência"]
    df2 = df2[
        ~df2["Referência"]
        .astype(str)
        .str.contains("Totais|Previs", case=False, na=False)
    ]

    # Previsão de venda: extrai apenas o número (ex.: '28 UN' -> 28)
    raw_forecast = df2[forecast_col].astype(str)
    nums_forecast = raw_forecast.str.extract(r"(\d+)")[0]
    qtd_projetada = pd.to_numeric(nums_forecast, errors="coerce").fillna(0)

    # Estoque atual: se a coluna existir, extrai número; se não, zera
    if estoque_col:
        raw_stock = df2[estoque_col].astype(str)
        nums_stock = raw_stock.str.extract(r"(\d+)")[0]
        qtd_estoque = pd.to_numeric(nums_stock, errors="coerce").fillna(0)
    else:
        qtd_estoque = pd.Series(0, index=df2.index, dtype="float64")

    base = pd.DataFrame(
        {
            "referencia": df2["Referência"].astype(str).str.strip(),
            "banho": df2["Produto"].map(infer_banho),
            "qtd_projetada": qtd_projetada,
            "qtd_estoque": qtd_estoque,
        }
    )

    return base


# ------------------------------------------------------------------ #
# Upload dos arquivos
# ------------------------------------------------------------------ #

st.subheader("1. Upload das planilhas")

col1, col2, col3 = st.columns(3)

with col1:
    file_retorno = st.file_uploader(
        "RETORNO DE BANHO (já enviado, ainda não voltou)",
        type=["xlsx", "xls", "csv"],
        key="retorno",
    )

with col2:
    file_producao = st.file_uploader(
        "PRODUÇÃO (já voltou do banho)",
        type=["xlsx", "xls", "csv"],
        key="producao",
    )

with col3:
    file_proj = st.file_uploader(
        "PROJEÇÃO (WM10 - .xls / HTML)",
        type=["xlsx", "xls", "csv"],
        key="proj",
    )

if not (file_retorno and file_producao and file_proj):
    st.info("Envie as **três planilhas** para iniciar o cálculo.")
    st.stop()

# ------------------------------------------------------------------ #
# Leitura das planilhas
# ------------------------------------------------------------------ #

df_retorno_raw = carregar_planilha(file_retorno)
df_producao_raw = carregar_planilha(file_producao)
df_proj_raw = carregar_planilha(file_proj)

if df_retorno_raw is None or df_producao_raw is None or df_proj_raw is None:
    st.stop()

st.success("Planilhas carregadas com sucesso!")

with st.expander("Pré-visualização rápida das planilhas"):
    st.markdown("### Retorno de Banho")
    st.dataframe(df_retorno_raw.head(), use_container_width=True)

    st.markdown("### Produção")
    st.dataframe(df_producao_raw.head(), use_container_width=True)

    st.markdown("### Projeção (WM10)")
    st.dataframe(df_proj_raw.head(), use_container_width=True)

# ------------------------------------------------------------------ #
# Processamento
# ------------------------------------------------------------------ #

st.subheader("2. Processar e calcular projeção")

if st.button("Calcular projeção de banho"):
    base_retorno = preparar_retorno_ou_producao(df_retorno_raw, "qtd_retorno")
    base_producao = preparar_retorno_ou_producao(df_producao_raw, "qtd_producao")
    base_proj = preparar_projecao(df_proj_raw)

    if base_retorno is None or base_producao is None or base_proj is None:
        st.stop()

    # Merge das bases
    df_merge = (
        base_proj.merge(
            base_producao, on=["referencia", "banho"], how="left"
        )
        .merge(
            base_retorno, on=["referencia", "banho"], how="left"
        )
    )

    # Trata NaN como 0
    for col in ["qtd_projetada", "qtd_producao", "qtd_retorno", "qtd_estoque"]:
        df_merge[col] = pd.to_numeric(df_merge[col], errors="coerce").fillna(0)

    # Cálculos finais
    df_merge["qtd_ja_coberta"] = (
        df_merge["qtd_producao"] + df_merge["qtd_retorno"] + df_merge["qtd_estoque"]
    )

    # Base sem margem (apenas para cálculo interno)
    qtd_a_enviar_base = (
        df_merge["qtd_projetada"] - df_merge["qtd_ja_coberta"]
    ).clip(lower=0)

    # Margem de 30% e arredondamento para cima
    df_merge["qtd_a_enviar_margem"] = np.ceil(
        qtd_a_enviar_base * 1.3
    ).astype(int)

    # Só mantém itens que realmente precisam ser enviados
    df_merge = df_merge[df_merge["qtd_a_enviar_margem"] > 0].reset_index(drop=True)

    # Ordenação
    df_merge = df_merge.sort_values(by=["banho", "referencia"]).reset_index(drop=True)

    # Renomear para exibição
    df_resultado = df_merge.rename(
        columns={
            "referencia": "Referência",
            "banho": "Tipo de banho",
            "qtd_projetada": "Quantidade projetada",
            "qtd_estoque": "Quantidade em estoque atual",
            "qtd_producao": "Quantidade em produção",
            "qtd_retorno": "Quantidade em retorno de banho",
            "qtd_ja_coberta": "Quantidade já coberta (estoque + produção + retorno)",
            "qtd_a_enviar_margem": "Quantidade a enviar (margem 30%)",
        }
    )

    st.subheader("3. Resultado consolidado")

    tipos_banho = sorted(df_resultado["Tipo de banho"].dropna().unique().tolist())
    filtro_banho = st.multiselect(
        "Filtrar por tipo de banho (opcional)",
        options=tipos_banho,
        default=tipos_banho,
    )

    df_mostrar = df_resultado.copy()
    if filtro_banho:
        df_mostrar = df_mostrar[df_mostrar["Tipo de banho"].isin(filtro_banho)]

    st.dataframe(df_mostrar, use_container_width=True)

    # Download em Excel
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_resultado.to_excel(writer, index=False, sheet_name="Projecao_Banho")

    st.download_button(
        label="Baixar resultado em Excel",
        data=buffer.getvalue(),
        file_name="projecao_banho_metais.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
