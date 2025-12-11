import io
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
   - Projeção de peças a serem banhadas.  
   - Colunas usadas: `Referência`, `Banho`, `Previsão de Venda`.

O app cruza tudo por **Referência + Tipo de Banho** e calcula:

- Quantidade projetada (Previsão de Venda)
- Quantidade em produção
- Quantidade em retorno de banho
- Quantidade já coberta
- Quantidade que ainda precisa ser enviada para o banho.
"""
)

# -------------------- Funções auxiliares --------------------


def carregar_planilha(uploaded_file):
    """Lê CSV/XLS/XLSX em um DataFrame."""
    if uploaded_file is None:
        return None

    nome = uploaded_file.name.lower()

    try:
        if nome.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file)
        elif nome.endswith(".csv"):
            # sep=None detecta ; , \t automaticamente
            df = pd.read_csv(uploaded_file, sep=None, engine="python")
        else:
            st.error(f"Formato não suportado: {uploaded_file.name}")
            return None
    except Exception as e:
        st.error(f"Erro ao ler {uploaded_file.name}: {e}")
        return None

    return df


def preparar_retorno_ou_producao(df, nome_qtd):
    """
    Prepara base de RETORNO ou PRODUÇÃO a partir de um layout padrão:

    - Produto (texto: "FO040 - Nome da peça")
    - Categoria (tipo de banho: "Ouro", "Ródio"...)
    - A Produzir (quantidade)
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


def preparar_projecao(df):
    """
    Prepara base de PROJEÇÃO (WM10):

    - Referência
    - Banho
    - Previsão de Venda
    """
    col_obrigatorias = {"Referência", "Banho", "Previsão de Venda"}
    faltando = col_obrigatorias.difference(df.columns)
    if faltando:
        st.error(
            f"A planilha de PROJEÇÃO não contém as colunas obrigatórias {faltando}. "
            "Confirme se exportou o relatório correto do WM10."
        )
        return None

    base = (
        df.assign(
            referencia=lambda d: d["Referência"].astype(str).str.strip(),
            banho=lambda d: d["Banho"].astype(str).str.strip(),
            qtd_projetada=lambda d: pd.to_numeric(
                d["Previsão de Venda"], errors="coerce"
            ).fillna(0),
        )[["referencia", "banho", "qtd_projetada"]]
    )

    return base


# -------------------- Upload dos arquivos --------------------

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
        "PROJEÇÃO (WM10)",
        type=["xlsx", "xls", "csv"],
        key="proj",
    )

if not (file_retorno and file_producao and file_proj):
    st.info("Envie as **três planilhas** para iniciar o cálculo.")
    st.stop()

# -------------------- Leitura das planilhas --------------------

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

# -------------------- Processamento --------------------

st.subheader("2. Processar e calcular projeção")

if st.button("Calcular projeção de banho"):
    # Bases consolidadas
    base_retorno = preparar_retorno_ou_producao(df_retorno_raw, "qtd_retorno")
    base_producao = preparar_retorno_ou_producao(df_producao_raw, "qtd_producao")
    base_proj = preparar_projecao(df_proj_raw)

    if base_retorno is None or base_producao is None or base_proj is None:
        st.stop()

    # Junção das bases
    df_merge = (
        base_proj.merge(
            base_producao, on=["referencia", "banho"], how="left"
        )
        .merge(
            base_retorno, on=["referencia", "banho"], how="left"
        )
    )

    # Trata NaN como 0
    for col in ["qtd_projetada", "qtd_producao", "qtd_retorno"]:
        df_merge[col] = pd.to_numeric(df_merge[col], errors="coerce").fillna(0)

    # Cálculos finais
    df_merge["qtd_ja_coberta"] = df_merge["qtd_producao"] + df_merge["qtd_retorno"]
    df_merge["qtd_a_enviar"] = df_merge["qtd_projetada"] - df_merge["qtd_ja_coberta"]
    df_merge["qtd_a_enviar"] = df_merge["qtd_a_enviar"].clip(lower=0)

    # Ordenação
    df_merge = df_merge.sort_values(by=["banho", "referencia"]).reset_index(drop=True)

    # Renomear para exibição
    df_resultado = df_merge.rename(
        columns={
            "referencia": "Referência",
            "banho": "Tipo de banho",
            "qtd_projetada": "Quantidade projetada",
            "qtd_producao": "Quantidade em produção",
            "qtd_retorno": "Quantidade em retorno de banho",
            "qtd_ja_coberta": "Quantidade já coberta",
            "qtd_a_enviar": "Quantidade a enviar para o banho",
        }
    )

    st.subheader("3. Resultado consolidado")

    # Filtro por tipo de banho
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
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )
