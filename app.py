"""
Dashboard de Elegibilidades / Efetivações — lê planilhas do Google Sheets
e monta um painel interativo com Streamlit + Plotly.
"""

import re
import calendar
from datetime import date
from collections import Counter

import pandas as pd
import streamlit as st
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ABA "ELETIVOS" (Novo layout)
# ---------------------------------------------------------------------------

CELL_MAP = {
    "internacoes_eletivas": "B1",
    "internacoes_eletivas_efetivadas": "D1",
    "pacientes_qt_autorizados": "B2:D2",
    "clinicos_ui": "B3:D3",
    "clinicos_ctia": "B4:D4",
    "clinicos_utip": "B5:D5",
    "locais_municipios_origem": "B6:Z6",
    "convenios": "B7:Z7",
    "acomodacoes_semi": "B8:D8",
    "acomodacoes_apto": "B9:D9",
}

# Campos cujas células NÃO são mescladas (valores em sequência pro lado)
SEQUENCE_FIELDS = ["locais_municipios_origem", "convenios"]

# Rótulos amigáveis para exibir na tela
LABELS = {
    "internacoes_eletivas": "Quantidade de internações eletivas",
    "internacoes_eletivas_efetivadas": "Quantas internações eletivas efetivadas",
    "pacientes_qt_autorizados": "Quantidade de pacientes QT autorizados",
    "clinicos_ui": "Clínicos para UI",
    "clinicos_ctia": "Clínicos para CTIA",
    "clinicos_utip": "Clínicos para UTIP",
    "acomodacoes_semi": "Acomodações SEMI",
    "acomodacoes_apto": "Acomodações APTO",
}

NUMERIC_FIELDS = [f for f in CELL_MAP if f not in SEQUENCE_FIELDS]

COMPARISONS = [
    ("Internações Eletivas: Solicitadas x Efetivadas", "internacoes_eletivas", "internacoes_eletivas_efetivadas"),
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ABA "ELEGÍVEIS" (formulário Google -> planilha de respostas)
# ---------------------------------------------------------------------------

ELEGIVEIS_WORKSHEET_GID = 2058249992

COL_ELEGIVEL = "ELEGÍVEL PARA HMV :"
COL_ACEITO = "PACIENTE ACEITO :"
COL_ORIGEM = "LOCAL DE ORIGEM (Hospital, Clínica, Pronto Atendimento, residência , etc)"
COL_CONVENIO = "CONVÊNIO DO PACIENTE:"
COL_LOCAL_ENTRADA = "LOCAL ENTRADA HMV :"
COL_ACOMOD_PRIVATIVO = "TIPO DE ACOMODAÇÃO: [Privativo]"
COL_ACOMOD_SEMI = "TIPO DE ACOMODAÇÃO: [Semi-privativo]"

# Esta constante será forçada no índice 20 (Coluna U) dentro do load_data_elegiveis
COL_CIDADE = "CIDADE_DE_PARA_COLUNA_U"

DIMENSIONS_ELEGIVEIS = {
    "Local de origem": COL_ORIGEM,
    "Cidade de origem": COL_CIDADE,
    "Convênio": COL_CONVENIO,
    "Local de entrada no HMV": COL_LOCAL_ENTRADA,
}

DATE_COLUMN_ELEGIVEIS = "Carimbo de data/hora"


def calc_pct(elegivel, efetivado):
    if elegivel and elegivel > 0:
        return (efetivado / elegivel) * 100
    return 0.0


def render_comparisons(source, container=st, key_prefix="cmp"):
    for label, campo_elegivel, campo_efetivado in COMPARISONS:
        elegivel = float(source[campo_elegivel])
        efetivado = float(source[campo_efetivado])
        pct = calc_pct(elegivel, efetivado)

        chart_key = f"{key_prefix}_{campo_elegivel}_{campo_efetivado}"

        container.markdown(f"**{label}**")
        c1, c2, c3 = container.columns(3)
        c1.metric("Solicitadas/Elegíveis", int(elegivel))
        c2.metric("Efetivadas", int(efetivado))
        c3.metric("% de execução", f"{pct:.0f}%")

        fig = px.bar(
            x=["Solicitadas/Elegíveis", "Efetivadas"],
            y=[elegivel, efetivado],
            text=[int(elegivel), int(efetivado)],
            labels={"x": "", "y": ""},
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(height=220, margin=dict(t=10, b=10, l=10, r=10), showlegend=False)
        container.plotly_chart(fig, use_container_width=True, key=chart_key)
        container.divider()


def render_value_counts(df, col, label, container=st):
    if col not in df.columns:
        container.warning(f"Coluna '{col}' não encontrada na planilha.")
        return

    valores = df[col].astype(str).str.strip()
    valores = valores[valores != ""]
    vc = valores.value_counts()

    container.markdown(f"**{label}** — {len(valores)} resposta(s) preenchida(s)")
    if vc.empty:
        container.info("Nenhuma resposta preenchida ainda.")
        return

    metric_cols = container.columns(len(vc))
    for mcol, (valor, qtd) in zip(metric_cols, vc.items()):
        mcol.metric(valor, int(qtd))


def render_group_breakdown(df, dim_col, value_col, container=st):
    """Conta ocorrências de value_col segregadas por dim_col e exibe em formato de tabela."""
    if value_col not in df.columns:
        container.warning(f"Coluna '{value_col}' não encontrada na planilha.")
        return

    data = df[[dim_col, value_col]].copy()
    data[dim_col] = data[dim_col].astype(str).str.strip()
    data[value_col] = data[value_col].astype(str).str.strip()
    
    # Remove as linhas onde a dimensão principal (ex: cidade) está vazia
    data = data[data[dim_col] != ""]
    data[value_col] = data[value_col].replace("", "(vazio)")

    if data.empty:
        container.info("Sem dados preenchidos para essa segregação.")
        return

    # Cria uma tabela cruzada (Pivot)
    # Linhas: a dimensão escolhida (ex: Local de origem)
    # Colunas: as respostas (ex: Sim, Não, (vazio))
    tabela = pd.crosstab(data[dim_col], data[value_col])
    
    # Adiciona uma coluna 'Total' para ordenar dos locais mais frequentes pros menos
    tabela['Total'] = tabela.sum(axis=1)
    tabela = tabela.sort_values('Total', ascending=False)
    
    # Tira o index para que a dimensão apareça como uma coluna limpa na tabela
    tabela = tabela.reset_index()
    tabela.columns.name = None
    
    # Exibe no Streamlit como uma tabela iterativa
    container.dataframe(tabela, use_container_width=True, hide_index=True)


def classify_acomodacao(row):
    priv = str(row.get(COL_ACOMOD_PRIVATIVO, "")).strip()
    semi = str(row.get(COL_ACOMOD_SEMI, "")).strip()
    if priv and semi:
        return "Privativo + Semi-privativo"
    if priv:
        return "Privativo"
    if semi:
        return "Semi-privativo"
    return "Não informado"


# ---------------------------------------------------------------------------
# CONEXÃO COM O GOOGLE SHEETS
# ---------------------------------------------------------------------------

@st.cache_resource(show_spinner=False)
def get_client():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_client()
    sheet_key = st.secrets.get("SPREADSHEET_KEY")
    sheet_url = st.secrets.get("SPREADSHEET_URL")
    if sheet_key:
        return client.open_by_key(sheet_key)
    if sheet_url:
        return client.open_by_url(sheet_url)
    raise RuntimeError("Defina SPREADSHEET_KEY ou SPREADSHEET_URL em st.secrets.")


@st.cache_resource(show_spinner=False)
def get_spreadsheet_elegiveis():
    sheet_key = st.secrets.get("SPREADSHEET_KEY_ELEGIVEIS")
    sheet_url = st.secrets.get("SPREADSHEET_URL_ELEGIVEIS")
    if not sheet_key and not sheet_url:
        return None
    client = get_client()
    if sheet_key:
        return client.open_by_key(sheet_key)
    return client.open_by_url(sheet_url)


def _cell_ref_to_rowcol(ref: str):
    return gspread.utils.a1_to_rowcol(ref)


def _read_field(values, ref: str) -> str:
    if ":" in ref:
        start_ref, end_ref = ref.split(":")
        r1, c1 = _cell_ref_to_rowcol(start_ref)
        r2, c2 = _cell_ref_to_rowcol(end_ref)
    else:
        r1, c1 = _cell_ref_to_rowcol(ref)
        r2, c2 = r1, c1

    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if r - 1 < len(values) and c - 1 < len(values[r - 1]):
                val = values[r - 1][c - 1].strip()
                if val:
                    return val
    return ""


def _read_field_sequence(values, ref: str, sep: str = "; ") -> str:
    if ":" not in ref:
        r, c = _cell_ref_to_rowcol(ref)
        if r - 1 < len(values) and c - 1 < len(values[r - 1]):
            return values[r - 1][c - 1].strip()
        return ""

    start_ref, end_ref = ref.split(":")
    r1, c1 = _cell_ref_to_rowcol(start_ref)
    r2, c2 = _cell_ref_to_rowcol(end_ref)

    encontrados = []
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            if r - 1 < len(values) and c - 1 < len(values[r - 1]):
                val = values[r - 1][c - 1].strip()
                if val:
                    encontrados.append(val)
    return sep.join(encontrados)


def parse_day_from_title(title: str, month: int, year: int):
    t = title.strip()
    m = re.fullmatch(r"(?:dia\s*)?0*(\d{1,2})º?", t, flags=re.IGNORECASE)
    if m:
        day = int(m.group(1))
        try:
            return date(year, month, day)
        except ValueError:
            return None

    m = re.fullmatch(r"(\d{1,2})[/\-](\d{1,2})(?:[/\-](\d{2,4}))?", t)
    if m:
        d, mo = int(m.group(1)), int(m.group(2))
        y = int(m.group(3)) if m.group(3) else year
        if y < 100:
            y += 2000
        try:
            return date(y, mo, d)
        except ValueError:
            return None
    return None


@st.cache_data(ttl=300, show_spinner="Lendo planilha de Eletivos...")
def load_data(month: int, year: int):
    ss = get_spreadsheet()
    rows = []
    skipped = []

    for ws in ss.worksheets():
        day = parse_day_from_title(ws.title, month, year)
        if day is None:
            skipped.append(ws.title)
            continue

        values = ws.get_all_values()
        record = {"data": day, "aba": ws.title}

        for field, ref in CELL_MAP.items():
            if field in SEQUENCE_FIELDS:
                record[field] = _read_field_sequence(values, ref)
            else:
                record[field] = _read_field(values, ref)

        rows.append(record)

    if not rows:
        return pd.DataFrame(), skipped

    df = pd.DataFrame(rows).sort_values("data").reset_index(drop=True)

    for field in NUMERIC_FIELDS:
        df[field] = (
            df[field]
            .astype(str)
            .str.replace(",", ".", regex=False)
            .str.extract(r"(-?\d+\.?\d*)")[0]
        )
        df[field] = pd.to_numeric(df[field], errors="coerce").fillna(0)

    return df, skipped


@st.cache_data(ttl=300, show_spinner="Lendo planilha de Elegíveis...")
def load_data_elegiveis(worksheet_name: str | None = None):
    ss = get_spreadsheet_elegiveis()
    if ss is None:
        return None

    if worksheet_name:
        ws = ss.worksheet(worksheet_name)
    elif ELEGIVEIS_WORKSHEET_GID is not None:
        ws = ss.get_worksheet_by_id(ELEGIVEIS_WORKSHEET_GID)
    else:
        ws = ss.sheet1

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    header = values[0]
    
    # GARANTE A LEITURA DA COLUNA U: Substitui o nome do cabeçalho da coluna 20 (U) 
    # pelo nosso rótulo padrão interno, ignorando qual texto está de fato na planilha.
    if len(header) > 20:
        header[20] = COL_CIDADE

    df = pd.DataFrame(values[1:], columns=header)

    if DATE_COLUMN_ELEGIVEIS and DATE_COLUMN_ELEGIVEIS in df.columns:
        df[DATE_COLUMN_ELEGIVEIS] = pd.to_datetime(
            df[DATE_COLUMN_ELEGIVEIS], errors="coerce", dayfirst=True
        )

    return df


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Dashboard de Elegibilidades", layout="wide")
st.title("📊 Dashboard de Elegibilidades e Efetivações")

top_eletivos, top_elegiveis = st.tabs(["🗂️ Eletivos", "📝 Elegíveis"])

# =============================================================================
# ABA "ELETIVOS" 
# =============================================================================
with top_eletivos:
    with st.sidebar:
        st.header("Filtros — Eletivos")
        today = date.today()
        ano = st.number_input("Ano", min_value=2020, max_value=2100, value=today.year, step=1)
        mes = st.selectbox(
            "Mês",
            options=list(range(1, 13)),
            index=today.month - 1,
            format_func=lambda m: calendar.month_name[m].capitalize(),
        )
        if st.button("🔄 Recarregar dados (Eletivos)"):
            load_data.clear()
            st.rerun()

    df, skipped_tabs = load_data(mes, ano)

    if df.empty:
        st.warning(
            "Nenhuma aba reconhecida como dia do mês foi encontrada para o "
            f"período selecionado ({calendar.month_name[mes]}/{ano}). "
        )
        st.stop()

    tab_geral, tab_dia, tab_local = st.tabs(["📈 Visão Geral", "📅 Por Dia", "📍 Locais e Convênios"])

    # --- VISÃO GERAL -----------------------------------------------------------
    with tab_geral:
        st.subheader(f"Totais de {calendar.month_name[mes].capitalize()}/{ano}")

        totals = df[NUMERIC_FIELDS].sum()

        kpi_fields = [
            "internacoes_eletivas",
            "internacoes_eletivas_efetivadas",
            "pacientes_qt_autorizados",
        ]
        cols = st.columns(len(kpi_fields))
        for col, field in zip(cols, kpi_fields):
            col.metric(LABELS[field], int(totals[field]))

        st.divider()

        st.markdown("### Comparativo Mensal")
        render_comparisons(totals, key_prefix="geral")

        st.markdown("**Totais detalhados do mês**")
        totals_df = pd.DataFrame({
            "Indicador": [LABELS[f] for f in NUMERIC_FIELDS],
            "Total": [int(totals[f]) for f in NUMERIC_FIELDS],
        }).sort_values("Total", ascending=False)
        st.plotly_chart(
            px.bar(totals_df, x="Total", y="Indicador", orientation="h"),
            use_container_width=True,
        )

        st.markdown("**Evolução diária**")
        campo_evolucao = st.selectbox(
            "Escolha o indicador",
            options=NUMERIC_FIELDS,
            format_func=lambda f: LABELS[f],
            key="campo_evolucao",
        )
        fig = px.line(df, x="data", y=campo_evolucao, markers=True,
                       labels={"data": "Dia", campo_evolucao: LABELS[campo_evolucao]})
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("Ver tabela completa"):
            display_df = df.copy()
            display_df = display_df.rename(columns={**LABELS, "data": "Data", "aba": "Aba"})
            st.dataframe(display_df, use_container_width=True)

    # --- POR DIA -----------------------------------------------------------
    with tab_dia:
        st.subheader("Consultar um dia específico")

        dia_escolhido = st.selectbox(
            "Selecione o dia",
            options=df["data"].tolist(),
            format_func=lambda d: d.strftime("%d/%m/%Y"),
        )

        row = df[df["data"] == dia_escolhido].iloc[0]
        st.caption(f"Aba de origem: **{row['aba']}**")

        c1, c2, c3 = st.columns(3)
        c1.metric(LABELS["internacoes_eletivas"], int(row["internacoes_eletivas"]))
        c2.metric(LABELS["internacoes_eletivas_efetivadas"], int(row["internacoes_eletivas_efetivadas"]))
        c3.metric(LABELS["pacientes_qt_autorizados"], int(row["pacientes_qt_autorizados"]))

        st.divider()
        st.markdown("### Comparativo do dia")
        render_comparisons(row, key_prefix=f"dia_{dia_escolhido.isoformat()}")

        detalhe = pd.DataFrame({
            "Indicador": [LABELS[f] for f in NUMERIC_FIELDS],
            "Valor": [int(row[f]) for f in NUMERIC_FIELDS],
        })
        st.dataframe(detalhe, use_container_width=True, hide_index=True)

        if row.get("locais_municipios_origem"):
            st.markdown(f"**Locais/Municípios de origem:** {row['locais_municipios_origem']}")
        if row.get("convenios"):
            st.markdown(f"**Convênios:** {row['convenios']}")

    # --- POR LOCAL / CONVÊNIO ------------------------------------------------
    with tab_local:
        c_loc, c_conv = st.columns(2)
        
        with c_loc:
            st.subheader("Locais / Municípios de Origem")
            origem_counter = Counter()
            for val in df["locais_municipios_origem"]:
                val = str(val).strip()
                if not val:
                    continue
                for parte in re.split(r"[;,/]", val):
                    parte = parte.strip()
                    if parte:
                        origem_counter[parte] += 1
                        
            if origem_counter:
                origem_df = pd.DataFrame(
                    sorted(origem_counter.items(), key=lambda x: -x[1]),
                    columns=["Local", "Ocorrências"],
                )
                st.plotly_chart(px.bar(origem_df, x="Ocorrências", y="Local", orientation="h"), use_container_width=True)
            else:
                st.info("Nenhum local preenchido.")

        with c_conv:
            st.subheader("Convênios")
            conv_counter = Counter()
            for val in df["convenios"]:
                val = str(val).strip()
                if not val:
                    continue
                for parte in re.split(r"[;,/]", val):
                    parte = parte.strip()
                    if parte:
                        conv_counter[parte] += 1
                        
            if conv_counter:
                conv_df = pd.DataFrame(
                    sorted(conv_counter.items(), key=lambda x: -x[1]),
                    columns=["Convênio", "Ocorrências"],
                )
                st.plotly_chart(px.bar(conv_df, x="Ocorrências", y="Convênio", orientation="h", color_discrete_sequence=["#00b894"]), use_container_width=True)
            else:
                st.info("Nenhum convênio preenchido.")


# =============================================================================
# ABA "ELEGÍVEIS" 
# =============================================================================
with top_elegiveis:
    st.subheader("Elegíveis")

    if get_spreadsheet_elegiveis() is None:
        st.info("Planilha não conectada. Configure as chaves SPREADSHEET_KEY_ELEGIVEIS no st.secrets.")
    else:
        if st.button("🔄 Recarregar dados (Elegíveis)"):
            load_data_elegiveis.clear()
            st.rerun()

        df_elegiveis_full = load_data_elegiveis()

        if df_elegiveis_full is None:
            st.warning("Não foi possível conectar à planilha de Elegíveis.")
        elif df_elegiveis_full.empty:
            st.warning("A planilha de Elegíveis foi encontrada, mas está vazia.")
        else:
            df_elegiveis = df_elegiveis_full

            if DATE_COLUMN_ELEGIVEIS and DATE_COLUMN_ELEGIVEIS in df_elegiveis.columns:
                datas_validas = df_elegiveis[DATE_COLUMN_ELEGIVEIS].dropna()
                if not datas_validas.empty:
                    periodo = st.date_input(
                        "Filtrar por período",
                        value=(datas_validas.min().date(), datas_validas.max().date()),
                    )
                    if isinstance(periodo, tuple) and len(periodo) == 2:
                        inicio, fim = periodo
                        mask = df_elegiveis[DATE_COLUMN_ELEGIVEIS].dt.date.between(inicio, fim)
                        df_elegiveis = df_elegiveis[mask]

            st.markdown("### Elegibilidade e aceite")
            c1, c2 = st.columns(2)
            with c1:
                render_value_counts(df_elegiveis, COL_ELEGIVEL, "Elegível para HMV", container=c1)
            with c2:
                render_value_counts(df_elegiveis, COL_ACEITO, "Paciente aceito", container=c2)

            st.divider()

            st.markdown("### Elegíveis e aceitos por grupo")
            dim_label = st.selectbox("Segregar por", options=list(DIMENSIONS_ELEGIVEIS.keys()), key="dim_elegiveis")
            dim_col = DIMENSIONS_ELEGIVEIS[dim_label]

            if dim_col not in df_elegiveis.columns:
                st.warning(f"Coluna '{dim_col}' (ou equivalente) não encontrada na planilha.")
            else:
                gc1, gc2 = st.columns(2)
                with gc1:
                    st.caption(f"Elegível para HMV, por {dim_label.lower()}")
                    render_group_breakdown(df_elegiveis, dim_col, COL_ELEGIVEL, container=gc1)
                with gc2:
                    st.caption(f"Paciente aceito, por {dim_label.lower()}")
                    render_group_breakdown(df_elegiveis, dim_col, COL_ACEITO, container=gc2)

            st.divider()

            st.markdown("### Tipo de acomodação")
            if COL_ACOMOD_PRIVATIVO in df_elegiveis.columns or COL_ACOMOD_SEMI in df_elegiveis.columns:
                acomod_series = df_elegiveis.apply(classify_acomodacao, axis=1)
                acomod_counts = (
                    acomod_series.value_counts()
                    .rename_axis("Tipo de acomodação")
                    .reset_index(name="Quantidade")
                )
                st.plotly_chart(px.bar(acomod_counts, x="Quantidade", y="Tipo de acomodação", orientation="h"), use_container_width=True)
            else:
                st.warning("Colunas de tipo de acomodação não encontradas na planilha.")

            st.divider()
            st.markdown("### Base completa de respostas")
            st.dataframe(df_elegiveis, use_container_width=True)
