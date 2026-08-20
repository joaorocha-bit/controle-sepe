"""
Dashboard de Elegibilidades / Efetivações — lê planilhas do Google Sheets
e monta um painel interativo com Streamlit + Plotly.
"""

import re
import calendar
from datetime import date, datetime
from collections import Counter

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
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
        fig.update_traces(textposition="outside", marker_color=["#00e5ff", "#00ff9d"])
        fig.update_layout(
            height=220, margin=dict(t=10, b=10, l=10, r=10), showlegend=False,
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font_color="#e6f1ff",
        )
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
# TEMA VISUAL — "COMMAND CENTER" (dark / alto contraste)
# ---------------------------------------------------------------------------

def inject_theme():
    st.markdown(
        """
        <style>
        :root{
            --cc-bg:#05070d;
            --cc-panel:#0d1526;
            --cc-border:#1c2740;
            --cc-cyan:#00e5ff;
            --cc-green:#00ff9d;
            --cc-orange:#ff9500;
            --cc-red:#ff3b5c;
            --cc-text:#e6f1ff;
            --cc-muted:#7e8aa3;
        }

        .stApp{
            background: radial-gradient(circle at top, #0b1220 0%, #05070d 65%);
            color: var(--cc-text);
        }

        section[data-testid="stSidebar"]{
            background-color: var(--cc-panel);
            border-right: 1px solid var(--cc-border);
        }

        h1, h2, h3, h4, .stMarkdown p{
            color: var(--cc-text) !important;
        }

        h1{
            letter-spacing: 1px;
            text-shadow: 0 0 18px rgba(0,229,255,0.25);
        }

        /* Abas */
        button[data-baseweb="tab"]{
            color: var(--cc-muted);
            font-weight: 600;
        }
        button[data-baseweb="tab"][aria-selected="true"]{
            color: var(--cc-cyan) !important;
            border-bottom-color: var(--cc-cyan) !important;
        }

        /* st.metric nativo */
        div[data-testid="stMetric"]{
            background: linear-gradient(145deg, #0d1526, #0a0f1c);
            border: 1px solid var(--cc-border);
            border-radius: 12px;
            padding: 12px 16px;
        }
        div[data-testid="stMetricValue"]{
            color: var(--cc-cyan);
            font-family: "Courier New", monospace;
        }
        div[data-testid="stMetricLabel"]{
            color: var(--cc-muted);
        }

        /* Cards de KPI customizados */
        .kpi-card{
            background: linear-gradient(145deg, #0d1526, #0a0f1c);
            border: 1px solid var(--cc-border);
            border-radius: 14px;
            padding: 18px 20px;
            box-shadow: 0 0 18px rgba(0,229,255,0.06);
            height: 100%;
        }
        .kpi-label{
            color: var(--cc-muted);
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 6px;
        }
        .kpi-value{
            font-family: "Courier New", monospace;
            font-size: 2.3rem;
            font-weight: 700;
            line-height: 1;
        }
        .kpi-sub{
            color: var(--cc-muted);
            font-size: 0.75rem;
            margin-top: 6px;
        }

        .cc-clock{
            font-family: "Courier New", monospace;
            color: var(--cc-green);
            font-size: 1rem;
            text-align: right;
        }

        .cc-divider{
            border-bottom: 1px solid var(--cc-border);
            margin: 4px 0 20px 0;
        }

        [data-testid="stDataFrame"]{
            border: 1px solid var(--cc-border);
            border-radius: 8px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def kpi_card(label, value, sublabel="", color="var(--cc-cyan)", container=st):
    container.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value" style="color:{color};">{value}</div>
            <div class="kpi-sub">{sublabel}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def status_color(pct):
    if pct >= 70:
        return "var(--cc-green)"
    if pct >= 40:
        return "var(--cc-orange)"
    return "var(--cc-red)"


# ---------------------------------------------------------------------------
# ABA "COMMAND CENTER"
# ---------------------------------------------------------------------------

def render_command_center(df_eletivos, skipped_tabs, df_elegiveis_full, mes, ano):
    now = datetime.now()

    head_l, head_r = st.columns([3, 1])
    with head_l:
        st.markdown(f"### 🎛️ Painel de Comando — {calendar.month_name[mes].capitalize()}/{ano}")
    with head_r:
        st.markdown(f"<div class='cc-clock'>🕒 {now.strftime('%d/%m/%Y %H:%M:%S')}</div>", unsafe_allow_html=True)
    st.markdown("<div class='cc-divider'></div>", unsafe_allow_html=True)

    cfg1, cfg2, cfg3 = st.columns([1, 2, 3])
    with cfg1:
        auto_refresh = st.toggle("🔁 Auto-atualizar", value=True, key="cc_auto_refresh")
    with cfg2:
        intervalo_min = st.slider("Intervalo (min)", min_value=1, max_value=30, value=5, key="cc_intervalo")
    with cfg3:
        if auto_refresh:
            st.caption(f"⏱️ Tela atualiza sozinha a cada {intervalo_min} min — ideal para telão/TV.")
        else:
            st.caption("Auto-atualização desligada. Use o botão de recarregar na barra lateral.")

    if auto_refresh:
        components.html(
            f"""
            <script>
            setTimeout(function() {{
                window.parent.location.reload();
            }}, {int(intervalo_min * 60 * 1000)});
            </script>
            """,
            height=0,
        )

    st.markdown("#### 📌 Indicadores-chave")

    if df_eletivos is None or df_eletivos.empty:
        st.info("Sem dados de Eletivos para o período selecionado — ajuste o filtro de Mês/Ano na barra lateral.")
    else:
        totals = df_eletivos[NUMERIC_FIELDS].sum()
        elegivel = float(totals["internacoes_eletivas"])
        efetivado = float(totals["internacoes_eletivas_efetivadas"])
        pct_execucao = calc_pct(elegivel, efetivado)

        c1, c2, c3, c4 = st.columns(4)
        kpi_card("Internações Eletivas", int(elegivel), "solicitadas no mês", container=c1)
        kpi_card("Efetivadas", int(efetivado), "internações realizadas", color="var(--cc-green)", container=c2)
        kpi_card("% Execução", f"{pct_execucao:.0f}%", "solicitadas → efetivadas",
                  color=status_color(pct_execucao), container=c3)
        kpi_card("Pacientes QT Autorizados", int(totals["pacientes_qt_autorizados"]), "no mês", container=c4)

    if df_elegiveis_full is not None and not df_elegiveis_full.empty:
        elegivel_counts = (
            df_elegiveis_full[COL_ELEGIVEL].astype(str).str.strip().value_counts()
            if COL_ELEGIVEL in df_elegiveis_full.columns else pd.Series(dtype=int)
        )
        aceito_counts = (
            df_elegiveis_full[COL_ACEITO].astype(str).str.strip().value_counts()
            if COL_ACEITO in df_elegiveis_full.columns else pd.Series(dtype=int)
        )

        total_registros = len(df_elegiveis_full)
        total_elegiveis_sim = int(elegivel_counts.get("Sim", 0))
        total_aceitos_sim = int(aceito_counts.get("Sim", 0))
        pct_aceite = calc_pct(total_elegiveis_sim, total_aceitos_sim)

        c5, c6, c7, c8 = st.columns(4)
        kpi_card("Respostas Recebidas", total_registros, "formulário de elegibilidade", container=c5)
        kpi_card("Elegíveis (Sim)", total_elegiveis_sim, color="var(--cc-green)", container=c6)
        kpi_card("Aceitos (Sim)", total_aceitos_sim, color="var(--cc-cyan)", container=c7)
        kpi_card("% Aceite s/ Elegíveis", f"{pct_aceite:.0f}%", "aceitos → elegíveis",
                  color=status_color(pct_aceite), container=c8)
    else:
        st.info("Planilha de Elegíveis não conectada — configure SPREADSHEET_KEY_ELEGIVEIS em st.secrets para ver esses indicadores aqui.")

    st.markdown("#### 🏥 Ocupação e Origem")
    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.markdown("**Tipo de acomodação (Elegíveis)**")
        if (
            df_elegiveis_full is not None and not df_elegiveis_full.empty
            and (COL_ACOMOD_PRIVATIVO in df_elegiveis_full.columns or COL_ACOMOD_SEMI in df_elegiveis_full.columns)
        ):
            acomod_series = df_elegiveis_full.apply(classify_acomodacao, axis=1)
            acomod_counts = acomod_series.value_counts().rename_axis("Tipo").reset_index(name="Qtd")
            fig = px.pie(
                acomod_counts, names="Tipo", values="Qtd", hole=0.55,
                color_discrete_sequence=["#00e5ff", "#00ff9d", "#ff9500", "#ff3b5c"],
            )
            fig.update_layout(
                margin=dict(t=10, b=10, l=10, r=10), height=260,
                paper_bgcolor="rgba(0,0,0,0)", font_color="#e6f1ff",
                legend=dict(orientation="h", y=-0.15),
            )
            st.plotly_chart(fig, use_container_width=True, key="cc_acomod")
        else:
            st.info("Sem dados de acomodação.")

    with col_b:
        st.markdown("**Top locais de origem (Eletivos)**")
        origem_counter = Counter()
        if df_eletivos is not None and not df_eletivos.empty:
            for val in df_eletivos["locais_municipios_origem"]:
                val = str(val).strip()
                if not val:
                    continue
                for parte in re.split(r"[;,/]", val):
                    parte = parte.strip()
                    if parte:
                        origem_counter[parte] += 1
        if origem_counter:
            top_origem = pd.DataFrame(
                sorted(origem_counter.items(), key=lambda x: -x[1])[:5],
                columns=["Local", "Ocorrências"],
            )
            fig = px.bar(top_origem, x="Ocorrências", y="Local", orientation="h",
                          color_discrete_sequence=["#00e5ff"])
            fig.update_layout(
                margin=dict(t=10, b=10, l=10, r=10), height=260,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font_color="#e6f1ff",
            )
            st.plotly_chart(fig, use_container_width=True, key="cc_origem")
        else:
            st.info("Sem dados de origem.")

    with col_c:
        st.markdown("**Top convênios (Eletivos)**")
        conv_counter = Counter()
        if df_eletivos is not None and not df_eletivos.empty:
            for val in df_eletivos["convenios"]:
                val = str(val).strip()
                if not val:
                    continue
                for parte in re.split(r"[;,/]", val):
                    parte = parte.strip()
                    if parte:
                        conv_counter[parte] += 1
        if conv_counter:
            top_conv = pd.DataFrame(
                sorted(conv_counter.items(), key=lambda x: -x[1])[:5],
                columns=["Convênio", "Ocorrências"],
            )
            fig = px.bar(top_conv, x="Ocorrências", y="Convênio", orientation="h",
                          color_discrete_sequence=["#00ff9d"])
            fig.update_layout(
                margin=dict(t=10, b=10, l=10, r=10), height=260,
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font_color="#e6f1ff",
            )
            st.plotly_chart(fig, use_container_width=True, key="cc_conv")
        else:
            st.info("Sem dados de convênio.")

    st.caption(
        f"Última atualização dos dados: {now.strftime('%d/%m/%Y %H:%M:%S')} · "
        f"cache renovado a cada 5 min · {len(skipped_tabs or [])} aba(s) ignorada(s) no período."
    )


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Dashboard de Elegibilidades", layout="wide", page_icon="📊")
inject_theme()
st.title("📊 Dashboard de Elegibilidades e Efetivações")

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

    st.divider()
    st.header("Elegíveis")
    elegiveis_conectado = get_spreadsheet_elegiveis() is not None
    if elegiveis_conectado:
        if st.button("🔄 Recarregar dados (Elegíveis)"):
            load_data_elegiveis.clear()
            st.rerun()
    else:
        st.caption("Não conectado — configure SPREADSHEET_KEY_ELEGIVEIS em st.secrets.")

df, skipped_tabs = load_data(mes, ano)
df_elegiveis_full = load_data_elegiveis() if elegiveis_conectado else None

top_command, top_eletivos, top_elegiveis = st.tabs(["🎛️ Command Center", "🗂️ Eletivos", "📝 Elegíveis"])

# =============================================================================
# ABA "COMMAND CENTER"
# =============================================================================
with top_command:
    render_command_center(df, skipped_tabs, df_elegiveis_full, mes, ano)

# =============================================================================
# ABA "ELETIVOS"
# =============================================================================
with top_eletivos:
    if df.empty:
        st.warning(
            "Nenhuma aba reconhecida como dia do mês foi encontrada para o "
            f"período selecionado ({calendar.month_name[mes]}/{ano})."
        )
    else:
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
            fig_totais = px.bar(totals_df, x="Total", y="Indicador", orientation="h")
            fig_totais.update_traces(marker_color="#00e5ff")
            fig_totais.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#e6f1ff"
            )
            st.plotly_chart(fig_totais, use_container_width=True)

            st.markdown("**Evolução diária**")
            campo_evolucao = st.selectbox(
                "Escolha o indicador",
                options=NUMERIC_FIELDS,
                format_func=lambda f: LABELS[f],
                key="campo_evolucao",
            )
            fig = px.line(df, x="data", y=campo_evolucao, markers=True,
                           labels={"data": "Dia", campo_evolucao: LABELS[campo_evolucao]})
            fig.update_traces(line_color="#00e5ff", marker_color="#00ff9d")
            fig.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#e6f1ff"
            )
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
                    fig_o = px.bar(origem_df, x="Ocorrências", y="Local", orientation="h")
                    fig_o.update_traces(marker_color="#00e5ff")
                    fig_o.update_layout(
                        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#e6f1ff"
                    )
                    st.plotly_chart(fig_o, use_container_width=True)
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
                    fig_c = px.bar(conv_df, x="Ocorrências", y="Convênio", orientation="h",
                                    color_discrete_sequence=["#00ff9d"])
                    fig_c.update_layout(
                        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#e6f1ff"
                    )
                    st.plotly_chart(fig_c, use_container_width=True)
                else:
                    st.info("Nenhum convênio preenchido.")


# =============================================================================
# ABA "ELEGÍVEIS"
# =============================================================================
with top_elegiveis:
    st.subheader("Elegíveis")

    if not elegiveis_conectado:
        st.info("Planilha não conectada. Configure as chaves SPREADSHEET_KEY_ELEGIVEIS no st.secrets.")
    elif df_elegiveis_full is None:
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
            fig_a = px.bar(acomod_counts, x="Quantidade", y="Tipo de acomodação", orientation="h")
            fig_a.update_traces(marker_color="#00e5ff")
            fig_a.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font_color="#e6f1ff"
            )
            st.plotly_chart(fig_a, use_container_width=True)
        else:
            st.warning("Colunas de tipo de acomodação não encontradas na planilha.")

        st.divider()
        st.markdown("### Base completa de respostas")
        st.dataframe(df_elegiveis, use_container_width=True)
