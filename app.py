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
    "clinicos_neo": "B6:D6",
    "locais_municipios_origem": "B7:Z7",
    "convenios": "B8:Z8",
    "acomodacoes_semi": "B9:D9",
    "acomodacoes_apto": "B10:D10",
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
    "clinicos_neo": "Clínicos para NEO",
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
COL_CIDADE = "CIDADE"

DIMENSIONS_ELEGIVEIS = {
    "Local de origem": COL_ORIGEM,
    "Cidade de origem": COL_CIDADE,
    "Convênio": COL_CONVENIO,
    "Local de entrada no HMV": COL_LOCAL_ENTRADA,
}

DATE_COLUMN_ELEGIVEIS = "Carimbo de data/hora"


# ---------------------------------------------------------------------------
# PALETA / TEMA — usada em CSS e nos gráficos para manter tudo consistente
# ---------------------------------------------------------------------------

COLOR_PRIMARY = "#00e5ff"   # ciano — indicador neutro / "solicitado"
COLOR_SUCCESS = "#00ff9d"   # verde — positivo / "efetivado" / "sim"
COLOR_WARNING = "#ff9500"   # laranja — atenção
COLOR_DANGER = "#ff3b5c"    # vermelho — negativo / "não"
COLOR_MUTED = "#7e8aa3"
TEXT_COLOR = "#e6f1ff"
PALETTE_SEQUENCE = [COLOR_PRIMARY, COLOR_SUCCESS, COLOR_WARNING, COLOR_DANGER]


# ---------------------------------------------------------------------------
# HELPERS DE NEGÓCIO
# ---------------------------------------------------------------------------

def calc_pct(elegivel, efetivado):
    if elegivel and elegivel > 0:
        return (efetivado / elegivel) * 100
    return 0.0


def status_color(pct):
    if pct >= 70:
        return COLOR_SUCCESS
    if pct >= 40:
        return COLOR_WARNING
    return COLOR_DANGER


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


def build_split_counter(series, seps=r"[;,/]"):
    """Conta ocorrências de valores separados por ; , / dentro de uma série de texto."""
    counter = Counter()
    for val in series:
        val = str(val).strip()
        if not val:
            continue
        for parte in re.split(seps, val):
            parte = parte.strip()
            if parte:
                counter[parte] += 1
    return counter


def ranked_counts_df(counter, label_col, count_col="Ocorrências", top=None):
    items = sorted(counter.items(), key=lambda x: -x[1])
    if top:
        items = items[:top]
    return pd.DataFrame(items, columns=[label_col, count_col])


# ---------------------------------------------------------------------------
# HELPERS DE UI — componentes padronizados reutilizados nas 3 abas
# ---------------------------------------------------------------------------

def section_title(text, container=st):
    """Cabeçalho de seção padrão (mesmo nível/estilo em todas as abas)."""
    container.markdown(f"#### {text}")


def kpi_card(label, value, sublabel="", color=COLOR_PRIMARY, tooltip="", container=st):
    tooltip_attr = f'data-tooltip="{tooltip}"' if tooltip else ""
    info_icon = '<span class="kpi-info-icon">i</span>' if tooltip else ""
    container.markdown(
        f"""
        <div class="kpi-card" {tooltip_attr}>
            <div class="kpi-label">{label}{info_icon}</div>
            <div class="kpi-value" style="color:{color};">{value}</div>
            <div class="kpi-sub">{sublabel}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def kpi_row(items, container=st):
    """Renderiza uma linha de kpi_cards a partir de uma lista de dicts
    {label, value, sublabel?, color?, tooltip?} — usado em todas as abas
    para que os indicadores tenham sempre a mesma cara."""
    if not items:
        return
    cols = container.columns(len(items))
    for col, item in zip(cols, items):
        kpi_card(
            item["label"],
            item["value"],
            item.get("sublabel", ""),
            color=item.get("color", COLOR_PRIMARY),
            tooltip=item.get("tooltip", ""),
            container=col,
        )


def apply_chart_theme(fig, height=260, showlegend=False):
    """Layout padrão (fundo transparente, fonte, margens) para todo gráfico Plotly."""
    fig.update_layout(
        height=height,
        margin=dict(t=10, b=10, l=10, r=10),
        showlegend=showlegend,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font_color=TEXT_COLOR,
    )
    return fig


def horizontal_bar_chart(df, x_col, y_col, color=COLOR_PRIMARY, height=280):
    fig = px.bar(df, x=x_col, y=y_col, orientation="h")
    fig.update_traces(marker_color=color)
    apply_chart_theme(fig, height=height)
    return fig


def render_ranked_bar(counter, label_col, container=st, color=COLOR_PRIMARY, top=None,
                       height=260, key="", empty_msg="Sem dados preenchidos."):
    """A partir de um Counter, monta e plota um ranking em barras horizontais —
    usado para 'locais de origem' e 'convênios' tanto no Command Center quanto
    na aba Eletivos."""
    if not counter:
        container.info(empty_msg)
        return
    df_rank = ranked_counts_df(counter, label_col, top=top)
    fig = horizontal_bar_chart(df_rank, "Ocorrências", label_col, color=color, height=height)
    container.plotly_chart(fig, use_container_width=True, key=key)


def render_acomodacao_pie(df, container=st, height=260, key="acomod"):
    """Gráfico de pizza de tipo de acomodação — mesmo componente usado no
    Command Center e na aba Elegíveis, para que a visão seja idêntica nos
    dois lugares."""
    if COL_ACOMOD_PRIVATIVO not in df.columns and COL_ACOMOD_SEMI not in df.columns:
        container.info("Sem dados de acomodação.")
        return
    acomod_series = df.apply(classify_acomodacao, axis=1)
    acomod_counts = acomod_series.value_counts().rename_axis("Tipo").reset_index(name="Qtd")
    fig = px.pie(
        acomod_counts, names="Tipo", values="Qtd", hole=0.55,
        color_discrete_sequence=PALETTE_SEQUENCE,
    )
    apply_chart_theme(fig, height=height, showlegend=True)
    fig.update_layout(legend=dict(orientation="h", y=-0.15))
    container.plotly_chart(fig, use_container_width=True, key=key)


def render_comparisons(source, container=st, key_prefix="cmp"):
    for label, campo_elegivel, campo_efetivado in COMPARISONS:
        elegivel = float(source[campo_elegivel])
        efetivado = float(source[campo_efetivado])
        pct = calc_pct(elegivel, efetivado)

        chart_key = f"{key_prefix}_{campo_elegivel}_{campo_efetivado}"

        container.markdown(f"**{label}**")
        kpi_row(
            [
                {"label": "Solicitadas/Elegíveis", "value": int(elegivel)},
                {"label": "Efetivadas", "value": int(efetivado), "color": COLOR_SUCCESS},
                {"label": "% de execução", "value": f"{pct:.0f}%", "color": status_color(pct)},
            ],
            container=container,
        )

        fig = px.bar(
            x=["Solicitadas/Elegíveis", "Efetivadas"],
            y=[elegivel, efetivado],
            text=[int(elegivel), int(efetivado)],
            labels={"x": "", "y": ""},
        )
        fig.update_traces(textposition="outside", marker_color=[COLOR_PRIMARY, COLOR_SUCCESS])
        apply_chart_theme(fig, height=220)
        container.plotly_chart(fig, use_container_width=True, key=chart_key)
        container.divider()


def render_value_counts(df, col, label, container=st, key_prefix="vc"):
    """Cartões de contagem para colunas categóricas (Sim/Não/etc.), com cor
    verde/vermelho/ciano padronizada conforme a resposta."""
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

    items = []
    for valor, qtd in vc.items():
        val_lower = valor.strip().lower()
        if val_lower == "sim":
            color = COLOR_SUCCESS
        elif val_lower in ("não", "nao"):
            color = COLOR_DANGER
        else:
            color = COLOR_PRIMARY
        items.append({"label": valor, "value": int(qtd), "color": color})

    kpi_row(items, container=container)


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
        if field not in df.columns:
            df[field] = 0
            continue
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
    if len(header) > 25:
        header[25] = COL_CIDADE

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

        /* st.metric nativo (usado só em locais pontuais) */
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

        /* Espaçamento entre colunas/cards */
        div[data-testid="stHorizontalBlock"]{
            gap: 20px;
            margin-bottom: 22px;
        }

        /* Cards de KPI customizados — componente único usado nas 3 abas */
        .kpi-card{
            position: relative;
            background: linear-gradient(145deg, #0d1526, #0a0f1c);
            border: 1px solid var(--cc-border);
            border-radius: 14px;
            padding: 20px 22px;
            box-shadow: 0 0 18px rgba(0,229,255,0.06);
            height: 100%;
            cursor: help;
            transition: border-color 0.15s ease, box-shadow 0.15s ease;
        }
        .kpi-card:hover{
            border-color: var(--cc-cyan);
            box-shadow: 0 0 22px rgba(0,229,255,0.18);
        }
        .kpi-label{
            display: flex;
            align-items: center;
            gap: 6px;
            color: var(--cc-muted);
            font-size: 0.75rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 8px;
        }
        .kpi-info-icon{
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 14px;
            height: 14px;
            border-radius: 50%;
            border: 1px solid var(--cc-muted);
            color: var(--cc-muted);
            font-size: 0.65rem;
            font-style: normal;
            flex-shrink: 0;
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
            margin-top: 8px;
        }

        /* Tooltip customizado ao passar o mouse no card */
        .kpi-card[data-tooltip]:hover::after{
            content: attr(data-tooltip);
            position: absolute;
            bottom: calc(100% + 10px);
            left: 50%;
            transform: translateX(-50%);
            background: #101a2e;
            border: 1px solid var(--cc-border);
            color: var(--cc-text);
            padding: 10px 12px;
            border-radius: 8px;
            font-size: 0.72rem;
            line-height: 1.35;
            white-space: normal;
            width: 230px;
            text-align: left;
            box-shadow: 0 6px 20px rgba(0,0,0,0.55);
            z-index: 100;
        }
        .kpi-card[data-tooltip]:hover::before{
            content: "";
            position: absolute;
            bottom: 100%;
            left: 50%;
            transform: translateX(-50%);
            border: 6px solid transparent;
            border-top-color: var(--cc-border);
            z-index: 100;
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

        /* Botões de atualização de dados na sidebar */
        .refresh-block button{
            width: 100%;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


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
            st.caption("Auto-atualização desligada. Use os botões de atualizar na barra lateral.")

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

    section_title("📌 Indicadores-chave — Eletivos")
    if df_eletivos is None or df_eletivos.empty:
        st.info("Sem dados de Eletivos para o período selecionado — ajuste o filtro de Mês/Ano na barra lateral.")
    else:
        totals = df_eletivos.reindex(columns=NUMERIC_FIELDS, fill_value=0).sum()
        elegivel = float(totals["internacoes_eletivas"])
        efetivado = float(totals["internacoes_eletivas_efetivadas"])
        pct_execucao = calc_pct(elegivel, efetivado)

        kpi_row([
            {
                "label": "Internações Eletivas", "value": int(elegivel), "sublabel": "solicitadas no mês",
                "tooltip": "Total de internações eletivas solicitadas no mês, somando os valores registrados em cada aba diária da planilha de Eletivos.",
            },
            {
                "label": "Efetivadas", "value": int(efetivado), "sublabel": "internações realizadas",
                "color": COLOR_SUCCESS,
                "tooltip": "Quantidade de internações eletivas que efetivamente aconteceram no mês (o paciente foi internado).",
            },
            {
                "label": "% Execução", "value": f"{pct_execucao:.0f}%", "sublabel": "solicitadas → efetivadas",
                "color": status_color(pct_execucao),
                "tooltip": "Percentual de internações efetivadas em relação às solicitadas. Verde ≥ 70%, laranja entre 40% e 69%, vermelho < 40%.",
            },
            {
                "label": "Pacientes QT Autorizados", "value": int(totals["pacientes_qt_autorizados"]), "sublabel": "no mês",
                "tooltip": "Quantidade de pacientes autorizados para quimioterapia (QT) no período selecionado.",
            },
        ])

    section_title("📌 Indicadores-chave — Elegíveis")
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
        total_elegiveis_sim = int(elegivel_counts.get("SIM", 0))
        total_aceitos_sim = int(aceito_counts.get("SIM", 0))
        pct_aceite = calc_pct(total_elegiveis_sim, total_aceitos_sim)

        kpi_row([
            {
                "label": "Respostas Recebidas", "value": total_registros, "sublabel": "formulário de elegibilidade",
                "tooltip": "Total de respostas registradas no formulário de elegibilidade dentro do período filtrado.",
            },
            {
                "label": "Elegíveis (Sim)", "value": total_elegiveis_sim, "color": COLOR_SUCCESS,
                "tooltip": "Quantidade de pacientes marcados como 'Sim' no campo 'Elegível para HMV' do formulário.",
            },
            {
                "label": "Aceitos (Sim)", "value": total_aceitos_sim, "color": COLOR_PRIMARY,
                "tooltip": "Quantidade de pacientes marcados como 'Sim' no campo 'Paciente aceito' do formulário.",
            },
            {
                "label": "% Aceite s/ Elegíveis", "value": f"{pct_aceite:.0f}%", "sublabel": "aceitos → elegíveis",
                "color": status_color(pct_aceite),
                "tooltip": "Percentual de pacientes aceitos em relação aos elegíveis. Mesma régua de cores do % de Execução.",
            },
        ])
    else:
        st.info("Planilha de Elegíveis não conectada — configure SPREADSHEET_KEY_ELEGIVEIS em st.secrets para ver esses indicadores aqui.")

    st.divider()
    section_title("🏥 Ocupação e Origem")
    col_a, col_b, col_c = st.columns(3)

    with col_a:
        st.markdown("**Tipo de acomodação (Elegíveis)**")
        if df_elegiveis_full is not None and not df_elegiveis_full.empty:
            render_acomodacao_pie(df_elegiveis_full, container=col_a, key="cc_acomod")
        else:
            col_a.info("Sem dados de acomodação.")

    with col_b:
        st.markdown("**Top locais de origem (Eletivos)**")
        if df_eletivos is not None and not df_eletivos.empty:
            origem_counter = build_split_counter(df_eletivos["locais_municipios_origem"])
            render_ranked_bar(
                origem_counter, "Local", container=col_b, color=COLOR_PRIMARY, top=5,
                key="cc_origem", empty_msg="Sem dados de origem.",
            )
        else:
            col_b.info("Sem dados de origem.")

    with col_c:
        st.markdown("**Top convênios (Eletivos)**")
        if df_eletivos is not None and not df_eletivos.empty:
            conv_counter = build_split_counter(df_eletivos["convenios"])
            render_ranked_bar(
                conv_counter, "Convênio", container=col_c, color=COLOR_SUCCESS, top=5,
                key="cc_conv", empty_msg="Sem dados de convênio.",
            )
        else:
            col_c.info("Sem dados de convênio.")

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
    st.header("📅 Filtros — Eletivos")
    today = date.today()
    ano = st.number_input("Ano", min_value=2020, max_value=2100, value=today.year, step=1)
    mes = st.selectbox(
        "Mês",
        options=list(range(1, 13)),
        index=today.month - 1,
        format_func=lambda m: calendar.month_name[m].capitalize(),
    )

    st.divider()
    st.header("🔄 Atualização de dados")
    st.markdown('<div class="refresh-block">', unsafe_allow_html=True)

    elegiveis_conectado = get_spreadsheet_elegiveis() is not None

    if elegiveis_conectado:
        r1, r2 = st.columns(2)
        with r1:
            if st.button("Eletivos", key="refresh_eletivos"):
                load_data.clear()
                st.rerun()
        with r2:
            if st.button("Elegíveis", key="refresh_elegiveis"):
                load_data_elegiveis.clear()
                st.rerun()
        if st.button("🔁 Atualizar tudo", key="refresh_all"):
            load_data.clear()
            load_data_elegiveis.clear()
            st.rerun()
    else:
        if st.button("🔄 Recarregar dados (Eletivos)", key="refresh_eletivos_only"):
            load_data.clear()
            st.rerun()
        st.caption("Elegíveis não conectado — configure SPREADSHEET_KEY_ELEGIVEIS em st.secrets.")

    st.markdown("</div>", unsafe_allow_html=True)

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
            section_title(f"📌 Indicadores-chave — {calendar.month_name[mes].capitalize()}/{ano}")

            totals = df.reindex(columns=NUMERIC_FIELDS, fill_value=0).sum()
            kpi_row([
                {"label": LABELS["internacoes_eletivas"], "value": int(totals["internacoes_eletivas"])},
                {"label": LABELS["internacoes_eletivas_efetivadas"], "value": int(totals["internacoes_eletivas_efetivadas"]), "color": COLOR_SUCCESS},
                {"label": LABELS["pacientes_qt_autorizados"], "value": int(totals["pacientes_qt_autorizados"])},
            ])

            st.divider()

            section_title("⚖️ Comparativo Mensal")
            render_comparisons(totals, key_prefix="geral")

            section_title("📊 Totais detalhados do mês")
            totals_df = pd.DataFrame({
                "Indicador": [LABELS[f] for f in NUMERIC_FIELDS],
                "Total": [int(totals[f]) for f in NUMERIC_FIELDS],
            }).sort_values("Total", ascending=False)
            fig_totais = horizontal_bar_chart(totals_df, "Total", "Indicador", color=COLOR_PRIMARY, height=320)
            st.plotly_chart(fig_totais, use_container_width=True, key="geral_totais")

            section_title("📈 Evolução diária")
            campo_evolucao = st.selectbox(
                "Escolha o indicador",
                options=NUMERIC_FIELDS,
                format_func=lambda f: LABELS[f],
                key="campo_evolucao",
            )
            fig = px.line(df, x="data", y=campo_evolucao, markers=True,
                           labels={"data": "Dia", campo_evolucao: LABELS[campo_evolucao]})
            fig.update_traces(line_color=COLOR_PRIMARY, marker_color=COLOR_SUCCESS)
            apply_chart_theme(fig, height=320)
            st.plotly_chart(fig, use_container_width=True, key="geral_evolucao")

            with st.expander("Ver tabela completa"):
                display_df = df.copy()
                display_df = display_df.rename(columns={**LABELS, "data": "Data", "aba": "Aba"})
                st.dataframe(display_df, use_container_width=True)

        # --- POR DIA -----------------------------------------------------------
        with tab_dia:
            section_title("📅 Consultar um dia específico")

            dia_escolhido = st.selectbox(
                "Selecione o dia",
                options=df["data"].tolist(),
                format_func=lambda d: d.strftime("%d/%m/%Y"),
            )

            row = df[df["data"] == dia_escolhido].iloc[0]
            st.caption(f"Aba de origem: **{row['aba']}**")

            kpi_row([
                {"label": LABELS["internacoes_eletivas"], "value": int(row["internacoes_eletivas"])},
                {"label": LABELS["internacoes_eletivas_efetivadas"], "value": int(row["internacoes_eletivas_efetivadas"]), "color": COLOR_SUCCESS},
                {"label": LABELS["pacientes_qt_autorizados"], "value": int(row["pacientes_qt_autorizados"])},
            ])

            st.divider()
            section_title("⚖️ Comparativo do dia")
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
            section_title("📍 Locais / Municípios de Origem e Convênios")
            c_loc, c_conv = st.columns(2)

            with c_loc:
                st.markdown("**Locais / Municípios de Origem**")
                origem_counter = build_split_counter(df["locais_municipios_origem"])
                render_ranked_bar(
                    origem_counter, "Local", container=c_loc, color=COLOR_PRIMARY,
                    height=320, key="local_origem", empty_msg="Nenhum local preenchido.",
                )

            with c_conv:
                st.markdown("**Convênios**")
                conv_counter = build_split_counter(df["convenios"])
                render_ranked_bar(
                    conv_counter, "Convênio", container=c_conv, color=COLOR_SUCCESS,
                    height=320, key="local_conv", empty_msg="Nenhum convênio preenchido.",
                )


# =============================================================================
# ABA "ELEGÍVEIS"
# =============================================================================
with top_elegiveis:
    section_title("📝 Elegíveis")

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

        st.divider()
        section_title("📌 Elegibilidade e aceite")
        c1, c2 = st.columns(2)
        with c1:
            render_value_counts(df_elegiveis, COL_ELEGIVEL, "Elegível para HMV", container=c1, key_prefix="eleg")
        with c2:
            render_value_counts(df_elegiveis, COL_ACEITO, "Paciente aceito", container=c2, key_prefix="aceito")

        st.divider()

        section_title("🌍 Elegíveis e aceitos por grupo")
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

        section_title("🏨 Tipo de acomodação")
        render_acomodacao_pie(df_elegiveis, container=st, height=300, key="elegiveis_acomod")

        st.divider()
        section_title("📄 Base completa de respostas")
        st.dataframe(df_elegiveis, use_container_width=True)
