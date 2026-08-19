"""
Dashboard de Elegibilidades / Efetivações — lê uma planilha do Google Sheets
(uma aba por dia do mês, sempre no mesmo layout) e monta um painel interativo
com Streamlit + Plotly.

Como funciona:
1. Conecta no Google Sheets via Service Account (credenciais em st.secrets).
2. Varre todas as abas da planilha, tenta interpretar o nome de cada aba
   como um dia do mês (ex: "1", "01", "01/08", "01/08/2026" ...).
3. Para cada aba reconhecida como um dia, lê os valores nas células fixas
   definidas em CELL_MAP (ajuste esse dicionário se o layout da sua
   planilha for um pouco diferente do modelo).
4. Monta um DataFrame com uma linha por dia e usa isso para alimentar
   3 visões: Visão Geral, Por Dia, Por Local.
"""

import re
import calendar
from datetime import date, datetime
from collections import Counter

import pandas as pd
import streamlit as st
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ajuste aqui se o layout da sua planilha mudar
# ---------------------------------------------------------------------------

# Endereço de cada informação dentro de CADA aba (uma aba = um dia).
# Se algum número aparecer "deslocado" no seu dashboard, o ajuste é aqui.
CELL_MAP = {
    "elegibilidades_realizadas": "C1",
    "pacientes_elegiveis": "B2",
    "pacientes_efetivados": "D2",
    "elegib_emerg_adulta": "B3",
    "efetivados_emerg_adulta": "D3",
    "elegib_emerg_pediatrica": "B4",
    "efetivados_emerg_pediatrica": "D4",
    "internacoes_eletivas": "B5",
    "internacoes_eletivas_efetivadas": "D5",
    "efetivados_co": "B6",
    "efetivados_mater": "D6",
    "pacientes_qt_autorizados": "C7",
    "efetivados_ui": "C8",
    "efetivados_ctia": "C9",
    "efetivados_utip": "C10",
    "locais_municipios_origem": "B11",
    "local_origem_transferencia": "C12",
}

# Rótulos amigáveis para exibir na tela
LABELS = {
    "elegibilidades_realizadas": "Elegibilidades realizadas",
    "pacientes_elegiveis": "Pacientes elegíveis",
    "pacientes_efetivados": "Pacientes efetivados",
    "elegib_emerg_adulta": "Elegibilidades emergência adulta",
    "efetivados_emerg_adulta": "Efetivados emergência adulta",
    "elegib_emerg_pediatrica": "Elegibilidades emergência pediátrica",
    "efetivados_emerg_pediatrica": "Efetivados emergência pediátrica",
    "internacoes_eletivas": "Internações eletivas",
    "internacoes_eletivas_efetivadas": "Internações eletivas efetivadas",
    "efetivados_co": "Pacientes efetivados no CO",
    "efetivados_mater": "Pacientes efetivados na Mater",
    "pacientes_qt_autorizados": "Pacientes QT autorizados",
    "efetivados_ui": "Pacientes efetivados na UI",
    "efetivados_ctia": "Pacientes efetivados na CTIA",
    "efetivados_utip": "Pacientes efetivados na UTIP",
}

NUMERIC_FIELDS = [f for f in CELL_MAP if f not in ("locais_municipios_origem", "local_origem_transferencia")]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


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
    raise RuntimeError(
        "Defina SPREADSHEET_KEY ou SPREADSHEET_URL em st.secrets."
    )


def _cell_ref_to_rowcol(ref: str):
    return gspread.utils.a1_to_rowcol(ref)


def parse_day_from_title(title: str, month: int, year: int):
    """Tenta interpretar o nome da aba como um dia do mês selecionado."""
    t = title.strip()

    # "1", "01", "1º", "dia 1" etc.
    m = re.fullmatch(r"(?:dia\s*)?0*(\d{1,2})º?", t, flags=re.IGNORECASE)
    if m:
        day = int(m.group(1))
        try:
            return date(year, month, day)
        except ValueError:
            return None

    # "01/08", "01-08", "01/08/2026"
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


@st.cache_data(ttl=300, show_spinner="Lendo planilha do Google Sheets...")
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
            r, c = _cell_ref_to_rowcol(ref)
            raw = ""
            if r - 1 < len(values) and c - 1 < len(values[r - 1]):
                raw = values[r - 1][c - 1]
            record[field] = raw.strip()

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


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.set_page_config(page_title="Dashboard de Elegibilidades", layout="wide")
st.title("📊 Dashboard de Elegibilidades e Efetivações")

with st.sidebar:
    st.header("Filtros")
    today = date.today()
    ano = st.number_input("Ano", min_value=2020, max_value=2100, value=today.year, step=1)
    mes = st.selectbox(
        "Mês",
        options=list(range(1, 13)),
        index=today.month - 1,
        format_func=lambda m: calendar.month_name[m].capitalize(),
    )
    if st.button("🔄 Recarregar dados"):
        st.cache_data.clear()
        st.rerun()

df, skipped_tabs = load_data(mes, ano)

if df.empty:
    st.warning(
        "Nenhuma aba reconhecida como dia do mês foi encontrada para o "
        f"período selecionado ({calendar.month_name[mes]}/{ano}). "
        "Verifique o nome das abas na planilha ou ajuste o parser "
        "`parse_day_from_title` em app.py."
    )
    st.stop()

tab_geral, tab_dia, tab_local = st.tabs(["📈 Visão Geral", "📅 Por Dia", "📍 Por Local"])

# --- VISÃO GERAL -----------------------------------------------------------
with tab_geral:
    st.subheader(f"Totais de {calendar.month_name[mes].capitalize()}/{ano}")

    totals = df[NUMERIC_FIELDS].sum()

    kpi_fields = [
        "elegibilidades_realizadas",
        "pacientes_elegiveis",
        "pacientes_efetivados",
        "internacoes_eletivas",
        "internacoes_eletivas_efetivadas",
    ]
    cols = st.columns(len(kpi_fields))
    for col, field in zip(cols, kpi_fields):
        col.metric(LABELS[field], int(totals[field]))

    st.divider()

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

    if skipped_tabs:
        with st.expander(f"{len(skipped_tabs)} aba(s) ignorada(s) (nome não reconhecido como dia)"):
            st.write(", ".join(skipped_tabs))

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
    c1.metric(LABELS["elegibilidades_realizadas"], int(row["elegibilidades_realizadas"]))
    c2.metric(LABELS["pacientes_elegiveis"], int(row["pacientes_elegiveis"]))
    c3.metric(LABELS["pacientes_efetivados"], int(row["pacientes_efetivados"]))

    st.divider()
    detalhe = pd.DataFrame({
        "Indicador": [LABELS[f] for f in NUMERIC_FIELDS],
        "Valor": [int(row[f]) for f in NUMERIC_FIELDS],
    })
    st.dataframe(detalhe, use_container_width=True, hide_index=True)

    if row.get("locais_municipios_origem"):
        st.markdown(f"**Locais/Municípios de origem:** {row['locais_municipios_origem']}")
    if row.get("local_origem_transferencia"):
        st.markdown(f"**Local de origem da transferência:** {row['local_origem_transferencia']}")

# --- POR LOCAL -----------------------------------------------------------
with tab_local:
    st.subheader("Distribuição por local de origem")

    origem_counter = Counter()
    for val in df["local_origem_transferencia"]:
        val = str(val).strip()
        if not val:
            continue
        # separa múltiplos locais escritos na mesma célula (";" ou ",")
        for parte in re.split(r"[;,/]", val):
            parte = parte.strip()
            if parte:
                origem_counter[parte] += 1

    if origem_counter:
        origem_df = pd.DataFrame(
            sorted(origem_counter.items(), key=lambda x: -x[1]),
            columns=["Local", "Ocorrências"],
        )
        st.plotly_chart(
            px.bar(origem_df, x="Ocorrências", y="Local", orientation="h"),
            use_container_width=True,
        )
        st.dataframe(origem_df, use_container_width=True, hide_index=True)

        st.divider()
        local_sel = st.selectbox("Ver dias em que esse local aparece", options=origem_df["Local"])
        filtro = df[df["local_origem_transferencia"].str.contains(local_sel, case=False, na=False)]
        st.dataframe(
            filtro[["data", "aba", "local_origem_transferencia"]].rename(
                columns={"data": "Data", "aba": "Aba", "local_origem_transferencia": "Local"}
            ),
            use_container_width=True,
            hide_index=True,
        )
    else:
        st.info("Nenhum local de origem preenchido no período selecionado.")

    municipios_counter = Counter()
    for val in df["locais_municipios_origem"]:
        val = str(val).strip()
        if not val:
            continue
        for parte in re.split(r"[;,/]", val):
            parte = parte.strip()
            if parte:
                municipios_counter[parte] += 1

    if municipios_counter:
        st.markdown("**Municípios de origem (campo separado)**")
        mun_df = pd.DataFrame(
            sorted(municipios_counter.items(), key=lambda x: -x[1]),
            columns=["Município", "Ocorrências"],
        )
        st.plotly_chart(
            px.bar(mun_df, x="Ocorrências", y="Município", orientation="h"),
            use_container_width=True,
        )
