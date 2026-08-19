"""
Dashboard de Elegibilidades / Efetivações — lê planilhas do Google Sheets
e monta um painel interativo com Streamlit + Plotly.

Estrutura do app:
- Aba "🗂️ Eletivos"   -> o painel original: uma aba por dia do mês, sempre
  no mesmo layout (CELL_MAP), com as visões Geral / Por Dia / Por Local.
- Aba "📝 Elegíveis"  -> planilha de respostas de um formulário do Google
  (uma linha = uma resposta, não uma aba por dia). Está PREPARADA mas ainda
  não configurada: falta apontar a planilha e os nomes reais das colunas
  assim que ela existir (ver bloco "CONFIGURAÇÃO — ABA ELEGÍVEIS" abaixo).
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
# CONFIGURAÇÃO — ABA "ELETIVOS" (ajuste aqui se o layout da sua planilha mudar)
# ---------------------------------------------------------------------------

CELL_MAP = {
    "elegibilidades_realizadas": "B1:D1",
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
    "pacientes_qt_autorizados": "B7:D7",
    "efetivados_ui": "B8:D8",
    "efetivados_ctia": "B9:D9",
    "efetivados_utip": "B10:D10",
    # AJUSTE AQUI: Aumentamos o range de P para Z, assim ele pega as cidades sequenciais "pro lado"
    "locais_municipios_origem": "B11:Z11",
    "local_origem_transferencia": "B12:Z12",
}

# Campos cujas células NÃO são mescladas: a informação é digitada em
# sequência, uma por célula, indo para o lado (ex: B11="Cidade A",
# C11="Cidade B", D11="Cidade C" ...). Para esses campos juntamos TODOS os
# valores não vazios do intervalo.
SEQUENCE_FIELDS = ["locais_municipios_origem", "local_origem_transferencia"]

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

NUMERIC_FIELDS = [f for f in CELL_MAP if f not in SEQUENCE_FIELDS]

# Pares "elegível x efetivado" para o comparativo com % de execução.
COMPARISONS = [
    ("Pacientes elegíveis x Efetivados", "pacientes_elegiveis", "pacientes_efetivados"),
    ("Emergência Adulta: Elegíveis x Efetivados", "elegib_emerg_adulta", "efetivados_emerg_adulta"),
    ("Emergência Pediátrica: Elegíveis x Efetivados", "elegib_emerg_pediatrica", "efetivados_emerg_pediatrica"),
    ("Internações Eletivas: Solicitadas x Efetivadas", "internacoes_eletivas", "internacoes_eletivas_efetivadas"),
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


# ---------------------------------------------------------------------------
# CONFIGURAÇÃO — ABA "ELEGÍVEIS" (formulário Google -> planilha de respostas)
# ---------------------------------------------------------------------------
# Planilha de respostas do formulário. Cada linha = uma resposta (não tem
# aba por dia). Aponte SPREADSHEET_KEY_ELEGIVEIS em st.secrets para o ID
# abaixo; reaproveita a mesma service account (gcp_service_account) já usada
# em Eletivos — não esqueça de compartilhar essa planilha com o e-mail da
# service account também.

# ID da aba específica dentro da planilha (parâmetro gid da URL). Se a
# planilha de respostas mudar de aba/gid, ajuste aqui.
ELEGIVEIS_WORKSHEET_GID = 2058249992

# Nomes exatos das colunas (copiados do cabeçalho da planilha).
COL_ELEGIVEL = "ELEGÍVEL PARA HMV :"
COL_ACEITO = "PACIENTE ACEITO :"
COL_ORIGEM = "LOCAL DE ORIGEM (Hospital, Clínica, Pronto Atendimento, residência , etc)"
COL_CIDADE = "CIDADE ORIGEM :"
COL_CONVENIO = "CONVÊNIO DO PACIENTE:"
COL_LOCAL_ENTRADA = "LOCAL ENTRADA HMV :"
COL_ACOMOD_PRIVATIVO = "TIPO DE ACOMODAÇÃO: [Privativo]"
COL_ACOMOD_SEMI = "TIPO DE ACOMODAÇÃO: [Semi-privativo]"

# Dimensões pelas quais dá pra segregar Elegível/Aceito na aba.
DIMENSIONS_ELEGIVEIS = {
    "Local de origem": COL_ORIGEM,
    "Cidade de origem": COL_CIDADE,
    "Convênio": COL_CONVENIO,
    "Local de entrada no HMV": COL_LOCAL_ENTRADA,
}

# Coluna de timestamp, usada no filtro de período.
DATE_COLUMN_ELEGIVEIS = "Carimbo de data/hora"


def calc_pct(elegivel, efetivado):
    """% de execução = efetivado / elegível * 100 (0 se elegível for 0)."""
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
        c1.metric("Elegíveis", int(elegivel))
        c2.metric("Efetivados", int(efetivado))
        c3.metric("% de execução", f"{pct:.0f}%")

        fig = px.bar(
            x=["Elegíveis", "Efetivados"],
            y=[elegivel, efetivado],
            text=[int(elegivel), int(efetivado)],
            labels={"x": "", "y": ""},
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(height=220, margin=dict(t=10, b=10, l=10, r=10), showlegend=False)
        container.plotly_chart(fig, use_container_width=True, key=chart_key)
        container.divider()


def render_value_counts(df, col, label, container=st):
    """Mostra, em metrics lado a lado, quantas respostas existem para cada
    valor distinto de uma coluna (ex: quantos 'Sim' / 'Não' em Elegível)."""
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
    """Conta ocorrências de value_col (ex: Elegível/Aceito) segregadas pelas
    alternativas presentes em dim_col (ex: origem, cidade, convênio...)."""
    if value_col not in df.columns:
        container.warning(f"Coluna '{value_col}' não encontrada na planilha.")
        return

    data = df[[dim_col, value_col]].copy()
    data[dim_col] = data[dim_col].astype(str).str.strip()
    data[value_col] = data[value_col].astype(str).str.strip()
    data = data[data[dim_col] != ""]
    data[value_col] = data[value_col].replace("", "(vazio)")

    if data.empty:
        container.info("Sem dados preenchidos para essa segregação.")
        return

    counts = data.groupby([dim_col, value_col]).size().reset_index(name="Quantidade")
    fig = px.bar(
        counts, x="Quantidade", y=dim_col, color=value_col,
        orientation="h", barmode="group",
        labels={dim_col: "", value_col: ""},
    )
    fig.update_layout(
        height=max(280, 42 * counts[dim_col].nunique()),
        legend_title_text="",
        margin=dict(t=10, b=10, l=10, r=10),
    )
    container.plotly_chart(fig, use_container_width=True)


def classify_acomodacao(row):
    """Classifica o tipo de acomodação a partir das colunas [Privativo] /
    [Semi-privativo]: se as duas estiverem vazias, considera 'Não informado'."""
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
    """Planilha de Eletivos (uma aba por dia)."""
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


@st.cache_resource(show_spinner=False)
def get_spreadsheet_elegiveis():
    """Planilha de Elegíveis (respostas de formulário). Retorna None se as
    secrets ainda não tiverem sido configuradas — a aba trata esse caso."""
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
    """Lê o valor de uma célula única ("B2") ou o primeiro valor não vazio
    dentro de um intervalo ("B1:D1")."""
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
    """Como _read_field, mas para linhas onde cada célula do intervalo tem
    UM valor diferente (ex: B11='Cidade A', C11='Cidade B', D11='Cidade C').
    Junta todos os valores não vazios encontrados."""
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
    """Lê a planilha de respostas do formulário de Elegíveis.

    Diferente de Eletivos, aqui NÃO existe uma aba por dia: é uma planilha
    "achatada" (uma linha = uma resposta do formulário), então lemos a
    primeira aba (ou a informada em `worksheet_name`) inteira e usamos a
    primeira linha como cabeçalho.

    Retorna None se a conexão ainda não foi configurada (secrets ausentes),
    para a UI diferenciar "não configurado" de "configurado mas vazio".
    """
    ss = get_spreadsheet_elegiveis()
    if ss is None:
        return None

    if worksheet_name:
        ws = ss.worksheet(worksheet_name)
    elif ELEGIVEIS_WORKSHEET_GID is not None:
        ws = ss.get_worksheet_by_id(ELEGIVEIS_WORKSHEET_GID)
    else:
        ws = ss.sheet1

    records = ws.get_all_records()
    df = pd.DataFrame(records)

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
# ABA "ELETIVOS" — painel original (uma aba por dia do mês)
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

        st.markdown("### Comparativo: Elegíveis x Efetivados (mês)")
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

        st.markdown("### Comparativo: Elegíveis x Efetivados (dia)")
        render_comparisons(row, key_prefix=f"dia_{dia_escolhido.isoformat()}")

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


# =============================================================================
# ABA "ELEGÍVEIS" — planilha de respostas de formulário (tabelão + KPIs)
# =============================================================================
with top_elegiveis:
    st.subheader("Elegíveis")

    if get_spreadsheet_elegiveis() is None:
        st.info(
            "Esta aba já está com a conexão preparada, mas ainda falta "
            "apontar a planilha. Assim que ela existir, defina em "
            "`st.secrets` a chave `SPREADSHEET_KEY_ELEGIVEIS` (ou "
            "`SPREADSHEET_URL_ELEGIVEIS`) e preencha "
            "`KEY_COUNT_COLUMNS_ELEGIVEIS` / `DATE_COLUMN_ELEGIVEIS` no "
            "topo do arquivo com os nomes reais das colunas do formulário."
        )
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

            # --- Filtro por data ------------------------------------------
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

            # --- KPIs: Elegível para HMV / Paciente aceito ------------------
            st.markdown("### Elegibilidade e aceite")
            c1, c2 = st.columns(2)
            with c1:
                render_value_counts(df_elegiveis, COL_ELEGIVEL, "Elegível para HMV", container=c1)
            with c2:
                render_value_counts(df_elegiveis, COL_ACEITO, "Paciente aceito", container=c2)

            st.divider()

            # --- Segregação por origem / cidade / convênio / local de entrada
            st.markdown("### Elegíveis e aceitos por grupo")
            dim_label = st.selectbox(
                "Segregar por",
                options=list(DIMENSIONS_ELEGIVEIS.keys()),
                key="dim_elegiveis",
            )
            dim_col = DIMENSIONS_ELEGIVEIS[dim_label]

            if dim_col not in df_elegiveis.columns:
                st.warning(f"Coluna '{dim_col}' não encontrada na planilha.")
            else:
                gc1, gc2 = st.columns(2)
                with gc1:
                    st.caption(f"Elegível para HMV, por {dim_label.lower()}")
                    render_group_breakdown(df_elegiveis, dim_col, COL_ELEGIVEL, container=gc1)
                with gc2:
                    st.caption(f"Paciente aceito, por {dim_label.lower()}")
                    render_group_breakdown(df_elegiveis, dim_col, COL_ACEITO, container=gc2)

            st.divider()

            # --- Tipo de acomodação ------------------------------------------
            st.markdown("### Tipo de acomodação")
            if COL_ACOMOD_PRIVATIVO in df_elegiveis.columns or COL_ACOMOD_SEMI in df_elegiveis.columns:
                acomod_series = df_elegiveis.apply(classify_acomodacao, axis=1)
                acomod_counts = (
                    acomod_series.value_counts()
                    .rename_axis("Tipo de acomodação")
                    .reset_index(name="Quantidade")
                )
                st.plotly_chart(
                    px.bar(acomod_counts, x="Quantidade", y="Tipo de acomodação", orientation="h"),
                    use_container_width=True,
                )
            else:
                st.warning("Colunas de tipo de acomodação não encontradas na planilha.")

            st.divider()

            # --- Tabelão completo -------------------------------------------
            st.markdown("### Base completa de respostas")
            st.dataframe(df_elegiveis, use_container_width=True)
