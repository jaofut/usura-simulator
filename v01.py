import os
import io
from datetime import date
import pandas as pd
import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_DISPONIVEL = True
except ImportError:
    GSHEETS_DISPONIVEL = False

# Configurações do Google Sheets
NOME_PLANILHA_GSHEETS = "Gerenciador de Empréstimos - Usura Simulator"
NOME_ABA_GSHEETS = "emprestimos"

COLUNAS = [
    "ID",
    "Nome_Emprestador",
    "Nome_Tomador",
    "Valor",                     # capital original emprestado (histórico, não muda)
    "Data_Emprestimo",
    "Taxa_Mensal (%)",
    "Data_Vencimento",           # empréstimo + 1 mês (referência para status "em atraso")
    "Saldo_Principal",           # capital ainda em aberto (some com pagamentos)
    "Juros_Travados",            # juros já acumulados e não pagos, congelados no último pagamento
    "Data_Base_Juros",           # data a partir da qual novos juros ainda não foram contados
    "Valor_Pago_Total",          # soma histórica de tudo que já foi recebido (informativo)
    "Data_Ultimo_Pagamento",
]


# Persistência


def _migrar_linha_antiga(row) -> dict:
    """Reaproveita empréstimos salvos em versões antigas do arquivo,
    calculando o saldo devedor total delas 'hoje' e convertendo tudo
    para Saldo_Principal, com Juros_Travados = 0 e a contagem de juros
    reiniciando a partir de hoje."""
    hoje = date.today()
    valor = row.get("Valor", 0.0) or 0.0
    taxa_mensal = row.get("Taxa_Mensal (%)", 0.0) or 0.0
    valor_pago_antigo = row.get("Valor_Pago", row.get("Valor_Pago_Total", 0.0)) or 0.0
    data_vencimento = row.get("Data_Vencimento", pd.NaT)

    montante_base = valor * (1 + taxa_mensal / 100)
    saldo_principal_antigo = max(montante_base - valor_pago_antigo, 0.0)

    if pd.notna(data_vencimento) and hoje > data_vencimento.date() and saldo_principal_antigo > 0:
        dias_atraso = (hoje - data_vencimento.date()).days
        taxa_diaria = (taxa_mensal / 100) / 30
        juros_atraso = saldo_principal_antigo * taxa_diaria * dias_atraso
    else:
        juros_atraso = 0.0

    return {
        "Saldo_Principal": saldo_principal_antigo + juros_atraso,
        "Juros_Travados": 0.0,
        "Data_Base_Juros": pd.Timestamp(hoje),
        "Valor_Pago_Total": valor_pago_antigo,
    }


def conectar_worksheet_gsheets():
    """Conecta ao Google Sheets usando credenciais (Service Account via st.secrets ou arquivo local)
    e retorna (worksheet, mensagem_status)."""
    if not GSHEETS_DISPONIVEL:
        return None, "Bibliotecas gspread/google-auth não estão instaladas. Execute: pip install gspread google-auth"

    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    creds = None
    origem_creds = ""

    # 1. Tenta carregar do st.secrets
    try:
        if hasattr(st, "secrets") and st.secrets:
            if "gcp_service_account" in st.secrets:
                creds_dict = dict(st.secrets["gcp_service_account"])
                creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                origem_creds = "st.secrets (gcp_service_account)"
            elif "connections" in st.secrets and "gsheets" in st.secrets["connections"]:
                cfg = st.secrets["connections"]["gsheets"]
                if "client_email" in cfg and "private_key" in cfg:
                    creds_dict = dict(cfg)
                    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                    origem_creds = "st.secrets (connections.gsheets)"
    except Exception:
        pass

    # 2. Se não achou no secrets, tenta arquivos locais JSON
    if creds is None:
        arquivos_candidatos = ["credentials.json", "service_account.json", "gsheets_credentials.json"]
        for arq in arquivos_candidatos:
            if os.path.exists(arq):
                try:
                    creds = Credentials.from_service_account_file(arq, scopes=scopes)
                    origem_creds = f"arquivo local ({arq})"
                    break
                except Exception:
                    continue

    if creds is None:
        return None, "Nenhuma credencial (Service Account JSON) encontrada no st.secrets ou arquivo credentials.json local."

    try:
        client = gspread.authorize(creds)
    except Exception as e:
        return None, f"Erro na autenticação com o Google: {str(e)}"

    # Identifica a planilha alvo (por URL, ID ou Título)
    spreadsheet_alvo = NOME_PLANILHA_GSHEETS
    try:
        if hasattr(st, "secrets") and st.secrets:
            if "spreadsheet_url" in st.secrets:
                spreadsheet_alvo = st.secrets["spreadsheet_url"]
            elif "connections" in st.secrets and "gsheets" in st.secrets["connections"] and "spreadsheet" in st.secrets["connections"]["gsheets"]:
                spreadsheet_alvo = st.secrets["connections"]["gsheets"]["spreadsheet"]
    except Exception:
        pass

    sh = None
    try:
        if spreadsheet_alvo.startswith("https://"):
            sh = client.open_by_url(spreadsheet_alvo)
        elif "/" not in spreadsheet_alvo and " " not in spreadsheet_alvo and len(spreadsheet_alvo) > 35:
            sh = client.open_by_key(spreadsheet_alvo)
        else:
            sh = client.open(spreadsheet_alvo)
    except gspread.exceptions.SpreadsheetNotFound:
        try:
            sh = client.create(spreadsheet_alvo)
        except Exception as e_create:
            return None, f"Planilha '{spreadsheet_alvo}' não encontrada e falha ao criá-la: {str(e_create)}. Verifique se compartilhou a planilha com: {getattr(creds, 'service_account_email', 'sua Service Account')}."
    except Exception as e_open:
        return None, f"Erro ao abrir a planilha '{spreadsheet_alvo}': {str(e_open)}"

    try:
        worksheet = sh.worksheet(NOME_ABA_GSHEETS)
    except gspread.exceptions.WorksheetNotFound:
        try:
            primeira_aba = sh.get_worksheet(0)
            if primeira_aba and primeira_aba.title in ["Página1", "Sheet1", "Página 1", "Sheet 1"] and primeira_aba.get_all_values() == []:
                primeira_aba.update_title(NOME_ABA_GSHEETS)
                worksheet = primeira_aba
            else:
                worksheet = sh.add_worksheet(title=NOME_ABA_GSHEETS, rows=1000, cols=25)
        except Exception as e_ws:
            return None, f"Erro ao acessar ou criar a aba '{NOME_ABA_GSHEETS}': {str(e_ws)}"

    return worksheet, f"Conectado via {origem_creds} — Planilha: {sh.title}"


@st.cache_data(ttl=30, show_spinner="Carregando dados do Google Sheets...")
def _ler_dados_gsheets() -> list:
    worksheet, erro = conectar_worksheet_gsheets()
    if worksheet is None:
        return []
    try:
        return worksheet.get_all_records()
    except Exception:
        return []


def carregar_dados() -> pd.DataFrame:
    worksheet, msg = conectar_worksheet_gsheets()
    if worksheet is None:
        if "aviso_gsheets_mostrado" not in st.session_state:
            st.warning(f"☁️ Google Sheets não conectado: {msg}")
            st.info(
                "💡 **Como conectar ao Google Sheets:**\n"
                "1. Baixe o JSON da sua Service Account no Google Cloud Console e salve como **`credentials.json`** nesta pasta (`usura simulator/`) OU configure no `.streamlit/secrets.toml`.\n"
                "2. Crie a planilha no Google Drive chamada **`Gerenciador de Empréstimos - Usura Simulator`** e **compartilhe como Editor** com o e-mail (`client_email`) do JSON.\n"
                "*(Enquanto a conexão não estiver ativa, os dados são armazenados temporariamente na memória da sessão)*"
            )
            st.session_state["aviso_gsheets_mostrado"] = True
        
        if "df_memoria_gsheets" in st.session_state:
            df = st.session_state["df_memoria_gsheets"].copy()
        else:
            df = pd.DataFrame(columns=COLUNAS)
    else:
        records = _ler_dados_gsheets()
        if not records:
            df = pd.DataFrame(columns=COLUNAS)
        else:
            df = pd.DataFrame(records)
        st.session_state["df_memoria_gsheets"] = df.copy()

    for col in ["Data_Emprestimo", "Data_Vencimento", "Data_Base_Juros", "Data_Ultimo_Pagamento"]:
        if col not in df.columns:
            df[col] = pd.NaT
        df[col] = pd.to_datetime(df[col], errors="coerce")

    if "Data_Vencimento" in df.columns:
        sem_vencimento = df["Data_Vencimento"].isna() & df["Data_Emprestimo"].notna()
        df.loc[sem_vencimento, "Data_Vencimento"] = df.loc[sem_vencimento, "Data_Emprestimo"] + pd.DateOffset(months=1)

    # Migração: arquivo de versão antiga não tem Saldo_Principal / Juros_Travados.
    schema_antigo = "Saldo_Principal" not in df.columns or "Juros_Travados" not in df.columns
    if schema_antigo:
        for col in ["Saldo_Principal", "Juros_Travados", "Data_Base_Juros", "Valor_Pago_Total"]:
            if col not in df.columns:
                df[col] = None
        for idx, row in df.iterrows():
            migrado = _migrar_linha_antiga(row)
            for col, val in migrado.items():
                df.at[idx, col] = val

    for col in ["Saldo_Principal", "Juros_Travados", "Valor_Pago_Total"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    if "ID" in df.columns:
        df["ID"] = pd.to_numeric(df["ID"], errors="coerce").fillna(0).astype(int)
    if "Valor" in df.columns:
        df["Valor"] = pd.to_numeric(df["Valor"], errors="coerce").fillna(0.0)
    if "Taxa_Mensal (%)" in df.columns:
        df["Taxa_Mensal (%)"] = pd.to_numeric(df["Taxa_Mensal (%)"], errors="coerce").fillna(0.0)

    for col in COLUNAS:
        if col not in df.columns:
            df[col] = None

    return df[COLUNAS]


def salvar_dados(df: pd.DataFrame) -> None:
    st.session_state["df_memoria_gsheets"] = df.copy()
    
    worksheet, msg = conectar_worksheet_gsheets()
    if worksheet is None:
        st.error(f"❌ Não foi possível salvar na planilha do Google Sheets ({msg}). Dados retidos em memória na sessão.")
        return

    df_salvar = df[COLUNAS].copy()

    # Formata datas para string YYYY-MM-DD
    for col in ["Data_Emprestimo", "Data_Vencimento", "Data_Base_Juros", "Data_Ultimo_Pagamento"]:
        if col in df_salvar.columns:
            df_salvar[col] = df_salvar[col].apply(
                lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) and isinstance(x, (pd.Timestamp, date)) else ""
            )

    # Trata nulos em colunas numéricas
    for col in ["Valor", "Taxa_Mensal (%)", "Saldo_Principal", "Juros_Travados", "Valor_Pago_Total"]:
        if col in df_salvar.columns:
            df_salvar[col] = pd.to_numeric(df_salvar[col], errors="coerce").fillna(0.0)

    if "ID" in df_salvar.columns:
        df_salvar["ID"] = pd.to_numeric(df_salvar["ID"], errors="coerce").fillna(0).astype(int)

    # Preenche o restante com strings vazias
    df_salvar = df_salvar.fillna("")

    rows_data = []
    for _, row in df_salvar.iterrows():
        linha_pronta = []
        for col in COLUNAS:
            val = row[col]
            if pd.isna(val) or val is None or str(val) in ("nan", "None", "<NA>"):
                val = ""
            elif isinstance(val, (int, float)):
                val = float(val) if isinstance(val, float) else int(val)
            else:
                val = str(val)
            linha_pronta.append(val)
        rows_data.append(linha_pronta)

    table_data = [COLUNAS] + rows_data

    try:
        worksheet.clear()
        try:
            worksheet.update(values=table_data, range_name="A1", value_input_option="USER_ENTERED")
        except TypeError:
            worksheet.update("A1", table_data, value_input_option="USER_ENTERED")
        _ler_dados_gsheets.clear()
    except Exception as e:
        st.error(f"❌ Erro ao atualizar o Google Sheets: {str(e)}")


def proximo_id(df: pd.DataFrame) -> int:
    if df.empty:
        return 1
    return int(df["ID"].max()) + 1


# ------------------------------------------------------------------
# Cálculo de juros e saldo
# ------------------------------------------------------------------

def taxa_diaria(taxa_mensal_pct: float) -> float:
    return (taxa_mensal_pct / 100) / 30


def calcular_estado_atual(row, data_referencia: date):
    """Retorna (saldo_principal, juros_ate_agora, saldo_devedor_total, dias_atraso)."""
    data_base = row["Data_Base_Juros"].date() if not pd.isna(row["Data_Base_Juros"]) else row["Data_Emprestimo"].date()
    dias_desde_base = max((data_referencia - data_base).days, 0)

    saldo_principal = row["Saldo_Principal"]
    juros_periodo = saldo_principal * taxa_diaria(row["Taxa_Mensal (%)"]) * dias_desde_base
    juros_total = row["Juros_Travados"] + juros_periodo
    saldo_devedor_total = saldo_principal + juros_total

    if pd.isna(row["Data_Vencimento"]) or data_referencia <= row["Data_Vencimento"].date():
        dias_atraso = 0
    else:
        dias_atraso = (data_referencia - row["Data_Vencimento"].date()).days

    return saldo_principal, juros_total, saldo_devedor_total, dias_atraso


def definir_status(saldo_principal: float, valor_pago_total: float, dias_atraso: int) -> str:
    if saldo_principal <= 0.005:
        return "Pago"
    if valor_pago_total > 0:
        return "Parcialmente pago"
    if dias_atraso > 0:
        return "Em atraso"
    return "Em dia"


def registrar_pagamento(df: pd.DataFrame, id_emprestimo: int, valor_pago: float, data_pagamento: date) -> float:
    """Abate primeiro os juros já acumulados e não pagos; o que sobrar
    abate o saldo de capital. Retorna o novo saldo principal."""
    idx = df.index[df["ID"] == id_emprestimo][0]
    row = df.loc[idx]

    _, juros_total, _, _ = calcular_estado_atual(row, data_pagamento)

    if valor_pago <= juros_total:
        novos_juros_travados = juros_total - valor_pago
        novo_saldo_principal = row["Saldo_Principal"]
    else:
        restante_para_capital = valor_pago - juros_total
        novos_juros_travados = 0.0
        novo_saldo_principal = max(row["Saldo_Principal"] - restante_para_capital, 0.0)

    df.at[idx, "Saldo_Principal"] = novo_saldo_principal
    df.at[idx, "Juros_Travados"] = novos_juros_travados
    df.at[idx, "Data_Base_Juros"] = pd.Timestamp(data_pagamento)
    df.at[idx, "Valor_Pago_Total"] = row["Valor_Pago_Total"] + valor_pago
    df.at[idx, "Data_Ultimo_Pagamento"] = pd.Timestamp(data_pagamento)

    return novo_saldo_principal


def montar_tabela_exibicao(df: pd.DataFrame) -> pd.DataFrame:
    hoje = date.today()
    linhas = []
    for _, row in df.iterrows():
        saldo_principal, juros_total, saldo_devedor, atraso = calcular_estado_atual(row, hoje)
        status = definir_status(saldo_principal, row["Valor_Pago_Total"], atraso)

        linhas.append(
            {
                "ID": row["ID"],
                "Emprestador": row["Nome_Emprestador"],
                "Tomador": row["Nome_Tomador"],
                "Capital original (R$)": round(row["Valor"], 2),
                "Data do empréstimo": row["Data_Emprestimo"].date() if not pd.isna(row["Data_Emprestimo"]) else None,
                "Taxa mensal (%)": row["Taxa_Mensal (%)"],
                "Vencimento (1 mês)": row["Data_Vencimento"].date() if not pd.isna(row["Data_Vencimento"]) else None,
                "Status": status,
                "Dias em atraso": atraso,
                "Já recebido (R$)": round(row["Valor_Pago_Total"], 2),
                "Juros acumulados (R$)": round(juros_total, 2),
                "Saldo devedor atual (R$)": round(saldo_devedor, 2),
            }
        )
    return pd.DataFrame(linhas)


# ------------------------------------------------------------------
# Interface Streamlit
# ------------------------------------------------------------------

st.set_page_config(page_title="Gerenciador de Empréstimos", page_icon="💰", layout="wide")
st.title("💰 Gerenciador de Empréstimos")
st.caption(
    "Juros simples diários (derivados da taxa mensal) incidem sobre o saldo de capital em aberto. "
    "Cada pagamento abate primeiro os juros já acumulados e, o que sobrar, abate o capital."
)

df = carregar_dados()

aba_novo, aba_lista, aba_pagar = st.tabs(
    ["➕ Novo empréstimo", "📋 Lista de empréstimos", "✅ Registrar pagamento"]
)

# --- Aba: novo empréstimo -------------------------------------------------
with aba_novo:
    st.subheader("Cadastrar novo empréstimo")
    with st.form("form_novo_emprestimo", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            nome_emprestador = st.text_input("Nome de quem emprestou o dinheiro")
            nome_tomador = st.text_input("Nome de quem recebeu o empréstimo")
            valor = st.number_input("Valor do empréstimo (capital) (R$)", min_value=0.0, step=50.0, format="%.2f")
        with col2:
            data_emprestimo = st.date_input("Data do empréstimo", value=date.today())
            taxa_mensal = st.number_input(
                "Taxa de juros simples mensal (%)",
                min_value=0.0,
                step=0.5,
                format="%.2f",
            )
            st.caption(
                "O vencimento (montante = capital + juros do mês) é calculado automaticamente para "
                "1 mês após a data do empréstimo. Não é preciso informar taxa diária: se não for pago, "
                "os juros de atraso são derivados dessa mesma taxa mensal."
            )

        enviado = st.form_submit_button("Cadastrar empréstimo")

        if enviado:
            if not nome_emprestador.strip():
                st.error("Informe o nome de quem emprestou o dinheiro.")
            elif not nome_tomador.strip():
                st.error("Informe o nome de quem recebeu o empréstimo.")
            elif valor <= 0:
                st.error("O valor do empréstimo deve ser maior que zero.")
            else:
                data_vencimento = pd.to_datetime(data_emprestimo) + pd.DateOffset(months=1)
                novo_registro = {
                    "ID": proximo_id(df),
                    "Nome_Emprestador": nome_emprestador.strip(),
                    "Nome_Tomador": nome_tomador.strip(),
                    "Valor": valor,
                    "Data_Emprestimo": pd.to_datetime(data_emprestimo),
                    "Taxa_Mensal (%)": taxa_mensal,
                    "Data_Vencimento": data_vencimento,
                    "Saldo_Principal": valor,
                    "Juros_Travados": 0.0,
                    "Data_Base_Juros": pd.to_datetime(data_emprestimo),
                    "Valor_Pago_Total": 0.0,
                    "Data_Ultimo_Pagamento": pd.NaT,
                }
                df = pd.concat([df, pd.DataFrame([novo_registro])], ignore_index=True)
                salvar_dados(df)
                montante = valor * (1 + taxa_mensal / 100)
                st.success(
                    f"Empréstimo de R$ {valor:,.2f} de {nome_emprestador} para {nome_tomador} cadastrado. "
                    f"Montante devido em {data_vencimento.date()}: R$ {montante:,.2f}."
                )

# --- Aba: lista de empréstimos --------------------------------------------
with aba_lista:
    st.subheader("Todos os empréstimos")
    if df.empty:
        st.info("Nenhum empréstimo cadastrado ainda.")
    else:
        tabela = montar_tabela_exibicao(df)

        col_filtro1, col_filtro2 = st.columns([1, 2])
        with col_filtro1:
            filtro_status = st.selectbox(
                "Filtrar por status", ["Todos", "Em dia", "Em atraso", "Parcialmente pago", "Pago"]
            )
        with col_filtro2:
            filtro_nome = st.text_input("Buscar por nome (emprestador ou tomador)")

        tabela_filtrada = tabela.copy()
        if filtro_status != "Todos":
            tabela_filtrada = tabela_filtrada[tabela_filtrada["Status"] == filtro_status]
        if filtro_nome.strip():
            termo = filtro_nome.strip().lower()
            tabela_filtrada = tabela_filtrada[
                tabela_filtrada["Emprestador"].str.lower().str.contains(termo, na=False)
                | tabela_filtrada["Tomador"].str.lower().str.contains(termo, na=False)
            ]

        st.dataframe(tabela_filtrada, use_container_width=True, hide_index=True)

        total_emprestado = df["Valor"].sum()
        total_devido = tabela["Saldo devedor atual (R$)"].sum()
        total_juros = tabela["Juros acumulados (R$)"].sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Total emprestado (capital histórico)", f"R$ {total_emprestado:,.2f}")
        m2.metric("Total ainda a receber (todos)", f"R$ {total_devido:,.2f}")
        m3.metric("Juros acumulados (todos)", f"R$ {total_juros:,.2f}")

        st.divider()
        st.subheader("☁️ Status & Backup dos Dados")
        worksheet, msg_status = conectar_worksheet_gsheets()
        if worksheet is not None:
            url_planilha = getattr(worksheet.spreadsheet, "url", "#")
            st.success(f"✔️ Conectado ao Google Sheets: **{worksheet.spreadsheet.title}** (Aba: `{worksheet.title}`)")
            if url_planilha != "#":
                st.markdown(f"🔗 [**Abrir Planilha no Google Sheets**]({url_planilha})")
        else:
            st.warning(f"☁️ Status do Google Sheets: {msg_status}")

        col_b1, col_b2 = st.columns(2)
        with col_b1:
            csv_data = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Baixar backup em CSV",
                data=csv_data,
                file_name="emprestimos_gsheets_backup.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with col_b2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name=NOME_ABA_GSHEETS)
            st.download_button(
                "⬇️ Baixar backup em Excel (.xlsx)",
                data=buffer.getvalue(),
                file_name="emprestimos_gsheets_backup.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.divider()
        st.subheader("🔎 Quanto falta receber, por quem emprestou")

        emprestadores = sorted(df["Nome_Emprestador"].dropna().unique().tolist())
        if not emprestadores:
            st.info("Nenhum emprestador cadastrado ainda.")
        else:
            pessoa_selecionada = st.selectbox("Quem emprestou", emprestadores, key="select_emprestador")

            loans_pessoa = tabela[tabela["Emprestador"] == pessoa_selecionada]
            total_a_receber_pessoa = loans_pessoa["Saldo devedor atual (R$)"].sum()

            st.metric(
                f"Total que {pessoa_selecionada} ainda tem a receber (todos os empréstimos)",
                f"R$ {total_a_receber_pessoa:,.2f}",
            )

            opcoes_emprestimo = {
                f"#{int(r['ID'])} - para {r['Tomador']} - venc. {r['Vencimento (1 mês)']} - {r['Status']}": r["ID"]
                for _, r in loans_pessoa.iterrows()
            }
            escolha_emprestimo = st.selectbox(
                "Selecione um empréstimo específico para ver os detalhes",
                list(opcoes_emprestimo.keys()),
                key="select_emprestimo_detalhe",
            )
            id_detalhe = opcoes_emprestimo[escolha_emprestimo]
            detalhe = loans_pessoa[loans_pessoa["ID"] == id_detalhe].iloc[0]

            d1, d2, d3 = st.columns(3)
            d1.metric("Já recebido", f"R$ {detalhe['Já recebido (R$)']:,.2f}")
            d2.metric("Juros acumulados", f"R$ {detalhe['Juros acumulados (R$)']:,.2f}")
            d3.metric("Falta receber", f"R$ {detalhe['Saldo devedor atual (R$)']:,.2f}")
            st.caption(f"Status atual: **{detalhe['Status']}** — {detalhe['Dias em atraso']} dia(s) em atraso.")

# --- Aba: registrar pagamento --------------------------------------------
with aba_pagar:
    st.subheader("Registrar pagamento de um empréstimo")

    tabela_atual = montar_tabela_exibicao(df) if not df.empty else pd.DataFrame()
    pendentes_ids = tabela_atual[tabela_atual["Status"] != "Pago"]["ID"].tolist() if not tabela_atual.empty else []
    pendentes = df[df["ID"].isin(pendentes_ids)]

    if pendentes.empty:
        st.info("Não há empréstimos pendentes de pagamento.")
    else:
        opcoes = {
            f"#{int(row['ID'])} - {row['Nome_Emprestador']} → {row['Nome_Tomador']} "
            f"- venc. {row['Data_Vencimento'].date()}": row["ID"]
            for _, row in pendentes.iterrows()
        }
        escolha = st.selectbox("Selecione o empréstimo", list(opcoes.keys()))
        id_selecionado = opcoes[escolha]
        linha = df[df["ID"] == id_selecionado].iloc[0]

        hoje = date.today()
        saldo_principal, juros_hoje, saldo_total_hoje, dias_atraso = calcular_estado_atual(linha, hoje)

        c1, c2, c3 = st.columns(3)
        c1.metric("Saldo de capital em aberto", f"R$ {saldo_principal:,.2f}")
        c2.metric("Juros acumulados até hoje", f"R$ {juros_hoje:,.2f}")
        c3.metric("Saldo devedor total hoje", f"R$ {saldo_total_hoje:,.2f}")

        st.write(f"Dias em atraso: **{dias_atraso}**")

        data_pagamento = st.date_input("Data do pagamento", value=hoje, key="data_pagto")
        valor_pago_agora = st.number_input(
            "Valor recebido agora (R$) — pode ser parcial",
            min_value=0.0,
            max_value=float(round(saldo_total_hoje, 2)) if saldo_total_hoje > 0 else 0.0,
            value=float(round(saldo_total_hoje, 2)),
            format="%.2f",
        )

        if st.button("Confirmar pagamento"):
            if valor_pago_agora <= 0:
                st.error("Informe um valor maior que zero.")
            else:
                novo_saldo = registrar_pagamento(df, id_selecionado, valor_pago_agora, data_pagamento)
                salvar_dados(df)

                if novo_saldo <= 0.005:
                    st.success("Pagamento registrado — empréstimo totalmente liquidado.")
                else:
                    st.success(
                        f"Pagamento registrado. Saldo devedor após este pagamento: "
                        f"R$ {novo_saldo:,.2f} (novos juros passam a contar a partir de hoje, sobre esse saldo)."
                    )
                st.rerun()