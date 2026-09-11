import datetime
import json
import os
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# 1. CONFIGURAÇÃO DA PÁGINA E CSS EXECUTIVO PREMIUM
st.set_page_config(
    page_title="Platform Governance & Compliance | Maringá Turismo",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .stApp { background-color: #f8fafc; }
    
    /* Cards KPI Executivos */
    .kpi-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 16px 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
        border-left: 4px solid #3b82f6;
    }
    .kpi-label { font-size: 0.8rem; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: 0.5px; }
    .kpi-value { font-size: 1.8rem; font-weight: 800; color: #0f172a; margin-top: 4px; }
    .kpi-subtext { font-size: 0.75rem; color: #10b981; font-weight: 600; margin-top: 2px; }

    /* Estilização Refinada dos Cartões Kanban com Separação Nítida */
    .kanban-box {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 18px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        border-top: 5px solid #cbd5e1;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    .kanban-box:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.08);
    }
    .card-critica { border-top-color: #ef4444 !important; }
    .card-alta { border-top-color: #f97316 !important; }
    .card-media { border-top-color: #eab308 !important; }
    .card-baixa { border-top-color: #22c55e !important; }

    /* Badges */
    .badge {
        font-size: 0.70rem; font-weight: 700; padding: 3px 10px;
        border-radius: 12px; color: white; display: inline-block; text-transform: uppercase;
    }
    .badge-critica { background-color: #ef4444; }
    .badge-alta { background-color: #f97316; }
    .badge-media { background-color: #eab308; }
    .badge-baixa { background-color: #22c55e; }

    /* Timeline de Comentários no Relatório e Modal */
    .comment-item {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #3b82f6;
        padding: 12px 14px;
        margin-bottom: 10px;
        border-radius: 8px;
        font-size: 0.88rem;
    }
    .comment-header {
        font-weight: 700;
        color: #1e293b;
        font-size: 0.78rem;
        margin-bottom: 4px;
    }
    
    /* Visão Relatório Tipo Cascata / DRE */
    .dre-card {
        background-color: #ffffff;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        padding: 16px;
        margin-bottom: 16px;
    }
    .dre-header {
        font-size: 1.1rem;
        font-weight: bold;
        color: #1e3a8a;
        border-bottom: 2px solid #e2e8f0;
        padding-bottom: 6px;
        margin-bottom: 12px;
    }
</style>
""",
    unsafe_allow_html=True,
)

PASTA_EVIDENCIAS = "evidencias_acumuladas"
ARQUIVO_MATRIZ = "matriz_auditoria.xlsx"
ARQUIVO_USUARIOS = "usuarios_sistema.xlsx"
ARQUIVO_LOGS = "logs_auditoria_sistema.xlsx"

if not os.path.exists(PASTA_EVIDENCIAS):
    os.makedirs(PASTA_EVIDENCIAS)


# 2. HELPER FUNCTIONS: LOGS E USUÁRIOS
def registrar_log(usuario, acao, detalhe):
    try:
        data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        novo_log = pd.DataFrame(
            [
                {
                    "Data_Hora": data_hora,
                    "Usuario": str(usuario),
                    "Acao": str(acao),
                    "Detalhe": str(detalhe),
                }
            ]
        )

        if os.path.exists(ARQUIVO_LOGS):
            df_logs = pd.read_excel(ARQUIVO_LOGS, dtype=str)
            df_logs = pd.concat([df_logs, novo_log], ignore_index=True)
        else:
            df_logs = novo_log

        df_logs.to_excel(ARQUIVO_LOGS, index=False)
    except Exception as e:
        st.warning(f"Não foi possível registrar o log do sistema: {e}")


def carregar_usuarios():
    try:
        if os.path.exists(ARQUIVO_USUARIOS):
            df_u = pd.read_excel(ARQUIVO_USUARIOS, dtype=str)
            for col in ["Usuario", "Senha", "Nome", "Nivel", "Status"]:
                if col in df_u.columns:
                    df_u[col] = df_u[col].fillna("").astype(str)
            return df_u
        else:
            usuarios_default = pd.DataFrame(
                [
                    {
                        "Usuario": "mribeiro1",
                        "Senha": "123",
                        "Nome": "Marcos Ribeiro (Master)",
                        "Nivel": "Master",
                        "Status": "Ativo",
                    },
                    {
                        "Usuario": "auditor1",
                        "Senha": "123",
                        "Nome": "Auditor Operacional",
                        "Nivel": "Gestor",
                        "Status": "Ativo",
                    },
                ]
            )
            for col in usuarios_default.columns:
                usuarios_default[col] = usuarios_default[col].astype(str)
            usuarios_default.to_excel(ARQUIVO_USUARIOS, index=False)
            return usuarios_default
    except Exception as e:
        st.error(f"Erro ao carregar dados de usuários: {e}")
        return pd.DataFrame(
            columns=["Usuario", "Senha", "Nome", "Nivel", "Status"]
        )


def salvar_usuarios(df_u):
    try:
        df_salvar = df_u.copy()
        for col in df_salvar.columns:
            df_salvar[col] = df_salvar[col].astype(str)
        df_salvar.to_excel(ARQUIVO_USUARIOS, index=False)
    except Exception as e:
        st.error(f"Erro ao salvar arquivo de usuários: {e}")


# 3. BASE DE DADOS E EXTRAÇÃO DE ÚLTIMA AÇÃO
def extrair_data_ultima_acao(json_str, data_inicio_fallback):
    try:
        timeline = json.loads(str(json_str))
        if timeline and len(timeline) > 0:
            datas = [
                item.get("Data", "")
                for item in timeline
                if item.get("Data", "")
            ]
            if datas:
                return datas[-1]
    except:
        pass
    return str(data_inicio_fallback)


def carregar_dados():
    try:
        if os.path.exists(ARQUIVO_MATRIZ):
            df_base = pd.read_excel(ARQUIVO_MATRIZ, dtype=str)
        else:
            dados_iniciais = {
                "ID": ["AUD-01", "AUD-02", "AUD-03", "AUD-04", "AUD-05"],
                "Cliente_Projeto": [
                    "Vale S.A. - Governança",
                    "Banco Itaú - Bilhetes",
                    "Ambev - Reconciliação",
                    "Petrobras - Passagens",
                    "Vale S.A. - Cartões de Crédito",
                ],
                "Agencia": [
                    "São Paulo - HQ",
                    "Maringá",
                    "Rio de Janeiro",
                    "Curitiba",
                    "São Paulo - HQ",
                ],
                "Etapa_SIPOC": [
                    "Emissão & Reserva",
                    "Faturamento & Cobrança",
                    "Atendimento",
                    "Governança & TI",
                    "Emissão & Reserva",
                ],
                "Categoria": [
                    "Caixa",
                    "Crédito",
                    "Segurança",
                    "Compliance",
                    "Operacional",
                ],
                "Achado": [
                    "Divergência no fechamento físico de caixa",
                    "Falta de assinatura em contrato de cliente corporativo",
                    "Câmera de CFTV inoperante na tesouraria",
                    "Treinamento de compliance pendente",
                    "Emissão de passagens sem bilhete de autorização",
                ],
                "Severidade": ["Alta", "Crítica", "Média", "Baixa", "Crítica"],
                "Acao_Corretiva": [
                    "Realizar contagem diária e redefinir alçada.",
                    "Coletar assinatura pendente ou reter crédito.",
                    "Trocar equipamento de gravação CFTV.",
                    "Agendar treinamento para equipe.",
                    "Bloquear emissão sem prévia alçada no sistema.",
                ],
                "Area_Responsavel": [
                    "Operações",
                    "Risco/Crédito",
                    "Infraestrutura",
                    "Recursos Humanos",
                    "Operações",
                ],
                "Nome_Responsavel": [
                    "Carlos Silva",
                    "Ana Souza",
                    "João Lima",
                    "Fernanda Costa",
                    "Carlos Silva",
                ],
                "Email_Responsavel": [
                    "carlos@maringaturismo.com.br",
                    "ana@maringaturismo.com.br",
                    "joao@maringaturismo.com.br",
                    "fernanda@maringaturismo.com.br",
                    "carlos@maringaturismo.com.br",
                ],
                "Data_Inicio": [
                    "2026-08-01",
                    "2026-08-05",
                    "2026-08-10",
                    "2026-08-15",
                    "2026-08-20",
                ],
                "Prazo": [
                    "2026-09-15",
                    "2026-08-28",
                    "2026-09-10",
                    "2026-10-01",
                    "2026-09-02",
                ],
                "Data_Conclusao": ["-", "-", "2026-09-02", "-", "-"],
                "Status": [
                    "A Fazer",
                    "Em Validação",
                    "Concluído",
                    "Em Andamento",
                    "Atrasado",
                ],
                "Timeline_JSON": ["[]", "[]", "[]", "[]", "[]"],
            }
            df_base = pd.DataFrame(dados_iniciais)
            df_base.to_excel(ARQUIVO_MATRIZ, index=False)

        colunas_obrigatorias = {
            "Cliente_Projeto": "Projeto Geral",
            "Etapa_SIPOC": "Geral",
            "Status": "A Fazer",
            "Severidade": "Média",
            "Data_Inicio": str(datetime.date.today()),
            "Prazo": str(datetime.date.today()),
            "Data_Conclusao": "-",
            "Timeline_JSON": "[]",
        }
        for col, default_val in colunas_obrigatorias.items():
            if col not in df_base.columns:
                df_base[col] = default_val

        df_base["Status"] = df_base["Status"].replace({"Congos": "Concluído"})

        # Calcula Dinamicamente a Data da Última Ação Registrada
        df_base["Ultima_Acao"] = df_base.apply(
            lambda r: extrair_data_ultima_acao(
                r.get("Timeline_JSON", "[]"), r.get("Data_Inicio", "-")
            ),
            axis=1,
        )

        return df_base
    except Exception as e:
        st.error(f"Erro ao carregar matriz de auditoria: {e}")
        return pd.DataFrame()


def salvar_dados(dataframe):
    try:
        # Recalcula Última Ação antes de salvar
        if "Timeline_JSON" in dataframe.columns:
            dataframe["Ultima_Acao"] = dataframe.apply(
                lambda r: extrair_data_ultima_acao(
                    r.get("Timeline_JSON", "[]"), r.get("Data_Inicio", "-")
                ),
                axis=1,
            )
        dataframe.to_excel(ARQUIVO_MATRIZ, index=False)
    except Exception as e:
        st.error(f"Erro ao salvar alterações da matriz: {e}")


# Autenticação
if "autenticado" not in st.session_state:
    st.session_state["autenticado"] = False
if "usuario_logado" not in st.session_state:
    st.session_state["usuario_logado"] = None

if not st.session_state["autenticado"]:
    st.title("🛡️ Portal de Governança & Compliance | Login")
    tab_login, tab_cadastro = st.tabs(
        ["🔑 Acesso ao Sistema", "📝 Solicitar Novo Acesso"]
    )
    df_users = carregar_usuarios()

    with tab_login:
        with st.form("form_login"):
            usr = st.text_input("Usuário:")
            pwd = st.text_input("Senha:", type="password")
            if st.form_submit_button("Entrar no Sistema"):
                try:
                    match = df_users[
                        (df_users["Usuario"].astype(str) == str(usr).strip())
                        & (df_users["Senha"].astype(str) == str(pwd).strip())
                    ]
                    if len(match) > 0:
                        if match.iloc[0]["Status"] == "Ativo":
                            st.session_state["autenticado"] = True
                            st.session_state["usuario_logado"] = match.iloc[
                                0
                            ].to_dict()
                            registrar_log(usr, "Login", "Acesso efetuado")
                            st.rerun()
                        else:
                            st.error("Usuário aguardando aprovação Master.")
                    else:
                        st.error("Usuário ou senha incorretos.")
                except Exception as e:
                    st.error(f"Falha no processo de autenticação: {e}")

    with tab_cadastro:
        with st.form("form_solicitar_acesso"):
            novo_usr = st.text_input("Nome de Usuário:")
            novo_nome = st.text_input("Nome Completo:")
            nova_pwd = st.text_input("Senha:", type="password")
            if st.form_submit_button("Solicitar Acesso"):
                try:
                    if str(novo_usr).strip() in df_users["Usuario"].values:
                        st.warning("Usuário já existente.")
                    elif novo_usr and nova_pwd:
                        novo_row = pd.DataFrame(
                            [
                                {
                                    "Usuario": str(novo_usr).strip(),
                                    "Senha": str(nova_pwd).strip(),
                                    "Nome": str(novo_nome).strip(),
                                    "Nivel": "Gestor",
                                    "Status": "Pendente",
                                }
                            ]
                        )
                        df_users = pd.concat(
                            [df_users, novo_row], ignore_index=True
                        )
                        salvar_usuarios(df_users)
                        st.success("Solicitação enviada com sucesso!")
                    else:
                        st.warning("Preencha todos os campos obrigatórios.")
                except Exception as e:
                    st.error(f"Erro ao registrar novo usuário: {e}")

    st.stop()

# ---------------------------------------------------------
# PAINEL PRINCIPAL
# ---------------------------------------------------------
user_info = st.session_state["usuario_logado"]
is_master = user_info["Nivel"] == "Master"

st.sidebar.markdown(f"**Usuário:** {user_info['Nome']}")
st.sidebar.markdown(f"**Nível:** `{user_info['Nivel']}`")

# TROCA DE SENHA SEGURA
with st.sidebar.popover("🔑 Trocar Minha Senha"):
    st.write("### Alterar Senha")
    senha_atual = st.text_input("Senha Atual:", type="password")
    nova_senha = st.text_input("Nova Senha:", type="password")

    if st.button("Confirmar Alteração"):
        if not senha_atual or not nova_senha:
            st.warning("Por favor, preencha a senha atual e a nova senha.")
        else:
            try:
                df_u = carregar_usuarios()
                usuario_alvo = str(user_info["Usuario"]).strip()
                mask = df_u["Usuario"].astype(str).str.strip() == usuario_alvo

                if not mask.any():
                    st.error("Usuário não localizado no banco de dados.")
                else:
                    idx_u = df_u[mask].index[0]
                    senha_armazenada = str(df_u.loc[idx_u, "Senha"]).strip()

                    if senha_armazenada == str(senha_atual).strip():
                        novas_senhas = list(df_u["Senha"].astype(str))
                        novas_senhas[idx_u] = str(nova_senha).strip()

                        df_u["Senha"] = novas_senhas
                        salvar_usuarios(df_u)

                        st.session_state["usuario_logado"]["Senha"] = str(
                            nova_senha
                        ).strip()

                        registrar_log(
                            user_info["Usuario"],
                            "Troca de Senha",
                            "Senha alterada com sucesso",
                        )
                        st.success("Senha alterada com sucesso!")
                    else:
                        st.error("Senha atual incorreta.")
            except Exception as e:
                st.error(
                    f"Ocorreu um erro ao tentar atualizar a senha: {str(e)}"
                )

if st.sidebar.button("🚪 Sair"):
    st.session_state["autenticado"] = False
    st.rerun()

if "df_auditoria" not in st.session_state:
    st.session_state["df_auditoria"] = carregar_dados()

df = st.session_state["df_auditoria"]

# Colunas do Kanban Personalizáveis
if "colunas_kanban_custom" not in st.session_state:
    st.session_state["colunas_kanban_custom"] = [
        "A Fazer / Atrasado",
        "Em Andamento",
        "Em Validação (Auditor)",
        "Concluído",
    ]

# Filtro por Projeto
st.sidebar.divider()
st.sidebar.title("🔍 Filtro por Projeto")
projetos_disponiveis = ["Todos os Projetos"] + list(
    df["Cliente_Projeto"].unique()
)
projeto_selecionado = st.sidebar.selectbox(
    "📌 Selecione o Projeto / Cliente:", options=projetos_disponiveis
)

if projeto_selecionado != "Todos os Projetos":
    df_filtrado = df[df["Cliente_Projeto"] == projeto_selecionado]
else:
    df_filtrado = df.copy()

# Estruturação de Abas
abas = [
    "📊 Painel Executivo (BI & Governança)",
    "⚙️ Operações e Gestão de Projetos (CRUD)",
    "📌 Central de Projetos (Kanban Interativo)",
    "📤 Mestre Central de Anexos e Histórico",
    "📥 Extrator de Relatórios",
]

if is_master:
    abas.append("🛡️ Auditoria do Sistema e Acessos")

tabs = st.tabs(abas)


# ---------------------------------------------------------
# DIALOG NATIVO COMPACTO E AMPLO (POPUP MODAL MODERNO)
# ---------------------------------------------------------
@st.dialog("📋 Detalhes & Rastreabilidade do Apontamento", width="large")
def exibir_modal_detalhes_projeto(row_item):
    st.markdown(f"### Projeto: {row_item['Cliente_Projeto']}")
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.write(f"**ID:** #{row_item['ID']}")
    c_m2.write(f"**Responsável:** {row_item['Nome_Responsavel']}")
    c_m3.write(f"**Área:** {row_item['Area_Responsavel']}")

    st.markdown(f"**Achado Mapeado:** {row_item['Achado']}")
    st.markdown(f"**Ação Corretiva:** {row_item['Acao_Corretiva']}")
    st.info(
        f"🛫 **Data Início:** {row_item.get('Data_Inicio', '-')} | 🎯 **Prazo:** {row_item.get('Prazo', '-')} | ⏱️ **Última Ação:** {row_item.get('Ultima_Acao', '-')}"
    )

    st.divider()
    st.markdown("##### 💬 Linha do Tempo & Histórico Registrado")

    try:
        timeline = json.loads(str(row_item.get("Timeline_JSON", "[]")))
    except:
        timeline = []

    if not timeline:
        st.caption("Nenhum comentário ou evidência anexada até o momento.")
    else:
        for item in timeline:
            anexo_html = ""
            if item.get("Anexo"):
                anexo_html = f"<br>📎 <b>Anexo Vinculado:</b> <code>{item['Anexo']}</code>"

            st.markdown(
                f"""
            <div class="comment-item">
                <div class="comment-header">👤 {item['Nome']} ({item['Usuario']}) - 📅 {item['Data']}</div>
                <div>{item['Texto']} {anexo_html}</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

            if item.get("Anexo"):
                p_a = os.path.join(PASTA_EVIDENCIAS, item["Anexo"])
                if os.path.exists(p_a):
                    if item["Anexo"].lower().endswith(
                        (".png", ".jpg", ".jpeg")
                    ):
                        st.image(p_a, use_container_width=True)

    st.divider()
    st.markdown("##### ➕ Registrar Novo Comentário + Anexo")

    c_f1, c_f2 = st.columns([2, 1])
    txt_coment = c_f1.text_area(
        "Comentário sobre a evolução:", key=f"dlg_txt_{row_item['ID']}"
    )
    file_coment = c_f2.file_uploader(
        "Upload de Evidência:", key=f"dlg_file_{row_item['ID']}"
    )

    c_s1, c_s2 = st.columns([1, 1])
    if c_s1.button("💾 Salvar Histórico", key=f"dlg_save_{row_item['ID']}"):
        if txt_coment or file_coment:
            try:
                nome_anexo = None
                if file_coment is not None:
                    nome_anexo = f"{row_item['ID']}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}_{file_coment.name}"
                    caminho = os.path.join(PASTA_EVIDENCIAS, nome_anexo)
                    with open(caminho, "wb") as f:
                        f.write(file_coment.getbuffer())

                novo_item = {
                    "Data": datetime.datetime.now().strftime("%d/%m/%Y %H:%M"),
                    "Nome": user_info["Nome"],
                    "Usuario": user_info["Usuario"],
                    "Texto": txt_coment,
                    "Anexo": nome_anexo,
                }
                timeline.append(novo_item)

                idx_k = df[df["ID"] == row_item["ID"]].index[0]
                df.loc[idx_k, "Timeline_JSON"] = json.dumps(
                    timeline, ensure_ascii=False
                )
                salvar_dados(df)
                st.session_state["df_auditoria"] = df
                st.success("Histórico atualizado com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao salvar histórico: {e}")

    # ITEM 3: EXCLUSÃO DE PROJETOS DENTRO DO MODAL
    if c_s2.button(
        "🗑️ Excluir Este Projeto",
        key=f"dlg_del_{row_item['ID']}",
        type="primary",
    ):
        try:
            df_novo = df[df["ID"] != row_item["ID"]].copy()
            salvar_dados(df_novo)
            st.session_state["df_auditoria"] = df_novo
            registrar_log(
                user_info["Usuario"], "Exclusão", f"Excluiu ID {row_item['ID']}"
            )
            st.success("Projeto excluído com sucesso!")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao excluir projeto: {e}")


# ---------------------------------------------------------
# TELA 1: PAINEL EXECUTIVO COMPLETO
# ---------------------------------------------------------
with tabs[0]:
    st.markdown("### 📊 Painel Executivo de Governança, Riscos & Compliance")
    st.caption(
        "Maringá Turismo — Visão Consolidada de Riscos Operacionais e Soluções"
    )

    total_achados = len(df_filtrado)
    concluidos = len(df_filtrado[df_filtrado["Status"] == "Concluído"])
    em_validacao = len(df_filtrado[df_filtrado["Status"] == "Em Validação"])
    em_andamento = len(
        df_filtrado[df_filtrado["Status"].isin(["Em Andamento", "A Fazer"])]
    )
    atrasados = len(df_filtrado[df_filtrado["Status"] == "Atrasado"])

    taxa_resolucao = (
        round((concluidos / total_achados) * 100, 1) if total_achados > 0 else 0
    )

    penalidade_atraso = atrasados * 15
    penalidade_critica = (
        len(df_filtrado[df_filtrado["Severidade"] == "Crítica"]) * 10
    )
    score_governança = max(
        0, 100 - (penalidade_atraso + penalidade_critica)
    )

    k1, k2, k3, k4, k5 = st.columns(5)

    with k1:
        st.markdown(
            f"""<div class="kpi-card" style="border-left-color: #3b82f6;">
            <div class="kpi-label">Total Apontamentos</div>
            <div class="kpi-value">{total_achados}</div>
            <div class="kpi-subtext" style="color:#64748b;">Geral Mapeado</div>
        </div>""",
            unsafe_allow_html=True,
        )

    with k2:
        st.markdown(
            f"""<div class="kpi-card" style="border-left-color: #22c55e;">
            <div class="kpi-label">Ações Concluídas</div>
            <div class="kpi-value">{concluidos}</div>
            <div class="kpi-subtext">Taxa: {taxa_resolucao}%</div>
        </div>""",
            unsafe_allow_html=True,
        )

    with k3:
        st.markdown(
            f"""<div class="kpi-card" style="border-left-color: #f97316;">
            <div class="kpi-label">Em Validação / Auditor</div>
            <div class="kpi-value">{em_validacao}</div>
            <div class="kpi-subtext" style="color:#f97316;">Pendente Aprovação</div>
        </div>""",
            unsafe_allow_html=True,
        )

    with k4:
        st.markdown(
            f"""<div class="kpi-card" style="border-left-color: #ef4444;">
            <div class="kpi-label">Ações Atrasadas</div>
            <div class="kpi-value">{atrasados}</div>
            <div class="kpi-subtext" style="color:#ef4444;">⚠️ Atenção Imediata</div>
        </div>""",
            unsafe_allow_html=True,
        )

    with k5:
        cor_score = (
            "#22c55e"
            if score_governança >= 80
            else ("#eab308" if score_governança >= 60 else "#ef4444")
        )
        st.markdown(
            f"""<div class="kpi-card" style="border-left-color: {cor_score};">
            <div class="kpi-label">Health Score Governança</div>
            <div class="kpi-value" style="color: {cor_score};">{score_governança}%</div>
            <div class="kpi-subtext" style="color: {cor_score};">Índice de Segurança</div>
        </div>""",
            unsafe_allow_html=True,
        )

    st.divider()

    g1, g2, g3 = st.columns([1.2, 1.5, 1.3])

    with g1:
        st.markdown("##### 🍩 Distribuição por Status")
        fig_donut = go.Figure(
            data=[
                go.Pie(
                    labels=df_filtrado["Status"].value_counts().index,
                    values=df_filtrado["Status"].value_counts().values,
                    hole=0.55,
                    marker=dict(
                        colors=[
                            "#22c55e",
                            "#3b82f6",
                            "#f97316",
                            "#ef4444",
                            "#64748b",
                        ]
                    ),
                    textinfo="label+percent",
                )
            ]
        )
        fig_donut.update_layout(
            showlegend=False,
            margin=dict(t=20, b=20, l=10, r=10),
            height=260,
            annotations=[
                dict(
                    text=f"<b>{total_achados}</b><br>Ações",
                    x=0.5,
                    y=0.5,
                    font_size=16,
                    showarrow=False,
                )
            ],
        )
        st.plotly_chart(fig_donut, use_container_width=True)

    with g2:
        st.markdown("##### 📊 Matriz de Riscos por Área e Severidade")
        fig_bar_sev = px.bar(
            df_filtrado,
            x="Area_Responsavel",
            color="Severidade",
            color_discrete_map={
                "Crítica": "#ef4444",
                "Alta": "#f97316",
                "Média": "#eab308",
                "Baixa": "#22c55e",
            },
            barmode="stack",
            labels={
                "Area_Responsavel": "Área Responsável",
                "count": "Qtd Apontamentos",
            },
        )
        fig_bar_sev.update_layout(
            margin=dict(t=20, b=20, l=10, r=10),
            height=260,
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
            ),
        )
        st.plotly_chart(fig_bar_sev, use_container_width=True)

    with g3:
        st.markdown("##### 🏢 Concentração de Riscos por Unidade")
        agencia_counts = df_filtrado["Agencia"].value_counts().reset_index()
        agencia_counts.columns = ["Agencia", "Quantidade"]

        fig_agencia = px.bar(
            agencia_counts,
            y="Agencia",
            x="Quantidade",
            orientation="h",
            color_discrete_sequence=["#0284c7"],
            text="Quantidade",
        )
        fig_agencia.update_layout(
            margin=dict(t=20, b=20, l=10, r=10),
            height=260,
            xaxis_title="Apontamentos",
            yaxis_title="",
        )
        st.plotly_chart(fig_agencia, use_container_width=True)

    st.divider()

    # ITEM 6: MATRIZ COM DATA DA ÚLTIMA AÇÃO E NAVEGAÇÃO RÁPIDA POR CLIQUE
    st.markdown(
        "### ⚠️ Matriz de Riscos Críticos e Ações em Atraso (Clique para Detalhar)"
    )
    st.caption(
        "💡 Selecione uma linha da tabela para abrir os detalhes completos e o histórico do projeto na hora."
    )

    cols_exibir_matriz = [
        "ID",
        "Cliente_Projeto",
        "Achado",
        "Severidade",
        "Area_Responsavel",
        "Nome_Responsavel",
        "Prazo",
        "Ultima_Acao",
        "Status",
    ]
    df_urgente = df_filtrado[
        (df_filtrado["Severidade"].isin(["Crítica", "Alta"]))
        | (df_filtrado["Status"] == "Atrasado")
    ][cols_exibir_matriz]

    if len(df_urgente) == 0:
        st.success("🎉 Nenhuma ação urgente ou atrasada no momento!")
    else:
        # Tabela com Seleção Interativa por Clique
        event = st.dataframe(
            df_urgente,
            use_container_width=True,
            selection_mode="single-row",
            on_select="rerun",
        )

        # Se o usuário clicar em uma linha, abre o modal na hora
        if event and event.selection and event.selection.rows:
            row_idx = event.selection.rows[0]
            item_selecionado = df_urgente.iloc[row_idx]
            row_completa = df[df["ID"] == item_selecionado["ID"]].iloc[0]
            exibir_modal_detalhes_projeto(row_completa)

# ---------------------------------------------------------
# TELA 2: OPERAÇÕES E GESTÃO DE PROJETOS (CRUD + EXCLUSÃO)
# ---------------------------------------------------------
with tabs[1]:
    st.markdown("### ⚙️ Gestão de Projetos e Apontamentos")

    sub_t1, sub_t2, sub_t3 = st.tabs(
        [
            "➕ Incluir Novo Projeto / Achado",
            "✏️ Editar Projeto Existente",
            "🗑️ Excluir Projeto",
        ]
    )

    with sub_t1:
        with st.form("form_inc_proj"):
            c_i1, c_i2 = st.columns(2)
            inc_proj = c_i1.text_input(
                "Nome do Cliente / Projeto:", value="Cliente Corporativo"
            )
            inc_ag = c_i2.text_input("Unidade / Agência:", value="São Paulo - HQ")

            inc_achado = st.text_area("Descrição do Achado / Desvio:")
            inc_acao = st.text_area("Ação Corretiva Recomendada:")

            c_i3, c_i4, c_i5 = st.columns(3)
            inc_sev = c_i3.selectbox(
                "Severidade:", ["Baixa", "Média", "Alta", "Crítica"]
            )
            inc_area = c_i4.text_input("Área Responsável:", value="Operações")
            inc_resp = c_i5.text_input("Nome do Responsável:")

            c_i6, c_i7, c_i8 = st.columns(3)
            inc_email = c_i6.text_input("E-mail do Responsável:")
            inc_dt_inicio = c_i7.date_input("Data de Início:")
            inc_prazo = c_i8.date_input("Prazo Limite:")

            if st.form_submit_button("➕ Criar Registro"):
                try:
                    novo_id = f"AUD-{len(df) + 1:02d}"
                    novo_registro = pd.DataFrame(
                        [
                            {
                                "ID": novo_id,
                                "Cliente_Projeto": inc_proj,
                                "Agencia": inc_ag,
                                "Etapa_SIPOC": "Geral",
                                "Categoria": "Compliance",
                                "Achado": inc_achado,
                                "Severidade": inc_sev,
                                "Acao_Corretiva": inc_acao,
                                "Area_Responsavel": inc_area,
                                "Nome_Responsavel": inc_resp,
                                "Email_Responsavel": inc_email,
                                "Data_Inicio": str(inc_dt_inicio),
                                "Prazo": str(inc_prazo),
                                "Data_Conclusao": "-",
                                "Status": "A Fazer",
                                "Timeline_JSON": "[]",
                                "Ultima_Acao": str(inc_dt_inicio),
                            }
                        ]
                    )
                    df_novo = pd.concat([df, novo_registro], ignore_index=True)
                    salvar_dados(df_novo)
                    st.session_state["df_auditoria"] = df_novo
                    registrar_log(
                        user_info["Usuario"], "Inclusão", f"Criou {novo_id}"
                    )
                    st.success(f"Apontamento {novo_id} criado!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao incluir novo projeto: {e}")

    with sub_t2:
        if not df.empty:
            proj_sel_ed = st.selectbox(
                "Selecione o Projeto / Cliente:",
                options=df["Cliente_Projeto"].unique(),
                key="sb_ed_proj",
            )
            df_sub_ed = df[df["Cliente_Projeto"] == proj_sel_ed]

            id_ed = st.selectbox(
                "Apontamento Vinculado:", options=df_sub_ed["ID"].unique()
            )
            row_ed = df[df["ID"] == id_ed].iloc[0]

            with st.form("form_ed_proj"):
                ed_proj = st.text_input(
                    "Nome do Cliente / Projeto:", value=row_ed["Cliente_Projeto"]
                )
                ed_achado = st.text_area("Achado:", value=row_ed["Achado"])
                ed_acao = st.text_area("Ação:", value=row_ed["Acao_Corretiva"])

                c_e1, c_e2, c_e3 = st.columns(3)
                ed_resp = c_e1.text_input(
                    "Responsável:", value=row_ed["Nome_Responsavel"]
                )
                ed_dt_inicio = c_e2.text_input(
                    "Data Início:", value=str(row_ed.get("Data_Inicio", "-"))
                )
                ed_prazo = c_e3.text_input(
                    "Novo Prazo:", value=str(row_ed["Prazo"])
                )

                ed_status = st.selectbox(
                    "Status:",
                    [
                        "A Fazer",
                        "Em Andamento",
                        "Em Validação",
                        "Concluído",
                        "Atrasado",
                    ],
                    index=[
                        "A Fazer",
                        "Em Andamento",
                        "Em Validação",
                        "Concluído",
                        "Atrasado",
                    ].index(
                        row_ed["Status"]
                        if row_ed["Status"]
                        in [
                            "A Fazer",
                            "Em Andamento",
                            "Em Validação",
                            "Concluído",
                            "Atrasado",
                        ]
                        else "A Fazer"
                    ),
                )

                if st.form_submit_button("💾 Salvar Alterações"):
                    try:
                        idx = df[df["ID"] == id_ed].index[0]
                        df.loc[idx, "Cliente_Projeto"] = ed_proj
                        df.loc[idx, "Achado"] = ed_achado
                        df.loc[idx, "Acao_Corretiva"] = ed_acao
                        df.loc[idx, "Nome_Responsavel"] = ed_resp
                        df.loc[idx, "Data_Inicio"] = ed_dt_inicio
                        df.loc[idx, "Prazo"] = ed_prazo
                        df.loc[idx, "Status"] = ed_status
                        if ed_status == "Concluído":
                            df.loc[idx, "Data_Conclusao"] = str(
                                datetime.date.today()
                            )

                        salvar_dados(df)
                        st.session_state["df_auditoria"] = df
                        st.success("Projeto atualizado!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar edição: {e}")

    # ITEM 3: ABA DEDICADA DE EXCLUSÃO DE PROJETOS
    with sub_t3:
        if not df.empty:
            st.warning("⚠️ Atenção: A exclusão de um projeto é irreversível.")
            proj_sel_del = st.selectbox(
                "Selecione o Projeto para Excluir:",
                options=df["Cliente_Projeto"].unique(),
                key="sb_del_proj",
            )
            df_sub_del = df[df["Cliente_Projeto"] == proj_sel_del]
            id_del = st.selectbox(
                "ID do Apontamento:",
                options=df_sub_del["ID"].unique(),
                key="sb_del_id",
            )

            if st.button("🗑️ Confirmar Exclusão do Apontamento", type="primary"):
                try:
                    df_novo = df[df["ID"] != id_del].copy()
                    salvar_dados(df_novo)
                    st.session_state["df_auditoria"] = df_novo
                    registrar_log(
                        user_info["Usuario"], "Exclusão", f"Excluiu ID {id_del}"
                    )
                    st.success(f"Apontamento {id_del} excluído com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao excluir o projeto: {e}")

# ---------------------------------------------------------
# TELA 3: CENTRAL DE PROJETOS (KANBAN COM POSICIONAMENTO E SEPARAÇÃO)
# ---------------------------------------------------------
with tabs[2]:
    st.markdown("### 📌 Quadro Visual de Projetos & Ações")

    # ITEM 1: LIBERDADE DE ESCOLHER A POSIÇÃO DAS COLUNAS
    with st.expander("🛠️ Personalizar e Reordenar Colunas do Kanban"):
        c_k1, c_k2, c_k3 = st.columns([1.5, 1, 1.2])

        nova_col_k = c_k1.text_input("Nome da Nova Coluna:")
        posicao_k = c_k2.number_input(
            "Posição (1 a N):",
            min_value=1,
            max_value=len(st.session_state["colunas_kanban_custom"]) + 1,
            value=len(st.session_state["colunas_kanban_custom"]) + 1,
        )

        if c_k3.button("➕ Criar na Posição Escolhida"):
            if (
                nova_col_k
                and nova_col_k not in st.session_state["colunas_kanban_custom"]
            ):
                idx_pos = int(posicao_k) - 1
                st.session_state["colunas_kanban_custom"].insert(
                    idx_pos, nova_col_k
                )
                st.rerun()

        st.divider()
        c_r1, c_r2 = st.columns(2)
        col_del_k = c_r1.selectbox(
            "Excluir Coluna:",
            options=st.session_state["colunas_kanban_custom"],
            key="sb_del_col",
        )
        if c_r2.button("❌ Remover Coluna"):
            if len(st.session_state["colunas_kanban_custom"]) > 1:
                st.session_state["colunas_kanban_custom"].remove(col_del_k)
                st.rerun()

    cols_st = st.columns(len(st.session_state["colunas_kanban_custom"]))

    def status_por_coluna(col_nome):
        if "A Fazer" in col_nome:
            return ["A Fazer", "Atrasado"]
        elif "Andamento" in col_nome:
            return ["Em Andamento"]
        elif "Validação" in col_nome:
            return ["Em Validação"]
        elif "Concluído" in col_nome:
            return ["Concluído"]
        else:
            return [col_nome]

    # ITEM 2: SEPARAÇÃO NITIDA E ESTILIZADA ENTRE OS CARDS
    for index, col_nome in enumerate(st.session_state["colunas_kanban_custom"]):
        with cols_st[index]:
            st.markdown(f"#### 📌 {col_nome}")
            st_alvo = status_por_coluna(col_nome)
            itens = df_filtrado[df_filtrado["Status"].isin(st_alvo)]

            for _, row in itens.iterrows():
                sev_class = f"card-{str(row['Severidade']).lower()}"

                # Card com Estilo Separador Refinado
                st.markdown(
                    f"""
                <div class="kanban-box {sev_class}">
                    <div style="font-size: 0.75rem; color: #0284c7; font-weight: bold; text-transform: uppercase; border-bottom: 1px solid #f1f5f9; padding-bottom: 4px; margin-bottom: 6px;">
                        🏢 {row['Cliente_Projeto']}
                    </div>
                    <div style="font-size: 0.95rem; font-weight: bold; color: #0f172a; margin: 4px 0;">{row['Achado']}</div>
                    <div style="font-size: 0.82rem; color: #475569; margin-bottom: 8px; line-height: 1.3;"><b>Ação:</b> {row['Acao_Corretiva']}</div>
                    <div style="font-size: 0.72rem; color: #64748b; margin-bottom: 8px;">
                        🛫 <b>Início:</b> {row.get('Data_Inicio', '-')} | 🎯 <b>Prazo:</b> {row.get('Prazo', '-')}
                    </div>
                    <div>
                        <span class="badge badge-{str(row['Severidade']).lower()}">{row['Severidade']}</span>
                        <span style="font-size: 0.75rem; color: #94a3b8; float: right; font-weight: bold;">#{row['ID']}</span>
                    </div>
                </div>
                """,
                    unsafe_allow_html=True,
                )

                c_btn1, c_btn2 = st.columns([1.2, 1])

                with c_btn1:
                    st_mudar = st.selectbox(
                        "Mover:",
                        st.session_state["colunas_kanban_custom"],
                        key=f"sb_st_c_{row['ID']}",
                        label_visibility="collapsed",
                    )
                    if st.button("🚀 Mover", key=f"btn_mv_c_{row['ID']}"):
                        try:
                            idx_k = df[df["ID"] == row["ID"]].index[0]
                            st_logico = "Concluído" if "Concluído" in st_mudar else (
                                "Em Validação" if "Validação" in st_mudar else (
                                    "Em Andamento" if "Andamento" in st_mudar else "A Fazer"
                                )
                            )
                            df.loc[idx_k, "Status"] = st_logico
                            salvar_dados(df)
                            st.session_state["df_auditoria"] = df
                            st.rerun()
                        except Exception as e:
                            st.error(f"Erro ao mover o cartão: {e}")

                with c_btn2:
                    # ITEM 3: CHAMADA DO NOVO MODAL / POPUP AMPLO
                    if st.button("🔍 Detalhes", key=f"btn_pop_{row['ID']}"):
                        exibir_modal_detalhes_projeto(row)

# ---------------------------------------------------------
# TELA 4: MESTRE CENTRAL DE ANEXOS E HISTÓRICO
# ---------------------------------------------------------
with tabs[3]:
    st.markdown("### 📤 Central Mestre de Anexos & Histórico Auditável")

    if not df.empty:
        proj_sel_h = st.selectbox(
            "📌 Selecione o Projeto / Cliente:",
            options=df["Cliente_Projeto"].unique(),
            key="sb_proj_h",
        )
        df_sub_h = df[df["Cliente_Projeto"] == proj_sel_h]

        id_h = st.selectbox(
            "Selecione o Apontamento:",
            options=df_sub_h["ID"].unique(),
            key="sb_id_h",
        )
        row_h = df[df["ID"] == id_h].iloc[0]

        st.info(
            f"**Projeto:** {row_h['Cliente_Projeto']} | **Achado:** {row_h['Achado']} | **Status:** {row_h['Status']}"
        )

        try:
            timeline_mestre = json.loads(str(row_h.get("Timeline_JSON", "[]")))
        except:
            timeline_mestre = []

        if len(timeline_mestre) == 0:
            st.info("Nenhum histórico registrado para este projeto.")
        else:
            for t_item in timeline_mestre:
                st.markdown(
                    f"""
                <div class="comment-item">
                    <div class="comment-header">👤 {t_item['Nome']} ({t_item['Usuario']}) - 📅 {t_item['Data']}</div>
                    <div>{t_item['Texto']}</div>
                </div>
                """,
                    unsafe_allow_html=True,
                )
                if t_item.get("Anexo"):
                    p_m = os.path.join(PASTA_EVIDENCIAS, t_item["Anexo"])
                    if os.path.exists(p_m):
                        st.write(f"📎 **Anexo Vinculado:** `{t_item['Anexo']}`")
                        if t_item["Anexo"].lower().endswith(
                            (".png", ".jpg", ".jpeg")
                        ):
                            st.image(p_m, use_container_width=True)

# ---------------------------------------------------------
# TELA 5: EXTRATOR DE RELATÓRIOS (FILTROS DE PERÍODO & VISÃO CASCATA DRE)
# ---------------------------------------------------------
with tabs[4]:
    st.markdown("### 📥 Extrator Inteligente de Relatórios Executivos")
    st.caption(
        "Filtre os dados por emissão ou conclusão e visualize a cascata auditável completa de apontamentos e ações."
    )

    # ITENS 4 e 5: FILTROS DE PERÍODO (EMISSÃO E CONCLUSÃO) + PROJETO
    f_c1, f_c2, f_c3 = st.columns(3)

    proj_filtro_rel = f_c1.selectbox(
        "Filtrar por Projeto:",
        options=["Todos os Projetos"] + list(df["Cliente_Projeto"].unique()),
        key="f_rel_proj",
    )

    range_emissao = f_c2.date_input(
        "Período de Emissão (Início):",
        value=(datetime.date(2026, 1, 1), datetime.date(2026, 12, 31)),
        key="f_rel_emissao",
    )

    range_conclusao = f_c3.date_input(
        "Período de Conclusão:",
        value=(datetime.date(2026, 1, 1), datetime.date(2026, 12, 31)),
        key="f_rel_conclusao",
    )

    # Lógica de Filtragem dos Dados
    df_rel = df.copy()

    if proj_filtro_rel != "Todos os Projetos":
        df_rel = df_rel[df_rel["Cliente_Projeto"] == proj_filtro_rel]

    # Filtro Data Emissão
    if isinstance(range_emissao, tuple) and len(range_emissao) == 2:
        dt_ini_e, dt_fim_e = range_emissao
        df_rel["dt_tmp_emissao"] = pd.to_datetime(
            df_rel["Data_Inicio"], errors="coerce"
        ).dt.date
        df_rel = df_rel[
            (df_rel["dt_tmp_emissao"] >= dt_ini_e)
            & (df_rel["dt_tmp_emissao"] <= dt_fim_e)
        ]

    st.divider()

    # DOWNLOAD FORMATADO CSV/EXCEL
    try:
        csv_data = df_rel.to_csv(index=False, sep=";").encode("utf-8-sig")
        st.download_button(
            label="📥 Download Planilha Completa para Excel (CSV)",
            data=csv_data,
            file_name=f"Relatorio_Executivo_Compliance_{datetime.date.today()}.csv",
            mime="text/csv",
        )
    except Exception as e:
        st.error(f"Erro ao gerar download do relatório: {e}")

    st.divider()

    # ITEM 5: VISÃO EM CASCATA TIPO DRE (ESTRUTURADA E AUDITÁVEL)
    st.markdown("### 📊 Visão Auditável em Cascata (DRE de Governança)")

    if df_rel.empty:
        st.info("Nenhum registro encontrado para os filtros selecionados.")
    else:
        for idx, row_r in df_rel.iterrows():
            with st.container():
                st.markdown(
                    f"""
                <div class="dre-card">
                    <div class="dre-header">
                        📌 {row_r['Cliente_Projeto']} — ID: #{row_r['ID']} 
                        <span style="float: right; font-size: 0.85rem; color: #64748b;">Status: <b>{row_r['Status']}</b></span>
                    </div>
                    <div style="font-size: 0.9rem; margin-bottom: 6px;"><b>Achado:</b> {row_r['Achado']}</div>
                    <div style="font-size: 0.9rem; margin-bottom: 6px;"><b>Ação Recomendada:</b> {row_r['Acao_Corretiva']}</div>
                    <div style="font-size: 0.82rem; color: #475569; margin-bottom: 10px;">
                        👤 <b>Responsável:</b> {row_r['Nome_Responsavel']} | 🏢 <b>Área:</b> {row_r['Area_Responsavel']} | 🎯 <b>Prazo:</b> {row_r['Prazo']} | ⏱️ <b>Última Ação:</b> {row_r.get('Ultima_Acao', '-')}
                    </div>
                """,
                    unsafe_allow_html=True,
                )

                # Histórico em Cascata
                try:
                    hist_items = json.loads(
                        str(row_r.get("Timeline_JSON", "[]"))
                    )
                except:
                    hist_items = []

                if hist_items:
                    st.markdown("**📜 Histórico de Ações & Evidências Registradas:**")
                    for h in hist_items:
                        anx_str = (
                            f" | 📎 Anexo: `{h['Anexo']}`"
                            if h.get("Anexo")
                            else ""
                        )
                        st.markdown(
                            f"   * ➔ **[{h['Data']}] {h['Nome']}:** {h['Texto']}{anx_str}"
                        )
                else:
                    st.caption("   * ➔ Nenhum histórico registrado até o momento.")

                st.markdown("</div>", unsafe_allow_html=True)

# ---------------------------------------------------------
# TELA 6: AUDITORIA MASTER
# ---------------------------------------------------------
if is_master:
    with tabs[5]:
        st.markdown("### 🛡️ Administração de Acessos & Logs do Sistema")
        t_acessos, t_logs = st.tabs(
            ["👥 Gestão de Acessos", "📜 Trilha de Auditoria (Logs)"]
        )
        df_u = carregar_usuarios()

        with t_acessos:
            st.subheader("Aprovação e Níveis de Usuários")

            df_u_exibicao = df_u.copy()
            if "Senha" in df_u_exibicao.columns:
                df_u_exibicao["Senha"] = "••••••••"

            st.dataframe(df_u_exibicao, use_container_width=True)

            if not df_u.empty:
                usr_aprovar = st.selectbox(
                    "Selecione Usuário para Editar Permissões / Status:",
                    options=df_u["Usuario"].unique(),
                )
                c_st, c_nv = st.columns(2)
                novo_st_u = c_st.selectbox(
                    "Status:", ["Ativo", "Pendente", "Bloqueado"]
                )
                novo_nv_u = c_nv.selectbox("Nível:", ["Gestor", "Master"])

                if st.button("Atualizar Permissões"):
                    try:
                        idx_u = df_u[df_u["Usuario"] == usr_aprovar].index[0]
                        df_u.loc[idx_u, "Status"] = novo_st_u
                        df_u.loc[idx_u, "Nivel"] = novo_nv_u
                        salvar_usuarios(df_u)
                        registrar_log(
                            user_info["Usuario"],
                            "Gestão Acessos",
                            f"Alterou {usr_aprovar} para {novo_st_u}/{novo_nv_u}",
                        )
                        st.success("Permissões atualizadas!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao atualizar permissões: {e}")

        with t_logs:
            st.subheader("Trilha de Auditoria do Sistema")
            if os.path.exists(ARQUIVO_LOGS):
                try:
                    df_l = pd.read_excel(ARQUIVO_LOGS, dtype=str)
                    st.dataframe(df_l, use_container_width=True)
                except Exception as e:
                    st.error(f"Erro ao carregar logs: {e}")