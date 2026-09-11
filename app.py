import datetime
import json
import os
import io
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

    /* Estilização Refinada dos Cartões Kanban */
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

    /* Timeline de Comentários */
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


# 2. HELPER FUNCTIONS: CONVERSÃO DE DATAS BRASIL
def formatar_data_br(val_data):
    if not val_data or str(val_data).strip() in ["-", "", "nan", "None"]:
        return "-"
    try:
        dt = pd.to_datetime(str(val_data).split(" ")[0])
        return dt.strftime("%d/%m/%Y")
    except:
        return str(val_data)


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
    return formatar_data_br(data_inicio_fallback)


def registrar_log(usuario, acao, detalhe):
    try:
        data_hora = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
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

        # Recalcula Última Ação
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


# HELPER FUNCTION PARA RENDERIZAR ANEXOS E BOTÃO DE DOWNLOAD
def renderizar_anexo_elemento(nome_anexo, key_prefix):
    if not nome_anexo or str(nome_anexo).strip() in ["None", "null", ""]:
        return

    caminho = os.path.join(PASTA_EVIDENCIAS, str(nome_anexo))
    if os.path.exists(caminho):
        st.markdown(f"📎 **Anexo Registrado:** `{nome_anexo}`")

        # Se for imagem, exibe inline
        if str(nome_anexo).lower().endswith((".png", ".jpg", ".jpeg")):
            st.image(caminho, use_container_width=True)

        # Para TODOS os arquivos (PDF, imagens, docs), disponibiliza o botão de Download
        with open(caminho, "rb") as file_data:
            st.download_button(
                label=f"📥 Download Anexo ({nome_anexo.split('.')[-1].upper()})",
                data=file_data,
                file_name=nome_anexo,
                mime="application/octet-stream",
                key=f"btn_dl_{key_prefix}_{nome_anexo}",
            )
    else:
        st.caption(f"📎 Anexo vinculado (`{nome_anexo}`), mas arquivo não localizado no servidor.")


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

# TROCA DE SENHA
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

# Colunas do Kanban
if "colunas_kanban_custom" not in st.session_state:
    st.session_state["colunas_kanban_custom"] = [
        "A Fazer / Atrasado",
        "Em Andamento",
        "Em Validação (Auditor)",
        "Concluído",
    ]

# ---------------------------------------------------------
# FILTROS PRINCIPAIS NA SIDEBAR (DATA DE/ATÉ SIMPLIFICADA)
# ---------------------------------------------------------
st.sidebar.divider()
st.sidebar.title("🔍 Filtros Gerais")

projetos_disponiveis = ["Todos os Projetos"] + list(
    df["Cliente_Projeto"].unique()
)
projeto_selecionado = st.sidebar.selectbox(
    "📌 Projeto / Cliente:", options=projetos_disponiveis
)

# FILTRO 2: FILTRO DE DATA SIMPLIFICADO DE / ATÉ
st.sidebar.write("📅 **Período de Emissão (Data Inicial):**")
c_sb1, c_sb2 = st.sidebar.columns(2)
dt_ini_sb = c_sb1.date_input(
    "De:", value=datetime.date(2026, 1, 1), key="sb_dt_de", format="DD/MM/YYYY"
)
dt_fim_sb = c_sb2.date_input(
    "Até:", value=datetime.date(2026, 12, 31), key="sb_dt_ate", format="DD/MM/YYYY"
)

df_filtrado = df.copy()

if projeto_selecionado != "Todos os Projetos":
    df_filtrado = df_filtrado[
        df_filtrado["Cliente_Projeto"] == projeto_selecionado
    ]

if dt_ini_sb and dt_fim_sb:
    df_filtrado["dt_tmp_inicio"] = pd.to_datetime(
        df_filtrado["Data_Inicio"], errors="coerce"
    ).dt.date
    df_filtrado = df_filtrado[
        (df_filtrado["dt_tmp_inicio"] >= dt_ini_sb)
        & (df_filtrado["dt_tmp_inicio"] <= dt_fim_sb)
    ]

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
# MODAL / DIALOG DE DETALHES COM VISUALIZAÇÃO/DOWNLOAD DE ANEXOS
# ---------------------------------------------------------
@st.dialog("📋 Detalhes & Rastreabilidade do Apontamento", width="large")
def modal_detalhes_dialog(id_projeto):
    match_row = df[df["ID"] == id_projeto]
    if match_row.empty:
        st.error("Projeto não encontrado.")
        return

    row_item = match_row.iloc[0]

    st.markdown(f"### Projeto: {row_item['Cliente_Projeto']}")
    c_m1, c_m2, c_m3 = st.columns(3)
    c_m1.write(f"**ID:** #{row_item['ID']}")
    c_m2.write(f"**Responsável:** {row_item['Nome_Responsavel']}")
    c_m3.write(f"**Área:** {row_item['Area_Responsavel']}")

    st.markdown(f"**Achado Mapeado:** {row_item['Achado']}")
    st.markdown(f"**Ação Corretiva:** {row_item['Acao_Corretiva']}")
    st.info(
        f"🛫 **Data Início:** {formatar_data_br(row_item.get('Data_Inicio'))} | 🎯 **Prazo:** {formatar_data_br(row_item.get('Prazo'))} | ⏱️ **Última Ação:** {row_item.get('Ultima_Acao', '-')}"
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
        for idx_t, item in enumerate(timeline):
            st.markdown(
                f"""
            <div class="comment-item">
                <div class="comment-header">👤 {item['Nome']} ({item['Usuario']}) - 📅 {item['Data']}</div>
                <div>{item['Texto']}</div>
            </div>
            """,
                unsafe_allow_html=True,
            )

            # FILTRO 1: EXIBIÇÃO E DOWNLOAD DE ANEXOS DIVERSOS
            if item.get("Anexo"):
                renderizar_anexo_elemento(item["Anexo"], key_prefix=f"mdl_{row_item['ID']}_{idx_t}")

    st.divider()
    st.markdown("##### ➕ Registrar Novo Comentário + Anexo")

    c_f1, c_f2 = st.columns([2, 1])
    txt_coment = c_f1.text_area(
        "Comentário sobre a evolução:", key=f"dlg_txt_{row_item['ID']}"
    )
    file_coment = c_f2.file_uploader(
        "Upload de Evidência:", key=f"dlg_file_{row_item['ID']}"
    )

    if st.button("💾 Salvar Histórico", key=f"dlg_save_{row_item['ID']}"):
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

    st.divider()
    with st.expander("⚠️ Área de Risco: Excluir Apontamento"):
        st.write("Para excluir permanentemente este projeto, confirme abaixo:")
        chk_del = st.checkbox(
            "Eu compreendo e desejo excluir permanentemente este registro.",
            key=f"chk_del_dlg_{row_item['ID']}",
        )
        if st.button(
            "🔥 CONFIRMAR EXCLUSÃO DEFINITIVA",
            key=f"btn_confirm_del_dlg_{row_item['ID']}",
            type="primary",
            disabled=not chk_del,
        ):
            try:
                df_novo = df[df["ID"] != row_item["ID"]].copy()
                salvar_dados(df_novo)
                st.session_state["df_auditoria"] = df_novo
                st.session_state["id_modal_aberto"] = None
                registrar_log(
                    user_info["Usuario"],
                    "Exclusão",
                    f"Excluiu ID {row_item['ID']}",
                )
                st.success("Projeto excluído com sucesso!")
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao excluir o projeto: {e}")


if (
    "id_modal_aberto" in st.session_state
    and st.session_state["id_modal_aberto"]
):
    modal_detalhes_dialog(st.session_state["id_modal_aberto"])

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
        if not df_filtrado.empty:
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
        else:
            st.info("Sem dados para o filtro aplicado.")

    with g2:
        st.markdown("##### 📊 Matriz de Riscos por Área e Severidade")
        if not df_filtrado.empty:
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
                    orientation="h",
                    yanchor="bottom",
                    y=1.02,
                    xanchor="right",
                    x=1,
                ),
            )
            st.plotly_chart(fig_bar_sev, use_container_width=True)

    with g3:
        st.markdown("##### 🏢 Concentração de Riscos por Unidade")
        if not df_filtrado.empty:
            agencia_counts = (
                df_filtrado["Agencia"].value_counts().reset_index()
            )
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

    st.markdown("### 📋 Matriz Geral de Apontamentos (Todos os Registros)")
    st.caption(
        "💡 Selecione uma linha da tabela para abrir os detalhes completos e o histórico no modal."
    )

    df_matriz_exibicao = df_filtrado.copy()
    if not df_matriz_exibicao.empty:
        # FORMATANDO DATAS PARA O PADRÃO BRASILEIRO NA MATRIZ
        df_matriz_exibicao["Data_Inicio"] = df_matriz_exibicao["Data_Inicio"].apply(formatar_data_br)
        df_matriz_exibicao["Prazo"] = df_matriz_exibicao["Prazo"].apply(formatar_data_br)

        cols_exibir_matriz = [
            "ID",
            "Cliente_Projeto",
            "Achado",
            "Severidade",
            "Area_Responsavel",
            "Nome_Responsavel",
            "Data_Inicio",
            "Prazo",
            "Ultima_Acao",
            "Status",
        ]

        event = st.dataframe(
            df_matriz_exibicao[cols_exibir_matriz],
            use_container_width=True,
            selection_mode="single-row",
            on_select="rerun",
            key="df_matriz_interativa",
        )

        if event and event.selection and event.selection.rows:
            row_idx = event.selection.rows[0]
            item_selecionado = df_matriz_exibicao.iloc[row_idx]
            st.session_state["id_modal_aberto"] = item_selecionado["ID"]
            st.rerun()

# ---------------------------------------------------------
# TELA 2: OPERAÇÕES E GESTÃO DE PROJETOS (CRUD)
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
            inc_dt_inicio = c_i7.date_input("Data de Início:", format="DD/MM/YYYY")
            inc_prazo = c_i8.date_input("Prazo Limite:", format="DD/MM/YYYY")

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
                                "Ultima_Acao": formatar_data_br(inc_dt_inicio),
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
                    "Data Início (YYYY-MM-DD):", value=str(row_ed.get("Data_Inicio", "-"))
                )
                ed_prazo = c_e3.text_input(
                    "Novo Prazo (YYYY-MM-DD):", value=str(row_ed["Prazo"])
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

            chk_crud_del = st.checkbox(
                "Confirmo que desejo apagar permanentemente este projeto do banco de dados.",
                key="chk_crud_del",
            )

            if st.button(
                "🗑️ CONFIRMAR EXCLUSÃO DO PROJETO",
                type="primary",
                disabled=not chk_crud_del,
            ):
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
# TELA 3: CENTRAL DE PROJETOS (KANBAN)
# ---------------------------------------------------------
with tabs[2]:
    st.markdown("### 📌 Quadro Visual de Projetos & Ações")

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
        c_r1, c_r2, c_r3 = st.columns([1.5, 1, 1])
        col_del_k = c_r1.selectbox(
            "Excluir Coluna:",
            options=st.session_state["colunas_kanban_custom"],
            key="sb_del_col",
        )
        chk_col_del = c_r2.checkbox("Confirmar exclusão da coluna", key="chk_col_del")

        if c_r3.button("❌ Remover Coluna", disabled=not chk_col_del):
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

    for index, col_nome in enumerate(st.session_state["colunas_kanban_custom"]):
        with cols_st[index]:
            st.markdown(f"#### 📌 {col_nome}")
            st_alvo = status_por_coluna(col_nome)
            itens = df_filtrado[df_filtrado["Status"].isin(st_alvo)]

            for _, row in itens.iterrows():
                sev_class = f"card-{str(row['Severidade']).lower()}"

                # FORMATANDO DATAS NO PADRÃO BRASILEIRO NO CARD KANBAN
                dt_ini_br = formatar_data_br(row.get('Data_Inicio'))
                dt_prz_br = formatar_data_br(row.get('Prazo'))

                st.markdown(
                    f"""
                <div class="kanban-box {sev_class}">
                    <div style="font-size: 0.75rem; color: #0284c7; font-weight: bold; text-transform: uppercase; border-bottom: 1px solid #f1f5f9; padding-bottom: 4px; margin-bottom: 6px;">
                        🏢 {row['Cliente_Projeto']}
                    </div>
                    <div style="font-size: 0.95rem; font-weight: bold; color: #0f172a; margin: 4px 0;">{row['Achado']}</div>
                    <div style="font-size: 0.82rem; color: #475569; margin-bottom: 8px; line-height: 1.3;"><b>Ação:</b> {row['Acao_Corretiva']}</div>
                    <div style="font-size: 0.72rem; color: #64748b; margin-bottom: 8px;">
                        🛫 <b>Início:</b> {dt_ini_br} | 🎯 <b>Prazo:</b> {dt_prz_br}
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
                    if st.button("🔍 Detalhes", key=f"btn_pop_{row['ID']}"):
                        st.session_state["id_modal_aberto"] = row["ID"]
                        st.rerun()

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
            for idx_tm, t_item in enumerate(timeline_mestre):
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
                    renderizar_anexo_elemento(t_item["Anexo"], key_prefix=f"mst_{row_h['ID']}_{idx_tm}")

# ---------------------------------------------------------
# TELA 5: EXTRATOR DE RELATÓRIOS (EXCEL MULTI-ABAS CASCATA)
# ---------------------------------------------------------
with tabs[4]:
    st.markdown("### 📥 Extrator Inteligente de Relatórios Executivos")
    st.caption(
        "Filtre os dados por emissão ou conclusão e exporte o relatório em formato Excel (.xlsx) com a visão em cascata e detalhes completos."
    )

    f_c1, f_c2, f_c3, f_c4 = st.columns(4)

    proj_filtro_rel = f_c1.selectbox(
        "Filtrar por Projeto:",
        options=["Todos os Projetos"] + list(df["Cliente_Projeto"].unique()),
        key="f_rel_proj",
    )

    dt_ini_rel_e = f_c2.date_input(
        "Emissão (De):", value=datetime.date(2026, 1, 1), format="DD/MM/YYYY", key="f_rel_e_de"
    )
    dt_fim_rel_e = f_c3.date_input(
        "Emissão (Até):", value=datetime.date(2026, 12, 31), format="DD/MM/YYYY", key="f_rel_e_ate"
    )

    df_rel = df.copy()

    if proj_filtro_rel != "Todos os Projetos":
        df_rel = df_rel[df_rel["Cliente_Projeto"] == proj_filtro_rel]

    if dt_ini_rel_e and dt_fim_rel_e:
        df_rel["dt_tmp_e"] = pd.to_datetime(
            df_rel["Data_Inicio"], errors="coerce"
        ).dt.date
        df_rel = df_rel[
            (df_rel["dt_tmp_e"] >= dt_ini_rel_e)
            & (df_rel["dt_tmp_e"] <= dt_fim_rel_e)
        ]

    st.divider()

    # FILTRO 4: GERADOR DE EXCEL MULTI-ABAS (CASCATA DRE + BASE DADOS)
    def gerar_excel_relatorio_cascata(df_export):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # 1. ABA DE BASE DE DADOS COMPLETA
            df_base_exp = df_export.copy()
            df_base_exp["Data_Inicio"] = df_base_exp["Data_Inicio"].apply(formatar_data_br)
            df_base_exp["Prazo"] = df_base_exp["Prazo"].apply(formatar_data_br)
            df_base_exp["Data_Conclusao"] = df_base_exp["Data_Conclusao"].apply(formatar_data_br)

            cols_limpas = [c for c in df_base_exp.columns if not c.startswith("dt_tmp")]
            df_base_exp[cols_limpas].to_excel(writer, sheet_name="Base_Dados_Completa", index=False)

            # 2. ABA VISÃO CASCATA ESTRUTURADA
            linhas_cascata = []
            for _, r in df_export.iterrows():
                # Linha Cabeçalho do Projeto
                linhas_cascata.append({
                    "Nivel": "PROJETO",
                    "ID": r["ID"],
                    "Cliente_Projeto": r["Cliente_Projeto"],
                    "Status": r["Status"],
                    "Severidade": r["Severidade"],
                    "Achado_Ou_Comentario": r["Achado"],
                    "Acao_Ou_Usuario": r["Acao_Corretiva"],
                    "Responsavel": r["Nome_Responsavel"],
                    "Data_Inicio": formatar_data_br(r["Data_Inicio"]),
                    "Prazo": formatar_data_br(r["Prazo"]),
                    "Anexo_Evidencia": "-",
                })

                # Linhas filhas (Histórico/Timeline)
                try:
                    hist_l = json.loads(str(r.get("Timeline_JSON", "[]")))
                except:
                    hist_l = []

                for h in hist_l:
                    linhas_cascata.append({
                        "Nivel": "   └─ AÇÃO/HISTÓRICO",
                        "ID": r["ID"],
                        "Cliente_Projeto": r["Cliente_Projeto"],
                        "Status": r["Status"],
                        "Severidade": "-",
                        "Achado_Ou_Comentario": f"Texto: {h.get('Texto', '')}",
                        "Acao_Ou_Usuario": f"Por: {h.get('Nome', '')} ({h.get('Usuario', '')})",
                        "Responsavel": "-",
                        "Data_Inicio": h.get("Data", ""),
                        "Prazo": "-",
                        "Anexo_Evidencia": h.get("Anexo") if h.get("Anexo") else "Sem anexo",
                    })

            df_cascata_excel = pd.DataFrame(linhas_cascata)
            df_cascata_excel.to_excel(writer, sheet_name="Visao_Cascata_Executiva", index=False)

        output.seek(0)
        return output

    try:
        excel_bytes = gerar_excel_relatorio_cascata(df_rel)
        st.download_button(
            label="📥 Download Relatório Profissional em Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Relatorio_Executivo_Compliance_{datetime.date.today().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Erro ao gerar planilha Excel: {e}")

    st.divider()

    st.markdown("### 📊 Visão Auditável em Cascata (DRE de Governança)")

    if df_rel.empty:
        st.info("Nenhum registro encontrado para os filtros selecionados.")
    else:
        for idx_r, row_r in df_rel.iterrows():
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
                        👤 <b>Responsável:</b> {row_r['Nome_Responsavel']} | 🏢 <b>Área:</b> {row_r['Area_Responsavel']} | 🛫 <b>Início:</b> {formatar_data_br(row_r.get('Data_Inicio'))} | 🎯 <b>Prazo:</b> {formatar_data_br(row_r.get('Prazo'))} | ⏱️ <b>Última Ação:</b> {row_r.get('Ultima_Acao', '-')}
                    </div>
                """,
                    unsafe_allow_html=True,
                )

                try:
                    hist_items = json.loads(
                        str(row_r.get("Timeline_JSON", "[]"))
                    )
                except:
                    hist_items = []

                if hist_items:
                    st.markdown("**📜 Histórico de Ações & Evidências Registradas:**")
                    for idx_h, h in enumerate(hist_items):
                        st.markdown(
                            f"   * ➔ **[{h['Data']}] {h['Nome']}:** {h['Texto']}"
                        )
                        if h.get("Anexo"):
                            renderizar_anexo_elemento(h["Anexo"], key_prefix=f"casc_{row_r['ID']}_{idx_h}")
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