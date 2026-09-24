import datetime
import json
import os
import io
import re
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# FUSO HORÁRIO BRASIL (UTC-3)
TZ_BR = datetime.timezone(datetime.timedelta(hours=-3))

def get_now_br():
    return datetime.datetime.now(TZ_BR)

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

    /* Cartões Kanban */
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

    /* Badges Executivos */
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
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-left: 4px solid #0284c7;
        padding: 12px 14px;
        margin-bottom: 8px;
        border-radius: 8px;
        font-size: 0.88rem;
    }
    .comment-header {
        font-weight: 700;
        color: #0f172a;
        font-size: 0.80rem;
        margin-bottom: 4px;
    }

    /* Card de Anexo para a Mestre Central */
    .file-card {
        background-color: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 8px;
        padding: 14px;
        margin-bottom: 12px;
        border-left: 4px solid #10b981;
    }
    
    /* IA Insights Box */
    .ai-insights-box {
        background: linear-gradient(135deg, #eff6ff 0%, #e0f2fe 100%);
        border: 1px solid #bae6fd;
        border-radius: 10px;
        padding: 16px;
        margin-bottom: 15px;
    }
</style>
""",
    unsafe_allow_html=True,
)

PASTA_EVIDENCIAS = "evidencias_acumuladas"
ARQUIVO_MATRIZ = "matriz_auditoria.xlsx"
ARQUIVO_USUARIOS = "usuarios_sistema.xlsx"
ARQUIVO_LOGS = "logs_auditoria_sistema.xlsx"
ARQUIVO_CONFIG = "config_kanban.json"

if not os.path.exists(PASTA_EVIDENCIAS):
    os.makedirs(PASTA_EVIDENCIAS)


# 2. HELPER FUNCTIONS & EXCLUSÃO EM CASCATA DE ANEXOS
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


def apagar_anexos_do_apontamento(id_apontamento, timeline_json_str=None):
    """Garante a exclusão física de todos os anexos atrelados ao projeto excluído."""
    removidos = 0
    try:
        # 1. Remove pelo padrão de nome ID_...
        prefixo = f"{str(id_apontamento).strip()}_"
        if os.path.exists(PASTA_EVIDENCIAS):
            for arq in os.listdir(PASTA_EVIDENCIAS):
                if arq.startswith(prefixo):
                    caminho_completo = os.path.join(PASTA_EVIDENCIAS, arq)
                    if os.path.isfile(caminho_completo):
                        os.remove(caminho_completo)
                        removidos += 1

        # 2. Varre o JSON de timeline para garantir arquivos nomeados diferentemente
        if timeline_json_str:
            try:
                timeline = json.loads(str(timeline_json_str))
                for item in timeline:
                    nome_anexo = item.get("Anexo")
                    if nome_anexo:
                        caminho_f = os.path.join(PASTA_EVIDENCIAS, str(nome_anexo))
                        if os.path.isfile(caminho_f):
                            os.remove(caminho_f)
                            removidos += 1
            except:
                pass
    except Exception as e:
        print(f"Aviso ao remover anexos: {e}")
    return removidos


def registrar_log(usuario, acao, detalhe):
    try:
        data_hora = get_now_br().strftime("%d/%m/%Y %H:%M:%S")
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
                if col not in df_u.columns:
                    df_u[col] = ""
                df_u[col] = df_u[col].fillna("").astype(str).str.strip()
            
            mask_master = df_u["Usuario"].str.lower() == "mribeiro1"
            if not mask_master.any():
                novo_master = pd.DataFrame([{
                    "Usuario": "mribeiro1",
                    "Senha": "123",
                    "Nome": "Marcos Ribeiro (Master)",
                    "Nivel": "Master",
                    "Status": "Ativo"
                }])
                df_u = pd.concat([df_u, novo_master], ignore_index=True)
                salvar_usuarios(df_u)
            else:
                idx_m = df_u[mask_master].index[0]
                if df_u.loc[idx_m, "Status"] != "Ativo":
                    df_u.loc[idx_m, "Status"] = "Ativo"
                    salvar_usuarios(df_u)

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
            df_salvar[col] = df_salvar[col].astype(str).str.strip()
        df_salvar.to_excel(ARQUIVO_USUARIOS, index=False)
    except Exception as e:
        st.error(f"Erro ao salvar arquivo de usuários: {e}")


def carregar_colunas_kanban():
    padrao = [
        "A Fazer / Atrasado",
        "Em Andamento",
        "Em Validação (Auditor)",
        "Concluído",
    ]
    if os.path.exists(ARQUIVO_CONFIG):
        try:
            with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data.get("colunas", padrao)
        except:
            pass
    return padrao


def salvar_colunas_kanban(colunas):
    try:
        with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
            json.dump({"colunas": colunas}, f, ensure_ascii=False, indent=2)
    except Exception as e:
        st.error(f"Erro ao salvar configuração do Kanban: {e}")


def gerar_novo_id(dataframe):
    if dataframe.empty or "ID" not in dataframe.columns:
        return "AUD-01"
    
    ids_existentes = dataframe["ID"].astype(str).tolist()
    numeros = []
    for id_str in ids_existentes:
        match = re.search(r'\d+', id_str)
        if match:
            numeros.append(int(match.group()))
    
    proximo_num = max(numeros) + 1 if numeros else 1
    return f"AUD-{proximo_num:02d}"


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
                    "A Fazer / Atrasado",
                    "Em Validação (Auditor)",
                    "Concluído",
                    "Em Andamento",
                    "A Fazer / Atrasado",
                ],
                "Timeline_JSON": ["[]", "[]", "[]", "[]", "[]"],
                "Campos_Custom_JSON": ["{}", "{}", "{}", "{}", "{}"],
            }
            df_base = pd.DataFrame(dados_iniciais)
            df_base.to_excel(ARQUIVO_MATRIZ, index=False)

        colunas_obrigatorias = {
            "Cliente_Projeto": "Projeto Geral",
            "Etapa_SIPOC": "Geral",
            "Status": "A Fazer / Atrasado",
            "Severidade": "Média",
            "Data_Inicio": str(get_now_br().date()),
            "Prazo": str(get_now_br().date()),
            "Data_Conclusao": "-",
            "Timeline_JSON": "[]",
            "Campos_Custom_JSON": "{}",
        }
        for col, default_val in colunas_obrigatorias.items():
            if col not in df_base.columns:
                df_base[col] = default_val

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
        df_salvar = dataframe.copy()
        if "Timeline_JSON" in df_salvar.columns:
            df_salvar["Ultima_Acao"] = df_salvar.apply(
                lambda r: extrair_data_ultima_acao(
                    r.get("Timeline_JSON", "[]"), r.get("Data_Inicio", "-")
                ),
                axis=1,
            )
        for col in df_salvar.columns:
            df_salvar[col] = df_salvar[col].astype(str)

        df_salvar.to_excel(ARQUIVO_MATRIZ, index=False)
    except Exception as e:
        st.error(f"Erro ao salvar alterações da matriz: {e}")


# 3. GERADOR DE EXCEL EXECUTIVO ESTILIZADO (OPENPYXL ENTERPRISE)
def gerar_excel_estilizado(df_export):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Matriz_Auditoria"

    df_base_exp = df_export.copy()
    
    # Colunas Amigáveis para Exportação
    colunas_visiveis = [
        "ID", "Cliente_Projeto", "Agencia", "Categoria", "Severidade",
        "Achado", "Acao_Corretiva", "Area_Responsavel", "Nome_Responsavel",
        "Email_Responsavel", "Data_Inicio", "Prazo", "Data_Conclusao", "Status"
    ]
    
    cols_existentes = [c for c in colunas_visiveis if c in df_base_exp.columns]
    df_base_exp = df_base_exp[cols_existentes]

    # Formatar datas
    for c_date in ["Data_Inicio", "Prazo", "Data_Conclusao"]:
        if c_date in df_base_exp.columns:
            df_base_exp[c_date] = df_base_exp[c_date].apply(formatar_data_br)

    # Escrever Título do Relatório
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    title_cell.value = "RELATÓRIO EXECUTIVO DE GOVERNANÇA, RISCOS & COMPLIANCE"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 40

    # Data de Geração
    ws.merge_cells("A2:N2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Gerado em: {get_now_br().strftime('%d/%m/%Y às %H:%M:%S')} | Maringá Turismo"
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="64748B")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # Escrever Cabeçalhos das Colunas (Linha 4)
    headers = list(df_base_exp.columns)
    ws.append([]) # Linha 3 vazia
    ws.append(headers) # Linha 4

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1")
    )

    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=4, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    ws.row_dimensions[4].height = 28

    # Estilos de Cores para Status e Severidade
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")
    font_red = Font(name="Calibri", size=10, color="991B1B", bold=True)
    
    fill_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    font_green = Font(name="Calibri", size=10, color="166534", bold=True)
    
    fill_yellow = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    font_yellow = Font(name="Calibri", size=10, color="854D0E", bold=True)

    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    # Preencher Linhas de Dados
    for row_idx, row_data in enumerate(df_base_exp.values, start=5):
        ws.append(list(row_data))
        ws.row_dimensions[row_idx].height = 22
        
        is_even = (row_idx % 2 == 0)

        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            cell.alignment = Alignment(vertical="center", horizontal="left")

            if is_even:
                cell.fill = fill_zebra

            col_header = headers[col_idx - 1]

            # Destaques Condicionais
            if col_header == "Severidade":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                val_s = str(value).lower()
                if "crítica" in val_s or "critica" in val_s:
                    cell.fill, cell.font = fill_red, font_red
                elif "alta" in val_s:
                    cell.fill, cell.font = fill_yellow, font_yellow
                elif "baixa" in val_s:
                    cell.fill, cell.font = fill_green, font_green

            elif col_header == "Status":
                cell.alignment = Alignment(horizontal="center", vertical="center")
                val_st = str(value).lower()
                if "atrasado" in val_st:
                    cell.fill, cell.font = fill_red, font_red
                elif "concluído" in val_st or "concluido" in val_st:
                    cell.fill, cell.font = fill_green, font_green
                elif "andamento" in val_st or "validação" in val_st:
                    cell.fill, cell.font = fill_yellow, font_yellow

    # Ajuste Automático da Largura das Colunas
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row > 2 and cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)

    wb.save(output)
    output.seek(0)
    return output


def renderizar_anexo_elemento(nome_anexo, key_prefix):
    if not nome_anexo or str(nome_anexo).strip() in ["None", "null", ""]:
        return

    caminho = os.path.join(PASTA_EVIDENCIAS, str(nome_anexo))
    if os.path.exists(caminho):
        st.markdown(f"📎 **Anexo Registrado:** `{nome_anexo}`")

        if str(nome_anexo).lower().endswith((".png", ".jpg", ".jpeg")):
            st.image(caminho, use_container_width=True)

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


def renderizar_painel_detalhes(id_modal):
    match_row = df[df["ID"] == id_modal]
    if match_row.empty:
        return

    row_item = match_row.iloc[0]

    st.markdown("---")
    with st.container():
        c_head1, c_head2 = st.columns([6, 1])
        c_head1.markdown(
            f"## 📋 Detalhes do Apontamento #{row_item['ID']} — {row_item['Cliente_Projeto']}"
        )
        if c_head2.button("❌ FECHAR", type="primary", key="btn_close_modal_direct"):
            st.session_state["id_modal_aberto"] = None
            st.rerun()

        c_m1, c_m2, c_m3 = st.columns(3)
        c_m1.write(f"**Responsável:** {row_item['Nome_Responsavel']}")
        c_m2.write(f"**Área:** {row_item['Area_Responsavel']}")
        c_m3.write(f"**Status Atual:** `{row_item['Status']}`")

        st.markdown(f"**Achado Mapeado:** {row_item['Achado']}")
        st.markdown(f"**Ação Corretiva:** {row_item['Acao_Corretiva']}")
        st.info(
            f"🛫 **Data Início:** {formatar_data_br(row_item.get('Data_Inicio'))} | 🎯 **Prazo:** {formatar_data_br(row_item.get('Prazo'))} | ⏱️ **Última Ação:** {row_item.get('Ultima_Acao', '-')}"
        )

        try:
            campos_c = json.loads(str(row_item.get("Campos_Custom_JSON", "{}")))
            if campos_c:
                st.markdown("##### 🧩 Campos Personalizados Adicionais")
                cols_custom = st.columns(2)
                for i, (k, v) in enumerate(campos_c.items()):
                    cols_custom[i % 2].write(f"**{k}:** {v}")
        except:
            pass

        st.markdown("##### 💬 Histórico de Registros")
        try:
            timeline = json.loads(str(row_item.get("Timeline_JSON", "[]")))
        except:
            timeline = []

        if not timeline:
            st.caption("Nenhum comentário ou evidência anexada até o momento.")
        else:
            usuario_atual = user_info["Usuario"]

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

                if item.get("Usuario") == usuario_atual or is_master:
                    with st.expander(f"✏️ Editar Comentário #{idx_t + 1}"):
                        novo_txt_edit = st.text_area(
                            "Editar texto do comentário:",
                            value=item["Texto"],
                            key=f"edt_txt_{row_item['ID']}_{idx_t}"
                        )
                        if st.button("💾 Salvar Edição de Comentário", key=f"btn_edt_c_{row_item['ID']}_{idx_t}"):
                            timeline[idx_t]["Texto"] = novo_txt_edit
                            idx_k = df[df["ID"] == row_item["ID"]].index[0]
                            df.loc[idx_k, "Timeline_JSON"] = json.dumps(timeline, ensure_ascii=False)
                            salvar_dados(df)
                            st.session_state["df_auditoria"] = df
                            registrar_log(usuario_atual, "Edição Comentário", f"Editou comentário #{idx_t + 1} no {row_item['ID']}")
                            st.success("Comentário atualizado!")
                            st.rerun()

                if item.get("Anexo"):
                    renderizar_anexo_elemento(
                        item["Anexo"], key_prefix=f"mdl_{row_item['ID']}_{idx_t}"
                    )

        st.markdown("##### ➕ Registrar Novo Comentário + Anexo")
        cnt = st.session_state["form_counter"]
        c_f1, c_f2 = st.columns([2, 1])
        txt_coment = c_f1.text_area(
            "Comentário sobre a evolução:", key=f"dlg_txt_{row_item['ID']}_{cnt}"
        )
        file_coment = c_f2.file_uploader(
            "Upload de Evidência:", key=f"dlg_file_{row_item['ID']}_{cnt}"
        )

        if st.button("💾 Salvar Histórico", key=f"dlg_save_btn_{row_item['ID']}"):
            if txt_coment or file_coment:
                try:
                    nome_anexo = None
                    if file_coment is not None:
                        nome_anexo = f"{row_item['ID']}_{get_now_br().strftime('%Y%m%d_%H%M%S')}_{file_coment.name}"
                        caminho = os.path.join(PASTA_EVIDENCIAS, nome_anexo)
                        with open(caminho, "wb") as f:
                            f.write(file_coment.getbuffer())

                    novo_item = {
                        "Data": get_now_br().strftime("%d/%m/%Y %H:%M"),
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
                    st.session_state["form_counter"] += 1
                    registrar_log(user_info["Usuario"], "Novo Histórico", f"Adicionou comentário em {row_item['ID']}")
                    st.success("Histórico atualizado com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao salvar histórico: {e}")

        with st.expander("⚠️ Área de Risco: Excluir Apontamento"):
            chk_del = st.checkbox(
                "Eu compreendo e desejo excluir permanentemente este registro e todos os seus anexos físicos.",
                key=f"chk_del_dlg_{row_item['ID']}",
            )
            if st.button(
                "🔥 CONFIRMAR EXCLUSÃO DEFINITIVA",
                key=f"btn_confirm_del_dlg_{row_item['ID']}",
                type="primary",
                disabled=not chk_del,
            ):
                try:
                    # Exclusão em Cascata dos Arquivos Físicos
                    qtd_anexos_del = apagar_anexos_do_apontamento(row_item["ID"], row_item.get("Timeline_JSON"))
                    
                    df_novo = df[df["ID"] != row_item["ID"]].copy()
                    salvar_dados(df_novo)
                    st.session_state["df_auditoria"] = df_novo
                    st.session_state["id_modal_aberto"] = None
                    registrar_log(user_info["Usuario"], "Exclusão Apontamento", f"Excluiu apontamento {row_item['ID']} e {qtd_anexos_del} anexo(s)")
                    st.success(f"Projeto e {qtd_anexos_del} anexo(s) associado(s) excluídos com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao excluir o projeto: {e}")

    st.markdown("---")


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
            pwd = st.text_input("Senha (Aceita Alfanumérico):", type="password")
            if st.form_submit_button("Entrar no Sistema"):
                try:
                    match = df_users[
                        (df_users["Usuario"].astype(str).str.strip() == str(usr).strip())
                        & (df_users["Senha"].astype(str).str.strip() == str(pwd).strip())
                    ]
                    if len(match) > 0:
                        if match.iloc[0]["Status"] == "Ativo":
                            st.session_state["autenticado"] = True
                            st.session_state["usuario_logado"] = match.iloc[0].to_dict()
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
            nova_pwd = st.text_input("Senha (Aceita Letras, Números e Símbolos):", type="password")
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
    nova_senha = st.text_input("Nova Senha (Alfanumérica):", type="password")

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
                        df_u.loc[idx_u, "Senha"] = str(nova_senha).strip()
                        salvar_usuarios(df_u)

                        st.session_state["usuario_logado"]["Senha"] = str(nova_senha).strip()

                        registrar_log(
                            user_info["Usuario"],
                            "Troca de Senha",
                            "Senha alfanumérica alterada com sucesso",
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

# CARREGAMENTO PERSISTENTE DAS COLUNAS DO KANBAN
if "colunas_kanban_custom" not in st.session_state:
    st.session_state["colunas_kanban_custom"] = carregar_colunas_kanban()

if "form_counter" not in st.session_state:
    st.session_state["form_counter"] = 0

# =========================================================
# FILTROS NA SIDEBAR & ASSISTENTE PREDITIVO DE COMPLIANCE
# =========================================================

# --- 1. ASSISTENTE IA DE COMPLIANCE (DIAGNÓSTICO PREDITIVO REFORMULADO) ---
st.sidebar.markdown("### 🤖 Assistente IA de Compliance")

with st.sidebar.expander("💡 Diagnóstico Preditivo de Riscos", expanded=True):
    # Cálculos dinâmicos baseados no DataFrame
    now_br = pd.to_datetime(get_now_br().strftime("%Y-%m-%d"))
    
    # Mapeamento de colunas de severidade e prazo
    col_sev = next((c for c in ["Severidade", "Criticidade", "Prioridade"] if c in df.columns), None)
    col_prazo = next((c for c in ["Prazo", "Data_Prazo", "Data_Fim"] if c in df.columns), None)
    col_status = next((c for c in ["Status", "Estado"] if c in df.columns), None)

    # Métricas Dinâmicas
    qtd_criticos = 0
    if col_sev:
        qtd_criticos = len(df[df[col_sev].astype(str).str.contains("Alta|Crítica|Alta Severidade", case=False, na=False)])

    qtd_atrasados = 0
    if col_prazo:
        df_prazos = pd.to_datetime(df[col_prazo], errors="coerce")
        qtd_atrasados = len(df[(df_prazos < now_br) & (df[col_status] != "Concluído")]) if col_status else len(df[df_prazos < now_br])

    # Apresentação Visual Estilizada em HTML/CSS
    cor_alerta_critico = "#ef4444" if qtd_criticos > 0 else "#22c55e"
    cor_alerta_atraso = "#f59e0b" if qtd_atrasados > 0 else "#22c55e"

    st.markdown(
        f"""
        <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 12px; margin-bottom: 8px;">
            <div style="font-size: 0.85rem; font-weight: bold; color: #1e293b; margin-bottom: 8px; border-bottom: 1px solid #cbd5e1; padding-bottom: 4px;">
                📊 Análise em Tempo Real
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px;">
                <span style="font-size: 0.8rem; color: #475569;">Riscos de Alta Severidade:</span>
                <span style="background-color: {cor_alerta_critico}; color: white; padding: 2px 8px; border-radius: 12px; font-weight: bold; font-size: 0.75rem;">
                    {qtd_criticos}
                </span>
            </div>
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">
                <span style="font-size: 0.8rem; color: #475569;">Ações Fora do SLA:</span>
                <span style="background-color: {cor_alerta_atraso}; color: white; padding: 2px 8px; border-radius: 12px; font-weight: bold; font-size: 0.75rem;">
                    {qtd_atrasados}
                </span>
            </div>
            <div style="background-color: #eff6ff; border-left: 3px solid #3b82f6; padding: 6px 8px; border-radius: 4px; font-size: 0.75rem; color: #1e40af; line-height: 1.3;">
                <b>💡 Recomendação:</b> {"Priorizar tratativas em atraso e itens de severidade alta no Kanban." if (qtd_criticos > 0 or qtd_atrasados > 0) else "Operação em conformidade. Nossos indicadores estão dentro da meta!"}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# --- 2. FILTROS GERAIS NA SIDEBAR ---
st.sidebar.divider()
st.sidebar.title("🔍 Filtros Gerais")

# Filtro por Projeto / Cliente
projetos_disponiveis = ["Todos os Projetos"] + list(df["Cliente_Projeto"].dropna().unique())
projeto_selecionado = st.sidebar.selectbox(
    "📌 Projeto / Cliente:", options=projetos_disponiveis
)

# Filtro por Período de Emissão
st.sidebar.write("📅 **Período de Emissão (Data Inicial):**")
c_sb1, c_sb2 = st.sidebar.columns(2)

dt_ini_sb = c_sb1.date_input(
    "De:", value=datetime.date(2020, 1, 1), key="sb_dt_de", format="DD/MM/YYYY"
)
dt_fim_sb = c_sb2.date_input(
    "Até:", value=datetime.date(2030, 12, 31), key="sb_dt_ate", format="DD/MM/YYYY"
)

# Filtro por Criador / Responsável (Garantia de Mapeamento)
col_resp = next((c for c in ["Nome_Responsavel", "Usuario_Criador", "Criador", "Usuario", "Responsavel"] if c in df.columns), None)

usuario_selecionado = "Todos os Responsáveis"
if col_resp:
    responsaveis_disponiveis = ["Todos os Responsáveis"] + sorted(
        df[col_resp].dropna().astype(str).unique().tolist()
    )
    usuario_selecionado = st.sidebar.selectbox(
        "👤 Criador / Responsável:",
        options=responsaveis_disponiveis,
        key="sb_usuario_sel"
    )
else:
    # Se nenhuma coluna exata for encontrada, exibe informação de apoio
    st.sidebar.info("💡 Coluna de responsável não detectada no banco de dados.")

# --- 3. APLICAÇÃO DOS FILTROS NO DATAFRAME ---
df_filtrado = df.copy()

if projeto_selecionado != "Todos os Projetos":
    df_filtrado = df_filtrado[df_filtrado["Cliente_Projeto"] == projeto_selecionado]

if dt_ini_sb and dt_fim_sb:
    df_filtrado["dt_tmp_inicio"] = pd.to_datetime(
        df_filtrado["Data_Inicio"], errors="coerce"
    ).dt.date
    df_filtrado = df_filtrado[
        (df_filtrado["dt_tmp_inicio"] >= dt_ini_sb)
        & (df_filtrado["dt_tmp_inicio"] <= dt_fim_sb)
    ]

if col_resp and usuario_selecionado != "Todos os Responsáveis":
    df_filtrado = df_filtrado[df_filtrado[col_resp] == usuario_selecionado]

# --- 4. NAVEGAÇÃO DE ABAS ---
abas = [
    "📊 Painel Executivo (BI & Governança)",
    "📈 Dashboard Analítico de KPIs",
    "⚙️ Operações e Gestão (CRUD)",
    "📌 Central de Projetos (Kanban)",
    "📤 Mestre Central de Anexos",
    "📥 Extrator de Relatórios",
]

if is_master:
    abas.append("🛡️ Auditoria Master")

tabs = st.tabs(abas)

# ---------------------------------------------------------
# TELA 1: PAINEL EXECUTIVO COMPLETO
# ---------------------------------------------------------
with tabs[0]:
    st.markdown("### 📊 Painel Executivo de Governança, Riscos & Compliance")
    st.caption("Maringá Turismo — Visão Consolidada de Riscos Operacionais e Soluções")

    total_achados = len(df_filtrado)
    concluidos = len(df_filtrado[df_filtrado["Status"] == "Concluído"])
    em_validacao = len(df_filtrado[df_filtrado["Status"].str.contains("Validação", case=False, na=False)])
    atrasados = len(df_filtrado[df_filtrado["Status"].str.contains("Atrasado", case=False, na=False)])

    taxa_resolucao = round((concluidos / total_achados) * 100, 1) if total_achados > 0 else 0
    penalidade_atraso = atrasados * 15
    penalidade_critica = len(df_filtrado[df_filtrado["Severidade"] == "Crítica"]) * 10
    score_governança = max(0, 100 - (penalidade_atraso + penalidade_critica))

    k1, k2, k3, k4, k5 = st.columns(5)

    with k1:
        st.markdown(f"""<div class="kpi-card" style="border-left-color: #3b82f6;"><div class="kpi-label">Total Apontamentos</div><div class="kpi-value">{total_achados}</div><div class="kpi-subtext" style="color:#64748b;">Geral Mapeado</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card" style="border-left-color: #22c55e;"><div class="kpi-label">Ações Concluídas</div><div class="kpi-value">{concluidos}</div><div class="kpi-subtext">Taxa: {taxa_resolucao}%</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card" style="border-left-color: #f97316;"><div class="kpi-label">Em Validação / Auditor</div><div class="kpi-value">{em_validacao}</div><div class="kpi-subtext" style="color:#f97316;">Pendente Aprovação</div></div>""", unsafe_allow_html=True)
    with k4:
        st.markdown(f"""<div class="kpi-card" style="border-left-color: #ef4444;"><div class="kpi-label">Ações Atrasadas</div><div class="kpi-value">{atrasados}</div><div class="kpi-subtext" style="color:#ef4444;">⚠️ Atenção Imediata</div></div>""", unsafe_allow_html=True)
    with k5:
        cor_score = "#22c55e" if score_governança >= 80 else ("#eab308" if score_governança >= 60 else "#ef4444")
        st.markdown(f"""<div class="kpi-card" style="border-left-color: {cor_score};"><div class="kpi-label">Health Score Governança</div><div class="kpi-value" style="color: {cor_score};">{score_governança}%</div><div class="kpi-subtext" style="color: {cor_score};">Índice de Segurança</div></div>""", unsafe_allow_html=True)

    st.divider()

    g1, g2, g3 = st.columns([1.2, 1.5, 1.3])

    with g1:
        st.markdown("##### 🍩 Distribuição por Status")
        if not df_filtrado.empty:
            fig_donut = go.Figure(data=[go.Pie(labels=df_filtrado["Status"].value_counts().index, values=df_filtrado["Status"].value_counts().values, hole=0.55, marker=dict(colors=["#22c55e", "#3b82f6", "#f97316", "#ef4444", "#64748b"]), textinfo="label+percent")])
            fig_donut.update_layout(showlegend=False, margin=dict(t=20, b=20, l=10, r=10), height=260, annotations=[dict(text=f"<b>{total_achados}</b><br>Ações", x=0.5, y=0.5, font_size=16, showarrow=False)])
            st.plotly_chart(fig_donut, use_container_width=True)
        else:
            st.info("Sem dados para o filtro aplicado.")

    with g2:
        st.markdown("##### 📊 Matriz de Riscos por Área e Severidade")
        if not df_filtrado.empty:
            fig_bar_sev = px.bar(df_filtrado, x="Area_Responsavel", color="Severidade", color_discrete_map={"Crítica": "#ef4444", "Alta": "#f97316", "Média": "#eab308", "Baixa": "#22c55e"}, barmode="stack", labels={"Area_Responsavel": "Área Responsável", "count": "Qtd Apontamentos"})
            fig_bar_sev.update_layout(margin=dict(t=20, b=20, l=10, r=10), height=260, legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1))
            st.plotly_chart(fig_bar_sev, use_container_width=True)

    with g3:
        st.markdown("##### 🏢 Concentração de Riscos por Unidade")
        if not df_filtrado.empty:
            agencia_counts = df_filtrado["Agencia"].value_counts().reset_index()
            agencia_counts.columns = ["Agencia", "Quantidade"]
            fig_agencia = px.bar(agencia_counts, y="Agencia", x="Quantidade", orientation="h", color_discrete_sequence=["#0284c7"], text="Quantidade")
            fig_agencia.update_layout(margin=dict(t=20, b=20, l=10, r=10), height=260, xaxis_title="Apontamentos", yaxis_title="")
            st.plotly_chart(fig_agencia, use_container_width=True)

    st.divider()

    st.markdown("### 📋 Matriz Geral de Apontamentos (Todos os Registros)")
    df_matriz_exibicao = df_filtrado.copy()
    if not df_matriz_exibicao.empty:
        df_matriz_exibicao["Data_Inicio"] = df_matriz_exibicao["Data_Inicio"].apply(formatar_data_br)
        df_matriz_exibicao["Prazo"] = df_matriz_exibicao["Prazo"].apply(formatar_data_br)
        cols_exibir_matriz = ["ID", "Cliente_Projeto", "Achado", "Severidade", "Area_Responsavel", "Nome_Responsavel", "Data_Inicio", "Prazo", "Ultima_Acao", "Status"]
        st.dataframe(df_matriz_exibicao[cols_exibir_matriz], use_container_width=True, key="df_matriz_interativa")


# ---------------------------------------------------------
# TELA 2: DASHBOARD ANALÍTICO DE KPIS & MATRIZ DE RISCO 5X5
# ---------------------------------------------------------
with tabs[1]:
    st.markdown("### 📈 Dashboard Analítico de KPIs & Matriz de Risco 5x5")
    st.caption("Visão Avançada de Performance Operacional, SLA e Correlação de Severidades.")

    if not df_filtrado.empty:
        df_kpi = df_filtrado.copy()
        df_kpi["dt_inicio_parsed"] = pd.to_datetime(df_kpi["Data_Inicio"], errors="coerce")
        df_kpi["dt_prazo_parsed"] = pd.to_datetime(df_kpi["Prazo"], errors="coerce")
        now_dt = pd.to_datetime(get_now_br().strftime("%Y-%m-%d"))

        df_kpi["Dias_Corridos"] = (now_dt - df_kpi["dt_inicio_parsed"]).dt.days.fillna(0)
        df_kpi["Dias_Atraso_Real"] = (now_dt - df_kpi["dt_prazo_parsed"]).dt.days.apply(lambda x: max(0, x))

        dk1, dk2, dk3, dk4 = st.columns(4)
        dk1.metric("⏱️ Lead Time Médio", f"{df_kpi['Dias_Corridos'].mean():.1f} Dias")
        dk2.metric("🎯 Conformidade SLA", f"{((len(df_kpi[df_kpi['Status'] == 'Concluído']) / len(df_kpi))*100):.1f}%")
        dk3.metric("⚠️ Atraso Médio das Pendências", f"{df_kpi['Dias_Atraso_Real'].mean():.1f} Dias")
        dk4.metric("📊 Índice Risco Crítico", f"{((len(df_kpi[df_kpi['Severidade'] == 'Crítica']) / len(df_kpi))*100):.1f}%")

        st.divider()

        d_col1, d_col2 = st.columns([1.5, 1])

        with d_col1:
            st.markdown("##### 🎯 Heatmap - Matriz de Riscos 5x5 (Severidade x SLA)")
            
            # Construção da Matriz de Risco 5x5
            df_kpi["Faixa_SLA"] = pd.cut(
                df_kpi["Dias_Corridos"],
                bins=[-1, 7, 15, 30, 60, 9999],
                labels=["0-7 dias", "8-15 dias", "16-30 dias", "31-60 dias", "> 60 dias"]
            )
            
            matriz_pivot = pd.crosstab(
                df_kpi["Severidade"],
                df_kpi["Faixa_SLA"],
                dropna=False
            ).reindex(["Baixa", "Média", "Alta", "Crítica"], fill_value=0)

            fig_heatmap = px.imshow(
                matriz_pivot,
                labels=dict(x="Tempo de Resposta em Aberto", y="Severidade do Achado", color="Qtd Ações"),
                x=matriz_pivot.columns,
                y=matriz_pivot.index,
                color_continuous_scale="Reds",
                text_auto=True
            )
            fig_heatmap.update_layout(height=320, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_heatmap, use_container_width=True)

        with d_col2:
            st.markdown("##### 🏆 Ranking de Gargalo por Responsável")
            df_resp_gargalo = df_kpi[df_kpi["Status"] != "Concluído"]["Nome_Responsavel"].value_counts().reset_index()
            df_resp_gargalo.columns = ["Responsável", "Ações Pendentes"]
            
            fig_resp = px.bar(
                df_resp_gargalo.head(5),
                x="Ações Pendentes",
                y="Responsável",
                orientation="h",
                color="Ações Pendentes",
                color_continuous_scale="Oranges"
            )
            fig_resp.update_layout(height=320, margin=dict(t=10, b=10, l=10, r=10), showlegend=False)
            st.plotly_chart(fig_resp, use_container_width=True)
    else:
        st.info("Sem dados para alimentar os indicadores.")


# ---------------------------------------------------------
# TELA 3: OPERAÇÕES E GESTÃO DE PROJETOS (CRUD + ID AUTO)
# ---------------------------------------------------------
with tabs[2]:
    st.markdown("### ⚙️ Gestão de Projetos e Apontamentos")
    st.caption("Ações diretas de Abertura, Modificação e Exclusão de Apontamentos de Auditoria.")

    sub_t1, sub_t2, sub_t3 = st.tabs(["➕ Incluir Novo Projeto", "✏️ Editar Projeto", "🗑️ Excluir Projeto"])

    with sub_t1:
        st.markdown("##### ➕ Formulário de Abertura de Apontamento")
        id_sugerido = gerar_novo_id(df)
        st.info(f"🆔 **Novo ID gerado automaticamente:** `{id_sugerido}`")

        with st.form("form_inc_proj_novo", clear_on_submit=True):
            c_i1, c_i2 = st.columns(2)
            inc_proj = c_i1.text_input("Nome do Cliente / Projeto:")
            inc_ag = c_i2.text_input("Unidade / Agência:", value="São Paulo - HQ")

            inc_achado = st.text_area("Descrição do Achado / Desvio:")
            inc_acao = st.text_area("Ação Corretiva Recomendada:")

            c_i3, c_i4, c_i5 = st.columns(3)
            inc_sev = c_i3.selectbox("Severidade:", ["Baixa", "Média", "Alta", "Crítica"])
            inc_area = c_i4.text_input("Área Responsável:", value="Operações")
            inc_resp = c_i5.text_input("Nome do Responsável:")

            c_i6, c_i7, c_i8 = st.columns(3)
            inc_email = c_i6.text_input("E-mail do Responsável:")
            inc_dt_inicio = c_i7.date_input("Data de Início:", value=get_now_br().date(), format="DD/MM/YYYY")
            inc_prazo = c_i8.date_input("Prazo Limite:", value=get_now_br().date() + datetime.timedelta(days=15), format="DD/MM/YYYY")

            st.markdown("---")
            st.markdown("##### 🧩 Criar Campo Personalizado Adicional")
            c_cust1, c_cust2 = st.columns(2)
            campo_cust_nome = c_cust1.text_input("Nome do Novo Campo:")
            campo_cust_valor = c_cust2.text_input("Valor do Campo:")

            if st.form_submit_button("➕ Criar Registro Definitivo"):
                if not inc_proj or not inc_achado:
                    st.error("Por favor, preencha o Nome do Projeto e o Achado.")
                else:
                    try:
                        novo_id = gerar_novo_id(df)
                        data_inicio_str = inc_dt_inicio.strftime("%Y-%m-%d")
                        prazo_str = inc_prazo.strftime("%Y-%m-%d")

                        dict_custom = {}
                        if campo_cust_nome and campo_cust_valor:
                            dict_custom[campo_cust_nome] = campo_cust_valor

                        novo_registro = pd.DataFrame([{
                            "ID": str(novo_id),
                            "Cliente_Projeto": str(inc_proj),
                            "Agencia": str(inc_ag),
                            "Etapa_SIPOC": "Geral",
                            "Categoria": "Compliance",
                            "Achado": str(inc_achado),
                            "Severidade": str(inc_sev),
                            "Acao_Corretiva": str(inc_acao),
                            "Area_Responsavel": str(inc_area),
                            "Nome_Responsavel": str(inc_resp),
                            "Email_Responsavel": str(inc_email),
                            "Data_Inicio": data_inicio_str,
                            "Prazo": prazo_str,
                            "Data_Conclusao": "-",
                            "Status": st.session_state["colunas_kanban_custom"][0],
                            "Timeline_JSON": "[]",
                            "Campos_Custom_JSON": json.dumps(dict_custom, ensure_ascii=False),
                            "Ultima_Acao": formatar_data_br(data_inicio_str),
                        }])
                        df_novo = pd.concat([df, novo_registro], ignore_index=True)
                        salvar_dados(df_novo)
                        st.session_state["df_auditoria"] = df_novo
                        registrar_log(user_info["Usuario"], "Inclusão", f"Criou o apontamento {novo_id}")
                        st.success(f"Apontamento {novo_id} criado com sucesso!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao incluir novo projeto: {e}")

    with sub_t2:
        if not df.empty:
            proj_sel_ed = st.selectbox("Selecione o Projeto / Cliente para Editar:", options=df["Cliente_Projeto"].unique(), key="sb_ed_proj")
            df_sub_ed = df[df["Cliente_Projeto"] == proj_sel_ed]
            id_ed = st.selectbox("Apontamento Vinculado:", options=df_sub_ed["ID"].unique())
            row_ed = df[df["ID"] == id_ed].iloc[0]

            with st.form("form_ed_proj"):
                ed_proj = st.text_input("Nome do Cliente / Projeto:", value=row_ed["Cliente_Projeto"])
                ed_achado = st.text_area("Achado:", value=row_ed["Achado"])
                ed_acao = st.text_area("Ação:", value=row_ed["Acao_Corretiva"])

                c_e1, c_e2, c_e3 = st.columns(3)
                ed_resp = c_e1.text_input("Responsável:", value=row_ed["Nome_Responsavel"])
                ed_dt_inicio = c_e2.text_input("Data Início (YYYY-MM-DD):", value=str(row_ed.get("Data_Inicio", "-")))
                ed_prazo = c_e3.text_input("Novo Prazo (YYYY-MM-DD):", value=str(row_ed["Prazo"]))

                idx_st = 0
                if row_ed["Status"] in st.session_state["colunas_kanban_custom"]:
                    idx_st = st.session_state["colunas_kanban_custom"].index(row_ed["Status"])

                ed_status = st.selectbox("Status:", st.session_state["colunas_kanban_custom"], index=idx_st)

                if st.form_submit_button("💾 Salvar Alterações"):
                    try:
                        idx = df[df["ID"] == id_ed].index[0]
                        df.loc[idx, "Cliente_Projeto"] = str(ed_proj)
                        df.loc[idx, "Achado"] = str(ed_achado)
                        df.loc[idx, "Acao_Corretiva"] = str(ed_acao)
                        df.loc[idx, "Nome_Responsavel"] = str(ed_resp)
                        df.loc[idx, "Data_Inicio"] = str(ed_dt_inicio)
                        df.loc[idx, "Prazo"] = str(ed_prazo)
                        df.loc[idx, "Status"] = str(ed_status)
                        if "Concluído" in ed_status:
                            df.loc[idx, "Data_Conclusao"] = str(get_now_br().date())

                        salvar_dados(df)
                        st.session_state["df_auditoria"] = df
                        registrar_log(user_info["Usuario"], "Edição Apontamento", f"Editou o apontamento {id_ed}")
                        st.success("Projeto atualizado com sucesso!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao salvar edição: {e}")

    with sub_t3:
        if not df.empty:
            st.warning("⚠️ Atenção: A exclusão de um projeto apaga permanentemente o registro e TODOS os anexos físicos vinculados!")
            proj_sel_del = st.selectbox("Selecione o Projeto para Excluir:", options=df["Cliente_Projeto"].unique(), key="sb_del_proj")
            df_sub_del = df[df["Cliente_Projeto"] == proj_sel_del]
            id_del = st.selectbox("ID do Apontamento:", options=df_sub_del["ID"].unique(), key="sb_del_id")

            chk_crud_del = st.checkbox("Confirmo que desejo apagar permanentemente este projeto e seus arquivos físicos.", key="chk_crud_del")

            if st.button("🗑️ CONFIRMAR EXCLUSÃO DO PROJETO E ANEXOS", type="primary", disabled=not chk_crud_del):
                try:
                    row_del = df[df["ID"] == id_del].iloc[0]
                    qtd_del = apagar_anexos_do_apontamento(id_del, row_del.get("Timeline_JSON"))

                    df_novo = df[df["ID"] != id_del].copy()
                    salvar_dados(df_novo)
                    st.session_state["df_auditoria"] = df_novo
                    registrar_log(user_info["Usuario"], "Exclusão Apontamento", f"Excluiu o apontamento {id_del} e {qtd_del} anexo(s)")
                    st.success(f"Apontamento {id_del} e {qtd_del} arquivo(s) de evidência excluídos com sucesso!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro ao excluir o projeto: {e}")


# ---------------------------------------------------------
# TELA 4: CENTRAL DE PROJETOS (KANBAN)
# ---------------------------------------------------------
with tabs[3]:
    st.markdown("### 📌 Quadro Visual de Projetos & Ações")

    if "id_modal_aberto" in st.session_state and st.session_state["id_modal_aberto"]:
        renderizar_painel_detalhes(st.session_state["id_modal_aberto"])

    # --- NOVO: FILTRO RÁPIDO DE RESPONSÁVEL / CRIADOR ---
    col_resp_existente = "Nome_Responsavel" if "Nome_Responsavel" in df_filtrado.columns else ("Usuario_Criador" if "Usuario_Criador" in df_filtrado.columns else None)
    
    if col_resp_existente:
        lista_responsaveis = ["Todos os Responsáveis"] + sorted(df[col_resp_existente].dropna().astype(str).unique().tolist())
        resp_sel = st.selectbox("👤 Filtrar Cartões por Criador / Responsável:", options=lista_responsaveis, key="f_kanban_resp")
        if resp_sel != "Todos os Responsáveis":
            df_filtrado = df_filtrado[df_filtrado[col_resp_existente] == resp_sel]

    with st.expander("🛠️ Personalizar e Reordenar Colunas do Kanban"):
        c_k1, c_k2, c_k3 = st.columns([1.5, 1, 1.2])

        if "novo_col_input" not in st.session_state:
            st.session_state["novo_col_input"] = ""

        # Callback seguro para criação de colunas
        def processar_nova_coluna():
            nome_input = st.session_state.get("novo_col_input", "").strip()
            if nome_input and nome_input not in st.session_state["colunas_kanban_custom"]:
                pos_val = st.session_state.get("posicao_col_input", len(st.session_state["colunas_kanban_custom"]) + 1)
                idx_pos = int(pos_val) - 1
                
                st.session_state["colunas_kanban_custom"].insert(idx_pos, nome_input)
                salvar_colunas_kanban(st.session_state["colunas_kanban_custom"])
                
                st.session_state["novo_col_input"] = ""
                st.session_state["sucesso_msg_col"] = f"Coluna '{nome_input}' criada com sucesso!"

        nova_col_k = c_k1.text_input("Nome da Nova Coluna:", key="novo_col_input")
        posicao_k = c_k2.number_input(
            "Posição (1 a N):", 
            min_value=1, 
            max_value=len(st.session_state["colunas_kanban_custom"]) + 1, 
            value=len(st.session_state["colunas_kanban_custom"]) + 1,
            key="posicao_col_input"
        )

        c_k3.button("➕ Criar Coluna Dinâmica", on_click=processar_nova_coluna)

        if "sucesso_msg_col" in st.session_state and st.session_state["sucesso_msg_col"]:
            st.success(st.session_state["sucesso_msg_col"])
            st.session_state["sucesso_msg_col"] = ""

        st.divider()
        c_r1, c_r2, c_r3 = st.columns([1.5, 1, 1])
        col_del_k = c_r1.selectbox("Excluir Coluna:", options=st.session_state["colunas_kanban_custom"], key="sb_del_col")
        chk_col_del = c_r2.checkbox("Confirmar exclusão da coluna", key="chk_col_del")

        if c_r3.button("❌ Remover Coluna", disabled=not chk_col_del):
            if len(st.session_state["colunas_kanban_custom"]) > 1:
                st.session_state["colunas_kanban_custom"].remove(col_del_k)
                salvar_colunas_kanban(st.session_state["colunas_kanban_custom"])
                st.rerun()

    cols_st = st.columns(len(st.session_state["colunas_kanban_custom"]))

    for index, col_nome in enumerate(st.session_state["colunas_kanban_custom"]):
        with cols_st[index]:
            st.markdown(f"#### 📌 {col_nome}")
            itens = df_filtrado[df_filtrado["Status"] == col_nome]

            for _, row in itens.iterrows():
                sev_class = f"card-{str(row['Severidade']).lower()}"
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
                    idx_col_atual = 0
                    if col_nome in st.session_state["colunas_kanban_custom"]:
                        idx_col_atual = st.session_state["colunas_kanban_custom"].index(col_nome)

                    st_mudar = st.selectbox("Mover:", st.session_state["colunas_kanban_custom"], index=idx_col_atual, key=f"sb_st_c_{row['ID']}", label_visibility="collapsed")
                    
                    if st_mudar != col_nome:
                        try:
                            idx_k = df[df["ID"] == row["ID"]].index[0]
                            df.loc[idx_k, "Status"] = str(st_mudar)
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
# TELA 5: MESTRE CENTRAL DE ANEXOS
# ---------------------------------------------------------
with tabs[4]:
    st.markdown("### 📤 Mestre Central de Anexos & Repositório Documental")
    st.caption("Centralização inteligente e auditoria visual de todas as evidências anexadas no sistema.")

    lista_arquivos_servidor = [f for f in os.listdir(PASTA_EVIDENCIAS) if os.path.isfile(os.path.join(PASTA_EVIDENCIAS, f))]
    total_anexos_count = len(lista_arquivos_servidor)
    tamanho_total_mb = sum([os.path.getsize(os.path.join(PASTA_EVIDENCIAS, f)) for f in lista_arquivos_servidor]) / (1024 * 1024) if total_anexos_count > 0 else 0

    m1, m2, m3 = st.columns(3)
    m1.metric("📁 Total de Evidências Armazenadas", f"{total_anexos_count} Arquivos")
    m2.metric("💾 Espaço Utilizado em Disco", f"{tamanho_total_mb:.2f} MB")
    m3.metric("🔒 Status do Repositório", "Ativo & Auditado")

    st.divider()

    busca_anexo = st.text_input("🔍 Pesquisar por Nome de Arquivo ou Código do Apontamento:")

    if not lista_arquivos_servidor:
        st.info("Nenhuma evidência física foi armazenada no repositório ainda.")
    else:
        st.markdown("#### 📂 Galeria Físico-Documental de Anexos")
        cols_galeria = st.columns(3)

        idx_c = 0
        for arq in lista_arquivos_servidor:
            if busca_anexo and busca_anexo.lower() not in arq.lower():
                continue

            with cols_galeria[idx_c % 3]:
                caminho_arq = os.path.join(PASTA_EVIDENCIAS, arq)
                tam_kb = os.path.getsize(caminho_arq) / 1024

                st.markdown(
                    f"""
                <div class="file-card">
                    <div style="font-weight: bold; color: #0f172a; word-break: break-all;">📎 {arq}</div>
                    <div style="font-size: 0.78rem; color: #64748b; margin-top: 4px;">Tamanho: <b>{tam_kb:.1f} KB</b></div>
                </div>
                """,
                    unsafe_allow_html=True,
                )

                if arq.lower().endswith((".png", ".jpg", ".jpeg")):
                    st.image(caminho_arq, use_column_width=True)

                with open(caminho_arq, "rb") as f_data:
                    st.download_button(
                        label=f"📥 Baixar Arquivo",
                        data=f_data,
                        file_name=arq,
                        mime="application/octet-stream",
                        key=f"dl_mst_gal_{arq}",
                    )
            idx_c += 1


# ---------------------------------------------------------
# TELA 6: EXTRATOR DE RELATÓRIOS COM EXCEL ENTERPRISE
# ---------------------------------------------------------
with tabs[5]:
    st.markdown("### 📥 Extrator Inteligente & Análise de SLA de Governança")
    st.caption("Geração de relatórios gerenciais e exportação formatada em Excel Executivo.")

    f_c1, f_c2, f_c3 = st.columns(3)
    proj_filtro_rel = f_c1.selectbox("Filtrar por Projeto:", options=["Todos os Projetos"] + list(df["Cliente_Projeto"].unique()), key="f_rel_proj")
    dt_ini_rel_e = f_c2.date_input("Emissão (De):", value=datetime.date(2020, 1, 1), format="DD/MM/YYYY", key="f_rel_e_de")
    dt_fim_rel_e = f_c3.date_input("Emissão (Até):", value=datetime.date(2030, 12, 31), format="DD/MM/YYYY", key="f_rel_e_ate")

    df_rel = df.copy()

    if proj_filtro_rel != "Todos os Projetos":
        df_rel = df_rel[df_rel["Cliente_Projeto"] == proj_filtro_rel]

    if dt_ini_rel_e and dt_fim_rel_e:
        df_rel["dt_tmp_e"] = pd.to_datetime(df_rel["Data_Inicio"], errors="coerce").dt.date
        df_rel = df_rel[(df_rel["dt_tmp_e"] >= dt_ini_rel_e) & (df_rel["dt_tmp_e"] <= dt_fim_rel_e)]

    st.divider()

    if not df_rel.empty:
        df_sla = df_rel.copy()
        df_sla["dt_inicio_parsed"] = pd.to_datetime(df_sla["Data_Inicio"], errors="coerce")
        df_sla["dt_prazo_parsed"] = pd.to_datetime(df_sla["Prazo"], errors="coerce")
        now_br_dt = pd.to_datetime(get_now_br().strftime("%Y-%m-%d"))
        
        df_sla["Dias_Corridos"] = (now_br_dt - df_sla["dt_inicio_parsed"]).dt.days
        df_sla["Dias_Para_Vencer"] = (df_sla["dt_prazo_parsed"] - now_br_dt).dt.days

        # Proteção contra divisão por zero
        total_acoes = len(df_sla)
        taxa_conformidade = ((len(df_sla[df_sla['Status'] == 'Concluído']) / total_acoes) * 100) if total_acoes > 0 else 0.0
        lead_time_medio = df_sla['Dias_Corridos'].mean() if total_acoes > 0 else 0.0

        st.markdown("#### 📊 Painel de Desempenho e Matriz de SLA do Projeto")
        s1, s2, s3 = st.columns(3)
        s1.metric("⏱️ Lead Time Médio de Ações", f"{lead_time_medio:.1f} Dias")
        s2.metric("🎯 Taxa de Conformidade no Prazo", f"{taxa_conformidade:.1f}%")
        s3.metric("⚠️ Apontamentos Críticos Ativos", f"{len(df_sla[df_sla['Severidade'] == 'Crítica'])} Ações")

        st.divider()

        st.markdown("#### 📋 Matriz Analítica Executiva de SLA")
        cols_sla_render = ["ID", "Cliente_Projeto", "Achado", "Severidade", "Nome_Responsavel", "Status", "Dias_Corridos", "Dias_Para_Vencer"]
        
        # Garante a exibição das colunas mapeadas que existem no DataFrame
        cols_existentes = [c for c in cols_sla_render if c in df_sla.columns]
        st.dataframe(df_sla[cols_existentes], use_container_width=True)

    st.divider()
    try:
        excel_bytes = gerar_excel_estilizado(df_rel)
        st.download_button(
            label="📊 Exportar Relatório Executivo Estilizado em Excel (.xlsx)",
            data=excel_bytes,
            file_name=f"Relatorio_Executivo_Compliance_{get_now_br().strftime('%d_%m_%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"Erro ao gerar planilha Excel: {e}")

# ---------------------------------------------------------
# TELA 7: AUDITORIA MASTER
# ---------------------------------------------------------
if is_master:
    with tabs[6]:
        st.markdown("### 🛡️ Central Master de Auditoria de Sistema & Logs de Rastreabilidade")
        st.caption("Visão exclusiva de governança de TI, controle de privilégios de acesso e histórico completo de auditoria do sistema.")

        df_logs_all = pd.DataFrame()
        if os.path.exists(ARQUIVO_LOGS):
            try:
                df_logs_all = pd.read_excel(ARQUIVO_LOGS, dtype=str)
            except:
                pass

        total_logs = len(df_logs_all)
        total_exclusoes = len(df_logs_all[df_logs_all["Acao"].astype(str).str.contains("Exclusão", case=False, na=False)]) if not df_logs_all.empty else 0
        total_edicoes = len(df_logs_all[df_logs_all["Acao"].astype(str).str.contains("Edição|Troca", case=False, na=False)]) if not df_logs_all.empty else 0

        a1, a2, a3 = st.columns(3)
        a1.metric("📜 Total de Interações Registradas", f"{total_logs} Eventos")
        a2.metric("⚠️ Exclusões Críticas Efetuadas", f"{total_exclusoes} Ações")
        a3.metric("✏️ Alterações de Registros", f"{total_edicoes} Operações")

        st.divider()

        t_acessos, t_logs = st.tabs(["👥 Níveis de Acesso & Privilégios", "📜 Trilha Auditável Completa de Logs"])
        df_u = carregar_usuarios()

        with t_acessos:
            st.markdown("##### 🔑 Gerenciamento de Privilégios e Aprovação de Usuários")
            df_u_exibicao = df_u.copy()
            if "Senha" in df_u_exibicao.columns:
                df_u_exibicao["Senha"] = "••••••••"

            st.dataframe(df_u_exibicao, use_container_width=True)

            if not df_u.empty:
                st.markdown("##### ⚙️ Alterar Nível de Acesso ou Status de Usuário")
                c_u1, c_u2, c_u3 = st.columns(3)
                usr_aprovar = c_u1.selectbox("Usuário:", options=df_u["Usuario"].unique())
                novo_st_u = c_u2.selectbox("Novo Status:", ["Ativo", "Pendente", "Bloqueado"])
                novo_nv_u = c_u3.selectbox("Novo Nível:", ["Gestor", "Master"])

                if st.button("💾 Confirmar Atualização de Acesso"):
                    try:
                        idx_u = df_u[df_u["Usuario"] == usr_aprovar].index[0]
                        df_u.loc[idx_u, "Status"] = novo_st_u
                        df_u.loc[idx_u, "Nivel"] = novo_nv_u
                        salvar_usuarios(df_u)
                        registrar_log(user_info["Usuario"], "Gestão Acessos", f"Alterou privilégios do usuário {usr_aprovar} para {novo_st_u}/{novo_nv_u}")
                        st.success("Permissões de usuário atualizadas com sucesso!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao atualizar permissões: {e}")

        with t_logs:
            st.markdown("##### 📜 Filtro Avançado da Trilha Auditável de Sistema")
            if df_logs_all.empty:
                st.info("Nenhum log de auditoria registrado no sistema até o momento.")
            else:
                c_fl1, c_fl2 = st.columns(2)
                filtro_usr_log = c_fl1.selectbox("Filtrar por Usuário:", options=["Todos"] + list(df_logs_all["Usuario"].unique()))
                filtro_acao_log = c_fl2.selectbox("Filtrar por Tipo de Ação:", options=["Todas"] + list(df_logs_all["Acao"].unique()))

                df_logs_filtrado = df_logs_all.copy()
                if filtro_usr_log != "Todos":
                    df_logs_filtrado = df_logs_filtrado[df_logs_filtrado["Usuario"] == filtro_usr_log]
                if filtro_acao_log != "Todas":
                    df_logs_filtrado = df_logs_filtrado[df_logs_filtrado["Acao"] == filtro_acao_log]

                st.dataframe(df_logs_filtrado, use_container_width=True)