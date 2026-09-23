import datetime
import os
import smtplib
import sys
from email.mime.text import MIMEText
from email.header import Header
import pandas as pd

# Caminho da matriz de auditoria
ARQUIVO_MATRIZ = "matriz_auditoria.xlsx"

if not os.path.exists(ARQUIVO_MATRIZ):
    print(f"❌ Erro: O arquivo '{ARQUIVO_MATRIZ}' não foi encontrado na pasta atual.")
    sys.exit(1)

# 1. Leitura e Preparação da Base
df = pd.read_excel(ARQUIVO_MATRIZ, dtype=str)

# Converter coluna Prazo
df["Prazo_dt"] = pd.to_datetime(df["Prazo"], errors="coerce")
hoje = pd.to_datetime(datetime.date.today())

# Garantir existência da coluna de controle de alertas
if "Ultimo_Alerta_Enviado" not in df.columns:
    df["Ultimo_Alerta_Enviado"] = None

# 2. Lógica de Classificação de Prazos e Status
# Considera atrasadas apenas ações com prazo vencido e status que não sejam de conclusão/validação final
status_ignorados = ["Concluído", "Em Validação", "Em Validação (Auditor)"]

df["Dias_Atraso"] = (hoje - df["Prazo_dt"]).dt.days

# Atualizar status automaticamente na memória se o prazo expirou
mask_atrasado = (~df["Status"].isin(status_ignorados)) & (df["Prazo_dt"] < hoje)
df.loc[mask_atrasado, "Status"] = "Atrasado"

# Filtrar apenas ações atrasadas com e-mail válido
atrasados = df[
    (df["Status"] == "Atrasado")
    & (df["Email_Responsavel"].notna())
    & (df["Email_Responsavel"].str.strip() != "")
]


# 3. Função de Envio de E-mail (Suporte a UTF-8 e HTML/Text Limpo)
def enviar_alerta(destinatario, nome, id_acao, acao, prazo_dt, dias):
    assunto = f"[AUDITORIA - ALERTA CRÍTICO] Plano de Ação Atrasado: {id_acao}"
    
    prazo_str = prazo_dt.strftime('%d/%m/%Y') if pd.notnull(prazo_dt) else "-"

    # Mensagem formatada em HTML limpo
    corpo_html = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333333; line-height: 1.6;">
        <p>Prezado(a) <strong>{nome}</strong>,</p>
        <p>Identificamos que o Plano de Ação referente ao trabalho de Auditoria Interna / Compliance encontra-se 
        <span style="color: #d9534f; font-weight: bold;">ATRASADO</span>.</p>
        
        <table style="border-collapse: collapse; width: 100%; max-width: 600px; margin: 15px 0;">
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd; background-color: #f9f9f9; font-weight: bold;">ID da Ação:</td>
            <td style="padding: 8px; border: 1px solid #ddd;">{id_acao}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd; background-color: #f9f9f9; font-weight: bold;">Ação Requerida:</td>
            <td style="padding: 8px; border: 1px solid #ddd;">{acao}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd; background-color: #f9f9f9; font-weight: bold;">Prazo Limite Acordado:</td>
            <td style="padding: 8px; border: 1px solid #ddd;">{prazo_str}</td>
          </tr>
          <tr>
            <td style="padding: 8px; border: 1px solid #ddd; background-color: #f9f9f9; font-weight: bold;">Dias em Atraso:</td>
            <td style="padding: 8px; border: 1px solid #ddd; color: #d9534f; font-weight: bold;">{dias} dia(s)</td>
          </tr>
        </table>
        
        <p>Solicitamos acessar imediatamente o <strong>Portal de Compliance</strong> para atualizar o status e anexar a evidência de conclusão.</p>
        <br>
        <p>Atenciosamente,<br>
        <strong>Equipe de Auditoria Interna & Compliance</strong><br>
        Maringá Turismo</p>
      </body>
    </html>
    """

    msg = MIMEText(corpo_html, "html", "utf-8")
    msg["Subject"] = Header(assunto, "utf-8")
    remetente = os.environ.get("SMTP_USER", "auditoria@maringaturismo.com.br")
    msg["From"] = remetente
    msg["To"] = destinatario

    # ENVIO REAL VIA SMTP (Configuração por Variáveis de Ambiente)
    smtp_server = os.environ.get("SMTP_SERVER", "smtp.office365.com")
    smtp_port = int(os.environ.get("SMTP_PORT", 587))
    smtp_pass = os.environ.get("SMTP_PASSWORD", "")

    if smtp_pass:
        try:
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(remetente, smtp_pass)
                server.sendmail(remetente, [destinatario], msg.as_string())
            print(f"✅ Alerta SMTP enviado para {destinatario} | ID: {id_acao}")
            return True
        except Exception as e:
            print(f"❌ Falha ao enviar e-mail via SMTP para {destinatario}: {e}")
            return False
    else:
        # Modo de simulação (Log local quando a senha não está configurada)
        print(f"🔹 [SIMULAÇÃO] Alerta para {destinatario} | Ação: {id_acao} | Atraso: {dias} dias.")
        return True

# 4. Disparo e Atualização da Base
enviados_count = 0
for idx, row in atrasados.iterrows():
    data_ultimo = row.get("Ultimo_Alerta_Enviado")
    
    # Evita reenvio no mesmo dia
    if pd.isna(data_ultimo) or str(data_ultimo).strip() in ["", "None", "nan"] or pd.to_datetime(data_ultimo, errors="coerce").date() < hoje.date():
        sucesso = enviar_alerta(
            destinatario=str(row["Email_Responsavel"]).strip(),
            nome=str(row.get("Nome_Responsavel", "Responsável")).strip(),
            id_acao=str(row["ID"]),
            acao=str(row.get("Acao_Corretiva", "")),
            prazo_dt=row["Prazo_dt"],
            dias=row["Dias_Atraso"],
        )
        if sucesso:
            df.loc[idx, "Ultimo_Alerta_Enviado"] = hoje.strftime("%Y-%m-%d")
            enviados_count += 1

# Remover colunas temporárias e salvar no Excel
if "Prazo_dt" in df.columns:
    df.drop(columns=["Prazo_dt"], inplace=True)

df.to_excel(ARQUIVO_MATRIZ, index=False)
print(f"📌 Processamento de cobrança concluído. Total de alertas processados hoje: {enviados_count}")