import datetime
import os
import smtplib
from email.mime.text import MIMEText
import pandas as pd

# Caminho da matriz de auditoria
ARQUIVO_MATRIZ = "matriz_auditoria.xlsx"

if not os.path.exists(ARQUIVO_MATRIZ):
    print(
        f"Erro: O arquivo '{ARQUIVO_MATRIZ}' não foi encontrado na pasta atual."
    )
    exit()

# 1. Leitura e Preparação da Base
df = pd.read_excel(ARQUIVO_MATRIZ)
hoje = pd.to_datetime(datetime.date.today())
df["Prazo"] = pd.to_datetime(df["Prazo"])

# Garantir existência da coluna de controle de alertas
if "Ultimo_Alerta_Enviado" not in df.columns:
    df["Ultimo_Alerta_Enviado"] = None

# 2. Lógica de Classificação de Prazos e Status
df["Dias_Atraso"] = (hoje - df["Prazo"]).dt.days

# Atualizar automaticamente para Atrasado se prazo passou e não está concluído
df.loc[
    (df["Status"] != "Concluído") & (df["Prazo"] < hoje), "Status"
] = "Atrasado"

# Filtrar apenas ações atrasadas com e-mail válido
atrasados = df[
    (df["Status"] == "Atrasado")
    & (df["Email_Responsavel"].notna())
    & (df["Email_Responsavel"] != "")
]


# 3. Função de Envio de E-mail
def enviar_alerta(destinatario, nome, id_acao, acao, prazo, dias):
    assunto = f"[AUDITORIA - ALERTA CRÍTICO] Plano de Ação Atrasado: {id_acao}"
    corpo = f"""
    Prezado(a) {nome},

    Identificamos que o Plano de Ação referente ao trabalho de Auditoria Interna / Compliance encontra-se ATRASADO.

    - ID da Ação: {id_acao}
    - Ação Requerida: {acao}
    - Prazo Limite Acordado: {prazo.strftime('%d/%m/%Y')}
    - Dias em Atraso: {dias} dia(s)

    Solicitamos acessar imediatamente o Portal de Compliance para atualizar o status e anexar a evidência de conclusão.

    Atenciosamente,
    Equipe de Auditoria Interna & Compliance
    Maringá Turismo
    """
    msg = MIMEText(corpo)
    msg["Subject"] = assunto
    msg["From"] = "auditoria@maringaturismo.com.br"
    msg["To"] = destinatario

    # DESCOMENTAR PARA HABILITAR ENVIO REAL SMTP:
    # with smtplib.SMTP('smtp.office365.com', 587) as server:
    #     server.starttls()
    #     server.login("auditoria@maringaturismo.com.br", "SuaSenhaAqui")
    #     server.sendmail(msg['From'], [destinatario], msg.as_string())

    print(
        f"✅ Alerta enviado para {destinatario} | Ação: {id_acao} | Atraso: {dias} dias."
    )

# 4. Disparo e Atualização da Base
for idx, row in atrasados.iterrows():
    # Evita enviar múltiplos alertas no mesmo dia
    data_ultimo = row["Ultimo_Alerta_Enviado"]
    if pd.isna(data_ultimo) or pd.to_datetime(data_ultimo).date() < hoje.date():
        enviar_alerta(
            destinatario=row["Email_Responsavel"],
            nome=row["Nome_Responsavel"],
            id_acao=row["ID"],
            acao=row["Acao_Corretiva"],
            prazo=row["Prazo"],
            dias=row["Dias_Atraso"],
        )
        df.loc[idx, "Ultimo_Alerta_Enviado"] = hoje.strftime("%Y-%m-%d")

# Salvar atualizações de volta no Excel
df.to_excel(ARQUIVO_MATRIZ, index=False)
print("📌 Processamento de cobrança concluído e base atualizada.")







#git add requirements.txt app.py
#git commit -m "fix: adiciona openpyxl no requirements e ajusta caminho do app.py"
#git push origin main