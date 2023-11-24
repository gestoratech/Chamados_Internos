import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import time

# Realizando a configuração do titulo e logo da empresa dentro do site
st.set_page_config(page_title="Gestora Contabilidade - Envio de Chamados", page_icon="logovetorizada.png", layout="wide")
logo = "logovetorizada.png"

# Retirando algumas particularidades do Streamlit=
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {display: none !important;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)

page_bg_img = """
<style>
[data-testid="stAppViewContainer"] {
background: linear-gradient(180deg, rgba(137, 137, 137, 0.48), rgba(255, 255, 255, 1), rgba(255,255,255,1), rgba(255, 255, 255, 1), rgba(186, 209, 255, 1))
}
</style>
"""
           
st.markdown(page_bg_img, unsafe_allow_html=True)

st.image("logovetorizada.png", width=90)

# Título da aplicação
st.title(":blue[Gestora Contabilidade]")

# Subtítulo e breve descrição
st.header(
    """
  Seja muito bem-vindo ao sistema de envio de chamados internos da :blue[Gestora Contabilidade].
  """
)

st.write("___")

st.header(
    """
  Preencha os dados a seguir:
    """
)

# Configuração da conexão com o Google Sheets
scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('gsheets.json', scope)
client = gspread.authorize(credentials)

# Nome da planilha e da guia
spreadsheet_name = "RQ TI 07 - Gestão de Chamados Internos"
worksheet_name = "Apontamentos"

# Tenta abrir a planilha
try:
    sheet = client.open(spreadsheet_name)

    # Tenta abrir a guia 'Apontamentos'
    try:
        worksheet = sheet.worksheet(worksheet_name)

        # Lê os dados da guia 'Apontamentos'
        data = worksheet.get_all_values()
        header = data[0]
        rows = data[1:]

        # Adapte a lógica para a integração com o Streamlit
        tickets_nao_enviados = []

        for i, row in enumerate(rows):
            ticket_envio = row[2]
            tecnico = row[40]
            observa = row[16]
            emailEnvi = row[41]
            solicitante = row[13]

            # Se a coluna AP (41º índice) não estiver preenchida, mostra as informações
            if not emailEnvi:
                st.subheader("\nTickets sem e-mail enviado:")
                st.subheader(f"Ticket: :blue[{ticket_envio}] - Técnico: :blue[{tecnico}]")
                st.markdown(f":blue[Observações do chamado:] {observa}")
                st.markdown(f":blue[Solicitante:] {solicitante}")

                email_ticket = st.text_input(f"Informe o e-mail para o Ticket {ticket_envio}: ")

                if email_ticket:
                    # Atualiza a planilha com o e-mail fornecido e marca como enviado
                    worksheet.update_cell(i + 2, 42, "Sim")  # Atualiza a coluna 42 (letra 'AP') com "Sim"

                    # Código para enviar o e-mail ao colaborador
                    username = "rpa@gestoracontabilidade.com.br"
                    password = "Tav33891"
                    mail_from = "rpa@gestoracontabilidade.com.br"
                    mail_to = email_ticket
                    mail_subject = f"Avaliação ticket de nº {ticket_envio}"

                    mimemsg = MIMEMultipart()
                    mimemsg['From'] = mail_from
                    mimemsg['To'] = mail_to
                    mimemsg['Subject'] = mail_subject
                    mimemsg.attach(MIMEText(
                        f'''
                            <!DOCTYPE html>
                            <html>
                            <body>
                                <div>
                                <p>Prezado(a), <u><b>{row[13]}</b></u></p>
                                </div>
                                <div>
                                    <div>
                                        <p> Atendimento referente a(o) <u><b>{row[16]}</b></u>, realizado pelo técnico <u><b>{row[40]}</b></u>.</p>
                                    </div>
                                    <div>
                                        <p>Por gentileza, fazer avaliação do ticket de nº <u><b>{row[2]}</b></u>, através do link abaixo.</p>
                                    </div>
                                    <div>
                                        <p><a title="link" href="https://forms.gle/oYqnGsMxqm7LhFVj8">Click aqui! Para avaliar o ticket.</a></p>
                                    </div>
                                    <div>
                                        <img src="https://i.ibb.co/qp9n0wR/rpa.jpg" alt="rpa" border="0">
                                    </div>
                                </div>
                            </body>
                            </html>
                        ''', "html", "utf-8"
                    ))

                    # Configuração do servidor SMTP
                    connection = smtplib.SMTP(host='smtp.outlook.com', port=587)
                    connection.starttls()
                    connection.login(username, password)

                    # Envio do e-mail
                    connection.send_message(mimemsg)
                    connection.quit()
                    
                    with st.spinner ('Enviando...'):
                        time.sleep(4)
                        st.success("E-mail do ticket enviado com sucesso!")
                        
                    time.sleep(2)
                    
                    # Recarrega a página para atualizar os dados
                    with st.spinner ('Carregando...'):
                        time.sleep(2)
                        st.experimental_rerun()

                else:
                    st.write("Não há nenhum e-mail fornecido.")
                    
                st.write("___")

            else:
                # Adiciona os tickets enviados à lista
                tickets_nao_enviados.append({
                    "ticket": ticket_envio,
                    "tecnico": tecnico
                })
        # Lendo a planilha inteira para identificar as variáveis
        if not any(row[41] == "" for row in rows):
          st.warning('Todos os tickets já foram enviados!')
    # Exeções direcionada as conexões
    except gspread.WorksheetNotFound:
        st.error(f"Erro: A guia '{worksheet_name}' não foi encontrada na planilha '{spreadsheet_name}'.")
    except Exception as e:
        st.error(f"Erro ao ler os dados da guia: {e}")

except gspread.SpreadsheetNotFound:
    st.error(f"Erro: A planilha '{spreadsheet_name}' não foi encontrada.")
except Exception as e:
    st.error(f"Erro ao abrir a planilha: {e}")

st.info('Recarregue a página para verificar novos chamados para envio.')