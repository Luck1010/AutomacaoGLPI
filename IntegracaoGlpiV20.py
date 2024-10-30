import requests
import json    
from datetime import datetime
from fpdf import FPDF
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import tempfile
import os
from bs4 import BeautifulSoup
import html

# Configurações da API
api_url = 'https://servicedesk.yhbrasil.com/apirest.php/'  # Substitua pelo endereço correto do seu servidor GLPI
app_token = ''  # Substitua pelo seu App Token
user_token = ''  # Substitua pelo seu User Token

# Configurações de email
smtp_server = ''  # Substitua pelo servidor SMTP do seu provedor de email
smtp_port = 587
smtp_user = ''  # Substitua pelo seu email
smtp_password = ''  # Substitua pela sua senha de email
to_emails = []  # Lista de destinatários
email_subject = 'PASSAGEM DE TURNO'
email_body = 'Prezados, boa noite tudo bem?\n\nSegue passagem de turno em anexo.'

# Assinatura HTML
html_signature = """
<html>
<head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8"></head>
<body>
    <br><br>
    <table cellpadding="0" cellspacing="0" class="sc-gPEVay eQYmiW" style="vertical-align: -webkit-baseline-middle; letter-spacing: 2px; font-size: large; font-family: sans-serif;">
        <tbody>
            <tr>
                <td><div style="width: 0px;"></div>
                    <table cellpadding="0" cellspacing="0" class="sc-gPEVay eQYmiW" style="vertical-align: -webkit-baseline-middle; letter-spacing: 2px; font-size: large; font-family: sans-serif;">
                        <tbody>
                            <tr>
                                <td width="0" style="vertical-align: middle;">
                                    <span class="sc-kgAjT cuzzPp" style="margin-right: 0px; display: block;">
                                        <img src= "https://servicedesk.yhbrasil.com/YHLOGO.png" role="presentation" width="120" class="sc-cHGsZl bHiaRe" style="max-width: 120px;">
                                    </span>
                                </td>
                                <td style="vertical-align: middle;">
                                    <h3 color="#2e2e2e" class="sc-fBuWsC eeihxG" style="margin: 10px; font-size: 14px; color: rgb(46, 46, 46);">
                                        <span>Ticket</span><span>&nbsp;</span><span></span>
                                    </h3>
                                    <p color="#2e2e2e" font-size="large" class="sc-fMiknA bxZCMx" style="margin: 10px; color: rgb(252, 164, 4); font-size: 14px; line-height: 2px;">
                                        <b><span>TECNOLOGIA</span></b>
                                    </p>
                                    <p color="#2e2e2e" font-size="large" class="sc-fMiknA bxZCMx" style="margin: 10px; color: rgb(12, 12, 12); font-size: 8px; line-height: 2px;">
                                        <b><span>CONTACT CENTER E MUITO </span></b><b><span color="#2e2e2e" font-size="large" class="sc-fMiknA bxZCMx" style="margin: 10px; color: rgb(0, 0, 0); font-size: 10px; line-height: 2px;">+</span></b>
                                    </p>
                                </td>
                                <td width="2">
                                    <div style="border-width: 1px;"></div>
                                </td>
                                <td color="#e1bd31" direction="vertical" width="1" height="1" class="sc-jhAzac hmXDXQ" style="border: 5px; border-bottom: none; border-left: 2px solid rgb(252, 164, 4);"></td>
                                <td style="vertical-align: middle;">
                                    <table cellpadding="0" cellspacing="0" class="sc-gPEVay eQYmiW" style="vertical-align: -webkit-baseline-middle; font-size: large; font-family: sans-serif;">
                                        <tbody>
                                            <tr height="0" style="vertical-align: middle;"></tr>
                                            <tr>
                                                <td width="360" style="vertical-align: middle;">
                                                    <span color="#2e2e2e" class="sc-gipzik iyhjGb" style="font-size: 10px; color: rgb(252, 164, 4); font-size: 10px;">
                                                        <br>
                                                        <b><span>&nbsp;&nbsp;&nbsp; &nbsp; Rua Bonnard, 980 Bloco 22 | Alphaville Empresarial - Barueri - SP</span></b>
                                                    </span>
                                                    <table cellpadding="0" cellspacing="0" class="sc-gPEVay eQYmiW" style="vertical-align: -webkit-baseline-middle; font-size: large; font-family: sans-serif;">
                                                        <tbody>
                                                            <tr>
                                                                <td width="0"></td>
                                                            </tr>
                                                            <tr height="10" style="vertical-align: middle;"></tr>
                                                        </tbody>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="padding: 2px; margin-left: 12px; margin-top: 2px; color: rgb(46, 46, 46);">
                                                    <span color="#2e2e2e" class="sc-csuQGl CQhxV" style="font-size: 10px; color: rgb(46, 46, 46); font-size: 10px;">
                                                        <span>ticket@yhbrasil.com.br</span>
                                                    </span>
                                                </td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <table cellpadding="0" cellspacing="" class="sc-gPEVay eQYmiW" style="vertical-align: -webkit-baseline-middle; font-size: large; font-family: sans-serif;">
                                        <tbody>
                                            <tr>
                                                <td style="padding: 2px; margin-left: 12px; margin-right: 10px; margin-top: 5px; color: rgb(46, 46, 46);">
                                                    <span color="#2e2e2e" class="sc-csuQGl CQhxV" style="font-size: 10px; color: rgb(46, 46, 46);">
                                                        <img src="https://servicedesk.yhbrasil.com/CERTIFICADOS.png" width="60" style="max-width: 100px;" align="right">
                                                        <span>11 3777.9577</span> | <a color="#2e2e2e" class="sc-gipzik iyhjGb" style="text-decoration: none; color: rgb(46, 46, 46); font-size: 10px;">
                                                            <span>11 3777.9577 &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;</span>
                                                        </a>
                                                    </span>
                                                </td>
                                                <td></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr>
                <td></td>
            </tr>
            <tr>
                <td></td>
            </tr>
        </tbody>
    </table>
</body>
</html>
"""

# Mapeamento dos campos
field_mapping = {
    "2": "id",
    "1": "titulo",
    "12": "status",
    "15": "data_abertura"
}

# Função para formatar data
def format_date(date_str):
    try:
        dt = datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%d-%m-%Y %H:%M:%S")
    except ValueError:
        return date_str

# Função para criar PDF
def create_pdf(tickets):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Relatório de Chamados Pendentes", ln=True, align="C")
    pdf.ln(10)
    
    for ticket in tickets:
        for key, value in ticket.items():
            pdf.cell(200, 10, txt=f"{key.capitalize()}: {value}", ln=True)
        pdf.ln(10)
    
    temp_dir = tempfile.gettempdir()
    pdf_output = os.path.join(temp_dir, "chamados_pendentes.pdf")
    pdf.output(pdf_output)
    return pdf_output

# Função para enviar email
def send_email(subject, body, to_emails, attachment_path=None):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = ', '.join(to_emails)
    msg['Subject'] = subject

    # Anexar corpo do email e assinatura HTML
    body_html = f"<html><body>{body.replace('\n', '<br>')}<br>{html_signature}</body></html>"
    msg.attach(MIMEText(body_html, 'html'))

    if attachment_path:
        with open(attachment_path, "rb") as attachment:
            part = MIMEApplication(attachment.read(), _subtype="pdf")
            part.add_header('Content-Disposition', 'attachment', filename='chamados_pendentes.pdf')
            msg.attach(part)

    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.sendmail(smtp_user, to_emails, msg.as_string())
        server.quit()
        print(f'Email enviado para {", ".join(to_emails)}')
    except Exception as e:
        print(f'Erro ao enviar email: {e}')

# Iniciar sessão
session_url = f"{api_url}/initSession"
headers = {
    'Content-Type': 'application/json',
    'App-Token': app_token,
    'Authorization': f"user_token {user_token}"
}

def get_last_comment(ticket_id, headers):
    comments_url = f"{api_url}/Ticket/{ticket_id}/TicketFollowup"
    response = requests.get(comments_url, headers=headers)
    if response.status_code == 200:
        comments = response.json()
        if comments:
            last_comment = comments[-1].get('content', 'Sem comentários')
            # Decodificar entidades HTML
            last_comment = html.unescape(last_comment)
            # Limpar HTML usando BeautifulSoup
            if "<" in last_comment and ">" in last_comment:
                soup = BeautifulSoup(last_comment, 'html.parser')
                cleaned_comment = soup.get_text()
                return cleaned_comment
            return last_comment
    return 'Sem comentários'

try:
    response = requests.get(session_url, headers=headers)
    response.raise_for_status()  # Verifica se houve erro na requisição
    session_token = response.json().get('session_token')

    # Verificar se a sessão foi iniciada com sucesso
    if not session_token:
        print('Erro ao iniciar sessão')
        exit()

    # Configurar headers com session token
    headers['Session-Token'] = session_token

    # Fazer busca por chamados pendentes no grupo técnico "Departamentos > TI > Field Services" (ID 9)
    search_url = f"{api_url}/search/Ticket"
    params = {
        'criteria[0][field]': '12',  # Código do campo de status
        'criteria[0][searchtype]': 'equals',
        'criteria[0][value]': '4',  # valor para status pendente
        'criteria[1][link]': 'AND',
        'criteria[1][field]': '8',  # Código do campo de grupo técnico atribuído
        'criteria[1][searchtype]': 'equals',
        'criteria[1][value]': '9',  # ID do grupo técnico "Departamentos > TI > Field Services"
        'forcedisplay[0]': '2',     # Código do campo de ID
        'forcedisplay[1]': '1',     # Código do campo de nome
        'forcedisplay[2]': '12',    # Código do campo de status
        'forcedisplay[3]': '15',    # Código do campo de data de criação
    }

    response = requests.get(search_url, headers=headers, params=params)
    
    # Tratar o status 206 como sucesso
    if response.status_code == 206 or response.status_code == 200:
        tickets = response.json().get('data', [])
        formatted_tickets = []
        for ticket in tickets:
            # Formatar os dados do ticket
            formatted_ticket = {field_mapping[key]: value for key, value in ticket.items() if key in field_mapping}
            
            # Ajustar campos específicos
            if 'status' in formatted_ticket:  # Campo de status
                formatted_ticket['status'] = 'pendente' if formatted_ticket['status'] == 4 else formatted_ticket['status']
            
            if 'data_abertura' in formatted_ticket:  # Campo de data de criação
                formatted_ticket['data_abertura'] = format_date(formatted_ticket['data_abertura'])
            
            # Obter último comentário
            last_comment = get_last_comment(formatted_ticket['id'], headers)
            formatted_ticket['ultimo_comentario'] = last_comment

            formatted_tickets.append(formatted_ticket)

        if formatted_tickets:
            # Criar PDF com os chamados formatados
            pdf_path = create_pdf(formatted_tickets)
            # Enviar email com o PDF anexado
            send_email(email_subject, email_body, to_emails, pdf_path)
        else:
            # Enviar email informando que não há chamados pendentes
            no_tickets_body = 'Prezados, boa noite tudo bem? \n \n Não há chamados pendentes para o grupo técnico "Field Services".'
            send_email(email_subject, no_tickets_body, to_emails)

    else:
        response.raise_for_status()  # Levanta um erro para outros códigos de status HTTP

except requests.exceptions.RequestException as e:
    print(f"Erro ao buscar chamados: {e}")

finally:
    # Fechar sessão
    logout_url = f"{api_url}/killSession"
    requests.get(logout_url, headers=headers)
