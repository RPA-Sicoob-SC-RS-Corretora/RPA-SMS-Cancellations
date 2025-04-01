from os import getenv
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
from src.variables import VARIABLES
import os
import win32com.client  # Biblioteca para integração com o Outlook

class MailConfig:
    def __init__(self) -> None:
        self._mails_worksheet = None
        self._smtp_server = "smtp.office365.com"
        self._smtp_port = 587
        self._smtp_user = getenv("OUTLOOK_USER")
        self._smtp_password = getenv("OUTLOOK_PASSWORD")
        self.message = f"""
            <html>
                <body>
                    <div style="font-family: Arial, Helvetica, sans-serif;">
                        <p>Prezados,</p>
                        <p></p>
                        <p>Encaminho em anexo o relatório de monitoramento referente aos envios de SMS de Cancelamento dos seus associados. O documento detalha os associados que foram notificados e aqueles que ainda não receberam a comunicação.</p>
                        <p></p>
                        <p>Em caso de dúvidas entrar em contato com <a href="mailto:pendencias@segurosicoob.com.br" data-linkindex="0">pendencias@segurosicoob.com.br</a></p>
                        <p><i>*essa é uma mensagem automática, favor não responder<i></p>
                        <p></p>
                        <h2>Gestão de Parcelas</h2>
                        <p>Sicoob Corretora SC/RS</p>
                        <p></p>
                        <p>R. Tenente Silveira, 94 - 7º andar</p>
                        <p>88010-300 | Florianópolis - SC</p>\
                        <p><strong>T 48 3085-9200 |</strong> <a href="https://sicoobsc.com.br">sicoobsc.com.br</a></p>
                    </div>
                </body>
            </html>
        """

    @property
    def mails_worksheet(self):
        return self._mails_worksheet
    
    @mails_worksheet.setter
    def mails_worksheet(self, worksheet):
        self._mails_worksheet = worksheet

    def send_email(self, recipients, cc, subject, body, attachment_path):
        msg = MIMEMultipart()
        msg['From'] = self._smtp_user
        msg['To'] = ", ".join(recipients)
        msg['Cc'] = ", ".join(cc)
        msg['Subject'] = subject

        msg.attach(MIMEText(body, 'html'))

        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(attachment_path)}")
            msg.attach(part)

        server = smtplib.SMTP(self._smtp_server, self._smtp_port)
        server.starttls()
        server.login(self._smtp_user, self._smtp_password)
        text = msg.as_string()
        server.sendmail(self._smtp_user, recipients + cc, text)
        server.quit()

    def send_email_with_outlook(self, recipients, cc, subject, body, attachment_path):
        """
        Envia um email utilizando o Outlook.

        :param recipients: Lista de destinatários.
        :param cc: Lista de emails em cópia.
        :param subject: Assunto do email.
        :param body: Corpo do email em HTML.
        :param attachment_path: Caminho do arquivo a ser anexado.
        """
        try:
            # Inicializar o Outlook
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)

            # Configurar os campos do email
            mail.To = "; ".join(recipients)
            mail.CC = "; ".join(cc)
            mail.Subject = subject
            mail.HTMLBody = body
            mail.SentOnBehalfOfName = "automacao@segurosicoob.com.br"  # Garantir que o email seja enviado deste endereço

            # Anexar o arquivo
            if os.path.exists(attachment_path):
                mail.Attachments.Add(attachment_path)
                print(f"Arquivo anexado: {attachment_path}")
            else:
                print(f"Arquivo não encontrado para anexar: {attachment_path}")

            # Enviar o email
            mail.Send()
            print(f"Email enviado com sucesso para: {', '.join(recipients)}")
        except Exception as e:
            print(f"Erro ao enviar email com Outlook: {e}")

    def get_recipients(self, cooperativa):
        """
        Lê qualquer arquivo XLSX na pasta Emails e retorna os destinatários e CC para a cooperativa especificada.

        :param cooperativa: Nome da cooperativa para buscar os emails.
        :return: Lista de destinatários e CC.
        """
        try:
            # Localizar qualquer arquivo XLSX na pasta Emails
            email_dir = os.path.join(VARIABLES["FILE_PATH"], "Emails")
            if not os.path.exists(email_dir):
                raise FileNotFoundError(f"Pasta de emails não encontrada: {email_dir}")

            email_files = [f for f in os.listdir(email_dir) if f.endswith(".xlsx")]
            if not email_files:
                raise FileNotFoundError(f"Nenhum arquivo XLSX encontrado na pasta: {email_dir}")

            # Usar o primeiro arquivo XLSX encontrado
            email_file_path = os.path.join(email_dir, email_files[0])
            print(f"Arquivo de emails encontrado: {email_file_path}")

            # Ler o arquivo XLSX
            df = pd.read_excel(email_file_path)
            row = df[df['PA'].str.strip() == cooperativa.strip()]
            if row.empty:
                raise Exception(f"Nenhuma informação de email encontrada para a cooperativa: {cooperativa}")

            recipients = row['E-mail'].dropna().tolist()
            cc = row['CC'].dropna().tolist()
            return recipients, cc
        except Exception as e:
            print(f"Erro ao obter destinatários: {e}")
            raise

    def get_recipients_by_filename(self, file_path):
        """
        Lê o arquivo XLSX na pasta Emails e retorna os destinatários e CC com base no nome do arquivo.

        :param file_path: Caminho da pasta onde o arquivo está localizado.
        :return: Lista de destinatários e CC.
        """
        try:
            # Localizar o arquivo XLSX na pasta Emails
            email_dir = os.path.join(file_path, "Emails")
            if not os.path.exists(email_dir):
                raise FileNotFoundError(f"Pasta de emails não encontrada: {email_dir}")

            email_files = [f for f in os.listdir(email_dir) if f.endswith(".xlsx")]
            if not email_files:
                raise FileNotFoundError(f"Nenhum arquivo XLSX encontrado na pasta: {email_dir}")

            email_file_path = os.path.join(email_dir, email_files[0])
            print(f"Arquivo de emails encontrado: {email_file_path}")

            # Ler o arquivo XLSX
            df = pd.read_excel(email_file_path)

            # Exibir o conteúdo do arquivo lido
            print("Conteúdo do arquivo XLSX lido:")
            print(df.head())  # Exibir as primeiras linhas do DataFrame

            # Validar se as colunas necessárias existem
            if "PA" not in df.columns or "E-mail" not in df.columns or "CC" not in df.columns:
                raise ValueError("O arquivo XLSX deve conter as colunas 'PA', 'E-mail' e 'CC'.")

            # Obter o nome do arquivo sem extensão
            files = [f for f in os.listdir(file_path) if f.endswith(".csv")]
            if not files:
                raise FileNotFoundError(f"Nenhum arquivo CSV encontrado na pasta: {file_path}")

            filename = os.path.splitext(files[0])[0].strip()
            print(f"Buscando informações para o arquivo: {filename}")

            # Localizar a linha correspondente ao nome do arquivo na coluna PA
            row = df[df['PA'].str.strip() == filename]
            if row.empty:
                raise Exception(f"Nenhuma informação de email encontrada para o arquivo: {filename}")

            # Obter destinatários e CC da linha correspondente
            recipients = row['E-mail'].dropna().tolist()
            cc = row['CC'].dropna().tolist()

            return recipients, cc
        except Exception as e:
            print(f"Erro ao obter destinatários e CC: {e}")
            raise

    def send_email_with_attachment(self, file_path):
        """
        Envia os arquivos gerados como anexo, usando os emails encontrados no relatório na pasta Emails.

        :param file_path: Caminho da pasta onde os arquivos gerados estão localizados.
        """
        try:
            # Localizar o relatório na pasta Emails
            email_dir = os.path.join(file_path, "Emails")
            if not os.path.exists(email_dir):
                raise FileNotFoundError(f"Pasta de emails não encontrada: {email_dir}")

            email_files = [f for f in os.listdir(email_dir) if f.endswith(".xlsx")]
            if not email_files:
                raise FileNotFoundError(f"Nenhum arquivo XLSX encontrado na pasta: {email_dir}")

            email_file_path = os.path.join(email_dir, email_files[0])
            print(f"Relatório de emails encontrado: {email_file_path}")

            # Ler o relatório de emails
            email_df = pd.read_excel(email_file_path)

            # Validar se as colunas necessárias existem
            if "PA" not in email_df.columns or "E-mail" not in email_df.columns or "CC" not in email_df.columns:
                raise ValueError("O relatório de emails deve conter as colunas 'PA', 'E-mail' e 'CC'.")

            # Localizar os arquivos gerados na pasta file_path
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Pasta não encontrada: {file_path}")

            files = [f for f in os.listdir(file_path) if f.endswith(".csv")]
            if not files:
                raise FileNotFoundError(f"Nenhum arquivo CSV encontrado na pasta: {file_path}")

            # Enviar cada arquivo gerado para os destinatários correspondentes
            for file in files:
                filename = os.path.splitext(file)[0].strip()
                attachment_path = os.path.join(file_path, file)
                print(f"Preparando envio do arquivo: {attachment_path}")

                # Localizar destinatários e CC com base no valor da coluna 'PA'
                row = email_df[email_df['PA'].str.strip() == filename]
                if row.empty:
                    print(f"Nenhuma informação de email encontrada para o arquivo: {filename}")
                    continue

                recipients = row['E-mail'].dropna().tolist()
                cc = row['CC'].dropna().tolist()

                # Configurar o email
                subject = f"Relatório de Monitoramento - {filename}"
                body = self.message

                # Enviar o email utilizando o Outlook
                self.send_email_with_outlook(recipients, cc, subject, body, attachment_path)
                print(f"Email enviado com sucesso para o arquivo: {filename}")
        except Exception as e:
            print(f"Erro ao enviar emails com anexos: {e}")