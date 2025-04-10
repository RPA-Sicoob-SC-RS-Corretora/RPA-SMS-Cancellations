import pandas as pd
import os
from src.variables import VARIABLES
from src.services.ConfigEmail import MailConfig  # Caminho relativo correto

class MailConfigExtended(MailConfig):
    def get_recipients(self, cooperativa):
        """
        Lê qualquer arquivo XLSX na pasta Emails e retorna os destinatários e CC para a cooperativa especificada.
        """
        try:
            email_dir = os.path.join(VARIABLES["FILE_PATH"], "Emails")
            if not os.path.exists(email_dir):
                raise FileNotFoundError(f"Pasta de emails não encontrada: {email_dir}")

            email_files = [f for f in os.listdir(email_dir) if f.endswith(".xlsx")]
            if not email_files:
                raise FileNotFoundError(f"Nenhum arquivo XLSX encontrado na pasta: {email_dir}")

            email_file_path = os.path.join(email_dir, email_files[0])

            df = pd.read_excel(email_file_path)
            row = df[df['PA'].str.strip() == cooperativa.strip()]
            if row.empty:
                raise Exception(f"Nenhuma informação de email encontrada para a cooperativa: {cooperativa}")

            recipients = row['E-mail'].dropna().tolist()
            cc = row['CC'].dropna().tolist()
            return recipients, cc
        except Exception as e:
            raise

    def get_recipients_by_filename(self, file_path):
        """
        Lê o arquivo XLSX na pasta Emails e retorna os destinatários e CC com base no nome do arquivo.
        """
        try:
            email_dir = os.path.join(file_path, "Emails")
            if not os.path.exists(email_dir):
                raise FileNotFoundError(f"Pasta de emails não encontrada: {email_dir}")

            email_files = [f for f in os.listdir(email_dir) if f.endswith(".xlsx")]
            if not email_files:
                raise FileNotFoundError(f"Nenhum arquivo XLSX encontrado na pasta: {email_dir}")

            email_file_path = os.path.join(email_dir, email_files[0])

            df = pd.read_excel(email_file_path)

            if "PA" not in df.columns or "E-mail" not in df.columns or "CC" not in df.columns:
                raise ValueError("O arquivo XLSX deve conter as colunas 'PA', 'E-mail' e 'CC'.")

            files = [f for f in os.listdir(file_path) if f.endswith(".csv")]
            if not files:
                raise FileNotFoundError(f"Nenhum arquivo CSV encontrado na pasta: {file_path}")

            filename = os.path.splitext(files[0])[0].strip()

            row = df[df['PA'].str.strip() == filename]
            if row.empty:
                raise Exception(f"Nenhuma informação de email encontrada para o arquivo: {filename}")

            recipients = row['E-mail'].dropna().tolist()
            cc = row['CC'].dropna().tolist()

            return recipients, cc
        except Exception as e:
            raise

    def send_email_with_attachment(self, file_path):
        """
        Envia os arquivos gerados como anexo, usando os emails encontrados no relatório na pasta Emails.
        """
        try:
            email_dir = os.path.join(file_path, "Emails")
            if not os.path.exists(email_dir):
                raise FileNotFoundError(f"Pasta de emails não encontrada: {email_dir}")

            email_files = [f for f in os.listdir(email_dir) if f.endswith(".xlsx")]
            if not email_files:
                raise FileNotFoundError(f"Nenhum arquivo XLSX encontrado na pasta: {email_dir}")

            email_file_path = os.path.join(email_dir, email_files[0])

            email_df = pd.read_excel(email_file_path)

            if "PA" not in email_df.columns or "E-mail" not in email_df.columns or "CC" not in email_df.columns:
                raise ValueError("O relatório de emails deve conter as colunas 'PA', 'E-mail' e 'CC'.")

            if not os.path.exists(file_path):
                raise FileNotFoundError(f"Pasta não encontrada: {file_path}")

            files = [f for f in os.listdir(file_path) if f.endswith(".csv")]
            if not files:
                raise FileNotFoundError(f"Nenhum arquivo CSV encontrado na pasta: {file_path}")

            for file in files:
                filename = os.path.splitext(file)[0].strip()
                attachment_path = os.path.join(file_path, file)

                row = email_df[email_df['PA'].str.strip() == filename]
                if row.empty:
                    continue

                recipients = row['E-mail'].dropna().tolist()
                cc = row['CC'].dropna().tolist()

                subject = f"Relatório de Monitoramento - Cancelamento "
                body = self.message

                self.send_email(
                    recipients, cc, subject, body, attachment_path
                )
        except Exception as e:
            raise