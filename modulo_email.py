import os
import ssl
import smtplib
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email import encoders


def sendEmail(nomeRemetente: str, destinatario: str, dataInicial, dataFinal, pathArquivo: str, nomeArquivo):
    try:
        email_sender = "" # Coloque aqui o email do destinatario
        email_password = os.environ.get("EMAIL_PASSWORD")
        email_receiver = destinatario

        subject = f"Notas Fiscais - {nomeRemetente}"
        body = f"""
        Este e-mail contém um arquivo em anexo em formato .zip contendo os XMLs das notas fiscais geradas do dia {dataInicial} ao dia {dataFinal}.
        """

        em = EmailMessage()
        em['from'] = email_sender
        em['to'] = email_receiver
        em['subject'] = subject
        em.set_content(body)

        em.add_alternative(body, subtype='html')

        with open(f'{pathArquivo}', 'rb') as attachment_file:
            file_data = attachment_file.read()
            file_name = f"{nomeArquivo}.zip"

        attachment = MIMEBase('application', 'octet-stream')
        attachment.set_payload(file_data)
        encoders.encode_base64(attachment)
        attachment.add_header('Content-Disposition', f'attachment; filename="{file_name}"')
        em.attach(attachment)

        context = ssl.create_default_context()

        with smtplib.SMTP_SSL('smtp.gmail.com', 465, context=context) as smtp:
            smtp.login(email_sender, email_password)
            smtp.sendmail(email_sender, email_receiver, em.as_string())

        print("\n")
        print("Email enviado com sucesso!")

    except Exception as e:
        print("\n")
        print(f"Erro ao enviar o email: {e}")