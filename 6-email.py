import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

#CONEXÃO COM O E-MAIL
password = open("senha", "r").read()
from_email = "grecoteste@gmail.com"
to_email = "grecoteste@gmail.com"
subject = "Automação Planilha"
body = """"
Olá. Segue em anexo a automação da planilha
para a empresa XYZ Automação.

Qualquer dúvida estou a disposição!
"""

message = EmailMessage()
message["From"] = from_email
message["To"] = to_email
message["Subject"] = subject

message.set_content(body)
safe = ssl.create_default_context()


# ADD A BASE DE DADOS COMO ANEXO
anexo = "teste.xlsx"
# print(mimetypes.guess_type(anexo)[0].slipt("/"))
mime_type, mime_subtype = mimetypes.guess_type(anexo)[0].slipt("/")
with open(anexo, "rb") as a:
    message.add_attachment(
        a.read(),
        maintype=mime_type,
        subtype=mime_subtype,
        filename=anexo
    )

# ENVIO O E-MAILZÃO
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=safe) as stmp:
    stmp.login(from_email, password)
    stmp.sendmail(
        from_email,
        to_email,
        message.as_string()
    )