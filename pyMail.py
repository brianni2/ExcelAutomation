import email, smtplib, ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

class pyMail:
    def __init__(self, userEmail, password):
        self.userEmail = userEmail
        self.password = password
        
    def compose(self, subject, body, receiverEmail):
        self.message = MIMEMultipart()
        self.message["From"] = self.userEmail
        self.message["Subject"] = subject
        self.message["To"] = receiverEmail
        self.message.attach(MIMEText(body, "plain"))
        
    def attach(self, filename):
        for file in filename:
            with open(file, "rb") as attachment:
                # Add file as application/octet-stream
                # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename= {file}")
            self.message.attach(part)
        self.text = self.message.as_string()
        
    def send(self):
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(self.userEmail, self.password)
            server.sendmail(self.userEmail, self.message['To'], self.text)
