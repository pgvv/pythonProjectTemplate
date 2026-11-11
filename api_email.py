# --------------------------- IMPORTS ---------------------------#
import os
import re
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from dotenv import load_dotenv

# --------------------------- CONSTANTS ---------------------------#
PATH_ENV = "api_keys.env"


# --------------------------- FUNCTIONS ---------------------------#
class Gmail:
    """GMail using SMTP and MIME"""

    def __init__(self):
        load_dotenv(PATH_ENV)
        self.port_no = 587
        self.gmail_server = "smtp.gmail.com"

    def send_gmail(
        self,
        sender: str,
        mail_body: str,
        subject: str = "",
        receiver_mail: str = "",
    ):
        """
        Send mail using Gmail\n
        Default receiver mail is sender mail address\n
        **mail_body:** Subject: subject_name \\n\\n body_message
        """
        if not receiver_mail:
            receiver_mail = sender

        # Check addresses
        s_pattern = r"^[a-zA-Z0-9_.+-]+@gmail\.com$"
        if re.match(s_pattern, sender) is None:
            print("Invalid sender address. Address must be ending with '@gmail.com'.")
            return
        r_pattern = r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        if re.match(r_pattern, receiver_mail) is None:
            print("Invalid receiver address. Enter correct email address.")
            return

        # setting multipart message
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = receiver_mail
        msg["Subject"] = subject

        msg.attach(MIMEText(mail_body, "html"))

        with smtplib.SMTP(host=self.gmail_server, port=self.port_no) as connection:
            # TLS: Transport Layer Security, secures connection to email server by encrypting the message
            connection.starttls()
            try:
                connection.login(user=sender, password=os.getenv("GMAIL_PERSONAL", ""))
            except smtplib.SMTPAuthenticationError as e:
                print(f"Incorrect sender email/password: {e}")
            else:
                connection.sendmail(
                    from_addr=sender,
                    to_addrs=receiver_mail,
                    msg=msg.as_string(),
                )


# --------------------------- MAIN CODE ---------------------------#
if __name__ == "__main__":
    email = Gmail()

    body = "<h1>This is a test mail</h1>"
    email.send_gmail(mail_body=body, sender="pgvvishal31@gmail.com", subject="Python Test Subject")
