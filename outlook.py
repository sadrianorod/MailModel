from mailmodel import MailModel
from config_maildriver import STORAGE_EMAIL,STORAGE_PASSWORD

class Outlook(MailModel):
    
    def __init__(self):
        super().__init__()
        self.IMAP_SERVER = "outlook.office365.com"
        self.IMAP_PORT = 993
        self.SMTP_SERVER = "smtp.office365.com"
        self.SMTP_PORT = 587

    def login(self):
        return self.__login_mail(STORAGE_EMAIL,STORAGE_PASSWORD,self.IMAP_SERVER,self.IMAP_PORT)

    def send(self,msg):
        self.__send(msg,self.SMTP_SERVER,self.SMTP_PORT,STORAGE_EMAIL,STORAGE_PASSWORD)