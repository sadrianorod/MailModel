import imaplib
import smtplib
from config_maildriver import *
import re

from utils import since_date,export_csv,export_excel,pd,io
from time import time
from datetime import datetime

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
import email

class MailModel:

    def __init__(self):
        self.__is_logged_in = False
        self.__email_adress = ''

    ####################################################
    #                                                  #
    #               PROTECTED METHODS                  #
    #                                                  #
    ####################################################

    def __check_extensions(self,filename):
        """
            param filename: Name of attachments
            param type: String
            return: True wheter the extension is correct
            rtype: Boolean
        """
        return filename.endswith(EXTENSIONS_ALLOWED_IN_ATTACHMENTS)

    def __check_is_ok(self,res):
        """
        param res: status response
        rtype: Boolean
        """ 
        return res == 'OK'

    def __parser_uid(self,data):
        match = re.compile('\\d+ \\(UID (?P<uid>\\d+)\\)').match(data)
        return match.group('uid')
    
    
    def __login_mail(self,adress,password,IMAP_SERVER,IMAP_PORT):
        """
            :param email: storage email.
            :type email:  string.
            :param email: storage password.
            :type email:  string.
        """
        MAX_LOGIN_ATTEMPTS = 3
        self.__email_adress = adress
        login_attempts = 0 
        
        while True:
            try:
                self.imap = imaplib.IMAP4_SSL(IMAP_SERVER,IMAP_PORT)
                resp, _ = self.imap.login(adress, password)
                if self.__check_is_ok(resp):
                    self.__is_logged_in = True
                    return self.__is_logged_in

            except Exception:
                login_attempts += 1
                if login_attempts < MAX_LOGIN_ATTEMPTS:
                    continue
                return self.__is_logged_in
    
    def __get_attachments(self,msg):
        """
            Internal Function to get Attachments from e-mails
            param msg: E-mail message
            ptype msg: type Message
            
            param save_attachments: Boolean which decides wheter will
            save or not file in path
            ptype save_attachments: boolean
            
            param path: Complete path where the file will be saved
            ptype path: String
            
            return: list of attachments where extensions are allowed
            rtype: list
        """
        attach_list = []
        for part in msg.walk():
            if part.get_content_maintype()=="multipart":
                continue
            if part.get('Content-Disposition') is None:
                continue
            if part.get_filename() is not None and self.__check_extensions(part.get_filename()):
                raw_data = part.get_payload(decode=True)
                filename = part.get_filename()
                if '.xlsx' in filename: df = pd.read_excel(raw_data,index_col=False)
                
                else : 
                    raw_data = raw_data.decode()

                    df = pd.read_csv(io.StringIO(raw_data),sep=";",index_col=False)    
                
                attach_list.append((part.get_filename(),df))

        return attach_list

    def __append_mail(self,path_mail_folder,msg):
        """
            Insert a object Message into mailBox named name_mail_box
           
            param name_mail_box: Name of mail box
            ptype name_mail_box: string
           
            param msg: Object Message with Subject, From, Body, Date, Attachments
            ptype msg: Object Message
        """
        self.imap.append(path_mail_folder,'\\Draft',imaplib.Time2Internaldate(time()),msg.as_bytes())

    def __search(self,key,value,path_mail_folder,return_only_list_ids = False):
        """
            Search for a key and a value 

            param key: Type of search
            ptype key: String
            
            param value: Value of you are searching for
            ptype value: String

            param path_mail_folder: Where you will do this search
            ptype path_mail_folder: String

            param  return_only_list_ids: type of return you will expect 
            ptype  return_only_list_ids: Boolean

            return: response status and list of messages
            rtype: Boolean,List[Tuple(ID,Raw Messages)]
        """
        try:
            self.imap.select(path_mail_folder)
            res, binary_list_emails_ids = self.imap.search(None,key,'"{}"'.format(value))
            binary_list_emails_ids = binary_list_emails_ids[0].split()
            if self.__check_is_ok(res):
                if not return_only_list_ids: 
                    
                    msgs = []
                    
                    for id in binary_list_emails_ids:
                        _, raw_data = self.imap.fetch(id,'(RFC822)')
                        msgs.append(self.__convert_message_to_dict(id,raw_data))

                    return msgs                        
                
                else : 

                    return binary_list_emails_ids
        except:
            raise ConnectionError("Error to search e-mails")

    def __convert_message_to_dict(self,id,raw_email):
        """
            Process email and returns a dict{ID,FROM,DATE,TIME,SUBJECT,PATH,ATTACHMENTS_LIST}
            
            param id: uid from email
            ptype id: bynary
            param raw_email: Email in binary format
            ptype raw_email: Binary
            return Dictionary with all relevant informations
            rtype: Dict

        """
        msg = email.message_from_bytes(raw_email[0][1])
        try:
            date_string = datetime.strptime(msg['Date'],"%d %b %Y %H:%M:%S %z").strftime(FORMAT_DATE)
            time_string = datetime.strptime(msg['Date'],"%d %b %Y %H:%M:%S %z").strftime(FORMAT_TIME)
        except ValueError:
            try:
                date_string = datetime.strptime(msg['Date'],"%a, %d %b %Y %H:%M:%S %z").strftime(FORMAT_DATE)
                time_string = datetime.strptime(msg['Date'],"%a, %d %b %Y %H:%M:%S %z").strftime(FORMAT_TIME)
            except ValueError:
                date_string = datetime.strptime(msg['Date'],"%a, %d %b %Y %H:%M:%S %z (UTC)").strftime(FORMAT_DATE)
                time_string = datetime.strptime(msg['Date'],"%a, %d %b %Y %H:%M:%S %z (UTC)").strftime(FORMAT_TIME)
        finally:
            attachments = self.__get_attachments(msg)
            name,address_email =  email.utils.parseaddr(msg['From'])
            return {
                CONVERSION_DICT['ID']  : id,
                CONVERSION_DICT['FROM']: address_email,
                CONVERSION_DICT['NAME']: name,
                CONVERSION_DICT['DATE']: date_string,
                CONVERSION_DICT['TIME']: time_string,
                CONVERSION_DICT['SUBJECT']: msg['SUBJECT'],
                CONVERSION_DICT['PATH']: msg['RETURN-PATH'],
                CONVERSION_DICT['ATTACHMENTS_LIST']: attachments
            }

    def __send(self, msg, SMTP_SERVER,SMTP_PORT,STORAGE_EMAIL,STORAGE_PASSWORD):
         
        try:
            server = smtplib.SMTP(SMTP_SERVER,SMTP_PORT)
            server.ehlo()
            server.starttls()
            server.login(STORAGE_EMAIL,STORAGE_PASSWORD)
            text = msg.as_string()
            server.sendmail(self.get_email(),msg['TO'], text)
            server.quit()
            return True

        except:
            raise ConnectionError("Error to send e-mail")


    ####################################################
    #                                                  #
    #                PUBLIC  METHODS                   #
    #                                                  #
    ####################################################

    def search_for(self,email_address,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        return self.__search('FROM', email_address,path_folder,return_only_list_ids)
    
    def search_all_emails_since(self, days,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        """
            Get all emails since n days before, considering today = 0
            param days: Number days before
            ptype days: Int
            param id_search: If True search return the ids from all emails 
                             Else search return the raw data in binary from all emails
            rtype : Boolean, List [Ids or Raw Data]
        """
        key = '(SINCE "'+since_date(days)+'")'
        value = 'ALL'
        return self.__search(key,value,path_folder,return_only_list_ids)

    def search_all_emails_today(self,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        return self.search_all_emails_since(0,path_folder,return_only_list_ids)

    def search_all_read_emails_since(self, days,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        key = '(SINCE "'+since_date(days)+'")'
        value = 'SEEN'
        return self.__search(key,value,path_folder,return_only_list_ids)

    def search_all_read_emails_today(self,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        return self.search_all_read_emails_since(0,path_folder,return_only_list_ids)

    def search_all_unread_emails_since(self, days,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        key = '(SINCE "'+since_date(days)+'")'
        value = 'UNSEEN'
        return self.__search(key,value,path_folder,return_only_list_ids)

    def search_all_unread_emails_today(self,path_folder=SEARCH_DEFAULT_PATH_FOLDER,return_only_list_ids = False):
        return self.search_all_unread_emails_since(0,path_folder,return_only_list_ids)
    
    def create_mail(self,recipient,subject,Cc='',body_message=('','plain'),attachments=[],path_folder='',change_format = False):
        """
            Create a mail with infos passed by arguments and return a message object
        """
        
        msg = MIMEMultipart()
        msg['From'] = self.__email_adress
        msg['To'] = recipient
        msg['Cc'] = Cc
        msg['Subject'] = subject
        msg['Date'] = datetime.now().strftime("%d/%m/%Y %H:%M")
        msg.attach(MIMEText(body_message[0],body_message[1]))
        
        if attachments:
            for (filename,df) in attachments:
                if filename.endswith('.xlsx'):
                    attachment = export_excel(df,change_format)
                    part = MIMEBase('application', "vnd.ms-excel")
                    part.set_payload(attachment)
                else:
                    attachment = export_csv(df)
                    part = MIMEBase('application', "octet-stream")
                    part.set_payload(attachment)

                email.encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment', filename=filename)                
                msg.attach(part)
        
        if path_folder:
            self.__append_mail(path_folder,msg)
        
        return msg
    
    def move(self,path_start_folder,path_final_folder,list_emails_ids,delete_after_move=False):
        self.imap.select(path_start_folder,readonly=False)
        res_fetch = False
        res_copy = False
        for email_id in list_emails_ids:
            res_fetch, data_fetch = self.imap.fetch(email_id,"(UID)")
            msg_uid = self.__parser_uid(data_fetch[0].decode("utf-8"))
            res_copy, _ = self.imap.uid('COPY',msg_uid,path_final_folder)
            if self.__check_is_ok(res_fetch) and self.__check_is_ok(res_copy):
                if delete_after_move:
                    res, _ = self.imap.uid('STORE', msg_uid , '+FLAGS', '(\\Deleted)')
                    self.imap.expunge()
                    return self.__check_is_ok(res)
            
        return self.__check_is_ok(res_fetch) and self.__check_is_ok(res_copy)
    
    def delete_all_emails_since(self,days,path_folder):
        res,list_emails_ids = self.search_all_emails_since(days,path_folder,return_only_list_ids=True)
        if res:
            for email_id in list_emails_ids:
                res_fetch, data_fetch = self.imap.fetch(email_id,"(UID)")
                msg_uid = self.__parser_uid(data_fetch[0].decode("utf-8"))
                if self.__check_is_ok(res_fetch):
                    _ = self.imap.uid('STORE', msg_uid , '+FLAGS', '(\\Deleted)')
            
            self.imap.expunge()

    def delete_all_emails_today(self,path_folder):
        self.delete_all_emails_since(0,path_folder)
    
    def delete_all_read_emails_since(self,days,path_folder):
        res,list_emails_ids = self.search_all_read_emails_since(days,path_folder,return_only_list_ids=True)
        if res:
            for email_id in list_emails_ids:
                res_fetch, data_fetch = self.imap.fetch(email_id,"(UID)")
                msg_uid = self.__parser_uid(data_fetch[0].decode("utf-8"))
                if self.__check_is_ok(res_fetch):
                    _ = self.imap.uid('STORE', msg_uid , '+FLAGS', '(\\Deleted)')
            
            self.imap.expunge()

    def delete_all_read_emails_today(self,path_folder):
        self.delete_all_read_emails_since(0,path_folder)


    def logout(self):
        """
            return imap.logout(): ('BYE', [b'Microsoft Exchange Server IMAP4 server signing off.'])
            rtype: Boolean, bytes
        """
        res, _ = self.imap.logout()
        self.__is_logged_in = (res == 'BYE')
        return self.__is_logged_in