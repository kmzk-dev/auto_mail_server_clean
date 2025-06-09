import imaplib
import email
import re
from email.header import decode_header
from email.utils import getaddresses
from email.utils import parsedate_to_datetime

class MailManager:
    def __init__(self, server, port, account, password):
        self.server = server
        self.port = port
        self.account = account
        self.password = password
        self.mail_connection = None
        self.mail_ids = []

    def connect(self):
        """IMAPサーバーに接続"""
        try:
            self.mail_connection = imaplib.IMAP4_SSL(self.server, self.port)
            self.mail_connection.login(self.account, self.password)
            print(f"connected")
        except Exception as e:
            print(f"ERROR-connect: {e}")
            self.mail_connection = None
    
    def disconnect(self):
        """IMAP接続を切断する"""
        try:
            self.mail_connection.logout()
            print("disconnected")
        except Exception as e:
            print(f"ERROR-disconnect: {e}")

    def getIds(self,directory="Inbox"):
        """
        フォルダのすべてのメールIDを取得する
        """
        self.mail_ids = [] #再取得のための初期化
        try:
            self.mail_connection.select(directory, readonly=False)
            status, data = self.mail_connection.search(None, 'ALL')
            mail_ids = data[0].split()
            if mail_ids:
                print(f"Get {len(mail_ids)} Items In the {directory}")
                self.mail_ids = mail_ids
            else:
                print(f"Empty In the {directory}")
            return self.mail_ids
        except Exception as e :
            print(f"ERROR-getIds: {e}")
            return self.mail_ids
    
    def fetchHeaders(self):
        """
        各メールIDについて日時・送信者・送信先・件名を取得する
        """
        results = [["date","fromAdress","toName","toAddress","subject"]]
        for mail_id in self.mail_ids:
            if not mail_id:
                continue
            status, data = self.mail_connection.fetch(mail_id, '(RFC822.HEADER)')
            if status != 'OK':
                print(f"メールID {mail_id.decode()} のヘッダー取得に失敗")
                continue
            msg = email.message_from_bytes(data[0][1])

            def decode(val):
                if val is None:
                    return ""
                try:
                    decoded, charset = decode_header(val)[0]
                    if isinstance(decoded, bytes):
                        return decoded.decode(charset or "utf-8", errors="replace")
                    return decoded
                except:
                    return ""
            
            from_value = msg.get("From")
            from_name,from_addr = getaddresses([from_value])[0]

            to_value = msg.get("To")
            to_name,to_addr = getaddresses([to_value])[0]
            to_name = decode(to_name)
            if to_name == "":
                to_name = "設定なし"
            
            date_str = msg.get("Date")
            try:
                dt = parsedate_to_datetime(date_str)
                date_fmt = dt.strftime('%Y-%m-%d %H:%M')
            except Exception:
                date_fmt = date_str

            info = [
                date_fmt,
                from_addr,
                to_name,
                to_addr,
                decode(msg.get("Subject")),
            ]
            info_clean = [item.replace('\u3000', '') for item in info]
            #print(info_clean)
            results.append(info_clean)
        print("fetchHeaders did")    
        return results
    
    def markDeleteFlag(self):
        """
        self.mail_idsに格納されたメールに削除フラグを付与する
        """
        try:
            #self.mail_connection.select(directory, readonly=False)
            for mail_id in self.mail_ids:
                self.mail_connection.store(mail_id, '+FLAGS', '\\Deleted')
            print(f"Marked DeleteFlag")
        except Exception as e:
            print(f"ERROR-markDeleteFlag: {e}")
    
    def Delete(self):
        """
        削除フラグが付いたメールを完全に削除する
        """
        try:
            #self.mail_connection.select(directory, readonly=False)
            self.mail_connection.expunge()
            print(f"Expunged")
        except Exception as e:
            print(f"ERROR-expungeDeleted: {e}")

    def GetFolders(self):
        """
        サーバー内のフォルダ名（メールボックス名）を取得する
        """
        try:
            status, folders = self.mail_connection.list()
            folder_names = []
            print(folders)
            for folder in folders:
                decoded = folder.decode()
                parts = re.split(r'"[./\\,]"', decoded)
                # フォルダ名は "b'(<flags>) \"<delimiter>\" <folder name>'" の形式
                parts = folder.decode().split(' "." ')
                parts = parts[-1].replace('"', '')
                folder_names.append(parts)  
            return folder_names
        except Exception as e:
            print(f"ERROR-listFolders: {e}")
            return []



