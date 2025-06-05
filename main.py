from mail_manager_function import MailManager
import csv
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

load_dotenv() 

# IMAPサーバーの情報
IMAP_SERVER = os.environ.get('IMAP_SERVER')
IMAP_PORT = 993
ACCOUNT_NAME = os.environ.get('ACCOUNT_NAME')
PASSWORD = os.environ.get('PASSWORD')

mail = MailManager(IMAP_SERVER,IMAP_PORT,ACCOUNT_NAME,PASSWORD)
mail.connect()
directory = input("Tnter directory name")
mail.getIds(directory)
mail_list = mail.fetchHeaders()


now_str = (datetime.now() + timedelta(hours=9)).strftime('%Y%m%d_%H%M%S')
csv_file_path = f'output_{now_str}.csv'
with open(csv_file_path, 'w', newline='', encoding='utf-8') as csvfile:
    csv_writer = csv.writer(csvfile)
    csv_writer.writerows(mail_list)

isDelete = input("Do you delete them ? ( y / n )")
if isDelete == "y":
    mail.markDeleteFlag()
    mail.Delete()

mail.disconnect()