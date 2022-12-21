import os, os.path
import win32com.client, pandas as pd
from config import *

outlook = win32com.client.Dispatch("Outlook.Application")
mapi = outlook.GetNamespace("MAPI")
your_folder = mapi.Folders.Item(2).Folders['Входящие'].Folders[folder_name]

path = os.path.expanduser(attachment_folder)

users_with_mails = {}
df = pd.DataFrame()
all_messages = 0
with_attachment = 0
attachment_without = 0 
attachment_errors = {}
unknown_errors = {}
for message in your_folder.Items:
    all_messages += 1

    
    try:
        excel_attach = False
        for attachment in message.Attachments:
            if '.xlsx' in attachment.DisplayName:
                users_with_mails.update({f'{message.Sender.GetExchangeUser().PrimarySmtpAddress}':f'{attachment}'})
                excel_attach = True
                attachment.SaveAsFile(os.path.join(path, f'{message.Sender}_{message.Sender.GetExchangeUser().PrimarySmtpAddress}.xlsx'))
            
        if excel_attach is not True:
            attachment_errors.update({message.Sender.GetExchangeUser().PrimarySmtpAddress:'нет xlsx вложения'})
            attachment_without += 1
        else:
            users_with_mails.update({f'{message.Sender.GetExchangeUser().PrimarySmtpAddress}': None})
            with_attachment += 1 
        print(f"{'-'*45}\nПисьмо {all_messages}\nemail: {message.Sender.GetExchangeUser().PrimarySmtpAddress}\nпроблемы: None")
    except AttributeError as er:
        print(f'''{'-'*45}\nПисьмо {all_messages}\nemail: {str(message.Body.encode('utf-16', 'ignore')).split('mailto:')[1].split('">')[0]}\nпроблемы: {message.Subject}''')
        attachment_without += 1 
        attachment_errors.update({f'''{str(message.Body.encode('utf-16', 'ignore')).split('mailto:')[1].split('">')[0]}''': message.Subject})
    except Exception as er:
        print(f"{'-'*45}\nПисьмо {all_messages}\nemail: None\nпроблемы: {er}")
        attachment_without += 1
        unknown_errors.update({f'message {all_messages}':er})
     
        
print(f'''{'#'*45}\nВсего сообщений: {all_messages}\nБез вложений: {attachment_without} ({round(attachment_without/all_messages * 100)}%)\nC вложениями: {with_attachment} ({round(with_attachment/all_messages * 100)}%)\n{'#'*45}''')

num = 0
print('Пользователи без вложений:')
if len(attachment_errors):
    for k, v in attachment_errors.items():
        num += 1
        print(f'{str(num)}) {v} - {k}')
else:
    print('Пусто')
print('#'*45)
num = 0
print('Пользователи c неизвестными ошибками:')
if len(unknown_errors):
    for k, v in unknown_errors.items():
        num += 1
        print(str(num)+') ', k,v)
else:
    print('Пусто')
print('#'*45)
