import imaplib
import email
import os

from email.utils import parsedate
from datetime import datetime
from dateutil import relativedelta

def save_mail(login,passw,srv,mailbox):
    # mailbox = "INBOX.Inviata" per Inviata e "INBOX" per Arrivata
    # caratteri da cancellare dai nomi dei file
    restr_cars = ['\\','/',':','*','?','>','<','|','"','\t','\r','\n','Per conto di ']
    
    save_directory = os.path.join(login,mailbox)
    
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)
        
    mail = imaplib.IMAP4_SSL(srv)
    mail.login(login,passw)
    
    #mail.list() to see the available folders
    
    mail.select(mailbox)
    
    result, data = mail.uid('search',None, 'ALL')
    ids = data[0] # data is a list.
    id_list = ids.split() # ids is a space separated string
    id_list.reverse()
    
    for ids in id_list:
        result, data = mail.uid('fetch',ids, '(BODY.PEEK[HEADER])')
        # (BODY.PEEK[HEADER]) to fetch just the headers, (RFC822) to fetch everything 
        # here's the body, which is raw text of the whole email
        raw_email = data[0][1]
        # including headers and alternate payloads
        msg = email.message_from_string(raw_email)
        
        msg_from = msg['from']
        msg_to = msg['to']
        msg_data = msg['date']
        
        msg_from = ''.join(x for x in msg_from if x not in restr_cars)
        msg_from = msg_from.replace('Per conto di ', '')
        msg_to = ''.join(x for x in msg_to if x not in restr_cars)
        
        message_file_name = os.path.join(save_directory,'ID{0} FROM {1} TO {2} DATA {3}.eml'.format(ids,msg_from,msg_to,msg_data.replace(':','')))
        
        if not os.path.exists(message_file_name):
            file_status = "downloading..."
            result, data = mail.uid('fetch',ids, '(RFC822)')
            raw_email = data[0][1]
            msg = email.message_from_string(raw_email)
            open(message_file_name,'wb+').write(msg.as_string())
        else:
            file_status = "exists!!!"
            
        print('[{0}]: file "{1}" {2}'.format(datetime.today().strftime('%Y-%m-%d %H:%M:%S'),message_file_name,file_status))
        
    print('END.')    
    mail.logout()
    
def delete_mail(login,passw,srv,mailbox,months_old):
    mail = imaplib.IMAP4_SSL(srv)
    mail.login(login,passw)
    mail.select(mailbox)
    result, data = mail.uid('search',None, 'ALL')
    ids = data[0] # data is a list.
    id_list = ids.split() # ids is a space separated string
    id_list.reverse()
    
    for ids in id_list:
        result, data = mail.uid('fetch',ids, '(BODY.PEEK[HEADER])') # (BODY.PEEK[HEADER]) to fetch just the headers, (RFC822) to fetch everything 
        # here's the body, which is raw text of the whole email
        raw_email = data[0][1]
        # including headers and alternate payloads
        msg = email.message_from_string(raw_email)
        msg_data = msg['date']
        time_tuple=parsedate(msg['date'])
        dt_obj = datetime(*time_tuple[0:7])
        r = relativedelta.relativedelta(datetime.today(), dt_obj)
        if(r.months >= months_old):
            print('[{0}] DELETED: ID={1}, date={2} '.format(datetime.today().strftime('%Y-%m-%d %H:%M:%S'),ids,msg['date']))
            mail.uid('STORE', ids , '+FLAGS', '(\Deleted)')
    result = mail.expunge() 
    print result
    mail.logout()

# Main functionality    
srv = 'imaps.pec.aruba.it'
logins = ['fonditalia@pec.fonditalia.org','avvisolombardiappss@pec.fonditalia.org','rendicontazione@pec.formasicuro.it']
passws = ['fondipec8','ppsslomb','rendicontazione']
mailboxes = ['INBOX' , 'INBOX.Inviata']

#'avvisolombardia@pec.fonditalia.org','avvlomb',

for i in range(len(logins)):
    for mailbox in mailboxes:
        save_mail(logins[i],passws[i],srv,mailbox)

for i in range(len(logins)):
    delete_mail(logins[i],passws[i],srv,"INBOX",2)
for i in range(len(logins)):
    delete_mail(logins[i],passws[i],srv,"INBOX.Inviata",6)


