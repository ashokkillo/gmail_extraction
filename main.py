import numpy as np
import pandas as pd
import re
import imaplib, email
from tqdm import tqdm



user = 'user email' # Input your gmail address here
password = 'password' # Input your gmail password here
imap_url = 'imap.gmail.com'

my_mail = imaplib.IMAP4_SSL(imap_url) # connect to gmail
my_mail.login(user, password) # sign in with your credentials
print(my_mail.list()) # gives list of select options available
my_mail.select('ProHires') # select the folder that you want to retrieve

key = 'FROM' # As long as I retrieve email addresses from inbox, I select the key 'FROM'
value = '@' # It should contain text with "@"
_, data = my_mail.search(None, key, value)

mail_id_list = data[0].split()

msgs = []
for num in tqdm(mail_id_list):
    typ, data = my_mail.fetch(num, '(RFC822)')
    msgs.append(data)

subj_inbox_all = []
mails_inbox_all = []
for msg in msgs[::-1]:
    for response_part in msg:
        if type(response_part) is tuple and type(email.message_from_bytes((response_part[1]))) is email.message.Message:
            my_msg=email.message_from_bytes((response_part[1]))
            mails_inbox_all.append(my_msg['from'])
            subj_inbox_all.append(my_msg['subject'])

mail_inbox_all_clean = list(filter(lambda a: type(a) != email.header.Header, mails_inbox_all)) # Keep only the email address and not the name of the user
mails_inbox_all_ = np.unique(mail_inbox_all_clean) # keep all the unique emails

mails_in = []
for i in range(len(mails_inbox_all_)):
    mail = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', mails_inbox_all_[i])
    if mail is not None:
        mails_in.append(mail[0])
    else:
        mails_in.append(mails_inbox_all_[i])
pd.DataFrame(mails_in).to_excel('address_list1.xlsx')
print('Mails Extracted')
