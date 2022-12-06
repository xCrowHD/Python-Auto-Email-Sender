import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from email.mime.application import MIMEApplication
import os
import time
import pandas as pd
from dotenv import load_dotenv

corrected_names = {
    'Antenore Energia':'Antenore Energia SPA',
    'GE.PA SSDARL - 20 HOURS':'GE.PA - 20 Hours SRL',
    'Semplice Gas & Luce':'Semplice Gas & Luce SRL',
    'ANYTIMESS FITNESS':'Anytime Fitness SRL',
    'EGOSISTEMA':'Egosistema SPA',
    'PALESTRE ITALIANE':'Palestre Italiane SRL',
    'Energia Locale':'Energia Locale SRL'
}
def pause(massage = 'press any key to continue'):  # this function will pause the script with a default massage or a custome one.
    print(massage)
    os.system('pause')  # this will pause untill any key is pressed.

#getting the dir the script is on
my_dir = os.path.dirname(os.path.realpath(__file__))

enavars = r"{}".format(f'{my_dir}\Settingss\data.env')
load_dotenv(enavars)

sender_e = os.environ.get('sender_email')
sender_p = os.environ.get('sender_email_password')
pdf_dir = os.environ.get('pdf_dir')
csv_name = os.environ.get('csv_name')

#opening the csv
csv_string = r"{}".format(f'{my_dir}\Settingss\{csv_name}')
csv = pd.read_csv(csv_string, sep=';')
dt = pd.DataFrame(csv)
dt['Azienda Creditrice']=dt['Azienda Creditrice'].replace(corrected_names)
print(dt.head())

#attaches can be any files, can be a list
def email_constr(sender_email, receiver_email, subject, body, attaches):
    message = MIMEMultipart()
    message['From'] = formataddr(('Ufficio Crediti', f'{sender_email}'))
    message['To'] = receiver_email
    message['Subject'] = subject
    message.attach(MIMEText(body, 'html'))

    
    if attaches is not None:
          
        # Check whether we have the
        # lists of attachments or not!
        if type(attaches) is not list:
            
              # if it isn't a list, make it one
            attaches = [attaches]  
  
        for one_attachment in attaches:
  
            with open(one_attachment, 'rb') as f:
                
                # Read in the attachment
                # using MIMEApplication
                p = MIMEApplication(
                    f.read(),
                    name=os.path.basename(one_attachment)
                )
            p['Content-Disposition'] = f'attachment;\
            filename="{os.path.basename(one_attachment)}"'
              
            # At last, Add the attachment to our message object
            message.attach(p)
    
    return message


#use smtps with port
session = smtplib.SMTP('smtps.aruba.it', 587)
#enable security
session.starttls()
session.ehlo()
#login with mail_id and password
session.login(sender_e, sender_p)

#open the starting_point.txt file in read and write mode
start_p = open(r"{}".format(f'{my_dir}\Settingss\starting_point.txt'), 'r+')
i = int(start_p.read())
max = i + 50

#check if we would be outOfBounds
if max > dt.shape[0] - 1:
    max = dt.shape[0]
    print('CSV FINITO!')
    #here we update where we finished so next time we dont start from 0 but from max
    start_p.seek(0)
    start_p.write(str(0))
    start_p.truncate()
    start_p.close()
else:
    #here we update where we finished so next time we dont start from 0 but from max
    start_p.seek(0)
    start_p.write(str(max))
    start_p.truncate()
    start_p.close()


#sending the emails
while i < max:

    dt_row = dt.iloc[i]
    nb = str(dt_row['Nome Debitore'])
    email = str(dt_row['Email']).lower()
    nb_slipt = nb.split()

    nb_full_name = ''
    nb_surname = ''

    for item in nb_slipt:
        if item != nb_slipt[len(nb_slipt) - 1]:
            nb_full_name += item + ' '
            name = nb_full_name.lower().capitalize()
        else:
            nb_surname = item
            surname = nb_surname.lower().capitalize()
            
    
    pdf = dt_row['Codice Debitore']
    azienda = str(dt_row['Azienda Creditrice'])
    s = f'Pratica {azienda} {pdf}'
    
    html = f"""
<html>
  <head></head>
  <body>
    <p>Buongiorno {name}{surname}.</p>
    <p>In riferimento alla pratica {azienda} Vi invitiamo a prendere visione di quanto allegato.</p>
    <br>
    <p>I nostri uffici restano a Vostra disposizione dal lunedì al venerdì dalle 09:00 alle 18:00 al numero 0283595077</p>
    <br>
    <p>Distinti saluti</p>
    <p>Sistema Cliente Srl</p>
    <p>Tel: 0283595077</p>
  </body>
</html>
"""
    text = email_constr(sender_e, email, s, html, r"{}".format(f'{pdf_dir}\{pdf}.pdf'))
    session.sendmail(sender_e, email, text.as_string())
    i += 1



#closing the session 
session.quit()
print(f'Lo starting point é stato settato a : {max}')
print('Mails Sent')
pause()
