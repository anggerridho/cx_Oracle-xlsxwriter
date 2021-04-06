#!/usr/bin/env python
# coding: utf-8

# In[1]:


import cx_Oracle
import os
import datetime
import csv
import paramiko
import pandas as pd
import xlsxwriter
import requests # Untuk Curl
from xlsxwriter.utility import xl_rowcol_to_cell

now = datetime.date.today()
kemarin = now - datetime.timedelta(days=1)
kemarin_str = str(kemarin.strftime('%Y%m%d'))
dir = '/home/palugada/Daily/'
TOKEN="1273628193:AAGfFDAO2res2DeaGHOLVxLFxz25E06LYrY"
CHAT_ID="214749655"

os.environ["ORACLE_HOME"] = "/usr/lib/oracle/12.2/client64"
# workbook = xlsxwriter.Workbook('/home/palugada/Daily/Demo.xlsx')
workbook = xlsxwriter.Workbook(dir+"RBT_Daily_Report_"+str(kemarin_str)+".xlsx",{'strings_to_numbers': True})
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect('192.168.0.7', 1115, username='trbt', password='.pteluon')
database = cx_Oracle.connect('RBTRPTN/RBTRPTN@192.168.0.7:1521/RBTRPTN')

def fyi(*text):
    import requests
    TOKEN="1273628193:AAGfFDAO2res2DeaGHOLVxLFxz25E06LYrY"
    CHAT_ID="214749655"
    CHID="-1001475662404"
    API_ENDPOINT = "https://api.telegram.org/bot"+ TOKEN +"/sendMessage"
    data = {'chat_id':CHID,'text':text}
    requests.post(url = API_ENDPOINT, data = data)

bold = workbook.add_format({'bold': True, 'bg_color': '#92D050', 'align': 'center', 'valign': 'vcenter', 'border': 1})
bold.set_border()
bold.set_text_wrap()
neat = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
neat.set_border()
neat.set_text_wrap()
percentage = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
percentage.set_border()
percentage.set_num_format(10)
percentage.set_text_wrap()
content = workbook.add_format({'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'border': 1})
content.set_num_format('#,##0')
content.set_border()
date_format = workbook.add_format({'font_color': 'black', 'align': 'center', 'valign': 'vcenter', 'border': 1})
date_format.set_num_format(0)
date_format.set_border()
title = workbook.add_format({'bold': True})

# Create worksheet summary
summary = workbook.add_worksheet('SUMMARY')
new_subs = workbook.add_worksheet('NEW_SUBS')
prov = workbook.add_worksheet('PROV')
totaltraf = workbook.add_worksheet('TOTAL TRAFFIC')
pricerevenue = workbook.add_worksheet('PRICE & REVENUE')
chargingcdr = workbook.add_worksheet('CHARGING CDR')
errorsubs = workbook.add_worksheet('ERROR SUBS')
httpgw = workbook.add_worksheet('HTTPGW')
smsconfirm = workbook.add_worksheet('SMS_CONFIRMATION')
songtop = workbook.add_worksheet('SONG TOP')
revenue = workbook.add_worksheet('REVENUE')
melondiy = workbook.add_worksheet('MELON DIY')
melonmynsp = workbook.add_worksheet('MELON MYNSP')
suspendtop20 = workbook.add_worksheet('SUSPEND_TOP20')
region = workbook.add_worksheet('REGION')
dtmf = workbook.add_worksheet('DTMF')
gp = workbook.add_worksheet('GRACE PERIODE')
smsrev = workbook.add_worksheet('SMS REVENUE')
revrenew = workbook.add_worksheet('REVENUE RENEWAL')
nspreco = workbook.add_worksheet('NSP REKOMENDASI')
succratemt = workbook.add_worksheet('SUCCESS RATE MT AND MO')
transumb = workbook.add_worksheet('TRANSACTION_UMB_TMENU_PROMO')

for TabCon in ['summary','new_subs','prov','totaltraf','pricerevenue','chargingcdr','errorsubs','httpgw','smsconfirm','songtop','revenue','melondiy','melonmynsp','suspendtop20','region','dtmf','gp','smsrev','revrenew','nspreco','succratemt','transumb']:
    TabCol = TabCon + ".set_tab_color('#00B050')"
    exec(TabCol)


# In[2]:


# Send Email:
def SendEmail():
    import email, smtplib, ssl

    from email import encoders
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

#     subject = "An email with attachment from Python"
    Yesterday = str(kemarin.strftime('for %B %d, %Y'))
    subject = "(Testing) RBT Daily Report " + Yesterday
    body = "Dear All,\nPlease kindly find RBT Daily Reports " + Yesterday + " in attachment.\n\nBest Regards,\nAngger Ridho (https://linktr.ee/anggerdho)\nNote: xlsx & this email message is an automation process that I created"
    sender_email = "Angger Ridho <angger@eluon.com>"
    receiver_email = ['angger@eluon.com', 'rizky.fauzi@eluon.com', 'randy@eluon.com', 'dhifa@eluon.com', 'handi@eluon.com', 'diyas@eluon.com',                       'operation.support@eluon.com','amanda@eluon.com']
#     password = input("Type your password and press enter:")
    password = "@nG9er28"

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(receiver_email)
    message["Subject"] = subject
#     message["Bcc"] = receiver_email  # Recommended for mass emails

    # Add body to email
    message.attach(MIMEText(body, "plain"))

#     f = ["RBT_Daily_Report_" + str(kemarin_str) + ".xlsx", \
#          "New_RBT_Daily_Report_" + str(kemarin_str) + ".xlsx"]  # In same directory as script
    f = ["RBT_Daily_Report_" + str(kemarin_str) + ".xlsx"]

    # Open PDF file in binary mode
    for filename in f:  # add files to the message
        with open(filename, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f"attachment; filename= {filename}",
            )
            message.attach(part)
            text = message.as_string()

    # Log in to server using secure context and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("gw.eluon.com", 465, context=context) as server:
        server.login("angger@eluon.com", "@nG9er28")
        server.sendmail(sender_email, receiver_email, text)
        chat = 'RBT Daily Report ' + str(kemarin_str) + 'was just emailed at '+datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        fyi(chat)


# In[3]:


# A All KPI
def A_KPI():
    cursor = database.cursor()
    with open('A_ALL_KPI.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()
    
    summary.set_zoom(70)
    summary.hide_gridlines(2)
    summary.write('A1','A All KPI',title)
    summary.merge_range('B3:B4','Calling Date',bold)
    summary.merge_range('C3:M3','User InterFace',bold)
    summary.write('C4','SMS',bold)
    summary.write('D4', 'UMB',bold)
    summary.write('E4', 'IVR',bold)
    summary.write('F4', 'HTTPGW',bold)
    summary.write('G4', 'WAP',bold)
    summary.write('H4', 'WEB',bold)
    summary.write('I4', 'OPENAPI',bold)
    summary.write('J4', 'MKIOS',bold)
    summary.write('K4', 'WCC',bold)
    summary.write('L4', 'CSWEB',bold)
    summary.write('M4', 'RENEWAL',bold)
    summary.merge_range('N3:N4','Total Purchase Request Received by RBT Eluon',bold)
    summary.set_column(1,12, 12)
    summary.set_column('N:N', 20) # 13
    summary.merge_range('O3:P3','RBT Eluon Request to INGW',bold)
    summary.write('O4', 'Charge',bold)
    summary.write('P4', 'Refund',bold)
    summary.set_column(14,15, 13)
    summary.merge_range('Q3:S3','Transaction Response from INGW',bold)
    summary.write('Q4', 'Success Charging Transaction',bold)
    summary.write('R4', 'Charging Time Out (802)',bold)
    summary.write('S4', 'Charging Transaction Success Rate',bold)
    summary.set_column(16,18, 12)
    summary.merge_range('T3:X3','Charging Related',bold)
    summary.write('T4', 'Success Charged/Billable',bold)
    summary.write('U4', 'Success Charged Rate',bold)
    summary.write('V4', 'Low Balance',bold)
    summary.write('W4', 'Subscriber Locked',bold)
    summary.write('X4', 'Unknown Error Code (502)',bold)
    summary.set_column(19,23, 12)
    summary.merge_range('Y3:Z3','Refund Process',bold)
    summary.write('Y4', 'Success Refund',bold)
    summary.write('Z4', 'Refund Success Rate',bold)
    summary.set_column(24,25, 12)
    summary.merge_range('AA3:AB3','RBT Eluon Request to ProvGW',bold)
    summary.write('AA4', 'Activation',bold)
    summary.write('AB4', 'Deactivation',bold)
    summary.set_column(26,27, 12)
    summary.merge_range('AC3:AG3','Activation Responses from ProvGW',bold)
    summary.write('AC4', 'Success',bold)
    summary.write('AD4', 'Prov Error',bold)
    summary.write('AE4', 'Time Out',bold)
    summary.set_column(28,30, 12)
    summary.write('AF4', 'Provisioning Transaction Success Rate',bold)
    summary.set_column('AF:AF', 15)
    summary.set_row(3, 50)
    summary.write('AG4', 'Provisioning Success Rate',bold)
    summary.set_column('AG:AG', 13)
    summary.merge_range('AH3:AJ3','Deactivation Responses from ProvGW',bold)
    summary.write('AH4', 'Success',bold)
    summary.write('AI4', 'Failed',bold)
    summary.write('AJ4', 'Success Rate',bold)
    summary.set_column(33,34, 12)
    summary.merge_range('AK3:AK4','Revenue',bold)
    summary.set_column('AK:AK', 18) # 35
    summary.merge_range('AL3:AL4','Average Purchase Service Latency (ms)',bold)
    summary.set_column('AL:AL', 20) # 36
    
    # for row in result:
    #     print(row[0], row[1], row[2], row[3], row[4], row[5], row[6], row[7], row[8], row[9], row[10], row[11], row[12], row[13], row[14], row[15], row[16], row[17], row[18], row[19], row[20], row[21], row[22], row[23], row[24], row[25], row[26], row[27], row[28], row[29], row[30], row[31], row[32], row[33], row[34], row[35], row[36])
    
    for r, row in enumerate(result, start=4):
        for c, col in enumerate(row):
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            summary.write(r,8, row[7], content)
            summary.write(r,9, row[8], content)
            summary.write(r,10, row[9], content)
            summary.write(r,11, row[10], content)
            summary.write(r,12, row[11], content)
            summary.write(r,13, row[12], content)
            summary.write(r,14, row[13], content)
            summary.write(r,15, row[14], content)
            summary.write(r,16, row[15], content)
            summary.write(r,17, row[16], content)
            summary.write(r,18, row[17], content)
            summary.write(r,19, row[18], content)
            summary.write(r,20, row[19], content)
            summary.write(r,21, row[20], content)
            summary.write(r,22, row[21], content)
            summary.write(r,23, row[22], content)
            summary.write(r,24, row[23], content)
            summary.write(r,25, row[24], content)
            summary.write(r,26, row[25], content)
            summary.write(r,27, row[26], content)
            summary.write(r,28, row[27], content)
            summary.write(r,29, row[28], content)
            summary.write(r,30, row[29], content)
            summary.write(r,31, row[30], content)
            summary.write(r,32, row[31], content)
            summary.write(r,33, row[32], content)
            summary.write(r,34, row[33], content)
            summary.write(r,35, row[34], content)
            summary.write(r,36, row[35], content)
            summary.write(r,37, row[36], content)
            
    cursor.close()


# In[4]:


# B Transaction
def B_Transaction():
    cursor = database.cursor()
    with open('B_Transaction.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()

    summary.write('A8','B Transaction',title)
    summary.merge_range('B10:B11','Calling Date',bold)
    summary.merge_range('C10:E10','PURCHASE AUTORENEWAL ON (A)',bold)
    summary.set_column('C:E', 15)
    summary.write('C11','PA_ATTEMPT',bold)
    summary.write('D11', 'PA_SUCCESS',bold)
    summary.write('E11', 'PA_REVENUE',bold)
    summary.merge_range('F10:H10','PURCHASE AUTORENEWAL OFF (B)',bold)
    summary.set_column('F:H', 15)
    summary.write('F11','PA_ATTEMPT',bold)
    summary.write('G11', 'PA_SUCCESS',bold)
    summary.write('H11', 'PA_REVENUE',bold)
    summary.merge_range('I10:K10','PURCHASE AUTORENEWAL OFF (B)',bold)
    summary.set_column('I:K', 15)
    summary.write('I11','PA_ATTEMPT',bold)
    summary.write('J11', 'PA_SUCCESS',bold)
    summary.write('K11', 'PA_REVENUE',bold)
    summary.merge_range('L10:N10','RENEWAL AUTORENEWAL OFF (B)',bold)
    summary.set_column('L:N', 15)
    summary.write('L11','PA_ATTEMPT',bold)
    summary.write('M11', 'PA_SUCCESS',bold)
    summary.write('N11', 'PA_REVENUE',bold)
    summary.merge_range('O10:O11','TOTAL ATTEMPT',bold)
    summary.set_column('O:O', 12)
    summary.merge_range('P10:P11','TOTAL SUCCESS',bold)
    summary.set_column('P:P', 12)
    summary.merge_range('Q10:Q11','TOTAL REVENUE',bold)
    summary.set_column('Q:Q', 18)

    for r, row in enumerate(result, start=11):
        for c, col in enumerate(row):
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            summary.write(r,8, row[7], content)
            summary.write(r,9, row[8], content)
            summary.write(r,10, row[9], content)
            summary.write(r,11, row[10], content)
            summary.write(r,12, row[11], content)
            summary.write(r,13, row[12], content)
            summary.write(r,14, row[13], content)
            summary.write(r,15, row[14], content)
            summary.write(r,16, row[15], content)
            
    cursor.close()


# In[5]:


# C_Subscriber
def C_Subscriber():
    cursor = database.cursor()
    with open('C_Subscriber.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()

    summary.write('A16','C_Subscriber',title)
    summary.write('B18','Calling Date',bold)
    summary.set_row(17, 30)
    summary.write('C18','TOTAL SUBSCRIBER',bold)
    summary.write('D18','TOT REQ ACTIVATION',bold)
    summary.write('E18','TOT SUCC ACTIVATION',bold)
    summary.write('F18','TOT REQ DEACTIVATION',bold)
    summary.write('G18','TOT SUCC DEACTIVATION',bold)
    summary.write('H18','SUSPEND TO ACTIVE',bold)
    summary.write('I18','ACTIVE TO SUSPEND',bold)
    summary.write('J18','Free',bold)
    summary.write('K18','Paid',bold)
    summary.set_column('B:K', 15)
    
    for r, row in enumerate(result, start=18):
        for c, col in enumerate(row):
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            summary.write(r,8, row[7], content)
            
    cursor.close()


# In[6]:


# C_Subscriber Free Paid
def C_FreePaid():
    cursor = database.cursor()
    with open('C_Subscriber_Free_Paid.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()
    
    for r, row in enumerate(result, start=18):
        for c, col in enumerate(row):
            summary.write(r,9, row[0], content)
            summary.write(r,10, row[1], content)
            
    cursor.close()


# In[7]:


# C_SUBS_WITHDRAWAL
def C_SUBS_WITHDRAWAL():
    cursor = database.cursor()
    with open('C_SUBS_WITHDRAWAL.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()
    summary.write('B23','Calling Date',bold)
    summary.set_row(22, 30)
    summary.write('C23','SUBS WITHDRAWAL',bold)
    summary.write('D23','DELETE CONTENT',bold)
    summary.write('E23','RENEWAL CONTENT DELETION',bold)
    summary.write('F23','RENEWAL CONTENT INSUF BALANCE',bold)
    summary.write('G23','RENEWAL CONTENT ERROR CHARGING',bold)
    summary.write('H23','GRAND TOTAL',bold)
    
    for r, row in enumerate(result, start=23):
        for c, col in enumerate(row):
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            
    cursor.close()


# In[8]:


# C_CHURNED_USER
def C_CHURNED_USER():
    cursor = database.cursor()
    with open('C_CHURNED_USER.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()
    summary.write('B28','Calling Date',bold)
    summary.set_column('C:C', 32)
    summary.write('C28','CHURNED USER',bold)
    summary.write('D28','SUMBAGUT',bold)
    summary.write('E28','SUMBAGTENG',bold)
    summary.write('F28','SUMBAGSEL',bold)
    summary.write('G28','JABOTABEK',bold)
    summary.write('H28','JABAR',bold)
    summary.write('I28','JATENG',bold)
    summary.write('J28','JATIM',bold)
    summary.write('K28','BALINUSRA',bold)
    summary.write('L28','KALIMANTAN',bold)
    summary.write('M28','SULAWESI',bold)
    summary.write('N28','PAPUA',bold)
    summary.write('O28','OTHER',bold)
    summary.write('P28','TOTAL',bold)

    for r, row in enumerate(result, start=28):
        for c, col in enumerate(row):
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            summary.write(r,8, row[7], content)
            summary.write(r,9, row[8], content)
            summary.write(r,10, row[9], content)
            summary.write(r,11, row[10], content)
            summary.write(r,12, row[11], content)
            summary.write(r,13, row[12], content)
            summary.write(r,14, row[13], content)
            summary.write(r,15, row[14], content)
            
    cursor.close()
    chat = 'SUMMARY Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[9]:


# A. NEW SUBSCRIPTION TOTAL
def A_NEW_SUBSCRIPTION_TOTAL():
    cursor = database.cursor()
    with open('A_NEW_SUBSCRIPTION_TOTAL.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    new_subs.set_zoom(70)
    new_subs.hide_gridlines(2)
    new_subs.write('A1','A. NEW SUBSCRIPTION TOTAL',title)
    new_subs.merge_range('B3:B4','Calling Date',bold)
    new_subs.merge_range('C3:C4','Attempt',bold)
    new_subs.merge_range('D3:D4','Success Come In System',bold)
    new_subs.merge_range('E3:E4','SR Attempt Come In System',bold)
    new_subs.merge_range('F3:F4','Success Charge',bold)
    new_subs.merge_range('G3:G4','SR Success Charge',bold)
    new_subs.merge_range('H3:I3','SUCCESS CHARGE',bold)
    new_subs.write('H4','Success',bold)
    new_subs.write('I4','Waiting Provisioning',bold)
    new_subs.merge_range('J3:J4','Not Subscriber',bold)
    new_subs.merge_range('K3:K4','Insufficient Balance',bold)
    new_subs.set_column(1,2, 12)
    new_subs.set_column('D:D', 30)
    new_subs.set_column('E:E', 30)
    new_subs.set_column('F:F', 20)
    new_subs.set_column('G:G', 25)
    new_subs.set_column('H:H', 15)
    new_subs.set_column('I:I', 15)
    new_subs.set_column('J:J', 12)
    new_subs.set_column('K:K', 28)
    new_subs.set_row(3, 48)
    
    for r, row in enumerate(result, start=4):
        for c, col in enumerate(row):
            new_subs.write(r,1, row[0], date_format)
            new_subs.write(r,2, row[1], content)
            new_subs.write(r,3, row[2], content)
            new_subs.write(r,4, row[3], content)
            new_subs.write(r,5, row[4], content)
            new_subs.write(r,6, row[5], content)
            new_subs.write(r,7, row[6], content)
            new_subs.write(r,8, row[7], content)
            new_subs.write(r,9, row[8], content)
            new_subs.write(r,10, row[9], content)
            
    cursor.close()


# In[10]:


# B. NEW SUBSCRIPTION CHANNEL
def B_NEW_SUBSCRIPTION_CHANNEL():    
    
    # A_SUCCESS_RATE Purchasing
    def A_SUCCESS_RATE():
        cursor = database.cursor()
        with open('B_NEW_SUBSCRIPTION_CHANNEL_A_SUCCESS_RATE.txt','r') as inserts:
            query = inserts.read()

        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        new_subs.write('A7','B. NEW SUBSCRIPTION CHANNEL',title)
        new_subs.write('A9','Purchasing',title)
        new_subs.write('B11','Calling Date',bold)
        new_subs.write('C11','NOTE',bold)
        new_subs.write('D11','SMS',bold)
        new_subs.write('E11','UMB',bold)
        new_subs.write('F11','HTTP',bold)
        new_subs.write('G11','DTMF',bold)
        new_subs.write('H11','OPENAPI',bold)
        new_subs.write('I11','WEB',bold)
        new_subs.write('J11','IVR',bold)
        new_subs.write('K11','WAP',bold)
        new_subs.write('L11','CSWEB',bold)
        new_subs.write('M11','WCC',bold)
        new_subs.write('N11','OTHERS',bold)
        new_subs.merge_range('B15:N15',' ',bold)
        new_subs.set_row(14, 8)
        new_subs.write('B14','',neat)
        new_subs.set_column('C:C', 20)
        new_subs.write('C14','SUCCESS RATE',neat)
        
        for line in ['D','E','F','G','H','I','J','K','L','M','N']:
            new_subs.write(line + '14','',neat)
            new_subs.write_formula(line + '14', '=IFERROR(' + line +'13'+'/'+ line +'12' + ',"-")',percentage)
        
        for r, row in enumerate(result, start=11):
            for c, col in enumerate(row):
                new_subs.write(r,1, row[0], date_format)
                new_subs.write(r,2, row[1], content)
                new_subs.write(r,3, row[2], content)
                new_subs.write(r,4, row[3], content)
                new_subs.write(r,5, row[4], content)
                new_subs.write(r,6, row[5], content)
                new_subs.write(r,7, row[6], content)
                new_subs.write(r,8, row[7], content)
                new_subs.write(r,9, row[8], content)
                new_subs.write(r,10, row[9], content)
                new_subs.write(r,11, row[10], content)
                new_subs.write(r,12, row[11], content)
                new_subs.write(r,13, row[12], content)

        cursor.close()

    # B_SUCCESS_RATE Purchasing
    def B_SUCCESS_RATE():
        cursor = database.cursor()
        with open('B_NEW_SUBSCRIPTION_CHANNEL_B_SUCCESS_RATE.txt','r') as inserts:
            query = inserts.read()

        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        new_subs.write('B18','',neat)
        new_subs.write('C18','SUCCESS RATE',neat)
        
        for line in ['D','E','F','G','H','I','J','K','L','M','N']:
            new_subs.write(line + '18','',neat)
            new_subs.write_formula(line + '18', '=IFERROR(' + line +'17'+'/'+ line +'16' + ',"-")',percentage)

        for r, row in enumerate(result, start=15):
            for c, col in enumerate(row):
                new_subs.write(r,1, row[0], date_format)
                new_subs.write(r,2, row[1], content)
                new_subs.write(r,3, row[2], content)
                new_subs.write(r,4, row[3], content)
                new_subs.write(r,5, row[4], content)
                new_subs.write(r,6, row[5], content)
                new_subs.write(r,7, row[6], content)
                new_subs.write(r,8, row[7], content)
                new_subs.write(r,9, row[8], content)
                new_subs.write(r,10, row[9], content)
                new_subs.write(r,11, row[10], content)
                new_subs.write(r,12, row[11], content)
                new_subs.write(r,13, row[12], content)
            
        cursor.close()
        
    # C_SUCCESS_RATE Renewal
    def C_SUCCESS_RATE():
        cursor = database.cursor()
        with open('B_NEW_SUBSCRIPTION_CHANNEL_C_SUCCESS_RATE.txt','r') as inserts:
            query = inserts.read()

        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        new_subs.write('A20','Purchasing',title)
        new_subs.write('B22','Calling Date',bold)
        new_subs.write('C22','NOTE',bold)
        new_subs.write('D22','SMS',bold)
        new_subs.write('E22','UMB',bold)
        new_subs.write('F22','HTTP',bold)
        new_subs.write('G22','DTMF',bold)
        new_subs.write('H22','OPENAPI',bold)
        new_subs.write('I22','WEB',bold)
        new_subs.write('J22','IVR',bold)
        new_subs.write('K22','WAP',bold)
        new_subs.write('L22','CSWEB',bold)
        new_subs.write('M22','WCC',bold)
        new_subs.write('N22','OTHERS',bold)
        new_subs.merge_range('B26:N26',' ',bold)
        new_subs.set_row(25, 8)
        new_subs.write('B25','',neat)
        new_subs.write('C25','SUCCESS RATE',neat)
        
        for line in ['D','E','F','G','H','I','J','K','L','M','N']:
            new_subs.write(line + '25','',neat)
            new_subs.write_formula(line + '25', '=IFERROR(' + line +'24'+'/'+ line +'23' + ',"-")',percentage)
        
        for r, row in enumerate(result, start=22):
            for c, col in enumerate(row):
                new_subs.write(r,1, row[0], date_format)
                new_subs.write(r,2, row[1], content)
                new_subs.write(r,3, row[2], content)
                new_subs.write(r,4, row[3], content)
                new_subs.write(r,5, row[4], content)
                new_subs.write(r,6, row[5], content)
                new_subs.write(r,7, row[6], content)
                new_subs.write(r,8, row[7], content)
                new_subs.write(r,9, row[8], content)
                new_subs.write(r,10, row[9], content)
                new_subs.write(r,11, row[10], content)
                new_subs.write(r,12, row[11], content)
                new_subs.write(r,13, row[12], content)

        cursor.close()
        
    # D_SUCCESS_RATE Renewal
    def D_SUCCESS_RATE():
        cursor = database.cursor()
        with open('B_NEW_SUBSCRIPTION_CHANNEL_D_SUCCESS_RATE.txt','r') as inserts:
            query = inserts.read()

        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        new_subs.write('B29','',neat)
        new_subs.write('C29','SUCCESS RATE',neat)
        
        for line in ['D','E','F','G','H','I','J','K','L','M','N']:
            new_subs.write(line + '29','',neat)
            new_subs.write_formula(line + '29', '=IFERROR(' + line +'28'+'/'+ line +'27' + ',"-")',percentage)

        for r, row in enumerate(result, start=26):
            for c, col in enumerate(row):
                new_subs.write(r,1, row[0], date_format)
                new_subs.write(r,2, row[1], content)
                new_subs.write(r,3, row[2], content)
                new_subs.write(r,4, row[3], content)
                new_subs.write(r,5, row[4], content)
                new_subs.write(r,6, row[5], content)
                new_subs.write(r,7, row[6], content)
                new_subs.write(r,8, row[7], content)
                new_subs.write(r,9, row[8], content)
                new_subs.write(r,10, row[9], content)
                new_subs.write(r,11, row[10], content)
                new_subs.write(r,12, row[11], content)
                new_subs.write(r,13, row[12], content)
            
        cursor.close()
            
    A_SUCCESS_RATE()
    B_SUCCESS_RATE()
    C_SUCCESS_RATE()
    D_SUCCESS_RATE()


# In[11]:


# C. NEW SUBSCRIPTION HOUR
def C_NEW_SUBSCRIPTION_HOUR():
    cursor = database.cursor()
    with open('C_NEW_SUBSCRIPTION_HOUR.txt','r') as inserts:
            query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    new_subs.write('A31','C. NEW SUBSCRIPTION HOUR',title)
    new_subs.write('B33','CALLING_DATE',bold)
    new_subs.write('C33','TIME',bold)
    new_subs.write('D33','ATTEMPT',bold)
    new_subs.write('E33','SUCC_COME',bold)
    new_subs.write('F33','SR_SUCC_COME',bold)
    new_subs.write('G33','SUCC_CHARGE',bold)
    new_subs.write('H33','SR',bold)
    new_subs.set_column('B:B', 20)
    
    for r, row in enumerate(result, start=33):
        for c, col in enumerate(row):
            new_subs.write(r,1, row[0], date_format)
            new_subs.write(r,2, row[1], content)
            new_subs.write(r,3, row[2], content)
            new_subs.write(r,4, row[3], content)
            new_subs.write(r,5, row[4], content)
            new_subs.write(r,6, row[5], content)
            new_subs.write(r,7, row[6], content)
            
    cursor.close()


# In[12]:


# D. NEW SUBSCRIPTION HOUR PER CHANNEL
def D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL():
    cursor = database.cursor()
    with open('D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL.txt','r') as inserts:
            query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    new_subs.write('A59','D. NEW SUBSCRIPTION HOUR PER CHANNEL',title)
    new_subs.write('B61','CALLING_DATE',bold)
    new_subs.write('C61','TIME',bold)
    new_subs.write('D61','CHANNEL',bold)
    new_subs.write('E61','ATTEMPT',bold)
    new_subs.write('F61','SUCC_COME',bold)
    new_subs.write('G61','SR_SUCC_COME',bold)
    new_subs.write('H61','SUCC_CHARGE',bold)
    new_subs.write('I61','SR',bold)
    
    for r, row in enumerate(result, start=61):
        for c, col in enumerate(row):
            new_subs.write(r,1, row[0], date_format)
            new_subs.write(r,2, row[1], content)
            new_subs.write(r,3, row[2], content)
            new_subs.write(r,4, row[3], content)
            new_subs.write(r,5, row[4], content)
            new_subs.write(r,6, row[5], content)
            new_subs.write(r,7, row[6], content)
            new_subs.write(r,8, row[7], content)
            
    cursor.close()
    chat = 'NEW_SUBS Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[13]:


# A. PROV TOTAL :
def A_PROV_TOTAL():
    cursor = database.cursor()
    with open('A_PROV_TOTAL.txt','r') as inserts:
            query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    prov.set_zoom(70)
    prov.hide_gridlines(2)
    prov.write('A1','A. PROV TOTAL',title)
    prov.write('B3','CALLING_DATE',bold)
    prov.set_column(1,4, 15)
    prov.write('C3','ATT',bold)
    prov.write('D3','SUCC',bold)
    prov.write('E3','SR',bold)
    
    for r, row in enumerate(result, start=3):
        for c, col in enumerate(row):
            prov.write(r,1, row[0], date_format)
            prov.write(r,2, row[1], content)
            prov.write(r,3, row[2], content)
            prov.write(r,4, row[3], content)
            
    cursor.close()


# In[14]:


# B. PROV CHANNEL :
def B_PROV_CHANNEL():
    cursor = database.cursor()
    with open('B_PROV_CHANNEL.txt','r') as inserts:
            query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    prov.write('A6','B. PROV CHANNEL',title)
    prov.merge_range('B8:B9','CALLING_DATE',bold)
    prov.merge_range('C8:C9','CHANNEL',bold)
    prov.set_column('C:C', 18)
    prov.merge_range('D8:E8','PURCHASE AUTORENEWAL ON (A)',bold)
    prov.write('D9','ATTEMPT',bold)
    prov.write('E9','SUCCESS',bold)
    prov.merge_range('F8:G8','PURCHASE AUTORENEWAL OFF (B)',bold)
    prov.write('F9','ATTEMPT',bold)
    prov.write('G9','SUCCESS',bold)
    prov.set_column(3,8, 16)
    prov.merge_range('H8:H9','SR (A)',bold)
    prov.merge_range('I8:I9','SR (B)',bold)

    for line in ['B','C','D','E','F','G','H','I']:
        for kolom in range(11):
            Alline = line+'{}'.format(*([kolom + 10] * 1))
            prov.write(Alline,'',neat)
            SRA_Source = 'D{}/E{}'.format(*([kolom + 10] * 11))
            SRA_Cetak = "H{}".format(*([kolom + 10] * 11))
            SRB_Source = 'F{}/G{}'.format(*([kolom + 10] * 11))
            SRB_Cetak = "I{}".format(*([kolom + 10] * 11))
            prov.write_formula(SRA_Cetak,'=IFERROR('+SRA_Source+',"-")',percentage)
            prov.write_formula(SRB_Cetak,'=IFERROR('+SRB_Source+',"-")',percentage)
    
    for r, row in enumerate(result, start=9):
        for c, col in enumerate(row):
            prov.write(r,1, row[0], date_format)
            prov.write(r,2, row[1], content)
            prov.write(r,3, row[2], content)
            prov.write(r,4, row[3], content)
            prov.write(r,5, row[4], content)
            prov.write(r,6, row[5], content)
            
    cursor.close()


# In[15]:


# C. PROV CHANNEL HOUR :
def C_PROV_CHANNEL_HOUR():
    cursor = database.cursor()
    with open('C_PROV_CHANNEL_HOUR.txt','r') as inserts:
            query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    prov.write('A22','C. PROV CHANNEL HOUR',title)
    prov.merge_range('B24:B25','CALLING_DATE',bold)
    prov.merge_range('C24:C25','HOUR',bold)
    prov.merge_range('D24:D25', 'CHANNEL',bold)
    prov.set_column('D:D', 18)
    prov.merge_range('E24:F24','PURCHASE AUTORENEWAL ON (A)',bold)
    prov.write('E25','ATTEMPT',bold)
    prov.write('F25','SUCCESS',bold)
    prov.merge_range('G24:H24','PURCHASE AUTORENEWAL OFF (B)',bold)
    prov.write('G25','ATTEMTPT',bold)
    prov.write('H25','SUCCESS',bold)
    prov.merge_range('I24:I25', 'SR (A)',bold)
    prov.merge_range('J24:J25', 'SR (B)',bold)
    prov.set_column('E:H', 18)
    prov.set_column('I:J', 10)
    
    for line in ['B','C','D','E','F','G','H','I','J']:
        for kolom in range(200):
            Alline = line+'{}'.format(*([kolom + 26] * 200))
            prov.write(Alline,'',neat)
            SRA_Source = 'E{}/F{}'.format(*([kolom + 26] * 200))
            SRA_Cetak = "I{}".format(*([kolom + 26] * 200))
            SRB_Source = 'G{}/H{}'.format(*([kolom + 26] * 200))
            SRB_Cetak = "J{}".format(*([kolom + 26] * 200))
            prov.write_formula(SRA_Cetak,'=IFERROR('+SRA_Source+',"-")',percentage)
            prov.write_formula(SRB_Cetak,'=IFERROR('+SRB_Source+',"-")',percentage)
    
    for r, row in enumerate(result, start=25):
        for c, col in enumerate(row):
            prov.write(r,1, row[0], date_format)
            prov.write(r,2, row[1], content)
            prov.write(r,3, row[2], content)
            prov.write(r,4, row[3], content)
            prov.write(r,5, row[4], content)
            prov.write(r,6, row[5], content)
            prov.write(r,7, row[6], content)
    
    cursor.close()
    chat = 'PROV Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[16]:


# A. TOTAL RBT TRAFFIC :
def A_TOTAL_RBT_TRAFFIC():
    cursor = database.cursor()
    with open('A_TOTAL_RBT_TRAFFIC.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    totaltraf.set_zoom(70)
    totaltraf.hide_gridlines(2)
    totaltraf.write('A1','A. TOTAL RBT TRAFFIC',title)
    totaltraf.write('B3','CALLING_DATE',bold)
    totaltraf.write('C3','CHANNEL',bold)
    totaltraf.write('D3','ATTEMPT',bold)
    totaltraf.write('E3','SUCCESS',bold)
    totaltraf.write('F3','SR',bold)
    totaltraf.set_column('B:F', 25)
    
    for line in ['B','C','D','E','F']:
        for kolom in range(20):
            Alline = line+'{}'.format(*([kolom + 4] * 20))
            totaltraf.write(Alline,'',neat)
    
    for r, row in enumerate(result, start=3):
        for c, col in enumerate(row):
            totaltraf.write(r,1, row[0], date_format)
            totaltraf.write(r,2, row[1], content)
            totaltraf.write(r,3, row[2], content)
            totaltraf.write(r,4, row[3], content)
            totaltraf.write(r,5, row[4], content)
    
    cursor.close()


# In[17]:


# TRAFFIC PURCHASE :
def TRAFFIC_PURCHASE():
    cursor = database.cursor()
    with open('TRAFFIC_PURCHASE.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    totaltraf.write('A25','* TRAFFIC PURCHASE',title)
    totaltraf.merge_range('B27:B28','CALLING_DATE',bold)
    totaltraf.merge_range('C27:C28','CHANNEL',bold)
    totaltraf.merge_range('D27:E27','PURCHASE AUTORENEWAL ON (A)',bold)
    totaltraf.write('D28','ATTEMPT',bold)
    totaltraf.write('E28','SUCCESS',bold)
    totaltraf.merge_range('F27:G27','PURCHASE AUTORENEWAL OFF (B)',bold)
    totaltraf.write('F28','ATTEMPT',bold)
    totaltraf.write('G28','SUCCESS',bold)
    totaltraf.merge_range('H27:H28','SR (A)',bold)
    totaltraf.merge_range('I27:I28','SR (B)',bold)
    totaltraf.set_column('G:I', 25)
    
    for line in ['B','C','D','E','F','G','H','I']:
        for kolom in range(20):
            Alline = line+'{}'.format(*([kolom + 29] * 20))
            totaltraf.write(Alline,'',neat)
            SRA_Source = 'E{}/D{}'.format(*([kolom + 29] * 20))
            SRA_Cetak = "H{}".format(*([kolom + 29] * 20))
            SRB_Source = 'G{}/F{}'.format(*([kolom + 29] * 20))
            SRB_Cetak = "I{}".format(*([kolom + 29] * 20))
            totaltraf.write_formula(SRA_Cetak,'=IFERROR('+SRA_Source+',"-")',percentage)
            totaltraf.write_formula(SRB_Cetak,'=IFERROR('+SRB_Source+',"-")',percentage)
            
    for r, row in enumerate(result, start=28):
        for c, col in enumerate(row):
            totaltraf.write(r,1, row[0], date_format)
            totaltraf.write(r,2, row[1], content)
            totaltraf.write(r,3, row[2], content)
            totaltraf.write(r,4, row[3], content)
            totaltraf.write(r,5, row[4], content)
            totaltraf.write(r,6, row[5], content)
    
    cursor.close()


# In[18]:


# TRAFFIC RENEWAL :
def TRAFFIC_RENEWAL():
    cursor = database.cursor()
    with open('TRAFFIC_RENEWAL.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    totaltraf.write('M25','* TRAFFIC RENEWAL',title)
    totaltraf.merge_range('M27:M28','CALLING_DATE',bold)
    totaltraf.merge_range('N27:N28','CHANNEL',bold)
    totaltraf.merge_range('O27:P27','RENEWAL AUTORENEWAL ON (A)',bold)
    totaltraf.write('O28','ATTEMPT',bold)
    totaltraf.write('P28','SUCCESS',bold)
    totaltraf.merge_range('Q27:R27','RENEWAL AUTORENEWAL OFF (B)',bold)
    totaltraf.write('Q28','ATTEMPT',bold)
    totaltraf.write('R28','SUCCESS',bold)
    totaltraf.merge_range('S27:S28','SR (A)',bold)
    totaltraf.merge_range('T27:T28','SR (B)',bold)
    totaltraf.set_column('M:T', 25)
    
    for line in ['M','N','O','P','Q','R','S','T']:
        for kolom in range(20):
            Alline = line+'{}'.format(*([kolom + 29] * 20))
            totaltraf.write(Alline,'',neat)
            SRA_Source = 'P{}/O{}'.format(*([kolom + 29] * 20))
            SRA_Cetak = "S{}".format(*([kolom + 29] * 20))
            SRB_Source = 'R{}/Q{}'.format(*([kolom + 29] * 20))
            SRB_Cetak = "T{}".format(*([kolom + 29] * 20))
            totaltraf.write_formula(SRA_Cetak,'=IFERROR('+SRA_Source+',"-")',percentage)
            totaltraf.write_formula(SRB_Cetak,'=IFERROR('+SRB_Source+',"-")',percentage)
            
    for r, row in enumerate(result, start=28):
        for c, col in enumerate(row):
            totaltraf.write(r,12, row[0], date_format)
            totaltraf.write(r,13, row[1], content)
            totaltraf.write(r,14, row[2], content)
            totaltraf.write(r,15, row[3], content)
            totaltraf.write(r,16, row[4], content)
            totaltraf.write(r,17, row[5], content)
    
    cursor.close()
    chat = 'TOTAL TRAFFIC Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[19]:


# A. PRICE TRAFFIC & REVENUE :
def A_PRICE_TRAFFIC_REVENUE():
    cursor = database.cursor()
    with open('A_PRICE_TRAFFIC_REVENUE.txt','r') as inserts:
        query = inserts.read()
        
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    pricerevenue.set_zoom(70)
    pricerevenue.hide_gridlines(2)
    pricerevenue.write('A1','A. PRICE TRAFFIC & REVENUE',title)
    pricerevenue.write('B3','CALLING_DATE',bold)
    pricerevenue.write('C3','APPS_ID',bold)
    pricerevenue.write('D3','CHG_AMOUNT',bold)
    pricerevenue.write('E3','CHG_REQ',bold)
    pricerevenue.write('F3','REFUND_AMOUNT',bold)
    pricerevenue.write('G3','REFUND_REQ',bold)
    pricerevenue.merge_range('B37:C37','Total',neat)
    pricerevenue.merge_range('B38:F38','REVENUE',neat)
    pricerevenue.merge_range('B39:F39','REVENUE with tax 10%',neat)
    pricerevenue.write_formula('G38','=D37-F37',content)
    pricerevenue.write_formula('G39','=G38+(G38*0.1)',content)
    pricerevenue.set_column('B:K', 25)
    
    for persum in ['D','E','F','G']:
        pricerevenue.write_formula(persum+'37','=SUM('+persum+'4:'+persum+'36)',content)
    
    for line in ['B','C','D','E','F','G']:
        for kolom in range(33):
            Alline = line+'{}'.format(*([kolom + 4] * 33))
            pricerevenue.write(Alline,'',neat)
            
    for r, row in enumerate(result, start=3):
        for c, col in enumerate(row):
            pricerevenue.write(r,1, row[0], date_format)
            pricerevenue.write(r,2, row[1], content)
            pricerevenue.write(r,3, row[2], content)
            pricerevenue.write(r,4, row[3], content)
            pricerevenue.write(r,5, row[4], content)
            pricerevenue.write(r,6, row[5], content)
            
    cursor.close()


# In[20]:


# B. PRICE TRAFFIC & REVENUE :
def B_PRICE_TRAFFIC_REVENUE():
    cursor = database.cursor()
    with open('B_PRICE_TRAFFIC_REVENUE.txt','r') as inserts:
        query = inserts.read()
        
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    pricerevenue.merge_range('B42:B43','CALLING_DATE',bold)
    pricerevenue.merge_range('C42:C43','CHARGE_AMOUNT',bold)
    pricerevenue.merge_range('D42:E42','PURCHASE AUTORENEWAL ON (A)',bold)
    pricerevenue.write('D43','CHG_REQUEST',bold)
    pricerevenue.write('E43','CHG_AMOUNT',bold)
    pricerevenue.merge_range('F42:G42','PURCHASE AUTORENEWAL OFF (B)',bold)
    pricerevenue.write('F43','CHG_REQUEST',bold)
    pricerevenue.write('G43','CHG_AMOUNT',bold)
    pricerevenue.merge_range('H42:I42','RENEWAL AUTORENEWAL ON (A)',bold)
    pricerevenue.write('H43','CHG_REQUEST',bold)
    pricerevenue.write('I43','CHG_AMOUNT',bold)
    pricerevenue.merge_range('J42:K42','RENEWAL AUTORENEWAL OFF (B)',bold)
    pricerevenue.write('J43','CHG_REQUEST',bold)
    pricerevenue.write('K43','CHG_AMOUNT',bold)
    pricerevenue.merge_range('B76:C76','Total',neat)
    
    for persum in ['D','E','F','G','H','I','J','K']:
        pricerevenue.write_formula(persum+'76','=SUM('+persum+'44:'+persum+'75)',content)
    
    for line in ['B','C','D','E','F','G','H','I','J','K']:
        for kolom in range(32):
            Alline = line+'{}'.format(*([kolom + 44] * 32))
            pricerevenue.write(Alline,'',neat)
    
    for r, row in enumerate(result, start=43):
        for c, col in enumerate(row):
            pricerevenue.write(r,1, row[0], date_format)
            pricerevenue.write(r,2, row[1], content)
            pricerevenue.write(r,3, row[2], content)
            pricerevenue.write(r,4, row[3], content)
            pricerevenue.write(r,5, row[4], content)
            pricerevenue.write(r,6, row[5], content)
            pricerevenue.write(r,7, row[6], content)
            pricerevenue.write(r,8, row[7], content)
            pricerevenue.write(r,9, row[8], content)
            pricerevenue.write(r,10, row[9], content)
            
    cursor.close()


# In[21]:


# C. PRICE TRAFFIC & REVENUE :
def C_PRICE_TRAFFIC_REVENUE():
    cursor = database.cursor()
    with open('C_PRICE_TRAFFIC_REVENUE.txt','r') as inserts:
        query = inserts.read()
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    pricerevenue.write('B79','CALLING_DATE',bold)
    pricerevenue.write('C79','TYPE',bold)
    pricerevenue.write('D79','CSWEB',bold)
    pricerevenue.write('E79','DTMF',bold)
    pricerevenue.write('F79','HTTPGW',bold)
    pricerevenue.write('G79','IVR',bold)
    pricerevenue.write('H79','MKIOS',bold)
    pricerevenue.write('I79','OPENAPI',bold)
    pricerevenue.write('J79','POINT',bold)
    pricerevenue.write('K79','RECOMENDATION',bold)
    pricerevenue.write('L79','RENEWAL',bold)
    pricerevenue.set_column('M:AJ', 25)
    pricerevenue.write('M79','SMS',bold)
    pricerevenue.write('N79','UMB',bold)
    pricerevenue.write('O79','WCC',bold)
    pricerevenue.write('P79','WEB',bold)
    pricerevenue.write('Q79','Media_1',bold)
    pricerevenue.write('R79','Media_2',bold)
    pricerevenue.write('S79','Media_3',bold)
    pricerevenue.write('T79','Media_4',bold)
    pricerevenue.write('U79','Media_5',bold)
    pricerevenue.write('V79','Media_6',bold)
    pricerevenue.write('W79','Media_7',bold)
    pricerevenue.write('X79','Media_8',bold)
    pricerevenue.write('Y79','Media_9',bold)
    pricerevenue.write('Z79','Media_A',bold)
    pricerevenue.write('AA79','Media_E',bold)
    pricerevenue.write('AB79','Media_F',bold)
    pricerevenue.write('AC79','Media_G',bold)
    pricerevenue.write('AD79','Media_H',bold)
    pricerevenue.write('AE79','Media_K',bold)
    pricerevenue.write('AF79','Media_L',bold)
    pricerevenue.write('AG79','Media_Q',bold)
    pricerevenue.write('AH79','Media_R',bold)
    pricerevenue.write('AI79','Media_T',bold)
    pricerevenue.write('AJ79','Grand Total',bold)
    pricerevenue.write_formula('AJ84','=SUM(AJ80:AJ83)',content)
    pricerevenue.merge_range('B84:C84','Total',neat)
    
    for persum in ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ']:
        pricerevenue.write_formula(persum+'84','=SUM('+persum+'80:'+persum+'83)',content)
    
    for line in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ']:
        for kolom in range(4):
            Alline = line+'{}'.format(*([kolom + 80] * 4))
            pricerevenue.write(Alline,'',neat)
    
    for r, row in enumerate(result, start=79):
        for c, col in enumerate(row):
            pricerevenue.write(r,1, row[0], date_format)
            pricerevenue.write(r,2, row[1], content)
            pricerevenue.write(r,3, row[2], content)
            pricerevenue.write(r,4, row[3], content)
            pricerevenue.write(r,5, row[4], content)
            pricerevenue.write(r,6, row[5], content)
            pricerevenue.write(r,7, row[6], content)
            pricerevenue.write(r,8, row[7], content)
            pricerevenue.write(r,9, row[8], content)
            pricerevenue.write(r,10, row[9], content)
            pricerevenue.write(r,11, row[10], content)
            pricerevenue.write(r,12, row[11], content)
            pricerevenue.write(r,13, row[12], content)
            pricerevenue.write(r,14, row[13], content)
            pricerevenue.write(r,15, row[14], content)
            pricerevenue.write(r,16, row[15], content)
            pricerevenue.write(r,17, row[16], content)
            pricerevenue.write(r,18, row[17], content)
            pricerevenue.write(r,19, row[18], content)
            pricerevenue.write(r,20, row[19], content)
            pricerevenue.write(r,21, row[20], content)
            pricerevenue.write(r,22, row[21], content)
            pricerevenue.write(r,23, row[22], content)
            pricerevenue.write(r,24, row[23], content)
            pricerevenue.write(r,25, row[24], content)
            pricerevenue.write(r,26, row[25], content)
            pricerevenue.write(r,27, row[26], content)
            pricerevenue.write(r,28, row[27], content)
            pricerevenue.write(r,29, row[28], content)
            pricerevenue.write(r,30, row[29], content)
            pricerevenue.write(r,31, row[30], content)
            pricerevenue.write(r,32, row[31], content)
            pricerevenue.write(r,33, row[32], content)
            pricerevenue.write(r,34, row[33], content)
            pricerevenue.write(r,35, row[34], content)
            
    cursor.close()
    chat = 'PRICE & REVENUE Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[22]:


# A. SUMMARY CHARGING GW :
def A_SUMMARY_CHARGING_GW():
    cursor = database.cursor()
    query = ("SELECT DECODE (chg_type,'C','Charging','Q','Query Balance','R','Refund'), SUM (amount), SUM(CASE WHEN chg_type = 'C' AND ( RESULT LIKE 'Succ%' OR RESULT LIKE 'Insuff%' OR RESULT LIKE '%subs block%' OR RESULT LIKE '%locked%') THEN amount WHEN chg_type = 'R' AND RESULT LIKE 'Succ%' THEN amount WHEN chg_type = 'Q' AND RESULT LIKE 'Succ%' THEN amount ELSE 0 END),     ROUND(SUM(CASE WHEN chg_type = 'C' AND ( RESULT LIKE 'Succ%' OR RESULT LIKE 'Insuff%' OR RESULT LIKE '%subs block%' OR RESULT LIKE '%locked%') THEN amount WHEN chg_type = 'R' AND RESULT LIKE 'Succ%' THEN amount WHEN chg_type = 'Q' AND RESULT LIKE 'Succ%' THEN amount ELSE 0 END) / SUM (amount) * 100, 2) || '%' FROM (SELECT COUNT (*) AS amount, chg_type, RESULT     FROM perfcdr_chg WHERE chg_type != ' ' AND calling_datetime LIKE '"+str(kemarin_str)+"'||'%' GROUP BY chg_type,RESULT) GROUP BY chg_type ORDER BY 1")
        
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    chargingcdr.set_zoom(70)
    chargingcdr.hide_gridlines(2)
    chargingcdr.write('A1','A. SUMMARY CHARGING GW',title)
    chargingcdr.merge_range('B3:E3','SUMMARY CHARGING GW',bold)
    chargingcdr.write('B4','CHARGING_TYPE',bold)
    chargingcdr.write('C4','TOTAL',bold)
    chargingcdr.write('D4','SUCCESS',bold)
    chargingcdr.write('E4','SUCCESS_RATE',bold)
    chargingcdr.set_column('B:B', 20)
    chargingcdr.set_column('C:C', 12)
    chargingcdr.set_column('D:D', 70)
    chargingcdr.set_column('E:E', 20)
    
    for line in ['B','C','D','E']:
        for kolom in range(3):
            Alline = line+'{}'.format(*([kolom + 5] * 3))
            chargingcdr.write(Alline,'',neat)
            
    for r, row in enumerate(result, start=4):
        for c, col in enumerate(row):
            chargingcdr.write(r,1, row[0], content)
            chargingcdr.write(r,2, row[1], content)
            chargingcdr.write(r,3, row[2], content)
            chargingcdr.write(r,4, row[3], content)
            
    cursor.close()


# In[23]:


# B. CHARGING CDR ERROR :
def B_CHARGING_CDR_ERROR():
    cursor = database.cursor()
    query = ("SELECT COUNT (*),chg_type,result FROM perfcdr_chg WHERE chg_type != ' ' AND calling_datetime LIKE '"+str(kemarin_str)+"'||'%' GROUP BY chg_type, result")
    
     # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    chargingcdr.write('A9','B. CHARGING CDR ERROR',title)
    chargingcdr.write('B11','AMOUNT',bold)
    chargingcdr.write('C11','chg_type',bold)
    chargingcdr.write('D11','RESULT',bold)
    
    for line in ['B','C','D']:
        for kolom in range(26):
            Alline = line+'{}'.format(*([kolom + 12] * 26))
            chargingcdr.write(Alline,'',neat)
            
    for r, row in enumerate(result, start=11):
        for c, col in enumerate(row):
            chargingcdr.write(r,1, row[0], content)
            chargingcdr.write(r,2, row[1], content)
            chargingcdr.write(r,3, row[2], content)
            
    cursor.close()
    chat = 'CHARGING CDR Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[24]:


# ERROR SUBS :
def ERROR_SUBS():
    
    # A. NEW SUBS ERROR CHANNEL
    def A_NEW_SUBS_ERROR_CHANNEL():
        cursor = database.cursor()
        with open('A_NEW_SUBS_ERROR_CHANNEL.txt','r') as inserts:
            query = inserts.read()
    
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        errorsubs.set_zoom(70)
        errorsubs.hide_gridlines(2)
        errorsubs.write('A1','A. NEW SUBS ERROR CHANNEL',title)
        errorsubs.write('B3','CALLING_DATE',bold)
        errorsubs.write('C3','CHANNEL',bold)
        errorsubs.write('D3','ERROR_RESULT',bold)
        errorsubs.write('E3','ERROR_COUNT',bold)
        errorsubs.set_column('B:E', 18)
            
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                errorsubs.write(r,1, row[0], date_format)
                errorsubs.write(r,2, row[1], content)
                errorsubs.write(r,3, row[2], content)
                errorsubs.write(r,4, row[3], content)
            
        cursor.close()
    
    # B. NEW SUBS ERROR HOUR CHANNEL
    def B_NEW_SUBS_ERROR_HOUR_CHANNEL():
        cursor = database.cursor()
        with open('B_NEW_SUBS_ERROR_HOUR_CHANNEL.txt','r') as inserts:
            query = inserts.read()
    
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        errorsubs.write('G1','B. NEW SUBS ERROR HOUR CHANNEL',title)
        errorsubs.write('H3','CALLING_DATE',bold)
        errorsubs.write('I3','HOUR',bold)
        errorsubs.write('J3','CHANNEL',bold)
        errorsubs.write('K3','ERROR_RESULT',bold)
        errorsubs.write('L3','ERROR_COUNT',bold)
        errorsubs.set_column('H:L', 18)
            
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                errorsubs.write(r,7, row[0], date_format)
                errorsubs.write(r,8, row[1], content)
                errorsubs.write(r,9, row[2], content)
                errorsubs.write(r,10, row[3], content)
                errorsubs.write(r,11, row[4], content)
            
        cursor.close()
        chat = 'ERROR SUBS Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
    
    A_NEW_SUBS_ERROR_CHANNEL()
    B_NEW_SUBS_ERROR_HOUR_CHANNEL()


# In[25]:


# HTTPGW :
def HTTPGW():
    
    # A. HTTPGW BREAKDOWN PURCHASE
    def A_HTTPGW_BREAKDOWN_PURCHASE():
        cursor = database.cursor()
        with open('A_HTTPGW_BREAKDOWN_PURCHASE.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        httpgw.set_zoom(70)
        httpgw.hide_gridlines(2)
        httpgw.write('A1','A. HTTPGW BREAKDOWN PURCHASE',title)
        httpgw.write('B3','CALLING_DATE',bold)
        httpgw.write('C3','DVC_ID',bold)
        httpgw.write('D3','ATT',bold)
        httpgw.write('E3','SUCC',bold)
        httpgw.write('F3','SR',bold)
        httpgw.set_column('B:C', 17)
        httpgw.set_column('D:F', 12)
            
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                httpgw.write(r,1, row[0], date_format)
                httpgw.write(r,2, row[1], content)
                httpgw.write(r,3, row[2], content)
                httpgw.write(r,4, row[3], content)
                httpgw.write(r,5, row[4], content)
            
        cursor.close()
    
    # B. HTTPGW BREAKDOWN RENEWAL
    def B_HTTPGW_BREAKDOWN_RENEWAL():
        cursor = database.cursor()
        with open('B_HTTPGW_BREAKDOWN_RENEWAL.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        httpgw.write('H1','B. HTTPGW BREAKDOWN RENEWAL',title)
        httpgw.write('H3','CALLING_DATE',bold)
        httpgw.write('I3','ATT',bold)
        httpgw.write('J3','SUCC',bold)
        httpgw.write('K3','SR',bold)
        httpgw.set_column('H:K', 15)
            
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                httpgw.write(r,7, row[0], date_format)
                httpgw.write(r,8, row[1], content)
                httpgw.write(r,9, row[2], content)
                httpgw.write(r,10, row[3], content)
            
        cursor.close()
        chat = 'HTTPGW Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
    
    A_HTTPGW_BREAKDOWN_PURCHASE()
    B_HTTPGW_BREAKDOWN_RENEWAL()


# In[26]:


# SMS_CONFIRMATION :
def SMS_CONFIRMATION():
    
    # Purchase
    def Purchase():
        cursor = database.cursor()
        with open('SMS_CONFIRMATION_PURCHASE.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        smsconfirm.set_zoom(70)
        smsconfirm.hide_gridlines(2)
        smsconfirm.write('B2','Purchase',bold)
        smsconfirm.write('B3','Calling Date',bold)
        smsconfirm.write('C3','Sub Request',bold)
        smsconfirm.write('D3','Sub Respond 1st Confirm',bold)
        smsconfirm.write('E3','Non Respond 1st Confirm',bold)
        smsconfirm.write('F3','Success Charged',bold)
        smsconfirm.write('G3','Fail Charged',bold)
        smsconfirm.set_column('B:C', 13)
        smsconfirm.set_column('D:D', 50)
        smsconfirm.set_column('E:E', 25)
        smsconfirm.set_column('F:G', 15)
        
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                smsconfirm.write(r,1, row[0], date_format)
                smsconfirm.write(r,2, row[1], content)
                smsconfirm.write(r,3, row[2], content)
                smsconfirm.write(r,4, row[3], content)
                smsconfirm.write(r,5, row[4], content)
                smsconfirm.write(r,6, row[5], content)
            
        cursor.close()
        
    # Error_Detail_Purchase
    def Error_Detail_Purchase():
        cursor = database.cursor()
        with open('SMS_CONFIRMATION_ERRORDETAIL_PURCHASE.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        smsconfirm.write('B7','ERROR DETAIL',bold)
        smsconfirm.write('B8','Calling Date',bold)
        smsconfirm.write('C8','Error Code',bold)
        smsconfirm.write('D8','Error Definition',bold)
        smsconfirm.write('E8','Amount',bold)
        
        for line in ['B','C','D','E']:
            for kolom in range(11):
                Alline = line+'{}'.format(*([kolom + 9] * 11))
                smsconfirm.write(Alline,'',neat)
            
        for r, row in enumerate(result, start=8):
            for c, col in enumerate(row):
                smsconfirm.write(r,1, row[0], date_format)
                smsconfirm.write(r,2, row[1], content)
                smsconfirm.write(r,3, row[2], content)
                smsconfirm.write(r,4, row[3], content)
            
        cursor.close()
        
    # Renewal
    def Renewal():
        cursor = database.cursor()
        with open('SMS_CONFIRMATION_RENEWAL.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        smsconfirm.write('B23','Renewal',bold)
        smsconfirm.write('B24','Calling Date',bold)
        smsconfirm.write('C24','Sub Request',bold)
        smsconfirm.write('D24','Sub Respond 1st Confirm',bold)
        smsconfirm.write('E24','Non Respond 1st Confirm',bold)
        smsconfirm.write('F24','Success Charged',bold)
        smsconfirm.write('G24','Fail Charged',bold)
            
        for r, row in enumerate(result, start=24):
            for c, col in enumerate(row):
                smsconfirm.write(r,1, row[0], date_format)
                smsconfirm.write(r,2, row[1], content)
                smsconfirm.write(r,3, row[2], content)
                smsconfirm.write(r,4, row[3], content)
                smsconfirm.write(r,5, row[4], content)
                smsconfirm.write(r,6, row[5], content)
            
        cursor.close()
        
    # Error_Detail_Purchase
    def Error_Detail_Renewal():
        cursor = database.cursor()
        with open('SMS_CONFIRMATION_ERRORDETAIL_RENEWAL.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        smsconfirm.write('B28','ERROR DETAIL',bold)
        smsconfirm.write('B29','Calling Date',bold)
        smsconfirm.write('C29','Error Code',bold)
        smsconfirm.write('D29','Error Definition',bold)
        smsconfirm.write('E29','Amount',bold)
        
        for line in ['B','C','D','E']:
            for kolom in range(11):
                Alline = line+'{}'.format(*([kolom + 30] * 11))
                smsconfirm.write(Alline,'',neat)
            
        for r, row in enumerate(result, start=29):
            for c, col in enumerate(row):
                smsconfirm.write(r,1, row[0], date_format)
                smsconfirm.write(r,2, row[1], content)
                smsconfirm.write(r,3, row[2], content)
                smsconfirm.write(r,4, row[3], content)
            
        cursor.close()
        chat = 'SMS_CONFIRMATION Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
        
    Purchase()
    Error_Detail_Purchase()
    Renewal()
    Error_Detail_Renewal()


# In[27]:


# SONG TOP :
def SONGTOP():
    cursor = database.cursor()
    with open('SONGTOP.txt','r') as inserts:
        query = inserts.read()

    cursor.execute(query)
    df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
    df.to_csv(r'SongTop.csv', sep=',',index=False, header=False, mode='a')
#     df.to_csv(r'SongTop.csv', index=False, mode='a')
    
    songtop.set_zoom(70)
    songtop.hide_gridlines(2)
    songtop.write('A1','SONG TOP 75 STATISTIC',title)
    songtop.write('B3','CALLING_DATE',bold)
    songtop.write('C3','RANK',bold)
    songtop.write('D3','CONTENT_SINGER',bold)
    songtop.write('E3','CONTENT_TITLE',bold)
    songtop.write('F3','TRAFFIC',bold)
    songtop.write('G3','REVENUE',bold)
    songtop.write('H3','CP_NAME',bold)
    songtop.set_column('B:B', 14)
    songtop.set_column('C:C', 6)
    songtop.set_column('D:D', 40)
    songtop.set_column('E:E', 35)
    songtop.set_column('F:F', 8)
    songtop.set_column('G:G', 10)
    songtop.set_column('H:H', 26)
    
    with open('SongTop.csv',encoding='ISO-8859-1') as csvfile:
        readCSV = csv.reader(csvfile, delimiter=',')
        for r, row in enumerate(readCSV, start=3):
            for c, col in enumerate(row):
                songtop.write(r,1, row[0], date_format)
                songtop.write(r,2, row[1], content)
                songtop.write(r,3, row[2], content)
                songtop.write(r,4, row[3], content)
                songtop.write(r,5, row[4], content)
                songtop.write(r,6, row[5], content)
                songtop.write(r,7, row[6], content)
                
    cursor.close()
    chat = 'SONG TOP Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
    fyi(chat)


# In[28]:


def GROSSREVENUE():
    
    def GETDAT():
        cursor = database.cursor()
        query = ("SELECT CALLING_DATE,TOTAL_COUNT,TOTAL_SUCCESS,FAIL_COUNT,SUCCESS_RATE,BILLABLE_COUNT,GROSS_REVENUE,MELON_REVENUE FROM(SELECT calling_date,         SUM (TOTAL_COUNT) AS TOTAL_COUNT,SUM (SUCCESS_COUNT) AS TOTAL_SUCCESS,SUM (FAIL_COUNT) AS FAIL_COUNT,ROUND( (SUM (SUCCESS_COUNT)/ (SUM (SUCCESS_COUNT)         + SUM (FAIL_COUNT))) * 100,2) ||'%' AS success_rate,SUM (BILLABLE_COUNT) AS BILLABLE_COUNT,SUM (CHARGE_AMOUNT) AS GROSS_REVENUE,SUM (CHARGE_AMOUNT)         * 75 /100 AS MELON_REVENUE FROM CDR_SUMMARY WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD') AND calling_type IN (1,2,7,64,93,77,78,65,95,96,99,100)         AND calling_type BETWEEN 0 AND 999 GROUP BY calling_date ORDER BY calling_date)")
        cursor.execute(query)
        Dump = 'Gross_Revenue_'+str(kmrnbulan[0:6])+'.csv'
        df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
        
        for line in ['B','C','D','E','F','G','H','I']:
            for kolom in range(33):
                Alline = line+'{}'.format(*([kolom + 4] * 33))
                revenue.write(Alline,'',neat)
        revenue.set_zoom(70)
        revenue.hide_gridlines(2)
        revenue.write('A1','Revenue Statistics',title)
        revenue.write('B3','CALLING_DATE',bold)
        revenue.write('C3','TOTAL_COUNT',bold)
        revenue.write('D3','TOTAL_SUCCESS',bold)
        revenue.write('E3','FAIL_COUNT',bold)
        revenue.write('F3','SUCCESS_RATE',bold)
        revenue.write('G3','BILLABLE_COUNT',bold)
        revenue.write('H3','GROSS_REVENUE',bold)
        revenue.write('I3','',bold)
        revenue.write('B36','TOTAL',neat)
        revenue.set_column('B:I', 20)
        revenue.write_formula('F36','=AVERAGE(F4:F35)',percentage)
                
        for persum in ['C','D','E']:
            revenue.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
            
        for persum in ['G','H','I']:
            revenue.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
            
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=3):
                for c, col in enumerate(row):
                    revenue.write(r,1, row[0], date_format)
                    revenue.write(r,2, row[1], content)
                    revenue.write(r,3, row[2], content)
                    revenue.write(r,4, row[3], content)
                    revenue.write(r,5, row[4], percentage)
                    revenue.write(r,6, row[5], content)
                    revenue.write(r,7, row[6], content)
                    revenue.write(r,8, row[7], content)
                    
        cursor.close()
        chat = 'REVENUE Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
    
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[29]:


def MELONDIY():
    
    def GETDAT():
        cursor = database.cursor()
        query = ("SELECT CALLING_DATE,TOTAL_COUNT,TOTAL_SUCCESS,FAIL_COUNT,SUCCESS_RATE,BILLABLE_COUNT,GROSS_REVENUE FROM (SELECT calling_date, SUM(TOTAL_COUNT) AS TOTAL_COUNT,SUM(SUCCESS_COUNT) AS TOTAL_SUCCESS,         SUM (FAIL_COUNT) AS FAIL_COUNT,ROUND( (SUM (SUCCESS_COUNT)/ (SUM (SUCCESS_COUNT) + SUM (FAIL_COUNT))) * 100,2) ||'%' AS success_rate, SUM(BILLABLE_COUNT) AS BILLABLE_COUNT,SUM (CHARGE_AMOUNT) AS GROSS_REVENUE         FROM CDR_SUMMARY WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD') AND calling_type IN (1,2,7,64,93,77,78,65,95,96,99,100) AND RATING_CONTENT_ID in ( select content_id from (select content_id from content where mcp_user_id ='23610'         union all select bundle_id  as content_id from bundle where mcp_user_id ='23610')) AND calling_type BETWEEN 0 AND 999 GROUP BY calling_date ORDER BY calling_date)")
        cursor.execute(query)
        Dump = 'Melon_Diy_'+str(kmrnbulan[0:6])+'.csv'
        df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
        for line in ['B','C','D','E','F','G','H']:
            for kolom in range(33):
                Alline = line+'{}'.format(*([kolom + 4] * 33))
                melondiy.write(Alline,'',neat)
        melondiy.set_zoom(60)
        melondiy.hide_gridlines(2)
        melondiy.write('A1','Melon DIY Statistic',title)
        melondiy.write('B3','CALLING_DATE',bold)
        melondiy.write('C3','TOTAL',bold)
        melondiy.write('D3','SUCCESS',bold)
        melondiy.write('E3','FAIL',bold)
        melondiy.write('F3','SUCCESS_RATE',bold)
        melondiy.write('G3','BILLABLE',bold)
        melondiy.write('H3','REVENUE',bold)
        melondiy.write('B36','TOTAL',neat)
        melondiy.set_column('B:H', 20)
        melondiy.write_formula('F36','=AVERAGE(F4:F35)',percentage)
        
        for persum in ['C','D','E']:
            melondiy.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
        for persum in ['G','H']:
            melondiy.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=3):
                for c, col in enumerate(row):
                    melondiy.write(r,1, row[0], date_format)
                    melondiy.write(r,2, row[1], content)
                    melondiy.write(r,3, row[2], content)
                    melondiy.write(r,4, row[3], content)
                    melondiy.write(r,5, row[4], percentage)
                    melondiy.write(r,6, row[5], content)
                    melondiy.write(r,7, row[6], content)
        
        cursor.close()
        chat = 'MELON DIY Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
    
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[30]:


def MELONMYNSP():
    def Statistic():
        def GETDAT():
            cursor = database.cursor()
            query = ("SELECT CALLING_DATE,TOTAL_COUNT,TOTAL_SUCCESS,FAIL_COUNT,SUCCESS_RATE,BILLABLE_COUNT,GROSS_REVENUE FROM (SELECT calling_date,SUM (TOTAL_COUNT) AS TOTAL_COUNT,SUM (SUCCESS_COUNT) AS TOTAL_SUCCESS,SUM (FAIL_COUNT) AS FAIL_COUNT,             ROUND( (SUM (SUCCESS_COUNT)/ (SUM (SUCCESS_COUNT) + SUM (FAIL_COUNT))) * 100,2) ||'%' AS success_rate,SUM (BILLABLE_COUNT) AS BILLABLE_COUNT,SUM (CHARGE_AMOUNT) AS GROSS_REVENUE FROM CDR_SUMMARY WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD')             AND calling_type IN (1,2,7,77,64,93,78,65,95,96,99,100) AND RATING_CONTENT_ID in ( select content_id from (select content_id from content where mcp_user_id ='23661' union all select bundle_id  as content_id from bundle where mcp_user_id ='23661' ))             AND calling_type BETWEEN 0 AND 999 GROUP BY calling_date ORDER BY calling_date)")
            cursor.execute(query)
            Dump = 'Melon_MyNsp_'+str(kmrnbulan[0:6])+'.csv'
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')

            for line in ['B','C','D','E','F','G','H']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 4] * 33))
                    melonmynsp.write(Alline,'',neat)

            melonmynsp.set_zoom(60)
            melonmynsp.hide_gridlines(2)
            melonmynsp.write('A1','Melon MyNSP Statistic',title)
            melonmynsp.write('B3','CALLING_DATE',bold)
            melonmynsp.write('C3','TOTAL',bold)
            melonmynsp.write('D3','SUCCESS',bold)
            melonmynsp.write('E3','FAIL',bold)
            melonmynsp.write('F3','SUCCESS_RATE',bold)
            melonmynsp.write('G3','BILLABLE',bold)
            melonmynsp.write('H3','REVENUE',bold)
            melonmynsp.write('B36','TOTAL',neat)
            melonmynsp.set_column('B:H', 20)
            melonmynsp.write_formula('F36','=AVERAGE(F4:F35)',percentage)
            for persum in ['C','D','E']:
                melonmynsp.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
            for persum in ['G','H']:
                melonmynsp.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for r, row in enumerate(readCSV, start=3):
                    for c, col in enumerate(row):
                        melonmynsp.write(r,1, row[0], date_format)
                        melonmynsp.write(r,2, row[1], content)
                        melonmynsp.write(r,3, row[2], content)
                        melonmynsp.write(r,4, row[3], content)
                        melonmynsp.write(r,5, row[4], percentage)
                        melonmynsp.write(r,6, row[5], content)
                        melonmynsp.write(r,7, row[6], content)
                        
            cursor.close()

        sys = '1'
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
    
    def ERROR_CHANNEL():
        cursor = database.cursor()
        with open('MELON_MYNSP_ERRORCHANNEL.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        melonmynsp.write('B38','ERROR CHANNEL',title)
        melonmynsp.write('B39','CALLING_DATE',bold)
        melonmynsp.write('C39','CHANNEL',bold)
        melonmynsp.write('D39','ERROR_RESULT',bold)
        melonmynsp.write('E39','ERROR_COUNT',bold)
        
        for r, row in enumerate(result, start=39):
            for c, col in enumerate(row):
                melonmynsp.write(r,1, row[0], date_format)
                melonmynsp.write(r,2, row[1], content)
                melonmynsp.write(r,3, row[2], content)
                melonmynsp.write(r,4, row[3], content)
        cursor.close()
    
    def ERROR_PRICE():
        cursor = database.cursor()
        with open('MELON_MYNSP_ERRORPRICE.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        melonmynsp.write('G38','ERROR CHANNEL',title)
        melonmynsp.write('G39','CALLING_DATE',bold)
        melonmynsp.write('H39','PRICE',bold)
        melonmynsp.write('I39','ERROR_RESULT',bold)
        melonmynsp.write('J39','ERROR_COUNT',bold)
        melonmynsp.set_column('I:J', 20)
        
        for r, row in enumerate(result, start=39):
            for c, col in enumerate(row):
                melonmynsp.write(r,6, row[0], date_format)
                melonmynsp.write(r,7, row[1], content)
                melonmynsp.write(r,8, row[2], content)
                melonmynsp.write(r,9, row[3], content)
        cursor.close()
        chat = 'MELON MYNSP Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)

    Statistic()
    ERROR_CHANNEL()
    ERROR_PRICE()


# In[31]:


def SUSPEND_TOP20():
        cursor = database.cursor()
        with open('SUSPEND_TOP20.txt','r') as inserts:
            query = inserts.read()
            
        # Exec query new_subs
        cursor.execute(query)
        result = cursor.fetchall()
        suspendtop20.set_zoom(70)
        suspendtop20.hide_gridlines(2)
        suspendtop20.write('B3','RANK',bold)
        suspendtop20.write('C3','Content ID',bold)
        suspendtop20.write('D3','JUDUL LAGU',bold)
        suspendtop20.write('E3','TOTAL MSISDN',bold)
        suspendtop20.set_column('B:B', 10)
        suspendtop20.set_column('C:C', 30)
        suspendtop20.set_column('D:D', 60)
        suspendtop20.set_column('E:E', 30)
        
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                suspendtop20.write(r,1, row[0], date_format)
                suspendtop20.write(r,2, row[1], content)
                suspendtop20.write(r,3, row[2], content)
                suspendtop20.write(r,4, row[3], content)
                
        cursor.close()
        chat = 'SUSPEND_TOP20 Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)


# In[32]:


def REGION():
    def TRAFFIC_AND_REVENUE_PER_REGION():
        def GETDAT():
            cursor = database.cursor()
            query = ("select calling_date ,T_Sumbagut,SUMBAGUT,T_SUMBAGTENG,Sumbagteng ,T_Sumbagsel ,Sumbagsel ,T_Metro_Jakarta, Metro_Jakarta ,                     T_Outer_Jakarta , Outer_Jakarta ,T_Jabar , Jabar , T_Jateng_DIY , Jateng_DIY ,T_Jatim , Jatim , T_Balinusra ,Balinusra ,                     T_Kalimantan , Kalimantan , T_Sulawesi , Sulawesi , T_Papua_Maluku , Papua_Maluku , T_unknown , unknown , Traffic_total ,                     revenue_total                     FROM                     (                     select calling_date,                         sum(case when region='Sumbagut' then TOTAL_SUCCESS else 0 end)  T_Sumbagut,                             sum(case when region='Sumbagut' then TOTAL_REVENUE else 0 end)  Sumbagut,                             sum(case when region='Sumbagteng' then TOTAL_SUCCESS else 0 end) T_Sumbagteng,                         sum(case when region='Sumbagteng' then TOTAL_REVENUE else 0 end) Sumbagteng,                             sum(case when region='Sumbagsel' then TOTAL_SUCCESS else 0 end)  T_Sumbagsel,                             sum(case when region='Sumbagsel' then TOTAL_REVENUE else 0 end)  Sumbagsel,                         sum(case when region='Jabotabek'  and sub_region = 'Jakarta' then TOTAL_SUCCESS else 0 end) T_Metro_Jakarta,                         sum(case when region='Jabotabek'  and sub_region = 'Jakarta' then TOTAL_REVENUE else 0 end) Metro_Jakarta,                             sum(case when region='Jabotabek'  and  sub_region <> 'Jakarta' then TOTAL_SUCCESS else 0 end) T_Outer_Jakarta,                         sum(case when region='Jabotabek'  and  sub_region <> 'Jakarta' then TOTAL_REVENUE else 0 end) Outer_Jakarta,                             sum(case when region='Jabar' then TOTAL_SUCCESS else 0 end) T_Jabar,                         sum(case when region='Jabar' then TOTAL_REVENUE else 0 end) Jabar,                             sum(case when region='Jateng' then TOTAL_SUCCESS else 0 end) T_Jateng_DIY,                         sum(case when region='Jateng' then TOTAL_REVENUE else 0 end) Jateng_DIY,                             sum(case when region='Jatim' then TOTAL_SUCCESS else 0 end) T_Jatim,                         sum(case when region='Jatim' then TOTAL_REVENUE else 0 end) Jatim,                             sum(case when region='Balinusra' then TOTAL_SUCCESS else 0 end) T_Balinusra,                         sum(case when region='Balinusra' then TOTAL_REVENUE else 0 end) Balinusra,                             sum(case when region='Kalimantan' then TOTAL_SUCCESS else 0 end) T_Kalimantan,                         sum(case when region='Kalimantan' then TOTAL_REVENUE else 0 end) Kalimantan,                             sum(case when region='Sulawesi' then TOTAL_SUCCESS else 0 end)   T_Sulawesi,                         sum(case when region='Sulawesi' then TOTAL_REVENUE else 0 end)   Sulawesi,                              sum(case when region='Papua' then TOTAL_SUCCESS else 0 end)  T_Papua_Maluku,                         sum(case when region='Papua' then TOTAL_REVENUE else 0 end)  Papua_Maluku,                         sum(case when region='UNKNOWN' then TOTAL_SUCCESS else 0 end)  T_unknown,                         sum(case when region='UNKNOWN' then TOTAL_REVENUE else 0 end)  unknown,                             sum(TOTAL_SUCCESS)  Traffic_total,                         sum(TOTAL_REVENUE )  revenue_total                               from CDR_ACTIVATION_SUMMARY_DAY                     where  calling_date = to_char(sysdate-"+sys+",'yyyymmdd')                     group by calling_date                     order by calling_date)")
            cursor.execute(query)
            Dump = 'TRAFFIC_AND_REVENUE_PER_REGION_'+str(kmrnbulan[0:6])+'.csv'
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')

            for line in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 5] * 33))
                    region.write(Alline,'',neat)

            region.set_zoom(60)
            region.hide_gridlines(2)
            region.write('A1','1.TRAFFIC AND REVENUE PER REGION',title)
            region.merge_range('A3:A4','CALLING_DATE',bold)
            region.merge_range('B3:C3','Sumbagut',bold)
            region.write('B4','TRAFFIC',bold)
            region.write('C4','REVENUE',bold)
            region.merge_range('D3:E3','Sumbagteng',bold)
            region.write('D4','TRAFFIC',bold)
            region.write('E4','REVENUE',bold)
            region.merge_range('F3:G3','Sumbagsel',bold)
            region.write('F4','TRAFFIC',bold)
            region.write('G4','REVENUE',bold)
            region.merge_range('H3:I3','Metro Jakarta',bold)
            region.write('H4','TRAFFIC',bold)
            region.write('I4','REVENUE',bold)
            region.merge_range('J3:K3','Outer Jakarta',bold)
            region.write('J4','TRAFFIC',bold)
            region.write('K4','REVENUE',bold)
            region.merge_range('L3:M3','Jabar',bold)
            region.write('L4','TRAFFIC',bold)
            region.write('M4','REVENUE',bold)
            region.merge_range('N3:O3','Jateng DIY',bold)
            region.write('N4','TRAFFIC',bold)
            region.write('O4','REVENUE',bold)
            region.merge_range('P3:Q3','Jatim',bold)
            region.write('P4','TRAFFIC',bold)
            region.write('Q4','REVENUE',bold)
            region.merge_range('R3:S3','Balinusra',bold)
            region.write('R4','TRAFFIC',bold)
            region.write('S4','REVENUE',bold)
            region.merge_range('T3:U3','Kalimantan',bold)
            region.write('T4','TRAFFIC',bold)
            region.write('U4','REVENUE',bold)
            region.merge_range('V3:W3','Sulawesi',bold)
            region.write('V4','TRAFFIC',bold)
            region.write('W4','REVENUE',bold)
            region.merge_range('X3:Y3','Maluku Papua',bold)
            region.write('X4','TRAFFIC',bold)
            region.write('Y4','REVENUE',bold)
            region.merge_range('Z3:AA3','Unknown',bold)
            region.write('Z4','TRAFFIC',bold)
            region.write('AA4','REVENUE',bold)
            region.merge_range('AB3:AC3','ALL_Total',bold)
            region.write('AB4','TRAFFIC',bold)
            region.write('AC4','REVENUE',bold)
            region.write('A37','TOTAL',neat)
            region.set_column('A:AC', 25)
            for persum in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC']:
                region.write_formula(persum+'37','=SUM('+persum+'5:'+persum+'36)',content)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for r, row in enumerate(readCSV, start=4):
                    for c, col in enumerate(row):
                        region.write(r,0, row[0], date_format)
                        region.write(r,1, row[1], content)
                        region.write(r,2, row[2], content)
                        region.write(r,3, row[3], content)
                        region.write(r,4, row[4], content)
                        region.write(r,5, row[5], content)
                        region.write(r,6, row[6], content)
                        region.write(r,7, row[7], content)
                        region.write(r,8, row[8], content)
                        region.write(r,9, row[9], content)
                        region.write(r,10, row[10], content)
                        region.write(r,11, row[11], content)
                        region.write(r,12, row[12], content)
                        region.write(r,13, row[13], content)
                        region.write(r,14, row[14], content)
                        region.write(r,15, row[15], content)
                        region.write(r,16, row[16], content)
                        region.write(r,17, row[17], content)
                        region.write(r,18, row[18], content)
                        region.write(r,19, row[19], content)
                        region.write(r,20, row[20], content)
                        region.write(r,21, row[21], content)
                        region.write(r,22, row[22], content)
                        region.write(r,23, row[23], content)
                        region.write(r,24, row[24], content)
                        region.write(r,25, row[25], content)
                        region.write(r,26, row[26], content)
                        region.write(r,27, row[27], content)
                        region.write(r,28, row[28], content)
                        
            cursor.close()

        sys = '1'
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
    
    def UNIQ_SUBS_TRANSACTION():
        def GETDAT():
            cursor = database.cursor()
            query = ("SELECT CALLING_DATE ,SUMBAGUT , SUMBAGTENG , SUMBAGSEL ,       METRO_JAKARTA ,OUTER_JAKARTA ,      JABAR,    JATENG_DIY,             JATIM ,  BALINUSRA ,       KALIMANTAN ,      SULAWESI ,PAPUA_MALUKU ,UNKNOWN     ,TOTAL_UNIQ_SUBS             FROM (             select calling_date,             count(distinct case when region='Sumbagut' then CALLING_MDN end)  Sumbagut,             count(distinct case when region='Sumbagteng' then CALLING_MDN end) Sumbagteng,             count(distinct case when region='Sumbagsel' then CALLING_MDN end)  Sumbagsel,             count(distinct case when region='Jabotabek'  and sub_region = 'Jakarta' then CALLING_MDN end) Metro_Jakarta,             count(distinct case when region='Jabotabek'  and  sub_region <> 'Jakarta' then CALLING_MDN end) Outer_Jakarta,             count(distinct case when region='Jabar' then CALLING_MDN end) Jabar,             count(distinct case when region='Jateng' then CALLING_MDN end) Jateng_DIY,             count(distinct case when region='Jatim' then CALLING_MDN end) Jatim,             count(distinct case when region='Balinusra' then CALLING_MDN end) Balinusra,             count(distinct case when region='Kalimantan' then CALLING_MDN end) Kalimantan,             count(distinct case when region='Sulawesi' then CALLING_MDN end)   Sulawesi,             count(distinct case when region='Papua' then CALLING_MDN end)  Papua_Maluku,             count(distinct case when region='UNKNOWN' then CALLING_MDN end)  unknown,             count(distinct calling_mdn )  total_uniq_subs             from CALL_DETAILED_RECORD a, prefix_region b             where  calling_date = to_char(sysdate-"+sys+",'yyyymmdd')             and calling_type in (1,2,7,77,78,64,93,65,95,96,99,100)             and result in (805,0)             and substr(a.calling_mdn,1,7) = b.prefix             group by calling_date             order by calling_date)")
            cursor.execute(query)
            Dump = 'UNIQ_SUBS_TRANSACTION_'+str(kmrnbulan[0:6])+'.csv'
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')

            for line in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 44] * 33))
                    region.write(Alline,'',neat)
                    
            region.write('A41','2. UNIQ SUBS TRANSACTION',title)
            region.write('A43','CALLING_DATE',bold)
            region.write('B43','SUMBAGUT',bold)
            region.write('C43','SUMBAGTENG',bold)
            region.write('D43','SUMBAGSEL',bold)
            region.write('E43','METRO_JAKARTA',bold)
            region.write('F43','OUTER_JAKARTA',bold)
            region.write('G43','JABAR',bold)
            region.write('H43','JATENG_DIY',bold)
            region.write('I43','JATIM',bold)
            region.write('J43','BALINUSRA',bold)
            region.write('K43','KALIMANTAN',bold)
            region.write('L43','SULAWESI',bold)
            region.write('M43','PAPUA_MALUKU',bold)
            region.write('N43','UNKNOWN',bold)
            region.write('O43','TOTAL_UNIQ_SUBS',bold)
            region.write('A76','TOTAL',neat)
            for persum in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O']:
                region.write_formula(persum+'76','=SUM('+persum+'44:'+persum+'74)',content)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for r, row in enumerate(readCSV, start=43):
                    for c, col in enumerate(row):
                        region.write(r,0, row[0], date_format)
                        region.write(r,1, row[1], content)
                        region.write(r,2, row[2], content)
                        region.write(r,3, row[3], content)
                        region.write(r,4, row[4], content)
                        region.write(r,5, row[5], content)
                        region.write(r,6, row[6], content)
                        region.write(r,7, row[7], content)
                        region.write(r,8, row[8], content)
                        region.write(r,9, row[9], content)
                        region.write(r,10, row[10], content)
                        region.write(r,11, row[11], content)
                        region.write(r,12, row[12], content)
                        region.write(r,13, row[13], content)
                        region.write(r,14, row[14], content)
                        
            cursor.close()

        sys = '1'
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
            
    def UNIQ_SUBS_TRANSACTION_PROCESS_TYPE():
        def GETDAT():
            cursor = database.cursor()
            query = ("select CALLING_DATE,PROCESS_TYPE,Sumbagut,Sumbagteng,Sumbagsel,Metro_Jakarta,Outer_Jakarta,Jabar,Jateng_DIY,Jatim,Balinusra,Kalimantan,Sulawesi,             Papua_Maluku,OTHERS FROM (select a.calling_date as CALLING_DATE,             CASE when a.calling_type in (1,2,7,77) then 'Purchase'             when a.calling_type in (64,78,65,93,95,96,99,100) then 'Renewal' end as PROCESS_TYPE,             count(distinct case when b.region = 'Sumbagut' then a.calling_mdn end) as Sumbagut,             count(distinct case when b.region = 'Sumbagteng' then  a.calling_mdn end) as Sumbagteng,             count(distinct case when b.region = 'Sumbagsel' then  a.calling_mdn end) as Sumbagsel,             count(distinct case when b.region = 'Jabotabek'  and sub_region = 'Jakarta' then CALLING_MDN end) Metro_Jakarta,             count(distinct case when b.region = 'Jabotabek'  and  sub_region <> 'Jakarta' then CALLING_MDN end) Outer_Jakarta,             count(distinct case when b.region = 'Jabar' then  a.calling_mdn end) as Jabar,             count(distinct case when b.region = 'Jateng' then  a.calling_mdn end) as Jateng_DIY,             count(distinct case when b.region = 'Jatim' then  a.calling_mdn end) as Jatim,             count(distinct case when b.region = 'Balinusra' then a.calling_mdn end) as Balinusra,             count(distinct case when b.region = 'Kalimantan' then a.calling_mdn end) as Kalimantan,             count(distinct case when b.region = 'Sulawesi' then a.calling_mdn end) as Sulawesi,             count(distinct case when b.region = 'Papua' then a.calling_mdn end) as Papua_Maluku,             count(distinct case when b.region not in ('Sumbagut','Sumbagsel','Sumbagteng','Jabotabek','Jabar','Jateng','Jatim','Balinusra','Sulawesi','Kalimantan','Papua') then             a.calling_mdn end) as Others             from call_detailed_record a, prefix_region b             where a.calling_date = ('"+kmrnbulan+"')             and substr(a.calling_mdn,1,7)=b.prefix             and calling_type in (1,2,7,77,64,65,78,93,95,96,99,100)             and result=0 group BY a.calling_date, case             when a.calling_type in (1,2,7,77) then 'Purchase'             when a.calling_type in (64,78,65,93,95,96,99,100) then 'Renewal'             end)")
            cursor.execute(query)
            Dump = 'UNIQ_SUBS_TRANSACTION_PROCESS_TYPE_'+str(kmrnbulan[0:6])+'.csv'
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')

            for line in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 83] * 33))
                    region.write(Alline,'',neat)
                    
            region.write('A80','3. UNIQ SUBS TRANSACTION PROCESS TYPE',title)
            region.write('A82','CALLING_DATE',bold)
            region.write('B82','PROCESS_TYPE',bold)
            region.write('C82','SUMBAGUT',bold)
            region.write('D82','SUMBAGTENG',bold)
            region.write('E82','SUMBAGSEL',bold)
            region.write('F82','METRO_JAKARTA',bold)
            region.write('G82','OUTER_JAKARTA',bold)
            region.write('H82','JABAR',bold)
            region.write('I82','JATENG_DIY',bold)
            region.write('J82','JATIM',bold)
            region.write('K82','BALINUSRA',bold)
            region.write('L82','KALIMANTAN',bold)
            region.write('M82','SULAWESI',bold)
            region.write('N82','PAPUA_MALUKU',bold)
            region.write('O82','UNKNOWN',bold)
            region.write('P82','TOTAL_UNIQ_SUBS',bold)
            region.write('A115','TOTAL',neat)
            for persum in ['B','C','D','E','F','G','H','I','J','K','L','M','N','O','P']:
                region.write_formula(persum+'115','=SUM('+persum+'83:'+persum+'114)',content)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for r, row in enumerate(readCSV, start=82):
                    for c, col in enumerate(row):
                        region.write(r,0, row[0], date_format)
                        region.write(r,1, row[1], content)
                        region.write(r,2, row[2], content)
                        region.write(r,3, row[3], content)
                        region.write(r,4, row[4], content)
                        region.write(r,5, row[5], content)
                        region.write(r,6, row[6], content)
                        region.write(r,7, row[7], content)
                        region.write(r,8, row[8], content)
                        region.write(r,9, row[9], content)
                        region.write(r,10, row[10], content)
                        region.write(r,11, row[11], content)
                        region.write(r,12, row[12], content)
                        region.write(r,13, row[13], content)
                        region.write(r,14, row[14], content)
                        
            cursor.close()
            chat = 'REGION Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
            fyi(chat)

        sys = '1'
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
            
    TRAFFIC_AND_REVENUE_PER_REGION()
    UNIQ_SUBS_TRANSACTION()
    UNIQ_SUBS_TRANSACTION_PROCESS_TYPE()


# In[33]:


def DTMF():
    def DTMF_ONLY():
        def GETDAT():
            cursor = database.cursor()
            query = ("select calling_date , sum(trans_count)  ,             sum(case when result in (0,805) then trans_count else 0 end) ,             sum(case when result in (0,805) then revenue else 0 end ) ,             sum(case when result not in (0) then trans_count else 0 end)             from cdr_report_goofree             where calling_type in (1,77)             and calling_date = to_char(sysdate-"+sys+",'yyyymmdd')             and media_code like '%D'             group by calling_date, (case when media_code like '%D' then 'dtmf' else             null end) order by 1 asc")
            cursor.execute(query)
            Dump = 'DTMF_ONLY_'+str(kmrnbulan[0:6])+'.csv'
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                     
            for line in ['A','B','C','D','E']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 4] * 33))
                    dtmf.write(Alline,'',neat)
            
            dtmf.set_zoom(60)
            dtmf.hide_gridlines(2)
            dtmf.write('A2','DTMF ONLY',title)
            dtmf.write('A3','CALLING_DATE',bold)
            dtmf.write('B3','TOTAL_HITS',bold)
            dtmf.write('C3','SUCCESS',bold)
            dtmf.write('D3','REVENUE',bold)
            dtmf.write('E3','FAILED',bold)
            dtmf.write('A37','TOTAL',neat)
            dtmf.set_column('A:E', 15)
            for persum in ['B','C','D','E']:
                dtmf.write_formula(persum+'37','=SUM('+persum+'4:'+persum+'36)',content)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for r, row in enumerate(readCSV, start=3):
                    for c, col in enumerate(row):
                        dtmf.write(r,0, row[0], date_format)
                        dtmf.write(r,1, row[1], content)
                        dtmf.write(r,2, row[2], content)
                        dtmf.write(r,3, row[3], content)
                        dtmf.write(r,4, row[4], content)
                        
            cursor.close()

        sys = '1'
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
    
    def DTMF_AREA():
        def GETDAT():
            cursor = database.cursor()
            with open('DTMF_AREA.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            dtmf.merge_range('G3:G4','AREA',bold)
            dtmf.merge_range('H3:I3','DTMF',bold)
            dtmf.write('H4','Traffic',bold)
            dtmf.write('I4','Revenue',bold)
            dtmf.set_column('G:I', 12)
            for r, row in enumerate(result, start=4):
                    for c, col in enumerate(row):
                        dtmf.write(r,6, row[0], date_format)
                        dtmf.write(r,7, row[1], content)
                        dtmf.write(r,8, row[2], content)
                        
            cursor.close()
            
        GETDAT()
    
    def DTMF_REGION():
        def GETDAT():
            cursor = database.cursor()
            query = ("select calling_date ,Trx_Sumbagut,Rev_Sumbagut,Trx_Sumbagteng, Rev_Sumbagteng ,Trx_Sumbagsel , Rev_Sumbagsel ,             Trx_Jabotabek, Rev_Jabotabek , Trx_Jabar , Rev_Jabar , Trx_Jateng_DIY , Rev_Jateng_DIY , Trx_Jatim ,  Rev_Jatim , Trx_Balinusra             ,Rev_Balinusra ,Trx_Kalimantan , Rev_Kalimantan , Trx_Sulawesi , Rev_Sulawesi , Trx_Papua_Maluku , Rev_Papua_Maluku, Trx_Other             , Rev_Other , Traffic_total , revenue_total FROM (             select calling_date,                 sum(case when region='Sumbagut' then trans_count else 0 end)  Trx_Sumbagut,                     sum(case when region='Sumbagut' then revenue else 0 end)  Rev_Sumbagut,                     sum(case when region='Sumbagteng' then trans_count else 0 end) Trx_Sumbagteng,                 sum(case when region='Sumbagteng' then revenue else 0 end) Rev_Sumbagteng,                     sum(case when region='Sumbagsel' then trans_count else 0 end)  Trx_Sumbagsel,                     sum(case when region='Sumbagsel' then revenue else 0 end)  Rev_Sumbagsel,                 sum(case when region='Jabotabek' then trans_count else 0 end) Trx_Jabotabek,                 sum(case when region='Jabotabek'  then revenue else 0 end) Rev_Jabotabek,                     sum(case when region='Jabar' then trans_count else 0 end) Trx_Jabar,                 sum(case when region='Jabar' then revenue else 0 end) Rev_Jabar,                     sum(case when region='Jateng' then trans_count else 0 end) Trx_Jateng_DIY,                 sum(case when region='Jateng' then revenue else 0 end) Rev_Jateng_DIY,                     sum(case when region='Jatim' then trans_count else 0 end) Trx_Jatim,                 sum(case when region='Jatim' then revenue else 0 end) Rev_Jatim,                     sum(case when region='Balinusra' then trans_count else 0 end) Trx_Balinusra,                 sum(case when region='Balinusra' then revenue else 0 end) Rev_Balinusra,                     sum(case when region='Kalimantan' then trans_count else 0 end) Trx_Kalimantan,                 sum(case when region='Kalimantan' then revenue else 0 end) Rev_Kalimantan,                     sum(case when region='Sulawesi' then trans_count else 0 end)   Trx_Sulawesi,                 sum(case when region='Sulawesi' then revenue else 0 end)   Rev_Sulawesi,                      sum(case when region='Papua' then trans_count else 0 end)  Trx_Papua_Maluku,                 sum(case when region='Papua' then revenue else 0 end)  Rev_Papua_Maluku,                 sum(case when region not in ('Sumbagut','Sumbagteng','Sumbagsel','Jabotabek','Jabar','Jateng','Jatim','Balinusra','Kalimantan','Sulawesi','Papua') then             trans_count else 0 end)  Trx_Other,                 sum(case when region not in ('Sumbagut','Sumbagteng','Sumbagsel','Jabotabek','Jabar','Jateng','Jatim','Balinusra','Kalimantan','Sulawesi','Papua') then             revenue else 0 end)  Rev_Other,                     sum(trans_count)  Traffic_total,                 sum(revenue )  revenue_total                       from cdr_report_goofree             where  calling_date = to_char(sysdate-"+sys+",'yyyymmdd')             and calling_type in (1,77)             and media_code like '%D'             and result in (0,805)             group by calling_date             order by calling_date)")
            cursor.execute(query)
            Dump = 'DTMF_REGION_'+str(kmrnbulan[0:6])+'.csv'
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                     
            for line in ['K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 4] * 33))
                    dtmf.write(Alline,'',neat)
            dtmf.write('K3','CALLING_DATE',bold)
            dtmf.write('L3','T_SUMBAGUT',bold)
            dtmf.write('M3','SUMBAGUT',bold)
            dtmf.write('N3','T_SUMBAGTENG',bold)
            dtmf.write('O3','SUMBAGTENG',bold)
            dtmf.write('P3','T_SUMBAGSEL',bold)
            dtmf.write('Q3','SUMBAGSEL',bold)
            dtmf.write('R3','T_JABOTABEK',bold)
            dtmf.write('S3','JABOTABEK',bold)
            dtmf.write('T3','T_JABAR',bold)
            dtmf.write('U3','JABAR',bold)
            dtmf.write('V3','T_JATENG DIY',bold)
            dtmf.write('W3','JATENG DIY',bold)
            dtmf.write('X3','T_JATIM',bold)
            dtmf.write('Y3','JATIM',bold)
            dtmf.write('Z3','T_BALINUSRA',bold)
            dtmf.write('AA3','BALINUSRA',bold)
            dtmf.write('AB3','T_KALIMANTAN',bold)
            dtmf.write('AC3','KALIMANTAN',bold)
            dtmf.write('AD3','T_SULAWESI',bold)
            dtmf.write('AE3','SULAWESI',bold)
            dtmf.write('AF3','T_MALUKU_PAPUA',bold)
            dtmf.write('AG3','MALUKU_PAPUA',bold)
            dtmf.write('AH3','T_Other',bold)
            dtmf.write('AI3','Rev_Other',bold)
            dtmf.write('AJ3','TRAFFIC_TOTAL',bold)
            dtmf.write('AK3','REVENUE_TOTAL',bold)
            dtmf.write('K36','TOTAL',neat)
            dtmf.set_column('K:AK', 20)
            for persum in ['L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK']:
                dtmf.write_formula(persum+'36','=SUM('+persum+'4:'+persum+'35)',content)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                readCSV = csv.reader(csvfile, delimiter=',')
                for r, row in enumerate(readCSV, start=3):
                    for c, col in enumerate(row):
                        dtmf.write(r,10, row[0], date_format)
                        dtmf.write(r,11, row[1], content)
                        dtmf.write(r,12, row[2], content)
                        dtmf.write(r,13, row[3], content)
                        dtmf.write(r,14, row[4], content)
                        dtmf.write(r,15, row[5], content)
                        dtmf.write(r,16, row[6], content)
                        dtmf.write(r,17, row[7], content)
                        dtmf.write(r,18, row[8], content)
                        dtmf.write(r,19, row[9], content)
                        dtmf.write(r,20, row[10], content)
                        dtmf.write(r,21, row[11], content)
                        dtmf.write(r,22, row[12], content)
                        dtmf.write(r,23, row[13], content)
                        dtmf.write(r,24, row[14], content)
                        dtmf.write(r,25, row[15], content)
                        dtmf.write(r,26, row[16], content)
                        dtmf.write(r,27, row[17], content)
                        dtmf.write(r,28, row[18], content)
                        dtmf.write(r,29, row[19], content)
                        dtmf.write(r,30, row[20], content)
                        dtmf.write(r,31, row[21], content)
                        dtmf.write(r,32, row[22], content)
                        dtmf.write(r,33, row[23], content)
                        dtmf.write(r,34, row[24], content)
                        dtmf.write(r,35, row[25], content)
                        dtmf.write(r,36, row[26], content)
                        
            cursor.close()
            chat = 'DTMF Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
            fyi(chat)
        
        sys = '1'
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
                     
    DTMF_ONLY()
    DTMF_AREA()
    DTMF_REGION()


# In[34]:


def GP():
    def GETDAT():
        cursor = database.cursor()
        query = ("select calling_date, sum(case when result = 0 then 1 end),         sum(case when result = 0 then charge_amount end) as revenue         from call_detailed_record_all where calling_date = to_char(sysdate-"+sys+",'YYYYMMDD')         and calling_type = '95' group by calling_date order by calling_date")
        cursor.execute(query)
        Dump = 'GP_'+str(kmrnbulan[0:6])+'.csv'
        df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                 
        for line in ['B','C','D']:
            for kolom in range(33):
                Alline = line+'{}'.format(*([kolom + 5] * 33))
                gp.write(Alline,'',neat)
        
        gp.set_zoom(80)
        gp.hide_gridlines(2)
        gp.write('B2','GRACE PERIODE',title)
        gp.write('B4','CALLING_DATE',bold)
        gp.write('C4','GP Subs Success',bold)
        gp.write('D4','GP REV',bold)
        gp.write('B37','TOTAL',neat)
        gp.set_column('B:B', 16)
        gp.set_column('C:C', 17)
        gp.set_column('D:D', 13)
        for persum in ['C','D']:
            gp.write_formula(persum+'37','=SUM('+persum+'5:'+persum+'36)',content)
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=4):
                for c, col in enumerate(row):
                    gp.write(r,1, row[0], date_format)
                    gp.write(r,2, row[1], content)
                    gp.write(r,3, row[2], content)
                    
        cursor.close()
        chat = 'GRACE PERIODE Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
        
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[35]:


def SMS_REVENUE():
    def GETDAT():
        cursor = database.cursor()
        query = ("select calling_date,count(*),sum(case when result=0 then 1 else 0 end),                 sum(case when result=0 then charge_amount else 0 end) from cdr_all_in_one                 where calling_date= to_char(sysdate-"+sys+", 'YYYYMMDD')                 and calling_type in (1,2,7,77) and media_code like '%0'                 group by calling_date")
        cursor.execute(query)
        Dump = 'SMS_REVENUE_'+str(kmrnbulan[0:6])+'.csv'
        df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                 
        for line in ['A','B','C','D']:
            for kolom in range(33):
                Alline = line+'{}'.format(*([kolom + 5] * 33))
                smsrev.write(Alline,'',neat)
        
        smsrev.set_zoom(80)
        smsrev.hide_gridlines(2)
        smsrev.merge_range('A2:D2','SMS REVENUE',title)
        smsrev.write('A4','CALLING_DATE',bold)
        smsrev.write('B4','SMS_ATTEMPT_COUNT',bold)
        smsrev.write('C4','SMS_SUCC_COUNT',bold)
        smsrev.write('D4','SMS_REVENUE',bold)
        smsrev.write('A37','TOTAL',neat)
        smsrev.set_column('A:D', 25)
        for persum in ['B','C','D']:
            smsrev.write_formula(persum+'37','=SUM('+persum+'5:'+persum+'36)',content)
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=4):
                for c, col in enumerate(row):
                    smsrev.write(r,0, row[0], date_format)
                    smsrev.write(r,1, row[1], content)
                    smsrev.write(r,2, row[2], content)
                    smsrev.write(r,3, row[3], content)
        cursor.close()
        chat = 'SMS REVENUE Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
        
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[36]:


def REVENUE_RENEWAL():
    def GETDAT():
        cursor = database.cursor()
        query = ("select calling_date,sum(charge_amount) from cdr_all_in_one where calling_date=to_char(sysdate-"+sys+",'YYYYMMDD')                 and result=0 and calling_type in (64,93,95,96,99,100) and rating_id in (select rating_id from rating_mcp_info                 where initial_amount=0 and RECURRENT_RATING_CYCLE_DAY=7 and recurrent_amount between 600 and 3000) group by calling_date")
        cursor.execute(query)
        Dump = 'REVENUE_RENEWAL_'+str(kmrnbulan[0:6])+'.csv'
        df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                 
        for line in ['B','C']:
            for kolom in range(33):
                Alline = line+'{}'.format(*([kolom + 3] * 33))
                revrenew.write(Alline,'',neat)
        
        revrenew.set_zoom(80)
        revrenew.hide_gridlines(2)
        revrenew.write('B2','CALLING_DATE',bold)
        revrenew.write('C2','REVENUE RENEWAL ON RP.0',bold)
        revrenew.write('B35','TOTAL',neat)
        revrenew.set_column('B:B', 16)
        revrenew.set_column('C:C', 35)
        for persum in ['C']:
            revrenew.write_formula(persum+'35','=SUM('+persum+'3:'+persum+'34)',content)
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=2):
                for c, col in enumerate(row):
                    revrenew.write(r,1, row[0], date_format)
                    revrenew.write(r,2, row[1], content)
        cursor.close()
        chat = 'REVENUE RENEWAL Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
        
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[37]:


def NSP_REKOMENDASI(): # https://pandas.pydata.org/docs/user_guide/merging.html
    def GETDAT():
        Dump = 'NSP_Rekomendasi_'+str(kmrnbulan[0:6])+'.csv'
        stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/'+kmrnbulan+'/25.NSP_Rekomendasi_'+kmrnbulan+'_Traffic.csv')
        df1 = pd.read_csv(stdout, sep=',')
        # df1 = pd.read_csv(stdout, sep=',',usecols=["CALLING_DATE"])
        stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/'+kmrnbulan+'/25.NSP_Rekomendasi_'+kmrnbulan+'.csv')
        df2 = pd.read_csv(stdout, sep=',',usecols=["Total_SMS_Recommendation_Sent","Total_Attempt_Purchased","Total_Success_Purchased","Total_New_Subscriber_Activated","Total_Revenue_Purchased"])
        df = pd.concat([df1, df2], axis=1)
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
        for line in ['B','C','D','E','F','G','H','I','J','K','L']:
            for kolom in range(350):
                Alline = line+'{}'.format(*([kolom + 2] * 350))
                nspreco.write(Alline,'',neat)
        nspreco.set_zoom(50)
        nspreco.hide_gridlines(2)
        nspreco.write('B1','CALLING_DATE',bold)
        nspreco.write('C1','REGION',bold)
        nspreco.write('D1','TELKOMSEL TRAFFIC',bold)
        nspreco.write('E1','UNIQ TELKOMSEL TRAFFIC',bold)
        nspreco.write('F1','Total SMS Recommendation Sent',bold)
        nspreco.write('G1','Total Attempt Purchased',bold)
        nspreco.write('H1','Total Success Purchased',bold)
        nspreco.write('I1','Total New Subscriber Activated',bold)
        nspreco.write('J1','Total Revenue Purchased',bold)
        nspreco.write('K1','SR Purchasing',bold)
        nspreco.write('L1','SR Charging',bold)
        nspreco.freeze_panes(1, 1)
        nspreco.set_column('B:C', 18)
        nspreco.set_column('D:L', 50)
        for kolom in range(350):
            SRP = '{}'.format(*([kolom + 2] * 350))
            nspreco.write_formula('K'+SRP,'=I'+SRP+'/H'+SRP+'',percentage)
            nspreco.write_formula('L'+SRP,'=H'+SRP+'/G'+SRP+'',percentage)
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=1):
                for c, col in enumerate(row):
                    nspreco.write(r,1, row[0], date_format)
                    nspreco.write(r,2, row[1], content)
                    nspreco.write(r,3, row[2], content)
                    nspreco.write(r,4, row[3], content)
                    nspreco.write(r,5, row[4], content)
                    nspreco.write(r,6, row[5], content)
                    nspreco.write(r,7, row[6], content)
                    nspreco.write(r,8, row[7], content)
                    nspreco.write(r,9, row[8], content)
        chat = 'NSP REKOMENDASI Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
        
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[38]:


def SUCCESS_RATE_MT_AND_MO():
    def GETDAT():
        Dump = 'success_rate_MT_MO_'+str(kmrnbulan[0:6])+'.csv'
        stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/'+kmrnbulan+'/26.success_rate_MT_MO_'+kmrnbulan+'.csv')
        df = pd.read_csv(stdout, sep=',')
        df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
        
        for line in ['B','C','D','E','F']:
            for kolom in range(70):
                Alline = line+'{}'.format(*([kolom + 4] * 70))
                succratemt.write(Alline,'',neat)
        succratemt.set_zoom(50)
        succratemt.hide_gridlines(2)
        succratemt.write('B3','CALLING_DATE',bold)
        succratemt.write('C3','Type',bold)
        succratemt.write('D3','Attempt',bold)
        succratemt.write('E3','Success',bold)
        succratemt.write('F3','SR',bold)
        succratemt.set_column('B:F', 15)
        for kolom in range(70):
            SRP = '{}'.format(*([kolom + 4] * 70))
            succratemt.write_formula('F'+SRP,'=E'+SRP+'/D'+SRP+'',percentage)
        with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
            readCSV = csv.reader(csvfile, delimiter=',')
            for r, row in enumerate(readCSV, start=3):
                for c, col in enumerate(row):
                    succratemt.write(r,1, row[0], date_format)
                    succratemt.write(r,2, row[1], content)
                    succratemt.write(r,3, row[2], content)
                    succratemt.write(r,4, row[3], content)
        chat = 'SUCCESS RATE MT AND MO Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)
        
    sys = '1'
    lusa = now - datetime.timedelta(days=2)
    storyday = now - datetime.timedelta(days=int(sys))
    kmrnlusa = str(lusa.strftime('%Y%m%d'))
    kmrnbulan = str(storyday.strftime('%Y%m%d'))
    if kmrnbulan[0:6] == kmrnlusa[0:6]:
        AsNeeded = 'a'
        GETDAT()
    else:
        AsNeeded = 'w'
        GETDAT()


# In[39]:


def TRANSACTION_UMB_TMENU():
    stdin, stdout, stderr = ssh.exec_command('cat /TRBT/script/LIST_PROMO_UMB_TMENU.txt')
    LIST_PROMO = ''.join(stdout.readlines())

    def PROMO():
        cursor = database.cursor()
        query = ("SELECT TANGGAL, PROMO, Total_Transaksi FROM (SELECT calling_date AS TANGGAL,dvc_id AS PROMO,count(dvc_id) AS Total_Transaksi                 FROM call_detailed_record WHERE calling_date = to_char(sysdate-1, 'YYYYMMDD') AND calling_type IN (1, 77)                 AND dvc_id IN ("+LIST_PROMO+") AND media_code LIKE '%U' AND RESULT = 0 GROUP BY calling_date,dvc_id ORDER BY 3 DESC)")
        cursor.execute(query)
        result = cursor.fetchall()
        transumb.set_zoom(50)
        transumb.hide_gridlines(2)
        transumb.write('A1','JUMLAH TRANSAKSI UMB T-MENU PROMO AND HITS',title)
        transumb.write('A3','CALLING_DATE',bold)
        transumb.write('B3','PROMO',bold)
        transumb.write('C3','Total_Transaksi',bold)
        transumb.set_column('A:C', 16)
        for r, row in enumerate(result, start=3):
            for c, col in enumerate(row):
                transumb.write(r,0, row[0], date_format)
                transumb.write(r,1, row[1], content)
                transumb.write(r,2, row[2], content)
        cursor.close()
    
    def HITS():
        cursor = database.cursor()
        query = ("select TANGGAL,HITS,Total_Transaksi From (select calling_date as TANGGAL,dvc_id as HITS,count(dvc_id) as Total_Transaksi from call_detailed_record         where calling_date = to_char(sysdate-1, 'YYYYMMdd') and calling_type in (1,77) and dvc_id not in ("+LIST_PROMO+") and media_code like '%U' and result=0         group by calling_date,dvc_id order by 3 desc)")
        cursor.execute(query)
        result = cursor.fetchall()
        transumb.write('A26','CALLING_DATE',bold)
        transumb.write('B26','HITS',bold)
        transumb.write('C26','Total_Transaksi',bold)
        for r, row in enumerate(result, start=26):
            for c, col in enumerate(row):
                transumb.write(r,0, row[0], date_format)
                transumb.write(r,1, row[1], content)
                transumb.write(r,2, row[2], content)
        cursor.close()
        chat = 'TRANSACTION_UMB_TMENU_PROMO Worksheet has just been written on xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
        fyi(chat)

    PROMO()
    HITS()


# In[40]:


A_KPI()
B_Transaction()
C_Subscriber()
C_FreePaid()
C_SUBS_WITHDRAWAL()
C_CHURNED_USER()
A_NEW_SUBSCRIPTION_TOTAL()
B_NEW_SUBSCRIPTION_CHANNEL()
C_NEW_SUBSCRIPTION_HOUR()
D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL()
A_PROV_TOTAL()
B_PROV_CHANNEL()
C_PROV_CHANNEL_HOUR()
A_TOTAL_RBT_TRAFFIC()
TRAFFIC_PURCHASE()
TRAFFIC_RENEWAL()
A_PRICE_TRAFFIC_REVENUE()
B_PRICE_TRAFFIC_REVENUE()
C_PRICE_TRAFFIC_REVENUE()
A_SUMMARY_CHARGING_GW()
B_CHARGING_CDR_ERROR()
ERROR_SUBS()
HTTPGW()
SMS_CONFIRMATION()
SONGTOP()
GROSSREVENUE()
MELONDIY()
MELONMYNSP()
SUSPEND_TOP20()
REGION()
DTMF()
GP()
SMS_REVENUE()
REVENUE_RENEWAL()
NSP_REKOMENDASI()
SUCCESS_RATE_MT_AND_MO()
TRANSACTION_UMB_TMENU()
workbook.close()
database.close()
SendEmail()
