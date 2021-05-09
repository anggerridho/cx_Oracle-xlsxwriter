#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import cx_Oracle
import os
import datetime
import csv
import paramiko
import pandas as pd
import xlsxwriter
import requests # Untuk Curl
import operator # Sorter
from xlsxwriter.utility import xl_rowcol_to_cell

now = datetime.date.today()
kemarin = now - datetime.timedelta(days=1)
kemarin_str = str(kemarin.strftime('%Y%m%d'))
dir = '/root/Daily/'

os.environ["ORACLE_HOME"] = "/usr/lib/oracle/12.2/client64"
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect('127.0.0.1', 1115, username='trbt', password='.pteluon')
database = cx_Oracle.connect('RBTRPTN/RBTRPTN@127.0.0.1:1521/RBTRPTN')

def fyi(*text):
    import requests
    with open('/root/.Tokegram.txt','r') as Tokegram:
        TOKEN = Tokegram.read().rstrip('\n')
        CHAT_ID="214749655" # Mine
        CHID="-1001475662404" # Group
        API_ENDPOINT = "https://api.telegram.org/bot"+ TOKEN +"/sendMessage"
        data = {'chat_id':CHAT_ID,'text':text}
        requests.post(url = API_ENDPOINT, data = data)


# In[ ]:


# Send Email:
def SendEmail():
    import email, smtplib, ssl
    from email import encoders
    from email.mime.base import MIMEBase
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

#     subject = "An email with attachment from Python"
    Yesterday = str(kemarin.strftime('for %B %d, %Y'))
    subject = "(Automation) RBT Daily Report " + Yesterday
    body = "Dear All,\nPlease kindly find RBT Daily Reports " + Yesterday + " in attachment.\n\nBest Regards,\nAngger Ridho (https://linktr.ee/anggerdho)\nNote: xlsx & this email message is an automation process that I created"
    sender_email = "Angger Ridho <operation.support@eluon.com>"
#     sender_email = "Angger Ridho <angger@eluon.com>"
    receiver_email = ['rizky.fauzi@eluon.com', 'randy@eluon.com', 'dhifa@eluon.com', 'handi@eluon.com', 'diyas@eluon.com',                       'operation.support@eluon.com','amanda@eluon.com','angger@eluon.com']
#     receiver_email = ['angger@eluon.com']
#     password = input("Type your password and press enter:")
    password = "An9gr3k!!"

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
        with open(dir+filename, "rb") as attachment:
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
        server.login("operation.support@eluon.com", password)
        server.sendmail(sender_email, receiver_email, text)
        chat = 'RBT Daily Report ' + str(kemarin_str) + ' was just emailed at '+datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(chat)
        fyi(chat)


# In[ ]:


def RBTDailyReport():
    workbook = xlsxwriter.Workbook(dir+"RBT_Daily_Report_"+str(kemarin_str)+".xlsx",{'strings_to_numbers': True})
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
    
    def SUMMARY():
        # A All KPI
        def A_KPI():
            GetQuery = 'A. All KPI is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_ALL_KPI.txt','r') as inserts:
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

            if result:
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
                kelar = 'A. All KPI has just been written on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. All KPI on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B Transaction
        def B_Transaction():
            GetQuery = 'B. Transaction is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'B_Transaction.txt','r') as inserts:
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

            if result:
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
                kelar = 'B. Transaction has just been written on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing B. Transaction on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # C_Subscriber
        def C_Subscriber():
            GetQuery = 'C. Subscriber is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_Subscriber.txt','r') as inserts:
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

            if result:
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
                kelar = 'C. Subscriber has just been written on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. Subscriber on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # C_Subscriber Free Paid
        def C_FreePaid():
            GetQuery = 'C. Subscriber Free_Paid is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_Subscriber_Free_Paid.txt','r') as inserts:
                query = inserts.read()

            # Exec query summary
            cursor.execute(query)
            result = cursor.fetchall()

            if result:
                for r, row in enumerate(result, start=18):
                    for c, col in enumerate(row):
                        summary.write(r,9, row[0], content)
                        summary.write(r,10, row[1], content)
                cursor.close()
                kelar = 'C. Subscriber Free_Paid has just been written on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. Subscriber Free_Paid on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # C_SUBS_WITHDRAWAL
        def C_SUBS_WITHDRAWAL():
            GetQuery = 'C. Subscriber SUBS WITHDRAWAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_SUBS_WITHDRAWAL.txt','r') as inserts:
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

            if result:
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
                kelar = 'C. Subscriber SUBS WITHDRAWAL has just been written on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. Subscriber SUBS WITHDRAWAL on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # C_CHURNED_USER
        def C_CHURNED_USER():
            GetQuery = 'C. Subscriber CHURNED USER is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_CHURNED_USER.txt','r') as inserts:
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

            if result:
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
                kelar = 'C. Subscriber CHURNED USER has just been written on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. Subscriber CHURNED USER on worksheet SUMMARY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_KPI()
        B_Transaction()
        C_Subscriber()
        C_FreePaid()
        C_SUBS_WITHDRAWAL()
        C_CHURNED_USER()

    def NEW_SUBS():
        # A. NEW SUBSCRIPTION TOTAL
        def A_NEW_SUBSCRIPTION_TOTAL():
            GetQuery = 'A. NEW SUBSCRIPTION TOTAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_NEW_SUBSCRIPTION_TOTAL.txt','r') as inserts:
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

            if result:
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
                kelar = 'A. NEW SUBSCRIPTION TOTAL has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. NEW SUBSCRIPTION TOTAL on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B. NEW SUBSCRIPTION CHANNEL
        def B_NEW_SUBSCRIPTION_CHANNEL():
            # A_SUCCESS_RATE Purchasing
            def A_SUCCESS_RATE():
                GetQuery = 'B. NEW SUBSCRIPTION CHANNEL A. SUCCESS_RATE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(GetQuery)
                fyi(GetQuery)
                cursor = database.cursor()
                with open(dir+'B_NEW_SUBSCRIPTION_CHANNEL_A_SUCCESS_RATE.txt','r') as inserts:
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

                if result:
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
                    kelar = 'B. NEW SUBSCRIPTION CHANNEL A. SUCCESS_RATE has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(kelar)
                    fyi(kelar)
                else:
                    chat = 'Skip writing B. NEW SUBSCRIPTION CHANNEL A. SUCCESS_RATE on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(chat)
                    fyi(chat)

            # B_SUCCESS_RATE Purchasing
            def B_SUCCESS_RATE():
                GetQuery = 'B. NEW SUBSCRIPTION CHANNEL B. SUCCESS_RATE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(GetQuery)
                fyi(GetQuery)
                cursor = database.cursor()
                with open(dir+'B_NEW_SUBSCRIPTION_CHANNEL_B_SUCCESS_RATE.txt','r') as inserts:
                    query = inserts.read()

                # Exec query new_subs
                cursor.execute(query)
                result = cursor.fetchall()
                new_subs.write('B18','',neat)
                new_subs.write('C18','SUCCESS RATE',neat)

                for line in ['D','E','F','G','H','I','J','K','L','M','N']:
                    new_subs.write(line + '18','',neat)
                    new_subs.write_formula(line + '18', '=IFERROR(' + line +'17'+'/'+ line +'16' + ',"-")',percentage)

                if result:
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
                    kelar = 'B. NEW SUBSCRIPTION CHANNEL B. SUCCESS_RATE has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(kelar)
                    fyi(kelar)
                else:
                    chat = 'Skip writing B. NEW SUBSCRIPTION CHANNEL B. SUCCESS_RATE on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(chat)
                    fyi(chat)

            # C_SUCCESS_RATE Renewal
            def C_SUCCESS_RATE():
                GetQuery = 'B. NEW SUBSCRIPTION CHANNEL C. SUCCESS_RATE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(GetQuery)
                fyi(GetQuery)
                cursor = database.cursor()
                with open(dir+'B_NEW_SUBSCRIPTION_CHANNEL_C_SUCCESS_RATE.txt','r') as inserts:
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

                if result:
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
                    kelar = 'B. NEW SUBSCRIPTION CHANNEL C. SUCCESS_RATE has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(kelar)
                    fyi(kelar)
                else:
                    chat = 'Skip writing B. NEW SUBSCRIPTION CHANNEL C. SUCCESS_RATE on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(chat)
                    fyi(chat)

            # D_SUCCESS_RATE Renewal
            def D_SUCCESS_RATE():
                GetQuery = 'B. NEW SUBSCRIPTION CHANNEL D. SUCCESS_RATE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(GetQuery)
                fyi(GetQuery)
                cursor = database.cursor()
                with open(dir+'B_NEW_SUBSCRIPTION_CHANNEL_D_SUCCESS_RATE.txt','r') as inserts:
                    query = inserts.read()

                # Exec query new_subs
                cursor.execute(query)
                result = cursor.fetchall()
                new_subs.write('B29','',neat)
                new_subs.write('C29','SUCCESS RATE',neat)

                for line in ['D','E','F','G','H','I','J','K','L','M','N']:
                    new_subs.write(line + '29','',neat)
                    new_subs.write_formula(line + '29', '=IFERROR(' + line +'28'+'/'+ line +'27' + ',"-")',percentage)

                if result:
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
                    kelar = 'B. NEW SUBSCRIPTION CHANNEL D. SUCCESS_RATE has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(kelar)
                    fyi(kelar)
                else:
                    chat = 'Skip writing B. NEW SUBSCRIPTION CHANNEL D. SUCCESS_RATE on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(chat)
                    fyi(chat)

            A_SUCCESS_RATE()
            B_SUCCESS_RATE()
            C_SUCCESS_RATE()
            D_SUCCESS_RATE()

        # C. NEW SUBSCRIPTION HOUR
        def C_NEW_SUBSCRIPTION_HOUR():
            GetQuery = 'C. NEW SUBSCRIPTION HOUR is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_NEW_SUBSCRIPTION_HOUR.txt','r') as inserts:
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

            if result:
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
                kelar = 'C. NEW SUBSCRIPTION HOUR has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. NEW SUBSCRIPTION HOUR on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # D. NEW SUBSCRIPTION HOUR PER CHANNEL
        def D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL():
            GetQuery = 'D. NEW SUBSCRIPTION HOUR PER CHANNEL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL.txt','r') as inserts:
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

            if result:
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
                kelar = 'D. NEW SUBSCRIPTION HOUR PER CHANNEL has just been written on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing D. NEW SUBSCRIPTION HOUR PER CHANNEL on worksheet NEW_SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_NEW_SUBSCRIPTION_TOTAL()
        B_NEW_SUBSCRIPTION_CHANNEL()
        C_NEW_SUBSCRIPTION_HOUR()
        D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL()

    def PROV():
        # A. PROV TOTAL :
        def A_PROV_TOTAL():
            GetQuery = 'A. PROV TOTAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_PROV_TOTAL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        prov.write(r,1, row[0], date_format)
                        prov.write(r,2, row[1], content)
                        prov.write(r,3, row[2], content)
                        prov.write(r,4, row[3], content)
                cursor.close()
                kelar = 'A. PROV TOTAL has just been written on worksheet PROV xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. PROV TOTAL on worksheet PROV xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B. PROV CHANNEL :
        def B_PROV_CHANNEL():
            GetQuery = 'B. PROV CHANNEL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'B_PROV_CHANNEL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=9):
                    for c, col in enumerate(row):
                        prov.write(r,1, row[0], date_format)
                        prov.write(r,2, row[1], content)
                        prov.write(r,3, row[2], content)
                        prov.write(r,4, row[3], content)
                        prov.write(r,5, row[4], content)
                        prov.write(r,6, row[5], content)
                cursor.close()
                kelar = 'B. PROV CHANNEL has just been written on worksheet PROV xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing B. PROV CHANNEL on worksheet PROV xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # C. PROV CHANNEL HOUR :
        def C_PROV_CHANNEL_HOUR():
            GetQuery = 'C. PROV CHANNEL HOUR is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_PROV_CHANNEL_HOUR.txt','r') as inserts:
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

            if result:
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
                kelar = 'C. PROV CHANNEL HOUR has just been written on worksheet PROV xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. PROV CHANNEL HOUR on worksheet PROV xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_PROV_TOTAL()
        B_PROV_CHANNEL()
        C_PROV_CHANNEL_HOUR()

    def TOTAL_TRAFFIC():
        # A. TOTAL RBT TRAFFIC :
        def A_TOTAL_RBT_TRAFFIC():
            GetQuery = 'A. TOTAL RBT TRAFFIC is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_TOTAL_RBT_TRAFFIC.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        totaltraf.write(r,1, row[0], date_format)
                        totaltraf.write(r,2, row[1], content)
                        totaltraf.write(r,3, row[2], content)
                        totaltraf.write(r,4, row[3], content)
                        totaltraf.write(r,5, row[4], content)
                cursor.close()
                kelar = 'A. TOTAL RBT TRAFFIC has just been written on worksheet TOTAL TRAFFIC xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. TOTAL RBT TRAFFIC on worksheet TOTAL TRAFFIC xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # TRAFFIC PURCHASE :
        def TRAFFIC_PURCHASE():
            GetQuery = 'TRAFFIC PURCHASE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'TRAFFIC_PURCHASE.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=28):
                    for c, col in enumerate(row):
                        totaltraf.write(r,1, row[0], date_format)
                        totaltraf.write(r,2, row[1], content)
                        totaltraf.write(r,3, row[2], content)
                        totaltraf.write(r,4, row[3], content)
                        totaltraf.write(r,5, row[4], content)
                        totaltraf.write(r,6, row[5], content)
                cursor.close()
                kelar = 'TRAFFIC PURCHASE has just been written on worksheet TOTAL TRAFFIC xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing TRAFFIC PURCHASE on worksheet TOTAL TRAFFIC xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # TRAFFIC RENEWAL :
        def TRAFFIC_RENEWAL():
            GetQuery = 'TRAFFIC RENEWAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'TRAFFIC_RENEWAL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=28):
                    for c, col in enumerate(row):
                        totaltraf.write(r,12, row[0], date_format)
                        totaltraf.write(r,13, row[1], content)
                        totaltraf.write(r,14, row[2], content)
                        totaltraf.write(r,15, row[3], content)
                        totaltraf.write(r,16, row[4], content)
                        totaltraf.write(r,17, row[5], content)
                cursor.close()
                kelar = 'TRAFFIC RENEWAL has just been written on worksheet TOTAL TRAFFIC xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing TRAFFIC RENEWAL on worksheet TOTAL TRAFFIC xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_TOTAL_RBT_TRAFFIC()
        TRAFFIC_PURCHASE()
        TRAFFIC_RENEWAL()

    def PRICEREVENUE():
        # A. PRICE TRAFFIC & REVENUE :
        def A_PRICE_TRAFFIC_REVENUE():
            GetQuery = 'A. PRICE TRAFFIC & REVENUE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_PRICE_TRAFFIC_REVENUE.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        pricerevenue.write(r,1, row[0], date_format)
                        pricerevenue.write(r,2, row[1], content)
                        pricerevenue.write(r,3, row[2], content)
                        pricerevenue.write(r,4, row[3], content)
                        pricerevenue.write(r,5, row[4], content)
                        pricerevenue.write(r,6, row[5], content)
                cursor.close()
                kelar = 'A. PRICE TRAFFIC & REVENUE has just been written on worksheet PRICE & REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. PRICE TRAFFIC & REVENUE on worksheet PRICE & REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B. PRICE TRAFFIC & REVENUE :
        def B_PRICE_TRAFFIC_REVENUE():
            GetQuery = 'B. PRICE TRAFFIC REVENUE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'B_PRICE_TRAFFIC_REVENUE.txt','r') as inserts:
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

            if result:
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
                kelar = 'B. PRICE TRAFFIC REVENUE has just been written on worksheet PRICE & REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing B. PRICE TRAFFIC REVENUE on worksheet PRICE & REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # C. PRICE TRAFFIC & REVENUE :
        def C_PRICE_TRAFFIC_REVENUE():
            GetQuery = 'C. PRICE TRAFFIC REVENUE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'C_PRICE_TRAFFIC_REVENUE.txt','r') as inserts:
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

            if result:
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
                kelar = 'C. PRICE TRAFFIC REVENUE has just been written on worksheet PRICE & REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing C. PRICE TRAFFIC REVENUE on worksheet PRICE & REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_PRICE_TRAFFIC_REVENUE()
        B_PRICE_TRAFFIC_REVENUE()
        C_PRICE_TRAFFIC_REVENUE()

    def CHARGINGCDR():
        # A. SUMMARY CHARGING GW :
        def A_SUMMARY_CHARGING_GW():
            GetQuery = 'A. SUMMARY CHARGING GW is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            query = ("SELECT DECODE (chg_type,'C','Charging','Q','Query Balance','R','Refund'), SUM (amount), SUM(CASE WHEN chg_type = 'C' AND ( RESULT LIKE 'Succ%' OR RESULT LIKE 'Insuff%' OR RESULT LIKE '%subs block%' OR RESULT LIKE '%locked%') THEN amount WHEN chg_type = 'R' AND RESULT LIKE 'Succ%' THEN amount WHEN chg_type = 'Q' AND RESULT LIKE 'Succ%' THEN amount ELSE 0 END),             ROUND(SUM(CASE WHEN chg_type = 'C' AND ( RESULT LIKE 'Succ%' OR RESULT LIKE 'Insuff%' OR RESULT LIKE '%subs block%' OR RESULT LIKE '%locked%') THEN amount WHEN chg_type = 'R' AND RESULT LIKE 'Succ%' THEN amount WHEN chg_type = 'Q' AND RESULT LIKE 'Succ%' THEN amount ELSE 0 END) / SUM (amount) * 100, 2) || '%' FROM (SELECT COUNT (*) AS amount, chg_type, RESULT             FROM perfcdr_chg WHERE chg_type != ' ' AND calling_datetime LIKE '"+str(kemarin_str)+"'||'%' GROUP BY chg_type,RESULT) GROUP BY chg_type ORDER BY 1")

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

            if result:
                for r, row in enumerate(result, start=4):
                    for c, col in enumerate(row):
                        chargingcdr.write(r,1, row[0], content)
                        chargingcdr.write(r,2, row[1], content)
                        chargingcdr.write(r,3, row[2], content)
                        chargingcdr.write(r,4, row[3], content)
                cursor.close()
                kelar = 'A. SUMMARY CHARGING GW has just been written on worksheet CHARGING CDR xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. SUMMARY CHARGING GW on worksheet CHARGING CDR xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B. CHARGING CDR ERROR :
        def B_CHARGING_CDR_ERROR():
            GetQuery = 'B. CHARGING CDR ERROR is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
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

            if result:
                for r, row in enumerate(result, start=11):
                    for c, col in enumerate(row):
                        chargingcdr.write(r,1, row[0], content)
                        chargingcdr.write(r,2, row[1], content)
                        chargingcdr.write(r,3, row[2], content)

                cursor.close()
                kelar = 'B. CHARGING CDR ERROR has just been written on worksheet CHARGING CDR xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing B. CHARGING CDR ERROR on worksheet CHARGING CDR xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_SUMMARY_CHARGING_GW()
        B_CHARGING_CDR_ERROR()

    def ERROR_SUBS():
        # A. NEW SUBS ERROR CHANNEL
        def A_NEW_SUBS_ERROR_CHANNEL():
            GetQuery = 'A. NEW SUBS ERROR CHANNEL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_NEW_SUBS_ERROR_CHANNEL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        errorsubs.write(r,1, row[0], date_format)
                        errorsubs.write(r,2, row[1], content)
                        errorsubs.write(r,3, row[2], content)
                        errorsubs.write(r,4, row[3], content)
                cursor.close()
                kelar = 'A. NEW SUBS ERROR CHANNEL has just been written on worksheet ERROR SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. NEW SUBS ERROR CHANNEL on worksheet ERROR SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B. NEW SUBS ERROR HOUR CHANNEL
        def B_NEW_SUBS_ERROR_HOUR_CHANNEL():
            GetQuery = 'B. NEW SUBS ERROR HOUR CHANNEL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'B_NEW_SUBS_ERROR_HOUR_CHANNEL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        errorsubs.write(r,7, row[0], date_format)
                        errorsubs.write(r,8, row[1], content)
                        errorsubs.write(r,9, row[2], content)
                        errorsubs.write(r,10, row[3], content)
                        errorsubs.write(r,11, row[4], content)
                cursor.close()
                kelar = 'B. NEW SUBS ERROR HOUR CHANNEL has just been written on worksheet ERROR SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing B. NEW SUBS ERROR HOUR CHANNEL on worksheet ERROR SUBS xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_NEW_SUBS_ERROR_CHANNEL()
        B_NEW_SUBS_ERROR_HOUR_CHANNEL()

    def HTTPGW():
        # A. HTTPGW BREAKDOWN PURCHASE
        def A_HTTPGW_BREAKDOWN_PURCHASE():
            GetQuery = 'A. HTTPGW BREAKDOWN PURCHASE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'A_HTTPGW_BREAKDOWN_PURCHASE.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        httpgw.write(r,1, row[0], date_format)
                        httpgw.write(r,2, row[1], content)
                        httpgw.write(r,3, row[2], content)
                        httpgw.write(r,4, row[3], content)
                        httpgw.write(r,5, row[4], content)
                cursor.close()
                kelar = 'A. HTTPGW BREAKDOWN PURCHASE has just been written on worksheet HTTPGW xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing A. HTTPGW BREAKDOWN PURCHASE on worksheet HTTPGW xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # B. HTTPGW BREAKDOWN RENEWAL
        def B_HTTPGW_BREAKDOWN_RENEWAL():
            GetQuery = 'B. HTTPGW BREAKDOWN RENEWAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'B_HTTPGW_BREAKDOWN_RENEWAL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        httpgw.write(r,7, row[0], date_format)
                        httpgw.write(r,8, row[1], content)
                        httpgw.write(r,9, row[2], content)
                        httpgw.write(r,10, row[3], content)
                cursor.close()
                kelar = 'B. HTTPGW BREAKDOWN RENEWAL has just been written on worksheet HTTPGW xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing B. HTTPGW BREAKDOWN RENEWAL on worksheet HTTPGW xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        A_HTTPGW_BREAKDOWN_PURCHASE()
        B_HTTPGW_BREAKDOWN_RENEWAL()

    def SMS_CONFIRMATION():
        # Purchase
        def Purchase():
            GetQuery = 'PURCHASE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'SMS_CONFIRMATION_PURCHASE.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        smsconfirm.write(r,1, row[0], date_format)
                        smsconfirm.write(r,2, row[1], content)
                        smsconfirm.write(r,3, row[2], content)
                        smsconfirm.write(r,4, row[3], content)
                        smsconfirm.write(r,5, row[4], content)
                        smsconfirm.write(r,6, row[5], content)
                cursor.close()
                kelar = 'PURCHASE has just been written on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing PURCHASE on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # Error_Detail_Purchase
        def Error_Detail_Purchase():
            GetQuery = 'ERROR DETAIL PURCHASE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'SMS_CONFIRMATION_ERRORDETAIL_PURCHASE.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=8):
                    for c, col in enumerate(row):
                        smsconfirm.write(r,1, row[0], date_format)
                        smsconfirm.write(r,2, row[1], content)
                        smsconfirm.write(r,3, row[2], content)
                        smsconfirm.write(r,4, row[3], content)
                cursor.close()
                kelar = 'ERROR DETAIL PURCHASE has just been written on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing ERROR DETAIL PURCHASE on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # Renewal
        def Renewal():
            GetQuery = 'Renewal is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'SMS_CONFIRMATION_RENEWAL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=24):
                    for c, col in enumerate(row):
                        smsconfirm.write(r,1, row[0], date_format)
                        smsconfirm.write(r,2, row[1], content)
                        smsconfirm.write(r,3, row[2], content)
                        smsconfirm.write(r,4, row[3], content)
                        smsconfirm.write(r,5, row[4], content)
                        smsconfirm.write(r,6, row[5], content)
                cursor.close()
                kelar = 'Renewal has just been written on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing Renewal on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        # Error_Detail_Purchase
        def Error_Detail_Renewal():
            GetQuery = 'ERROR DETAIL RENEWAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'SMS_CONFIRMATION_ERRORDETAIL_RENEWAL.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=29):
                    for c, col in enumerate(row):
                        smsconfirm.write(r,1, row[0], date_format)
                        smsconfirm.write(r,2, row[1], content)
                        smsconfirm.write(r,3, row[2], content)
                        smsconfirm.write(r,4, row[3], content)
                cursor.close()
                kelar = 'ERROR DETAIL RENEWAL has just been written on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(kelar)
                fyi(kelar)
            else:
                chat = 'Skip writing ERROR DETAIL RENEWAL on worksheet SMS_CONFIRMATION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(chat)
                fyi(chat)

        Purchase()
        Error_Detail_Purchase()
        Renewal()
        Error_Detail_Renewal()

    def SONGTOP():
        GetQuery = '75 Song Top is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'SongTop_'+str(kmrnbulan[0:6])+'.csv'
            cursor = database.cursor()
            query = ("SELECT CALLING_DATE,RANK,CONTENT_SINGER,CONTENT_TITLE,TRAFFIC,REVENUE,CP_NAME FROM (SELECT a.calling_date, a.rating_content_id,                      SUM (a.total_success) AS traffic,SUM (a.total_revenue) AS revenue,b.promo_code AS content_title,b.content_singer AS content_singer,                      c.mcp_user_disp_name AS cp_name,RANK() OVER (PARTITION BY calling_date ORDER BY SUM (a.total_revenue) DESC) RANK FROM                      CDR_ACTIVATION_SUMMARY_DAY a, tot_content b, mcp_user_info c WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD') AND                      a.rating_content_id = b.content_id AND c.mcp_user_id = b.mcp_user_id GROUP BY a.rating_content_id,promo_code,content_singer,                      mcp_user_disp_name,a.calling_date ORDER BY 1, revenue DESC) WHERE RANK BETWEEN 1 AND 75")

            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append 75 Song Top '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                cursor.close()

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

            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                if csvfile:
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
                    chat = '75 Song Top has just been written on worksheet SONG TOP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing 75 Song Top on worksheet SONG TOP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def GROSSREVENUE():
        GetQuery = 'Revenue is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'Gross_Revenue_'+str(kmrnbulan[0:6])+'.csv'
            cursor = database.cursor()
            query = ("SELECT CALLING_DATE,TOTAL_COUNT,TOTAL_SUCCESS,FAIL_COUNT,SUCCESS_RATE,BILLABLE_COUNT,GROSS_REVENUE,MELON_REVENUE FROM(SELECT calling_date,             SUM (TOTAL_COUNT) AS TOTAL_COUNT,SUM (SUCCESS_COUNT) AS TOTAL_SUCCESS,SUM (FAIL_COUNT) AS FAIL_COUNT,ROUND( (SUM (SUCCESS_COUNT)/ (SUM (SUCCESS_COUNT)             + SUM (FAIL_COUNT))) * 100,2) ||'%' AS success_rate,SUM (BILLABLE_COUNT) AS BILLABLE_COUNT,SUM (CHARGE_AMOUNT) AS GROSS_REVENUE,SUM (CHARGE_AMOUNT)             * 75 /100 AS MELON_REVENUE FROM CDR_SUMMARY WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD') AND calling_type IN (1,2,7,64,93,77,78,65,95,96,99,100)             AND calling_type BETWEEN 0 AND 999 GROUP BY calling_date ORDER BY calling_date)")
            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append Gross Revenue '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                cursor.close()

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
                if csvfile:
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
                    chat = 'Revenue Statistics has just been written on worksheet REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing Revenue Statistics on worksheet REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def MELONDIY():
        GetQuery = 'MELON DIY is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'Melon_Diy_'+str(kmrnbulan[0:6])+'.csv'
            cursor = database.cursor()
            query = ("SELECT CALLING_DATE,TOTAL_COUNT,TOTAL_SUCCESS,FAIL_COUNT,SUCCESS_RATE,BILLABLE_COUNT,GROSS_REVENUE FROM (SELECT calling_date, SUM(TOTAL_COUNT) AS TOTAL_COUNT,SUM(SUCCESS_COUNT) AS TOTAL_SUCCESS,             SUM (FAIL_COUNT) AS FAIL_COUNT,ROUND( (SUM (SUCCESS_COUNT)/ (SUM (SUCCESS_COUNT) + SUM (FAIL_COUNT))) * 100,2) ||'%' AS success_rate, SUM(BILLABLE_COUNT) AS BILLABLE_COUNT,SUM (CHARGE_AMOUNT) AS GROSS_REVENUE             FROM CDR_SUMMARY WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD') AND calling_type IN (1,2,7,64,93,77,78,65,95,96,99,100) AND RATING_CONTENT_ID in ( select content_id from (select content_id from content where mcp_user_id ='23610'             union all select bundle_id  as content_id from bundle where mcp_user_id ='23610')) AND calling_type BETWEEN 0 AND 999 GROUP BY calling_date ORDER BY calling_date)")
            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append MELON DIY '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                cursor.close()

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
                if csvfile:
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
                    chat = 'MELON DIY has just been written on worksheet MELON DIY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing MELON DIY on worksheet MELON DIY xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def MELONMYNSP():
        def Statistic():
            GetQuery = 'Melon MyNSP is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'Melon_MyNsp_'+str(kmrnbulan[0:6])+'.csv'
                cursor = database.cursor()
                query = ("SELECT CALLING_DATE,TOTAL_COUNT,TOTAL_SUCCESS,FAIL_COUNT,SUCCESS_RATE,BILLABLE_COUNT,GROSS_REVENUE FROM (SELECT calling_date,SUM (TOTAL_COUNT) AS TOTAL_COUNT,SUM (SUCCESS_COUNT) AS TOTAL_SUCCESS,SUM (FAIL_COUNT) AS FAIL_COUNT,                 ROUND( (SUM (SUCCESS_COUNT)/ (SUM (SUCCESS_COUNT) + SUM (FAIL_COUNT))) * 100,2) ||'%' AS success_rate,SUM (BILLABLE_COUNT) AS BILLABLE_COUNT,SUM (CHARGE_AMOUNT) AS GROSS_REVENUE FROM CDR_SUMMARY WHERE CALLING_DATE = TO_CHAR(SYSDATE-"+sys+",'YYYYMMDD')                 AND calling_type IN (1,2,7,77,64,93,78,65,95,96,99,100) AND RATING_CONTENT_ID in ( select content_id from (select content_id from content where mcp_user_id ='23661' union all select bundle_id  as content_id from bundle where mcp_user_id ='23661' ))                 AND calling_type BETWEEN 0 AND 999 GROUP BY calling_date ORDER BY calling_date)")
                cursor.execute(query)
                df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append Melon MyNSP '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                    cursor.close()

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
                    if csvfile:
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
                        chat = 'Melon MyNSP has just been written on worksheet MELON MYNSP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing Melon MyNSP Statistic on worksheet MELON MYNSP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

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
            GetQuery = 'ERROR CHANNEL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'MELON_MYNSP_ERRORCHANNEL.txt','r') as inserts:
                query = inserts.read()

            # Exec query new_subs
            cursor.execute(query)
            result = cursor.fetchall()
            melonmynsp.write('B38','ERROR CHANNEL',title)
            melonmynsp.write('B39','CALLING_DATE',bold)
            melonmynsp.write('C39','CHANNEL',bold)
            melonmynsp.write('D39','ERROR_RESULT',bold)
            melonmynsp.write('E39','ERROR_COUNT',bold)

            if result:
                for r, row in enumerate(result, start=39):
                    for c, col in enumerate(row):
                        melonmynsp.write(r,1, row[0], date_format)
                        melonmynsp.write(r,2, row[1], content)
                        melonmynsp.write(r,3, row[2], content)
                        melonmynsp.write(r,4, row[3], content)
                cursor.close()
                chat = 'ERROR CHANNEL has just been written on worksheet MELON MYNSP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing ERROR CHANNEL on worksheet MELON MYNSP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def ERROR_PRICE():
            GetQuery = 'ERROR PRICE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open(dir+'MELON_MYNSP_ERRORPRICE.txt','r') as inserts:
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

            if result:
                for r, row in enumerate(result, start=39):
                    for c, col in enumerate(row):
                        melonmynsp.write(r,6, row[0], date_format)
                        melonmynsp.write(r,7, row[1], content)
                        melonmynsp.write(r,8, row[2], content)
                        melonmynsp.write(r,9, row[3], content)
                cursor.close()
                chat = 'ERROR PRICE has just been written on worksheet MELON MYNSP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing ERROR PRICE on worksheet MELON MYNSP xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        Statistic()
        ERROR_CHANNEL()
        ERROR_PRICE()

    def SUSPEND_TOP20():
        GetQuery = 'SUSPEND_TOP20 is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        cursor = database.cursor()
        with open(dir+'SUSPEND_TOP20.txt','r') as inserts:
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

        if result:
            for r, row in enumerate(result, start=3):
                for c, col in enumerate(row):
                    suspendtop20.write(r,1, row[0], date_format)
                    suspendtop20.write(r,2, row[1], content)
                    suspendtop20.write(r,3, row[2], content)
                    suspendtop20.write(r,4, row[3], content)
            cursor.close()
            chat = 'SUSPEND_TOP20 has just been written on worksheet SUSPEND_TOP20 xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(chat)
            fyi(chat)
        else:
            Alert = 'Skip writing SUSPEND_TOP20 on worksheet SUSPEND_TOP20 xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
            print(Alert)
            fyi(Alert)

    def REGION():
        def TRAFFIC_AND_REVENUE_PER_REGION():
            GetQuery = '1. TRAFFIC AND REVENUE PER REGION is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'TRAFFIC_AND_REVENUE_PER_REGION_'+str(kmrnbulan[0:6])+'.csv'
                cursor = database.cursor()
                query = ("select calling_date ,T_Sumbagut,SUMBAGUT,T_SUMBAGTENG,Sumbagteng ,T_Sumbagsel ,Sumbagsel ,T_Metro_Jakarta, Metro_Jakarta ,                         T_Outer_Jakarta , Outer_Jakarta ,T_Jabar , Jabar , T_Jateng_DIY , Jateng_DIY ,T_Jatim , Jatim , T_Balinusra ,Balinusra ,                         T_Kalimantan , Kalimantan , T_Sulawesi , Sulawesi , T_Papua_Maluku , Papua_Maluku , T_unknown , unknown , Traffic_total ,                         revenue_total                         FROM                         (                         select calling_date,                             sum(case when region='Sumbagut' then TOTAL_SUCCESS else 0 end)  T_Sumbagut,                                 sum(case when region='Sumbagut' then TOTAL_REVENUE else 0 end)  Sumbagut,                                 sum(case when region='Sumbagteng' then TOTAL_SUCCESS else 0 end) T_Sumbagteng,                             sum(case when region='Sumbagteng' then TOTAL_REVENUE else 0 end) Sumbagteng,                                 sum(case when region='Sumbagsel' then TOTAL_SUCCESS else 0 end)  T_Sumbagsel,                                 sum(case when region='Sumbagsel' then TOTAL_REVENUE else 0 end)  Sumbagsel,                             sum(case when region='Jabotabek'  and sub_region = 'Jakarta' then TOTAL_SUCCESS else 0 end) T_Metro_Jakarta,                             sum(case when region='Jabotabek'  and sub_region = 'Jakarta' then TOTAL_REVENUE else 0 end) Metro_Jakarta,                                 sum(case when region='Jabotabek'  and  sub_region <> 'Jakarta' then TOTAL_SUCCESS else 0 end) T_Outer_Jakarta,                             sum(case when region='Jabotabek'  and  sub_region <> 'Jakarta' then TOTAL_REVENUE else 0 end) Outer_Jakarta,                                 sum(case when region='Jabar' then TOTAL_SUCCESS else 0 end) T_Jabar,                             sum(case when region='Jabar' then TOTAL_REVENUE else 0 end) Jabar,                                 sum(case when region='Jateng' then TOTAL_SUCCESS else 0 end) T_Jateng_DIY,                             sum(case when region='Jateng' then TOTAL_REVENUE else 0 end) Jateng_DIY,                                 sum(case when region='Jatim' then TOTAL_SUCCESS else 0 end) T_Jatim,                             sum(case when region='Jatim' then TOTAL_REVENUE else 0 end) Jatim,                                 sum(case when region='Balinusra' then TOTAL_SUCCESS else 0 end) T_Balinusra,                             sum(case when region='Balinusra' then TOTAL_REVENUE else 0 end) Balinusra,                                 sum(case when region='Kalimantan' then TOTAL_SUCCESS else 0 end) T_Kalimantan,                             sum(case when region='Kalimantan' then TOTAL_REVENUE else 0 end) Kalimantan,                                 sum(case when region='Sulawesi' then TOTAL_SUCCESS else 0 end)   T_Sulawesi,                             sum(case when region='Sulawesi' then TOTAL_REVENUE else 0 end)   Sulawesi,                                  sum(case when region='Papua' then TOTAL_SUCCESS else 0 end)  T_Papua_Maluku,                             sum(case when region='Papua' then TOTAL_REVENUE else 0 end)  Papua_Maluku,                             sum(case when region='UNKNOWN' then TOTAL_SUCCESS else 0 end)  T_unknown,                             sum(case when region='UNKNOWN' then TOTAL_REVENUE else 0 end)  unknown,                                 sum(TOTAL_SUCCESS)  Traffic_total,                             sum(TOTAL_REVENUE )  revenue_total                                   from CDR_ACTIVATION_SUMMARY_DAY                         where  calling_date = to_char(sysdate-"+sys+",'yyyymmdd')                         group by calling_date                         order by calling_date)")
                cursor.execute(query)
                df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append 1. TRAFFIC AND REVENUE PER REGION '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                    cursor.close()

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
                    if csvfile:
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
                        chat = '1. TRAFFIC AND REVENUE PER REGION has just been written on worksheet REGION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing 1. TRAFFIC AND REVENUE PER REGION on worksheet REGION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

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
            GetQuery = '2. UNIQ SUBS TRANSACTION is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'UNIQ_SUBS_TRANSACTION_'+str(kmrnbulan[0:6])+'.csv'
                cursor = database.cursor()
                query = ("SELECT CALLING_DATE ,SUMBAGUT , SUMBAGTENG , SUMBAGSEL ,       METRO_JAKARTA ,OUTER_JAKARTA ,      JABAR,    JATENG_DIY,                 JATIM ,  BALINUSRA ,       KALIMANTAN ,      SULAWESI ,PAPUA_MALUKU ,UNKNOWN     ,TOTAL_UNIQ_SUBS                 FROM (                 select calling_date,                 count(distinct case when region='Sumbagut' then CALLING_MDN end)  Sumbagut,                 count(distinct case when region='Sumbagteng' then CALLING_MDN end) Sumbagteng,                 count(distinct case when region='Sumbagsel' then CALLING_MDN end)  Sumbagsel,                 count(distinct case when region='Jabotabek'  and sub_region = 'Jakarta' then CALLING_MDN end) Metro_Jakarta,                 count(distinct case when region='Jabotabek'  and  sub_region <> 'Jakarta' then CALLING_MDN end) Outer_Jakarta,                 count(distinct case when region='Jabar' then CALLING_MDN end) Jabar,                 count(distinct case when region='Jateng' then CALLING_MDN end) Jateng_DIY,                 count(distinct case when region='Jatim' then CALLING_MDN end) Jatim,                 count(distinct case when region='Balinusra' then CALLING_MDN end) Balinusra,                 count(distinct case when region='Kalimantan' then CALLING_MDN end) Kalimantan,                 count(distinct case when region='Sulawesi' then CALLING_MDN end)   Sulawesi,                 count(distinct case when region='Papua' then CALLING_MDN end)  Papua_Maluku,                 count(distinct case when region='UNKNOWN' then CALLING_MDN end)  unknown,                 count(distinct calling_mdn )  total_uniq_subs                 from CALL_DETAILED_RECORD a, prefix_region b                 where  calling_date = to_char(sysdate-"+sys+",'yyyymmdd')                 and calling_type in (1,2,7,77,78,64,93,65,95,96,99,100)                 and result in (805,0)                 and substr(a.calling_mdn,1,7) = b.prefix                 group by calling_date                 order by calling_date)")
                cursor.execute(query)
                df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append 2. UNIQ SUBS TRANSACTION '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                    cursor.close()

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
                    if csvfile:
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
                        chat = '2. UNIQ SUBS TRANSACTION has just been written on worksheet REGION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing 2. UNIQ SUBS TRANSACTION on worksheet REGION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

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
            GetQuery = '3. UNIQ SUBS TRANSACTION PROCESS TYPE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'UNIQ_SUBS_TRANSACTION_PROCESS_TYPE_'+str(kmrnbulan[0:6])+'.csv'
                cursor = database.cursor()
                query = ("select CALLING_DATE,PROCESS_TYPE,Sumbagut,Sumbagteng,Sumbagsel,Metro_Jakarta,Outer_Jakarta,Jabar,Jateng_DIY,Jatim,Balinusra,Kalimantan,Sulawesi,                 Papua_Maluku,OTHERS FROM (select a.calling_date as CALLING_DATE,                 CASE when a.calling_type in (1,2,7,77) then 'Purchase'                 when a.calling_type in (64,78,65,93,95,96,99,100) then 'Renewal' end as PROCESS_TYPE,                 count(distinct case when b.region = 'Sumbagut' then a.calling_mdn end) as Sumbagut,                 count(distinct case when b.region = 'Sumbagteng' then  a.calling_mdn end) as Sumbagteng,                 count(distinct case when b.region = 'Sumbagsel' then  a.calling_mdn end) as Sumbagsel,                 count(distinct case when b.region = 'Jabotabek'  and sub_region = 'Jakarta' then CALLING_MDN end) Metro_Jakarta,                 count(distinct case when b.region = 'Jabotabek'  and  sub_region <> 'Jakarta' then CALLING_MDN end) Outer_Jakarta,                 count(distinct case when b.region = 'Jabar' then  a.calling_mdn end) as Jabar,                 count(distinct case when b.region = 'Jateng' then  a.calling_mdn end) as Jateng_DIY,                 count(distinct case when b.region = 'Jatim' then  a.calling_mdn end) as Jatim,                 count(distinct case when b.region = 'Balinusra' then a.calling_mdn end) as Balinusra,                 count(distinct case when b.region = 'Kalimantan' then a.calling_mdn end) as Kalimantan,                 count(distinct case when b.region = 'Sulawesi' then a.calling_mdn end) as Sulawesi,                 count(distinct case when b.region = 'Papua' then a.calling_mdn end) as Papua_Maluku,                 count(distinct case when b.region not in ('Sumbagut','Sumbagsel','Sumbagteng','Jabotabek','Jabar','Jateng','Jatim','Balinusra','Sulawesi','Kalimantan','Papua') then                 a.calling_mdn end) as Others                 from call_detailed_record a, prefix_region b                 where a.calling_date = ('"+kmrnbulan+"')                 and substr(a.calling_mdn,1,7)=b.prefix                 and calling_type in (1,2,7,77,64,65,78,93,95,96,99,100)                 and result=0 group BY a.calling_date, case                 when a.calling_type in (1,2,7,77) then 'Purchase'                 when a.calling_type in (64,78,65,93,95,96,99,100) then 'Renewal'                 end)")
                cursor.execute(query)
                df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append 3. UNIQ SUBS TRANSACTION PROCESS TYPE '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                    cursor.close()

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
                    if csvfile:
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
                        chat = '3. UNIQ SUBS TRANSACTION PROCESS TYPE has just been written on worksheet REGION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing 3. UNIQ SUBS TRANSACTION PROCESS TYPE on worksheet REGION xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

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

    def DTMF():
        def DTMF_ONLY():
            GetQuery = 'DTMF ONLY is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'DTMF_ONLY_'+str(kmrnbulan[0:6])+'.csv'
                cursor = database.cursor()
                query = ("select calling_date , sum(trans_count)  ,                 sum(case when result in (0,805) then trans_count else 0 end) ,                 sum(case when result in (0,805) then revenue else 0 end ) ,                 sum(case when result not in (0) then trans_count else 0 end)                 from cdr_report_goofree                 where calling_type in (1,77)                 and calling_date = to_char(sysdate-"+sys+",'yyyymmdd')                 and media_code like '%D'                 group by calling_date, (case when media_code like '%D' then 'dtmf' else                 null end) order by 1 asc")
                cursor.execute(query)
                df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append DTMF ONLY '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                    cursor.close()

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
                    if csvfile:
                        readCSV = csv.reader(csvfile, delimiter=',')
                        for r, row in enumerate(readCSV, start=3):
                            for c, col in enumerate(row):
                                dtmf.write(r,0, row[0], date_format)
                                dtmf.write(r,1, row[1], content)
                                dtmf.write(r,2, row[2], content)
                                dtmf.write(r,3, row[3], content)
                                dtmf.write(r,4, row[4], content)
                        chat = 'DTMF ONLY has just been written on worksheet DTMF xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing DTMF ONLY on worksheet DTMF xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

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
            GetQuery = 'DTMF ONLY is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                cursor = database.cursor()
                with open(dir+'DTMF_AREA.txt','r') as inserts:
                    query = inserts.read()
                cursor.execute(query)
                result = cursor.fetchall()
                dtmf.merge_range('G3:G4','AREA',bold)
                dtmf.merge_range('H3:I3','DTMF',bold)
                dtmf.write('H4','Traffic',bold)
                dtmf.write('I4','Revenue',bold)
                dtmf.set_column('G:I', 12)
                if result:
                    for r, row in enumerate(result, start=4):
                            for c, col in enumerate(row):
                                dtmf.write(r,6, row[0], date_format)
                                dtmf.write(r,7, row[1], content)
                                dtmf.write(r,8, row[2], content)
                    cursor.close()
                    chat = 'DTMF AREA has just been written on worksheet DTMF xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing DTMF AREA on worksheet DTMF xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

            GETDAT()

        def DTMF_REGION():
            GetQuery = 'DTMF REGION is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'DTMF_REGION_'+str(kmrnbulan[0:6])+'.csv'
                cursor = database.cursor()
                query = ("select calling_date ,Trx_Sumbagut,Rev_Sumbagut,Trx_Sumbagteng, Rev_Sumbagteng ,Trx_Sumbagsel , Rev_Sumbagsel ,                 Trx_Jabotabek, Rev_Jabotabek , Trx_Jabar , Rev_Jabar , Trx_Jateng_DIY , Rev_Jateng_DIY , Trx_Jatim ,  Rev_Jatim , Trx_Balinusra                 ,Rev_Balinusra ,Trx_Kalimantan , Rev_Kalimantan , Trx_Sulawesi , Rev_Sulawesi , Trx_Papua_Maluku , Rev_Papua_Maluku, Trx_Other                 , Rev_Other , Traffic_total , revenue_total FROM (                 select calling_date,                     sum(case when region='Sumbagut' then trans_count else 0 end)  Trx_Sumbagut,                         sum(case when region='Sumbagut' then revenue else 0 end)  Rev_Sumbagut,                         sum(case when region='Sumbagteng' then trans_count else 0 end) Trx_Sumbagteng,                     sum(case when region='Sumbagteng' then revenue else 0 end) Rev_Sumbagteng,                         sum(case when region='Sumbagsel' then trans_count else 0 end)  Trx_Sumbagsel,                         sum(case when region='Sumbagsel' then revenue else 0 end)  Rev_Sumbagsel,                     sum(case when region='Jabotabek' then trans_count else 0 end) Trx_Jabotabek,                     sum(case when region='Jabotabek'  then revenue else 0 end) Rev_Jabotabek,                         sum(case when region='Jabar' then trans_count else 0 end) Trx_Jabar,                     sum(case when region='Jabar' then revenue else 0 end) Rev_Jabar,                         sum(case when region='Jateng' then trans_count else 0 end) Trx_Jateng_DIY,                     sum(case when region='Jateng' then revenue else 0 end) Rev_Jateng_DIY,                         sum(case when region='Jatim' then trans_count else 0 end) Trx_Jatim,                     sum(case when region='Jatim' then revenue else 0 end) Rev_Jatim,                         sum(case when region='Balinusra' then trans_count else 0 end) Trx_Balinusra,                     sum(case when region='Balinusra' then revenue else 0 end) Rev_Balinusra,                         sum(case when region='Kalimantan' then trans_count else 0 end) Trx_Kalimantan,                     sum(case when region='Kalimantan' then revenue else 0 end) Rev_Kalimantan,                         sum(case when region='Sulawesi' then trans_count else 0 end)   Trx_Sulawesi,                     sum(case when region='Sulawesi' then revenue else 0 end)   Rev_Sulawesi,                          sum(case when region='Papua' then trans_count else 0 end)  Trx_Papua_Maluku,                     sum(case when region='Papua' then revenue else 0 end)  Rev_Papua_Maluku,                     sum(case when region not in ('Sumbagut','Sumbagteng','Sumbagsel','Jabotabek','Jabar','Jateng','Jatim','Balinusra','Kalimantan','Sulawesi','Papua') then                 trans_count else 0 end)  Trx_Other,                     sum(case when region not in ('Sumbagut','Sumbagteng','Sumbagsel','Jabotabek','Jabar','Jateng','Jatim','Balinusra','Kalimantan','Sulawesi','Papua') then                 revenue else 0 end)  Rev_Other,                         sum(trans_count)  Traffic_total,                     sum(revenue )  revenue_total                           from cdr_report_goofree                 where  calling_date = to_char(sysdate-"+sys+",'yyyymmdd')                 and calling_type in (1,77)                 and media_code like '%D'                 and result in (0,805)                 group by calling_date                 order by calling_date)")
                cursor.execute(query)
                df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append DTMF REGION '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                    cursor.close()

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
                    if csvfile:
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
                        chat = 'DTMF REGION has just been written on worksheet DTMF xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing DTMF REGION on worksheet DTMF xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

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

    def GP():
        GetQuery = 'GRACE PERIODE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'GP_'+str(kmrnbulan[0:6])+'.csv'
            cursor = database.cursor()
            query = ("select calling_date, sum(case when result = 0 then 1 end),             sum(case when result = 0 then charge_amount end) as revenue             from call_detailed_record_all where calling_date = to_char(sysdate-"+sys+",'YYYYMMDD')             and calling_type = '95' group by calling_date order by calling_date")
            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append GRACE PERIODE '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                cursor.close()

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
                if csvfile:
                    readCSV = csv.reader(csvfile, delimiter=',')
                    for r, row in enumerate(readCSV, start=4):
                        for c, col in enumerate(row):
                            gp.write(r,1, row[0], date_format)
                            gp.write(r,2, row[1], content)
                            gp.write(r,3, row[2], content)
                    chat = 'GRACE PERIODE has just been written on worksheet GRACE PERIODE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing GRACE PERIODE on worksheet GRACE PERIODE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def SMS_REVENUE():
        GetQuery = 'SMS REVENUE is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'SMS_REVENUE_'+str(kmrnbulan[0:6])+'.csv'
            cursor = database.cursor()
            query = ("select calling_date,count(*),sum(case when result=0 then 1 else 0 end),                     sum(case when result=0 then charge_amount else 0 end) from cdr_all_in_one                     where calling_date= to_char(sysdate-"+sys+", 'YYYYMMDD')                     and calling_type in (1,2,7,77) and media_code like '%0'                     group by calling_date")
            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append SMS REVENUE '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                cursor.close()

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
                if csvfile:
                    readCSV = csv.reader(csvfile, delimiter=',')
                    for r, row in enumerate(readCSV, start=4):
                        for c, col in enumerate(row):
                            smsrev.write(r,0, row[0], date_format)
                            smsrev.write(r,1, row[1], content)
                            smsrev.write(r,2, row[2], content)
                            smsrev.write(r,3, row[3], content)
                    chat = 'SMS REVENUE has just been written on worksheet SMS REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing SMS REVENUE on worksheet SMS REVENUE xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def REVENUE_RENEWAL():
        GetQuery = 'REVENUE RENEWAL is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'REVENUE_RENEWAL_'+str(kmrnbulan[0:6])+'.csv'
            cursor = database.cursor()
            query = ("select calling_date,sum(charge_amount) from cdr_all_in_one where calling_date=to_char(sysdate-"+sys+",'YYYYMMDD')                     and result=0 and calling_type in (64,93,95,96,99,100) and rating_id in (select rating_id from rating_mcp_info                     where initial_amount=0 and RECURRENT_RATING_CYCLE_DAY=7 and recurrent_amount between 600 and 3000) group by calling_date")
            cursor.execute(query)
            df = pd.DataFrame.from_records(cursor.fetchall(),columns = [desc[0] for desc in cursor.description])
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append REVENUE RENEWAL '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',',index=False, header=False, mode=''+ AsNeeded +'')
                cursor.close()

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
                if csvfile:
                    readCSV = csv.reader(csvfile, delimiter=',')
                    for r, row in enumerate(readCSV, start=2):
                        for c, col in enumerate(row):
                            revrenew.write(r,1, row[0], date_format)
                            revrenew.write(r,2, row[1], content)
                    chat = 'REVENUE RENEWAL has just been written on worksheet REVENUE RENEWAL xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing REVENUE RENEWAL on worksheet REVENUE RENEWAL xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def NSP_REKOMENDASI(): # https://pandas.pydata.org/docs/user_guide/merging.html
        GetQuery = 'NSP REKOMENDASI is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'NSP_Rekomendasi_'+str(kmrnbulan[0:6])+'.csv'
            stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/'+kmrnbulan+'/25.NSP_Rekomendasi_'+kmrnbulan+'_Traffic.csv')
            df1 = pd.read_csv(stdout, sep=',')
            # df1 = pd.read_csv(stdout, sep=',',usecols=["CALLING_DATE"])
            stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/'+kmrnbulan+'/25.NSP_Rekomendasi_'+kmrnbulan+'.csv')
            df2 = pd.read_csv(stdout, sep=',',usecols=["Total_SMS_Recommendation_Sent","Total_Attempt_Purchased","Total_Success_Purchased","Total_New_Subscriber_Activated","Total_Revenue_Purchased"])
            df = pd.concat([df1, df2], axis=1)
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append NSP REKOMENDASI '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
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
                if csvfile:
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
                    chat = 'NSP REKOMENDASI has just been written on worksheet NSP REKOMENDASI xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing NSP REKOMENDASI on worksheet NSP REKOMENDASI xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def SUCCESS_RATE_MT_AND_MO():
        GetQuery = 'success rate MT MO is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'success_rate_MT_MO_'+str(kmrnbulan[0:6])+'.csv'
            stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/'+kmrnbulan+'/26.success_rate_MT_MO_'+kmrnbulan+'.csv')
            df = pd.read_csv(stdout, sep=',')
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append success rate MT MO '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                fyi(append)
            else:
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
                if csvfile:
                    readCSV = csv.reader(csvfile, delimiter=',')
                    for r, row in enumerate(readCSV, start=3):
                        for c, col in enumerate(row):
                            succratemt.write(r,1, row[0], date_format)
                            succratemt.write(r,2, row[1], content)
                            succratemt.write(r,3, row[2], content)
                            succratemt.write(r,4, row[3], content)
                    chat = 'success rate MT MO has just been written on worksheet SUCCESS RATE MT AND MO xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing success rate MT MO on worksheet SUCCESS RATE MT AND MO xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)

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

    def TRANSACTION_UMB_TMENU():
        GetQuery = 'JUMLAH TRANSAKSI UMB T-MENU PROMO is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)

        stdin, stdout, stderr = ssh.exec_command('cat /TRBT/script/LIST_PROMO_UMB_TMENU.txt')
        LIST_PROMO = ''.join(stdout.readlines())

        def PROMO():
            cursor = database.cursor()
            query = ("SELECT TANGGAL, PROMO, Total_Transaksi FROM (SELECT calling_date AS TANGGAL,dvc_id AS PROMO,count(dvc_id) AS Total_Transaksi                     FROM call_detailed_record WHERE calling_date = to_char(sysdate-1, 'YYYYMMDD') AND calling_type IN (1, 77)                     AND dvc_id IN ("+LIST_PROMO+") AND media_code LIKE '%U' AND RESULT = 0 GROUP BY calling_date,dvc_id ORDER BY 3 DESC)")
            cursor.execute(query)
            result = cursor.fetchall()
            transumb.set_zoom(50)
            transumb.hide_gridlines(2)
            transumb.write('A1','JUMLAH TRANSAKSI UMB T-MENU PROMO AND HITS',title)
            transumb.write('A3','CALLING_DATE',bold)
            transumb.write('B3','PROMO',bold)
            transumb.write('C3','Total_Transaksi',bold)
            transumb.set_column('A:C', 16)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        transumb.write(r,0, row[0], date_format)
                        transumb.write(r,1, row[1], content)
                        transumb.write(r,2, row[2], content)
                cursor.close()
                chat = 'JUMLAH TRANSAKSI UMB T-MENU PROMO has just been written on worksheet TRANSACTION_UMB_TMENU_PROMO xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing JUMLAH TRANSAKSI UMB T-MENU PROMO on worksheet TRANSACTION_UMB_TMENU_PROMO xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def HITS():
            GetQuery = 'JUMLAH TRANSAKSI UMB T-MENU HITS is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            query = ("select TANGGAL,HITS,Total_Transaksi From (select calling_date as TANGGAL,dvc_id as HITS,count(dvc_id) as Total_Transaksi from call_detailed_record             where calling_date = to_char(sysdate-1, 'YYYYMMdd') and calling_type in (1,77) and dvc_id not in ("+LIST_PROMO+") and media_code like '%U' and result=0             group by calling_date,dvc_id order by 3 desc)")
            cursor.execute(query)
            result = cursor.fetchall()
            transumb.write('A26','CALLING_DATE',bold)
            transumb.write('B26','HITS',bold)
            transumb.write('C26','Total_Transaksi',bold)
            if result:
                for r, row in enumerate(result, start=26):
                    for c, col in enumerate(row):
                        transumb.write(r,0, row[0], date_format)
                        transumb.write(r,1, row[1], content)
                        transumb.write(r,2, row[2], content)
                cursor.close()
                chat = 'JUMLAH TRANSAKSI UMB T-MENU HITS has just been written on worksheet TRANSACTION_UMB_TMENU_PROMO xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing JUMLAH TRANSAKSI UMB T-MENU HITS on worksheet TRANSACTION_UMB_TMENU_PROMO xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        PROMO()
        HITS()

    SUMMARY()
    NEW_SUBS()
    PROV()
    TOTAL_TRAFFIC()
    PRICEREVENUE()
    CHARGINGCDR()
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


# In[ ]:


def NewRBTDailyReport():
    workbook = xlsxwriter.Workbook(dir+"New_RBT_Daily_Report_"+str(kemarin_str)+".xlsx",{'strings_to_numbers': True})
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
    
    subs = workbook.add_worksheet('SUBS')
    subs_hourly = workbook.add_worksheet('SUBS_HOURLY')
    umb = workbook.add_worksheet('UMB')
    trxz = workbook.add_worksheet('TRX')
    trxz_hourly = workbook.add_worksheet('TRX_HOURLY')
    historicaldatsum = workbook.add_worksheet('HISTORICAL DATA SUMMARY')
    nspmembersince = workbook.add_worksheet('NSP_MEMBER_SINCE_2015')

    for TabCon in ['subs','subs_hourly','umb','trxz','trxz_hourly','historicaldatsum','nspmembersince']:
        TabCol = TabCon + ".set_tab_color('#00B050')"
        exec(TabCol)

    def SUB():
        def Summary():
            GetQuery = 'Summary for worksheet SUBS on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            def GETDAT():
                Dump = dir+'subs_summary_new_'+str(kmrnbulan)+'.csv'
                stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/new_report/'+kmrnbulan+'/1_subs_summary_new_'+kmrnbulan+'.csv')
                df = pd.read_csv(stdout, sep=',')
                number_rows = len(df.index)
                if number_rows == 0:
                    append = 'Failed to append Summary for worksheet SUBS on (New)RBTDailyReport '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',', header=False, mode='w')

                for line in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']:
                    for kolom in range(1):
                        Alline = line+'{}'.format(*([kolom + 5] * 1))
                        subs.write(Alline,'',neat)

                subs.set_zoom(55)
                subs.hide_gridlines(2)
                subs.write('A1','Summary :',title)
                subs.merge_range('A2:A4','Calling Date',bold)
                subs.merge_range('B2:F2','SUBSCRIBER',bold)
                subs.merge_range('B3:C3','NEW SUBSCRIBER',bold)
                subs.write('B4','NEW_FREE_USER',bold)
                subs.write('C4','NEW_PAID_USER',bold)
                subs.merge_range('D3:D4','STAY_USER',bold)
                subs.merge_range('E3:E4','ACTIVE_TO_SUSPEND_USER',bold)
                subs.merge_range('F3:F4','SUSPEND_TO_ACTIVE_USER',bold)
                subs.merge_range('G2:G4','TOTAL_ACTIVE_USER',bold)
                subs.merge_range('H2:H4','DAILY_NET_ADD',bold)
                subs.merge_range('I2:I4','TOTAL_ACTIVE_THIS_YEAR',bold)
                subs.merge_range('J2:P2','CHURN USER',bold)
                subs.merge_range('J3:J4','VOLUNTARY_CHURN',bold)
                subs.merge_range('K3:K4','FULL_DELETE_CONTENT',bold)
                subs.merge_range('L3:L4','ACTIVE_DELETE_CONTENT',bold)
                subs.merge_range('M3:M4','CHURN_INSUFFICIENT_BALANCE',bold)
                subs.merge_range('N3:N4','CHURN_ERROR_CHARGING',bold)
                subs.merge_range('O3:O4','OTHERS',bold)
                subs.merge_range('P3:P4','SUSPEND_IN_THREE_DAYS',bold)
                subs.merge_range('Q2:Q4','TOTAL_CHURN_USER',bold)
                subs.merge_range('R2:R4','ALL_MEMBERS_THIS_YEAR',bold)        
                subs.set_column('A:A', 20)
                subs.set_column('B:B', 23)
                subs.set_column('C:R', 35)
                subs.freeze_panes(0, 1)

                with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                    if csvfile:
                        readCSV = csv.reader(csvfile, delimiter=',')
                        for r, row in enumerate(readCSV, start=4):
                            for c, col in enumerate(row):
                                subs.write(r,0, row[0], date_format)
                                subs.write(r,1, row[1], content)
                                subs.write(r,2, row[2], content)
                                subs.write(r,3, row[3], content)
                                subs.write(r,4, row[4], content)
                                subs.write(r,5, row[5], content)
                                subs.write(r,6, row[6], content)
                                subs.write(r,7, row[7], content)
                                subs.write(r,8, row[8], content)
                                subs.write(r,9, row[9], content)
                                subs.write(r,10, row[10], content)
                                subs.write(r,11, row[11], content)
                                subs.write(r,12, row[12], content)
                                subs.write(r,13, row[13], content)
                                subs.write(r,14, row[14], content)
                                subs.write(r,15, row[15], content)
                                subs.write(r,16, row[16], content)
                                subs.write(r,17, row[17], content)
                        chat = 'Summary has just been written on worksheet SUBS (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                        print(chat)
                        fyi(chat)
                    else:
                        Alert = 'Skip writing Summary on worksheet SUBS (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                        print(Alert)
                        fyi(Alert)

            sys = '1'
            lusa = now - datetime.timedelta(days=2)
            storyday = now - datetime.timedelta(days=int(sys))
            kmrnlusa = str(lusa.strftime('%Y%m%d'))
            kmrnbulan = str(storyday.strftime('%Y%m%d'))
            if kmrnbulan[0:6] == kmrnlusa[0:6]:
                GETDAT()
            else:
                GETDAT()

        def DataPrice():
            for line in ['A','B','C','D','E']:
                for kolom in range(33):
                    Alline = line+'{}'.format(*([kolom + 10] * 33))
                    subs.write(Alline,'',neat)

            subs.write('A8','Data per-Price :',title)
            subs.write('A9','PRICE',bold)
            subs.write('B9','NEW_SUBS',bold)
            subs.write('C9','SUSPEND_TO_ACTIVE_USER',bold)
            subs.write('D9','CHURN_INSUFFICIENT_BALANCE',bold)
            subs.write('E9','CHURN_ERROR_CHARGING',bold)
            subs.write('A10','0',content)
            subs.write('A11','1',content)
            subs.write('A12','10',content)
            subs.write('A13','100',content)
            subs.write('A14','200',content)
            subs.write('A15','300',content)
            subs.write('A16','400',content)
            subs.write('A17','500',content)
            subs.write('A18','600',content)
            subs.write('A19','700',content)
            subs.write('A20','800',content)
            subs.write('A21','900',content)
            subs.write('A22','1000',content)
            subs.write('A23','1500',content)
            subs.write('A24','1818',content)
            subs.write('A25','2000',content)
            subs.write('A26','2500',content)
            subs.write('A27','2727',content)
            subs.write('A28','3000',content)
            subs.write('A29','3500',content)
            subs.write('A30','4000',content)
            subs.write('A31','4444',content)
            subs.write('A32','4500',content)
            subs.write('A33','4545',content)
            subs.write('A34','5000',content)
            subs.write('A35','5500',content)
            subs.write('A36','6000',content)
            subs.write('A37','6500',content)
            subs.write('A38','7000',content)
            subs.write('A39','7500',content)
            subs.write('A40','8000',content)
            subs.write('A41','8500',content)
            subs.write('A42','9000',content)

        def DataPerChannel():
            for line in ['A','B','C','D','E','F','G','H','I','J']:
                for kolom in range(39):
                    Alline = line+'{}'.format(*([kolom + 46] * 39))
                    subs.write(Alline,'',neat)

            subs.write('A44','Data Per Channel :',title)
            subs.write('A45','CHANNEL',bold)
            subs.write('B45','NEW_FREE_USER',bold)
            subs.write('C45','NEW_PAID_USER',bold)
            subs.write('D45','ACTIVE_TO_SUSPEND_USER',bold)
            subs.write('E45','SUSPEND_TO_ACTIVE_USER',bold)
            subs.write('F45','VOLUNTARY_CHURN',bold)
            subs.write('G45','FULL_DELETE_CONTENT',bold)
            subs.write('H45','ACTIVE_DELETE_CONTENT',bold)
            subs.write('I45','CHURN_INSUFFICIENT_BALANCE',bold)
            subs.write('J45','CHURN_ERROR_CHARGING',bold)
            subs.write('A46','CSWeb',content)
            subs.write('A47','DTMF',content)
            subs.write('A48','HTTPGW',content)
            subs.write('A49','IVR',content)
            subs.write('A50','MKios',content)
            subs.write('A51','Music Max',content)
            subs.write('A52','My Telkomsel',content)
            subs.write('A53','MyNSP',content)
            subs.write('A54','OpenAPI',content)
            subs.write('A55','PCU',content)
            subs.write('A56','Recomendation',content)
            subs.write('A57','Referal',content)
            subs.write('A58','Renewal',content)
            subs.write('A59','SMS',content)
            subs.write('A60','Sms Broadcast',content)
            subs.write('A61','UMB',content)
            subs.write('A62','WCC',content)
            subs.write('A63','WEB',content)
            subs.write('A64','Media_1',content)
            subs.write('A65','Media_2',content)
            subs.write('A66','Media_3',content)
            subs.write('A67','Media_4',content)
            subs.write('A68','Media_5',content)
            subs.write('A69','Media_6',content)
            subs.write('A70','Media_7',content)
            subs.write('A71','Media_8',content)
            subs.write('A72','Media_9',content)
            subs.write('A73','Media_A',content)
            subs.write('A74','Media_E',content)
            subs.write('A75','Media_F',content)
            subs.write('A76','Media_G',content)
            subs.write('A77','Media_H',content)
            subs.write('A78','Media_K',content)
            subs.write('A79','Media_L',content)
            subs.write('A80','Media_Q',content)
            subs.write('A81','Media_T',content)
            subs.write('A82','Media_F',content)
            subs.write('A83','Media_R',content)
            subs.write('A84','Point',content)

        def DataperRegion():
            for line in ['A','B','C','D','E','F','G','H','I','J','K','L','M','N']:
                for kolom in range(11):
                    Alline = line+'{}'.format(*([kolom + 89] * 11))
                    subs.write(Alline,'',neat)

            subs.write('A87','Data per-Region :',title)
            subs.write('A88','REGION',bold)
            subs.write('B88','NEW_FREE_USER',bold)
            subs.write('C88','NEW_PAID_USER',bold)
            subs.write('D88','STAY_USER',bold)
            subs.write('E88','ACTIVE_TO_SUSPEND_USER',bold)
            subs.write('F88','SUSPEND_TO_ACTIVE_USER',bold)
            subs.write('G88','TOTAL_ACTIVE_THIS_YEAR',bold)
            subs.write('H88','VOLUNTARY_CHURN',bold)
            subs.write('I88','FULL_DELETE_CONTENT',bold)
            subs.write('J88','ACTIVE_DELETE_CONTENT',bold)
            subs.write('K88','CHURN_INSUFFICIENT_BALANCE',bold)
            subs.write('L88','CHURN_ERROR_CHARGING',bold)
            subs.write('M88','SUSPEND_IN_THREE_DAYS',bold)
            subs.write('N88','ACTIVE_USER',bold)
            subs.write('A89','Balinusra',content)
            subs.write('A90','Jabar',content)
            subs.write('A91','Jabotabek',content)
            subs.write('A92','Jateng',content)
            subs.write('A93','Jatim',content)
            subs.write('A94','Kalimantan',content)
            subs.write('A95','Papua',content)
            subs.write('A96','Sulawesi',content)
            subs.write('A97','Sumbagsel',content)
            subs.write('A98','Sumbagteng',content)
            subs.write('A99','Sumbagut',content)

        Summary()
        DataPrice()
        DataPerChannel()
        DataperRegion()

    def SUB_HOURLY():
        def NEW_SUBSCRIPTION_TOTAL():
            GetQuery = 'NEW_SUBSCRIPTION_TOTAL for worksheet SUBS_HOURLY on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('NEW_SUBSCRIPTION_TOTAL.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            subs_hourly.set_zoom(70)
            subs_hourly.hide_gridlines(2)
            subs_hourly.write('A1','A. NEW SUBSCRIPTION TOTAL',title)
            subs_hourly.write('A2','Calling Date',bold)
            subs_hourly.write('B2','CALLING_HOUR',bold)
            subs_hourly.write('C2','CHANNEL',bold)
            subs_hourly.write('D2','NEW_FREE_USER',bold)
            subs_hourly.write('E2','NEW_PAID_USER',bold)
            subs_hourly.set_column('A:A', 12)
            subs_hourly.set_column('B:C', 15)
            subs_hourly.set_column('D:E', 18)
            subs_hourly.set_column('F:F', 1)
            if result:
                for r, row in enumerate(result, start=2):
                    for c, col in enumerate(row):
                        subs_hourly.write(r,0, row[0], date_format)
                        subs_hourly.write(r,1, row[1], content)
                        subs_hourly.write(r,2, row[2], content)
                        subs_hourly.write(r,3, row[3], content)
                        subs_hourly.write(r,4, row[4], content)
                cursor.close()
                chat = 'NEW_SUBSCRIPTION_TOTAL has just been written for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing NEW_SUBSCRIPTION_TOTAL for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def NEW_SUBSCRIPTION_CHANNEL():
            GetQuery = 'NEW_SUBSCRIPTION_CHANNEL for worksheet SUBS_HOURLY on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('NEW_SUBSCRIPTION_CHANNEL.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            subs_hourly.write('G1','B. NEW SUBSCRIPTION CHANNEL',title)
            subs_hourly.write('G2','Calling Date',bold)
            subs_hourly.write('H2','CALLING_HOUR',bold)
            subs_hourly.write('I2','CHANNEL',bold)
            subs_hourly.write('J2','ACTIVE_TO_SUSPEND_USER',bold)
            subs_hourly.set_column('G:G', 12)
            subs_hourly.set_column('H:I', 15)
            subs_hourly.set_column('J:J', 25)
            subs_hourly.set_column('K:K', 1)
            if result:
                for r, row in enumerate(result, start=2):
                    for c, col in enumerate(row):
                        subs_hourly.write(r,6, row[0], date_format)
                        subs_hourly.write(r,7, row[1], content)
                        subs_hourly.write(r,8, row[2], content)
                        subs_hourly.write(r,9, row[3], content)
                cursor.close
                chat = 'NEW_SUBSCRIPTION_CHANNEL has just been written for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing NEW_SUBSCRIPTION_CHANNEL for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def NEW_SUBSCRIPTION_HOUR():
            GetQuery = 'NEW_SUBSCRIPTION_HOUR for worksheet SUBS_HOURLY on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('NEW_SUBSCRIPTION_HOUR.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            subs_hourly.write('L1','C. NEW SUBSCRIPTION HOUR',title)
            subs_hourly.write('L2','Calling Date',bold)
            subs_hourly.write('M2','CALLING_HOUR',bold)
            subs_hourly.write('N2','CHANNEL',bold)
            subs_hourly.write('O2','SUSPEND_TO_ACTIVE_USER',bold)
            subs_hourly.set_column('L:L', 12)
            subs_hourly.set_column('M:N', 15)
            subs_hourly.set_column('O:O', 25)
            subs_hourly.set_column('P:P', 1)
            if result:
                for r, row in enumerate(result, start=2):
                    for c, col in enumerate(row):
                        subs_hourly.write(r,11, row[0], date_format)
                        subs_hourly.write(r,12, row[1], content)
                        subs_hourly.write(r,13, row[2], content)
                        subs_hourly.write(r,14, row[3], content)
                cursor.close()
                chat = 'NEW_SUBSCRIPTION_HOUR has just been written for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing NEW_SUBSCRIPTION_HOUR for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def NEW_SUBSCRIPTION_HOUR_CHANNEL():
            GetQuery = 'NEW_SUBSCRIPTION_HOUR_CHANNEL for worksheet SUBS_HOURLY on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('NEW_SUBSCRIPTION_HOUR_CHANNEL.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            subs_hourly.write('Q1','D. NEW SUBSCRIPTION HOUR CHANNEL',title)
            subs_hourly.write('Q2','Calling Date',bold)
            subs_hourly.write('R2','CALLING_HOUR',bold)
            subs_hourly.write('S2','CHANNEL',bold)
            subs_hourly.write('T2','VOLUNTARY_CHURN',bold)
            subs_hourly.write('U2','FULL_DELETE_CONTENT',bold)
            subs_hourly.write('V2','ACTIVE_DELETE_CONTENT',bold)
            subs_hourly.write('W2','CHURN_INSUFFICIENT_BALANCE',bold)
            subs_hourly.write('X2','CHURN_ERROR_CHARGING',bold)
            subs_hourly.set_column('Q:Q', 12)
            subs_hourly.set_column('R:S', 15)
            subs_hourly.set_column('T:X', 37)
            if result:
                for r, row in enumerate(result, start=2):
                    for c, col in enumerate(row):
                        subs_hourly.write(r,16, row[0], date_format)
                        subs_hourly.write(r,17, row[1], content)
                        subs_hourly.write(r,18, row[2], content)
                        subs_hourly.write(r,19, row[3], content)
                        subs_hourly.write(r,20, row[4], content)
                        subs_hourly.write(r,21, row[5], content)
                        subs_hourly.write(r,22, row[6], content)
                        subs_hourly.write(r,23, row[7], content)
                cursor.close()
                chat = 'NEW_SUBSCRIPTION_HOUR_CHANNEL has just been written for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing NEW_SUBSCRIPTION_HOUR_CHANNEL for worksheet SUBS_HOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        NEW_SUBSCRIPTION_TOTAL()
        NEW_SUBSCRIPTION_CHANNEL()
        NEW_SUBSCRIPTION_HOUR()
        NEW_SUBSCRIPTION_HOUR_CHANNEL()

    def UMBS():
        def AREA1():
            GetQuery = 'UMBS_AREA1 for worksheet UMB on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('UMBS_AREA1.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            umb.set_zoom(60)
            umb.hide_gridlines(2)
            umb.merge_range('B1:G1','AREA 1',bold)
            umb.merge_range('B2:B3','UMB',bold)
            umb.merge_range('C2:D2','ATTEMPT',bold)
            umb.write('C3','Free',bold)
            umb.write('D3','Paid',bold)
            umb.merge_range('E2:F2','SUCCESS',bold)
            umb.write('E3','Free',bold)
            umb.write('F3','Paid',bold)
            umb.merge_range('G2:G3','REVENUE',bold)
            umb.set_column('A:A', 2)
            umb.set_column('B:B', 17)
            umb.set_column('C:F', 9)
            umb.set_column('G:G', 13)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        umb.write(r,1, row[1], date_format)
                        umb.write(r,2, row[2], content)
                        umb.write(r,3, row[3], content)
                        umb.write(r,4, row[4], content)
                        umb.write(r,5, row[5], content)
                        umb.write(r,6, row[6], content)
                cursor.close()
                chat = 'UMBS_AREA1 has just been written for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing UMBS_AREA1 for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def AREA2():
            GetQuery = 'UMBS_AREA2 for worksheet UMB on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('UMBS_AREA2.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            umb.merge_range('I1:N1','AREA 2',bold)
            umb.merge_range('I2:I3','UMB',bold)
            umb.merge_range('J2:K2','ATTEMPT',bold)
            umb.write('J3','Free',bold)
            umb.write('K3','Paid',bold)
            umb.merge_range('L2:M2','SUCCESS',bold)
            umb.write('L3','Free',bold)
            umb.write('M3','Paid',bold)
            umb.merge_range('N2:N3','REVENUE',bold)
            umb.set_column('H:H', 2)
            umb.set_column('I:I', 17)
            umb.set_column('J:M', 9)
            umb.set_column('N:N', 13)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        umb.write(r,8, row[1], date_format)
                        umb.write(r,9, row[2], content)
                        umb.write(r,10, row[3], content)
                        umb.write(r,11, row[4], content)
                        umb.write(r,12, row[5], content)
                        umb.write(r,13, row[6], content)
                cursor.close()
                chat = 'UMBS_AREA2 has just been written for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing UMBS_AREA2 for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def AREA3():
            GetQuery = 'UMBS_AREA3 for worksheet UMB on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('UMBS_AREA3.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            umb.merge_range('P1:U1','AREA 3',bold)
            umb.merge_range('P2:P3','UMB',bold)
            umb.merge_range('Q2:R2','ATTEMPT',bold)
            umb.write('Q3','Free',bold)
            umb.write('R3','Paid',bold)
            umb.merge_range('S2:T2','SUCCESS',bold)
            umb.write('S3','Free',bold)
            umb.write('T3','Paid',bold)
            umb.merge_range('U2:U3','REVENUE',bold)
            umb.set_column('O:O', 2)
            umb.set_column('P:P', 17)
            umb.set_column('Q:T', 9)
            umb.set_column('U:U', 13)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        umb.write(r,15, row[1], date_format)
                        umb.write(r,16, row[2], content)
                        umb.write(r,17, row[3], content)
                        umb.write(r,18, row[4], content)
                        umb.write(r,19, row[5], content)
                        umb.write(r,20, row[6], content)
                cursor.close()
                chat = 'UMBS_AREA3 has just been written for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing UMBS_AREA3 for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def AREA4():
            GetQuery = 'UMBS_AREA4 for worksheet UMB on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('UMBS_AREA4.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            umb.merge_range('W1:AB1','AREA 4',bold)
            umb.merge_range('W2:W3','UMB',bold)
            umb.merge_range('X2:Y2','ATTEMPT',bold)
            umb.write('X3','Free',bold)
            umb.write('Y3','Paid',bold)
            umb.merge_range('Z2:AA2','SUCCESS',bold)
            umb.write('Z3','Free',bold)
            umb.write('AA3','Paid',bold)
            umb.merge_range('AB2:AB3','REVENUE',bold)
            umb.set_column('V:V', 2)
            umb.set_column('W:W', 17)
            umb.set_column('X:AA', 9)
            umb.set_column('AB:AB', 13)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        umb.write(r,22, row[1], date_format)
                        umb.write(r,23, row[2], content)
                        umb.write(r,24, row[3], content)
                        umb.write(r,25, row[4], content)
                        umb.write(r,26, row[5], content)
                        umb.write(r,27, row[6], content)
                cursor.close()
                chat = 'UMBS_AREA4 has just been written for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing UMBS_AREA4 for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def TOTAL():
            GetQuery = 'TOTAL for worksheet UMB on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('UMBS_TOTAL.txt','r') as inserts:
                query = inserts.read()
            cursor.execute(query)
            result = cursor.fetchall()
            umb.merge_range('AD1:AI1','TOTAL',bold)
            umb.merge_range('AD2:AD3','UMB',bold)
            umb.merge_range('AE2:AF2','ATTEMPT',bold)
            umb.write('AE3','Free',bold)
            umb.write('AF3','Paid',bold)
            umb.merge_range('AG2:AH2','SUCCESS',bold)
            umb.write('AG3','Free',bold)
            umb.write('AH3','Paid',bold)
            umb.merge_range('AI2:AI3','REVENUE',bold)
            umb.set_column('AC:AC', 2)
            umb.set_column('AD:AD', 17)
            umb.set_column('AE:AH', 9)
            umb.set_column('AI:AI', 13)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        umb.write(r,29, row[0], date_format)
                        umb.write(r,30, row[1], content)
                        umb.write(r,31, row[2], content)
                        umb.write(r,32, row[3], content)
                        umb.write(r,33, row[4], content)
                        umb.write(r,34, row[5], content)
                cursor.close()
                chat = 'TOTAL has just been written for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing TOTAL for worksheet UMB on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        AREA1()
        AREA2()
        AREA3()
        AREA4()
        TOTAL()

    def TRX():
        Colm = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y']
        def PURCHASE_AUTO():
            GetQuery = 'PURCHASE_AUTO for worksheet TRX on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('PURCHASE_AUTO.txt','r') as inserts:
                query = inserts.read()

            cursor.execute(query)
            result = cursor.fetchall()

            for line in Colm:
                    for kolom in range(1):
                        Alline = line+'{}'.format(*([kolom + 4] * 1))
                        trxz.write(Alline,'',neat)

            trxz.set_zoom(70)
            trxz.hide_gridlines(2)
            trxz.merge_range('A1:A3','DATE',bold)
            trxz.merge_range('B1:G1','PURCHASE AUTORENEWAL ON',bold)
            trxz.merge_range('B2:C2','ATTEMPT',bold)
            trxz.write('B3','FREE',bold)
            trxz.write('C3','PAID',bold)
            trxz.merge_range('D2:E2','SUCCESS',bold)
            trxz.write('D3','FREE',bold)
            trxz.write('E3','PAID',bold)
            trxz.merge_range('F2:F3','SR (%)',bold)
            trxz.merge_range('G2:G3','REVENUE',bold)
            trxz.merge_range('H1:M1','PURCHASE AUTORENEWAL OFF',bold)
            trxz.merge_range('H2:I2','ATTEMPT',bold)
            trxz.write('H3','FREE',bold)
            trxz.write('I3','PAID',bold)
            trxz.merge_range('J2:K2','SUCCESS',bold)
            trxz.write('J3','FREE',bold)
            trxz.write('K3','PAID',bold)
            trxz.merge_range('L2:L3','SR (%)',bold)
            trxz.merge_range('M2:M3','REVENUE',bold)
            trxz.merge_range('N1:S1','RENEWAL AUTORENEWAL ON',bold)
            trxz.merge_range('N2:O2','ATTEMPT',bold)
            trxz.write('N3','FREE',bold)
            trxz.write('O3','PAID',bold)
            trxz.merge_range('P2:Q2','SUCCESS',bold)
            trxz.write('P3','FREE',bold)
            trxz.write('Q3','PAID',bold)
            trxz.merge_range('R2:R3','SR (%)',bold)
            trxz.merge_range('S2:S3','REVENUE',bold)
            trxz.merge_range('T1:Y1','RENEWAL AUTORENEWAL OFF',bold)
            trxz.merge_range('T2:U2','ATTEMPT',bold)
            trxz.write('T3','FREE',bold)
            trxz.write('U3','PAID',bold)
            trxz.merge_range('V2:W2','SUCCESS',bold)
            trxz.write('V3','FREE',bold)
            trxz.write('W3','PAID',bold)
            trxz.merge_range('X2:X3','SR (%)',bold)
            trxz.merge_range('Y2:Y3','REVENUE',bold)
            trxz.set_column('A:A', 20)
            trxz.set_column('B:E', 25)
            trxz.set_column('F:G', 14)
            trxz.set_column('H:K', 25)
            trxz.set_column('L:M', 14)
            trxz.set_column('N:Q', 25)
            trxz.set_column('R:S', 14)
            trxz.set_column('T:W', 25)
            trxz.set_column('X:Y', 14)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        trxz.write(r,0, row[0], date_format)
                        trxz.write(r,1, row[1], content)
                        trxz.write(r,2, row[2], content)
                        trxz.write(r,3, row[3], content)
                        trxz.write(r,4, row[4], content)
                        trxz.write(r,5, row[5], content)
                        trxz.write(r,6, row[6], content)
                        trxz.write(r,7, row[7], content)
                        trxz.write(r,8, row[8], content)
                        trxz.write(r,9, row[9], content)
                        trxz.write(r,10, row[10], content)
                        trxz.write(r,11, row[11], content)
                        trxz.write(r,12, row[12], content)
                        trxz.write(r,13, row[13], content)
                        trxz.write(r,14, row[14], content)
                        trxz.write(r,15, row[15], content)
                        trxz.write(r,16, row[16], content)
                        trxz.write(r,17, row[17], content)
                        trxz.write(r,18, row[18], content)
                        trxz.write(r,19, row[19], content)
                        trxz.write(r,20, row[20], content)
                        trxz.write(r,21, row[21], content)
                        trxz.write(r,22, row[22], content)
                        trxz.write(r,23, row[23], content)
                        trxz.write(r,24, row[24], content)
                cursor.close()
                chat = 'PURCHASE_AUTO has just been written for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing PURCHASE_AUTO for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def PRICE():
            GetQuery = 'PRICE for worksheet TRX on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('PRICE.txt','r') as inserts:
                query = inserts.read()

            cursor.execute(query)
            result = cursor.fetchall()

            for line in Colm:
                    for kolom in range(33):
                        Alline = line+'{}'.format(*([kolom + 10] * 33))
                        trxz.write(Alline,'',neat)

            trxz.merge_range('A7:A9','PRICE',bold)
            trxz.merge_range('B7:G7','PURCHASE AUTORENEWAL ON',bold)
            trxz.merge_range('B8:C8','ATTEMPT',bold)
            trxz.write('B9','FREE',bold)
            trxz.write('C9','PAID',bold)
            trxz.merge_range('D8:E8','SUCCESS',bold)
            trxz.write('D9','FREE',bold)
            trxz.write('E9','PAID',bold)
            trxz.merge_range('F8:F9','SR (%)',bold)
            trxz.merge_range('G8:G9','REVENUE',bold)
            trxz.merge_range('H7:M7','PURCHASE AUTORENEWAL OFF',bold)
            trxz.merge_range('H8:I8','ATTEMPT',bold)
            trxz.write('H9','FREE',bold)
            trxz.write('I9','PAID',bold)
            trxz.merge_range('J8:K8','SUCCESS',bold)
            trxz.write('J9','FREE',bold)
            trxz.write('K9','PAID',bold)
            trxz.merge_range('L8:L9','SR (%)',bold)
            trxz.merge_range('M8:M9','REVENUE',bold)
            trxz.merge_range('N7:S7','RENEWAL AUTORENEWAL ON',bold)
            trxz.merge_range('N8:O8','ATTEMPT',bold)
            trxz.write('N9','FREE',bold)
            trxz.write('O9','PAID',bold)
            trxz.merge_range('P8:Q8','SUCCESS',bold)
            trxz.write('P9','FREE',bold)
            trxz.write('Q9','PAID',bold)
            trxz.merge_range('R8:R9','SR (%)',bold)
            trxz.merge_range('S8:S9','REVENUE',bold)
            trxz.merge_range('T7:Y7','RENEWAL AUTORENEWAL OFF',bold)
            trxz.merge_range('T8:U8','ATTEMPT',bold)
            trxz.write('T9','FREE',bold)
            trxz.write('U9','PAID',bold)
            trxz.merge_range('V8:W8','SUCCESS',bold)
            trxz.write('V9','FREE',bold)
            trxz.write('W9','PAID',bold)
            trxz.merge_range('X8:X9','SR (%)',bold)
            trxz.merge_range('Y8:Y9','REVENUE',bold)
            if result:
                for r, row in enumerate(result, start=9):
                    for c, col in enumerate(row):
                        trxz.write(r,0, row[0], date_format)
                        trxz.write(r,1, row[1], content)
                        trxz.write(r,2, row[2], content)
                        trxz.write(r,3, row[3], content)
                        trxz.write(r,4, row[4], content)
                        trxz.write(r,5, row[5], content)
                        trxz.write(r,6, row[6], content)
                        trxz.write(r,7, row[7], content)
                        trxz.write(r,8, row[8], content)
                        trxz.write(r,9, row[9], content)
                        trxz.write(r,10, row[10], content)
                        trxz.write(r,11, row[11], content)
                        trxz.write(r,12, row[12], content)
                        trxz.write(r,13, row[13], content)
                        trxz.write(r,14, row[14], content)
                        trxz.write(r,15, row[15], content)
                        trxz.write(r,16, row[16], content)
                        trxz.write(r,17, row[17], content)
                        trxz.write(r,18, row[18], content)
                        trxz.write(r,19, row[19], content)
                        trxz.write(r,20, row[20], content)
                        trxz.write(r,21, row[21], content)
                        trxz.write(r,22, row[22], content)
                        trxz.write(r,23, row[23], content)
                        trxz.write(r,24, row[24], content)
                cursor.close()
                chat = 'PRICE has just been written for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing PRICE for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def CHANNEL():
            GetQuery = 'CHANNEL for worksheet TRX on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('CHANNEL.txt','r') as inserts:
                query = inserts.read()

            cursor.execute(query)
            result = cursor.fetchall()

            for line in Colm:
                    for kolom in range(19):
                        Alline = line+'{}'.format(*([kolom + 47] * 19))
                        trxz.write(Alline,'',neat)

            trxz.merge_range('A44:A46','CHANNEL',bold)
            trxz.merge_range('B44:G44','PURCHASE AUTORENEWAL ON',bold)
            trxz.merge_range('B45:C45','ATTEMPT',bold)
            trxz.write('B46','FREE',bold)
            trxz.write('C46','PAID',bold)
            trxz.merge_range('D45:E45','SUCCESS',bold)
            trxz.write('D46','FREE',bold)
            trxz.write('E46','PAID',bold)
            trxz.merge_range('F45:F46','SR (%)',bold)
            trxz.merge_range('G45:G46','REVENUE',bold)
            trxz.merge_range('H44:M44','PURCHASE AUTORENEWAL OFF',bold)
            trxz.merge_range('H45:I45','ATTEMPT',bold)
            trxz.write('H46','FREE',bold)
            trxz.write('I46','PAID',bold)
            trxz.merge_range('J45:K45','SUCCESS',bold)
            trxz.write('J46','FREE',bold)
            trxz.write('K46','PAID',bold)
            trxz.merge_range('L45:L46','SR (%)',bold)
            trxz.merge_range('M45:M46','REVENUE',bold)
            trxz.merge_range('N44:S44','RENEWAL AUTORENEWAL ON',bold)
            trxz.merge_range('N45:O45','ATTEMPT',bold)
            trxz.write('N46','FREE',bold)
            trxz.write('O46','PAID',bold)
            trxz.merge_range('P45:Q45','SUCCESS',bold)
            trxz.write('P46','FREE',bold)
            trxz.write('Q46','PAID',bold)
            trxz.merge_range('R45:R46','SR (%)',bold)
            trxz.merge_range('S45:S46','REVENUE',bold)
            trxz.merge_range('T44:Y44','RENEWAL AUTORENEWAL OFF',bold)
            trxz.merge_range('T45:U45','ATTEMPT',bold)
            trxz.write('T46','FREE',bold)
            trxz.write('U46','PAID',bold)
            trxz.merge_range('V45:W45','SUCCESS',bold)
            trxz.write('V46','FREE',bold)
            trxz.write('W46','PAID',bold)
            trxz.merge_range('X45:X46','SR (%)',bold)
            trxz.merge_range('Y45:Y46','REVENUE',bold)
            if result:
                for r, row in enumerate(result, start=46):
                    for c, col in enumerate(row):
                        trxz.write(r,0, row[0], date_format)
                        trxz.write(r,1, row[1], content)
                        trxz.write(r,2, row[2], content)
                        trxz.write(r,3, row[3], content)
                        trxz.write(r,4, row[4], content)
                        trxz.write(r,5, row[5], content)
                        trxz.write(r,6, row[6], content)
                        trxz.write(r,7, row[7], content)
                        trxz.write(r,8, row[8], content)
                        trxz.write(r,9, row[9], content)
                        trxz.write(r,10, row[10], content)
                        trxz.write(r,11, row[11], content)
                        trxz.write(r,12, row[12], content)
                        trxz.write(r,13, row[13], content)
                        trxz.write(r,14, row[14], content)
                        trxz.write(r,15, row[15], content)
                        trxz.write(r,16, row[16], content)
                        trxz.write(r,17, row[17], content)
                        trxz.write(r,18, row[18], content)
                        trxz.write(r,19, row[19], content)
                        trxz.write(r,20, row[20], content)
                        trxz.write(r,21, row[21], content)
                        trxz.write(r,22, row[22], content)
                        trxz.write(r,23, row[23], content)
                        trxz.write(r,24, row[24], content)
                cursor.close()
                chat = 'CHANNEL has just been written for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing CHANNEL for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        def REGION():
            GetQuery = 'REGION for worksheet TRX on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('REGION.txt','r') as inserts:
                query = inserts.read()

            cursor.execute(query)
            result = cursor.fetchall()

            for line in Colm:
                    for kolom in range(11):
                        Alline = line+'{}'.format(*([kolom + 71] * 11))
                        trxz.write(Alline,'',neat)

            trxz.merge_range('A68:A70','REGION',bold)
            trxz.merge_range('B68:G68','PURCHASE AUTORENEWAL ON',bold)
            trxz.merge_range('B69:C69','ATTEMPT',bold)
            trxz.write('B70','FREE',bold)
            trxz.write('C70','PAID',bold)
            trxz.merge_range('D69:E69','SUCCESS',bold)
            trxz.write('D70','FREE',bold)
            trxz.write('E70','PAID',bold)
            trxz.merge_range('F69:F70','SR (%)',bold)
            trxz.merge_range('G69:G70','REVENUE',bold)
            trxz.merge_range('H68:M68','PURCHASE AUTORENEWAL OFF',bold)
            trxz.merge_range('H69:I69','ATTEMPT',bold)
            trxz.write('H70','FREE',bold)
            trxz.write('I70','PAID',bold)
            trxz.merge_range('J69:K69','SUCCESS',bold)
            trxz.write('J70','FREE',bold)
            trxz.write('K70','PAID',bold)
            trxz.merge_range('L69:L70','SR (%)',bold)
            trxz.merge_range('M69:M70','REVENUE',bold)
            trxz.merge_range('N68:S68','RENEWAL AUTORENEWAL ON',bold)
            trxz.merge_range('N69:O69','ATTEMPT',bold)
            trxz.write('N70','FREE',bold)
            trxz.write('O70','PAID',bold)
            trxz.merge_range('P69:Q69','SUCCESS',bold)
            trxz.write('P70','FREE',bold)
            trxz.write('Q70','PAID',bold)
            trxz.merge_range('R69:R70','SR (%)',bold)
            trxz.merge_range('S69:S70','REVENUE',bold)
            trxz.merge_range('T68:Y68','RENEWAL AUTORENEWAL OFF',bold)
            trxz.merge_range('T69:U69','ATTEMPT',bold)
            trxz.write('T70','FREE',bold)
            trxz.write('U70','PAID',bold)
            trxz.merge_range('V69:W69','SUCCESS',bold)
            trxz.write('V70','FREE',bold)
            trxz.write('W70','PAID',bold)
            trxz.merge_range('X69:X70','SR (%)',bold)
            trxz.merge_range('Y69:Y70','REVENUE',bold)
            if result:
                for r, row in enumerate(result, start=70):
                    for c, col in enumerate(row):
                        trxz.write(r,0, row[0], date_format)
                        trxz.write(r,1, row[1], content)
                        trxz.write(r,2, row[2], content)
                        trxz.write(r,3, row[3], content)
                        trxz.write(r,4, row[4], content)
                        trxz.write(r,5, row[5], content)
                        trxz.write(r,6, row[6], content)
                        trxz.write(r,7, row[7], content)
                        trxz.write(r,8, row[8], content)
                        trxz.write(r,9, row[9], content)
                        trxz.write(r,10, row[10], content)
                        trxz.write(r,11, row[11], content)
                        trxz.write(r,12, row[12], content)
                        trxz.write(r,13, row[13], content)
                        trxz.write(r,14, row[14], content)
                        trxz.write(r,15, row[15], content)
                        trxz.write(r,16, row[16], content)
                        trxz.write(r,17, row[17], content)
                        trxz.write(r,18, row[18], content)
                        trxz.write(r,19, row[19], content)
                        trxz.write(r,20, row[20], content)
                        trxz.write(r,21, row[21], content)
                        trxz.write(r,22, row[22], content)
                        trxz.write(r,23, row[23], content)
                        trxz.write(r,24, row[24], content)
                cursor.close()
                chat = 'REGION has just been written for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing REGION for worksheet TRX on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)

        PURCHASE_AUTO()
        PRICE()
        CHANNEL()
        REGION()

    def TRXHOURLY():
        Colm = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA']
        def HOURLY():
            GetQuery = 'HOURLY for worksheet TRXHOURLY on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
            print(GetQuery)
            fyi(GetQuery)
            cursor = database.cursor()
            with open('TRXHOURLY.txt','r') as inserts:
                query = inserts.read()

            cursor.execute(query)
            result = cursor.fetchall()

            for line in Colm:
                    for kolom in range(1):
                        Alline = line+'{}'.format(*([kolom + 4] * 1))
                        trxz_hourly.write(Alline,'',neat)

            trxz_hourly.set_zoom(70)
            trxz_hourly.hide_gridlines(2)
            trxz_hourly.merge_range('A1:A3','DATE',bold)
            trxz_hourly.merge_range('B1:B3','HOUR',bold)
            trxz_hourly.merge_range('C1:C3','CHANNEL',bold)
            trxz_hourly.merge_range('D1:I1','PURCHASE AUTORENEWAL ON',bold)
            trxz_hourly.merge_range('D2:E2','ATTEMPT',bold)
            trxz_hourly.write('D3','FREE',bold)
            trxz_hourly.write('E3','PAID',bold)
            trxz_hourly.merge_range('F2:G2','SUCCESS',bold)
            trxz_hourly.write('F3','FREE',bold)
            trxz_hourly.write('G3','PAID',bold)
            trxz_hourly.merge_range('H2:H3','SR (%)',bold)
            trxz_hourly.merge_range('I2:I3','REVENUE',bold)
            trxz_hourly.merge_range('J1:O1','PURCHASE AUTORENEWAL OFF',bold)
            trxz_hourly.merge_range('J2:K2','ATTEMPT',bold)
            trxz_hourly.write('J3','FREE',bold)
            trxz_hourly.write('K3','PAID',bold)
            trxz_hourly.merge_range('L2:M2','SUCCESS',bold)
            trxz_hourly.write('L3','FREE',bold)
            trxz_hourly.write('M3','PAID',bold)
            trxz_hourly.merge_range('N2:N3','SR (%)',bold)
            trxz_hourly.merge_range('O2:O3','REVENUE',bold)
            trxz_hourly.merge_range('P1:U1','RENEWAL AUTORENEWAL ON',bold)
            trxz_hourly.merge_range('P2:Q2','ATTEMPT',bold)
            trxz_hourly.write('P3','FREE',bold)
            trxz_hourly.write('Q3','PAID',bold)
            trxz_hourly.merge_range('R2:S2','SUCCESS',bold)
            trxz_hourly.write('R3','FREE',bold)
            trxz_hourly.write('S3','PAID',bold)
            trxz_hourly.merge_range('T2:T3','SR (%)',bold)
            trxz_hourly.merge_range('U2:U3','REVENUE',bold)
            trxz_hourly.merge_range('V1:AA1','RENEWAL AUTORENEWAL OFF',bold)
            trxz_hourly.merge_range('V2:W2','ATTEMPT',bold)
            trxz_hourly.write('V3','FREE',bold)
            trxz_hourly.write('W3','PAID',bold)
            trxz_hourly.merge_range('X2:Y2','SUCCESS',bold)
            trxz_hourly.write('X3','FREE',bold)
            trxz_hourly.write('Y3','PAID',bold)
            trxz_hourly.merge_range('Z2:Z3','SR (%)',bold)
            trxz_hourly.merge_range('AA2:AA3','REVENUE',bold)
            trxz_hourly.set_column('A:A', 13)
            trxz_hourly.set_column('B:B', 12)
            trxz_hourly.set_column('C:C', 17)
            trxz_hourly.set_column('D:E', 24)
            trxz_hourly.set_column('F:G', 24)
            trxz_hourly.set_column('H:I', 17)
            trxz_hourly.set_column('J:K', 24)
            trxz_hourly.set_column('L:M', 24)
            trxz_hourly.set_column('N:O', 17)
            trxz_hourly.set_column('P:Q', 24)
            trxz_hourly.set_column('R:S', 24)
            trxz_hourly.set_column('T:U', 17)
            trxz_hourly.set_column('V:W', 24)
            trxz_hourly.set_column('X:Y', 24)
            trxz_hourly.set_column('Z:AA', 17)
            if result:
                for r, row in enumerate(result, start=3):
                    for c, col in enumerate(row):
                        trxz_hourly.write(r,0, row[0], date_format)
                        trxz_hourly.write(r,1, row[1], content)
                        trxz_hourly.write(r,2, row[2], content)
                        trxz_hourly.write(r,3, row[3], content)
                        trxz_hourly.write(r,4, row[4], content)
                        trxz_hourly.write(r,5, row[5], content)
                        trxz_hourly.write(r,6, row[6], content)
                        trxz_hourly.write(r,7, row[7], content)
                        trxz_hourly.write(r,8, row[8], content)
                        trxz_hourly.write(r,9, row[9], content)
                        trxz_hourly.write(r,10, row[10], content)
                        trxz_hourly.write(r,11, row[11], content)
                        trxz_hourly.write(r,12, row[12], content)
                        trxz_hourly.write(r,13, row[13], content)
                        trxz_hourly.write(r,14, row[14], content)
                        trxz_hourly.write(r,15, row[15], content)
                        trxz_hourly.write(r,16, row[16], content)
                        trxz_hourly.write(r,17, row[17], content)
                        trxz_hourly.write(r,18, row[18], content)
                        trxz_hourly.write(r,19, row[19], content)
                        trxz_hourly.write(r,20, row[20], content)
                        trxz_hourly.write(r,21, row[21], content)
                        trxz_hourly.write(r,22, row[22], content)
                        trxz_hourly.write(r,23, row[23], content)
                        trxz_hourly.write(r,24, row[24], content)
                        trxz_hourly.write(r,25, row[25], content)
                        trxz_hourly.write(r,26, row[26], content)
                cursor.close()
                chat = 'HOURLY has just been written for worksheet TRXHOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                print(chat)
                fyi(chat)
            else:
                Alert = 'Skip writing HOURLY for worksheet TRXHOURLY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(Alert)
                fyi(Alert)
        HOURLY()

    def HISTORICALDATASUMMARY():
        GetQuery = 'TRX HISTORICAL for worksheet HISTORICAL DATA SUMMARY on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        fyi(GetQuery)
        def GETDAT():
            Dump = dir+'historicaldatasummary_'+str(kmrnbulan[0:6])+'.csv'
            Sub = dir+'subs_summary_new_'+str(kmrn)+'.csv'
            with open(''+Sub+'',encoding='ISO-8859-1') as csvfile:
                df = pd.read_csv(csvfile, sep=',', header=None, index_col=0)
                number_rows = len(df)
                if number_rows == 0:
                    append = 'Failed to append TRX HISTORICAL for worksheet HISTORICAL DATA SUMMARY on (New)RBTDailyReport '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(append)
                    fyi(append)
                else:
                    df.to_csv(r''+Dump+'', sep=',', header=False, mode=''+ AsNeeded +'')
                    os.remove(dir+'subs_summary_new_'+str(kmrn)+'.csv')

            historicaldatsum.set_zoom(70)
            historicaldatsum.hide_gridlines(2)
            historicaldatsum.merge_range('A1:A4','CALLING_DATE',bold)
            historicaldatsum.merge_range('B1:F1','SUBSCRIBER',bold)
            historicaldatsum.merge_range('B2:C2','NEW SUBSCRIBER',bold)
            historicaldatsum.merge_range('B3:B4','NEW_FREE_USER',bold)
            historicaldatsum.merge_range('C3:C4','NEW_PAID_USER',bold)
            historicaldatsum.merge_range('D2:D4','STAY_USER',bold)
            historicaldatsum.merge_range('E2:E4','ACTIVE_TO_SUSPEND_USER',bold)
            historicaldatsum.merge_range('F2:F4','SUSPEND_TO_ACTIVE_USER',bold)
            historicaldatsum.merge_range('G1:G4','TOTAL_ACTIVE_USER',bold)
            historicaldatsum.merge_range('H1:H4','DAILY_NET_ADD',bold)
            historicaldatsum.merge_range('I1:I4','TOTAL_ACTIVE_THIS_YEAR',bold)
            historicaldatsum.merge_range('J1:P1','CHURN USER',bold)
            historicaldatsum.merge_range('J2:J4','VOLUNTARY_CHURN',bold)
            historicaldatsum.merge_range('K2:K4','FULL_DELETE_CONTENT',bold)
            historicaldatsum.merge_range('L2:L4','ACTIVE_DELETE_CONTENT',bold)
            historicaldatsum.merge_range('M2:M4','CHURN_INSUFFICIENT_BALANCE',bold)
            historicaldatsum.merge_range('N2:N4','CHURN_ERROR_CHARGING',bold)
            historicaldatsum.merge_range('O2:O4','OTHERS',bold)
            historicaldatsum.merge_range('P2:P4','SUSPEND_IN_THREE_DAYS',bold)
            historicaldatsum.merge_range('Q1:Q4','TOTAL_CHURN_USER',bold)
            historicaldatsum.merge_range('R1:R4','ALL_MEMBERS_THIS_YEAR',bold)
            historicaldatsum.set_column('A:D', 18)
            historicaldatsum.set_column('E:F', 30)
            historicaldatsum.set_column('G:H', 23)
            historicaldatsum.set_column('I:I', 30)
            historicaldatsum.set_column('J:P', 35)
            historicaldatsum.set_column('Q:Q', 22)
            historicaldatsum.set_column('R:R', 30)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                if csvfile:
                    readCSV = csv.reader(csvfile, delimiter=',')
                    # sortedlist = sorted(readCSV, key=operator.itemgetter(0), reverse=True) # Sorter
                    sortedlist = sorted(readCSV, key=lambda row: row[0], reverse=True) # Sorter
                    for r, row in enumerate(sortedlist, start=4):
                        for c, col in enumerate(row):
                            historicaldatsum.write(r,0, row[0], date_format)
                            historicaldatsum.write(r,1, row[1], content)
                            historicaldatsum.write(r,2, row[2], content)
                            historicaldatsum.write(r,3, row[3], content)
                            historicaldatsum.write(r,4, row[4], content)
                            historicaldatsum.write(r,5, row[5], content)
                            historicaldatsum.write(r,6, row[6], content)
                            historicaldatsum.write(r,7, row[7], content)
                            historicaldatsum.write(r,8, row[8], content)
                            historicaldatsum.write(r,9, row[9], content)
                            historicaldatsum.write(r,10, row[10], content)
                            historicaldatsum.write(r,11, row[11], content)
                            historicaldatsum.write(r,12, row[12], content)
                            historicaldatsum.write(r,13, row[13], content)
                            historicaldatsum.write(r,14, row[14], content)
                            historicaldatsum.write(r,15, row[15], content)
                            historicaldatsum.write(r,16, row[16], content)
                            historicaldatsum.write(r,17, row[17], content)
                    chat = 'TRX HISTORICAL has just been written for worksheet HISTORICAL DATA SUMMARY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    fyi(chat)
                else:
                    Alert = 'Skip writing TRX HISTORICAL for worksheet HISTORICAL DATA SUMMARY on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    fyi(Alert)
        sys = '1'
        ystrdy = now - datetime.timedelta(days=1)
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrn = str(ystrdy.strftime('%Y%m%d'))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()
        
    def NSP_MEMBER_SINCE_2015():
        GetQuery = 'NSP MEMBER for worksheet NSP_MEMBER_SINCE_2015 on (New)RBTDailyReport is running queries at '+datetime.datetime.now().strftime("%H:%M:%S")
        print(GetQuery)
        # fyi(GetQuery)
        def GETDAT():
            Dump = dir+'member_since2k15_'+str(kmrnbulan[0:6])+'.csv'
            stdin, stdout, stderr = ssh.exec_command('cat /TRBT/rpt/daily/new_report/'+kmrnbulan+'/7_member_since2015_'+kmrnbulan+'.csv')
            df = pd.read_csv(stdout, sep=',', index_col=0)
            number_rows = len(df.index)
            if number_rows == 0:
                append = 'Failed to append NSP MEMBER for worksheet NSP_MEMBER_SINCE_2015 on (New)RBTDailyReport '+ kmrnbulan +' at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                print(append)
                # fyi(append)
            else:
                df.to_csv(r''+Dump+'', sep=',', header=False, mode=''+ AsNeeded +'')
            
            nspmembersince.set_zoom(85)
            nspmembersince.hide_gridlines(2)
            nspmembersince.merge_range('A1:C1','*starting point 1st January 2015',title)
            nspmembersince.write('A2','Date',bold)
            nspmembersince.write('B2','Total_NSP_Member',bold)
            nspmembersince.write('C2','Member_Increment',bold)
            nspmembersince.set_column('A:C', 18)
            for kolom in range(70):
                SRP = '{}'.format(*([kolom + 4] * 70))
                succratemt.write_formula('C'+SRP,'=E'+SRP+'/D'+SRP+'',title)
            with open(''+Dump+'',encoding='ISO-8859-1') as csvfile:
                if csvfile:
                    readCSV = csv.reader(csvfile, delimiter=',')
                    # sortedlist = sorted(readCSV, key=operator.itemgetter(0), reverse=True) # Sorter
                    sortedlist = sorted(readCSV, key=lambda row: row[0], reverse=True) # Sorter
                    for r, row in enumerate(sortedlist, start=2):
                        for c, col in enumerate(row):
                            nspmembersince.write(r,0, row[0], date_format)
                            nspmembersince.write(r,1, row[1], content)
                            nspmembersince.write(r,2, row[2], content)
                    chat = 'NSP MEMBER has just been written for worksheet NSP_MEMBER_SINCE_2015 on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")
                    print(chat)
                    # fyi(chat)
                else:
                    Alert = 'Skip writing NSP MEMBER for worksheet NSP_MEMBER_SINCE_2015 on (New)RBTDailyReport xlsx at '+datetime.datetime.now().strftime("%H:%M:%S")+', because the data is empty from the database'
                    print(Alert)
                    # fyi(Alert)
        sys = '1'
        ystrdy = now - datetime.timedelta(days=1)
        lusa = now - datetime.timedelta(days=2)
        storyday = now - datetime.timedelta(days=int(sys))
        kmrn = str(ystrdy.strftime('%Y%m%d'))
        kmrnlusa = str(lusa.strftime('%Y%m%d'))
        kmrnbulan = str(storyday.strftime('%Y%m%d'))
        if kmrnbulan[0:6] == kmrnlusa[0:6]:
            AsNeeded = 'a'
            GETDAT()
        else:
            AsNeeded = 'w'
            GETDAT()

    # SUB()
    # SUB_HOURLY()
    # UMBS()
    # TRX()
    # TRXHOURLY()
    # HISTORICALDATASUMMARY()
    # NSP_MEMBER_SINCE_2015()
    workbook.close()


# In[ ]:


RBTDailyReport()
# NewRBTDailyReport()
database.close()
SendEmail()

