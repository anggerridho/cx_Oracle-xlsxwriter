#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import cx_Oracle
import os
import xlsxwriter
import datetime

now = datetime.datetime.now()
kemarin = now - datetime.timedelta(1)
kemarin_str = str(kemarin.strftime('%Y%m%d'))
dir = '/home/palugada/Daily/'

os.environ["ORACLE_HOME"] = "/usr/lib/oracle/12.2/client64"
# workbook = xlsxwriter.Workbook('/home/palugada/Daily/Demo.xlsx')
workbook = xlsxwriter.Workbook(dir+"RBT_Daily_Report_"+str(kemarin_str)+".xlsx")
database = cx_Oracle.connect('RBTRPTN/RBTRPTN@192.168.0.7:1521/RBTRPTN')

bold = workbook.add_format({'bold': True, 'bg_color': '#92D050', 'align': 'center', 'valign': 'vcenter', 'border': 1})
bold.set_border()
bold.set_text_wrap()
neat = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
neat.set_border()
neat.set_text_wrap()
percentage = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1})
percentage.set_border()
percentage.set_num_format(10)
percentage.set_text_wrap()
title = workbook.add_format({'bold': True})

# Create worksheet summary
summary = workbook.add_worksheet('SUMMARY')
new_subs = workbook.add_worksheet('NEW_SUBS')
prov = workbook.add_worksheet('PROV')


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
    subject = "(Testing) RBT Daily Report " + Yesterday
    body = "Dear All,\nPlease kindly find RBT Daily Reports " + Yesterday + " in attachment.\n\nBest Regards,\nAngger Ridho (https://linktr.ee/anggerdho)"
    sender_email = "Angger Ridho <angger@eluon.com>"
    receiver_email = ['angger@eluon.com',                       'operation.support@eluon.com']
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


# In[ ]:


# A All KPI
def A_KPI():
    cursor = database.cursor()
    query = ("SELECT calling_date, sms, umb, ivr, http, wap, web, openapi, mobileapps, wcc, csweb, renewal, rbt_req, chg_req, refund_req,     chg_trans_succ, chg_timeout, decode(chg_req, 0, 0,(((chg_req-chg_timeout)/ chg_req)) * 100) AS chgSuccessRate, chg_succ,     decode(chg_req, 0, 0,(chg_succ / chg_req)* 100) AS billableRate, chg_insf, chg_subs_locked, chg_other_err, refund_succ,     decode(refund_req, 0, 0,(refund_succ / refund_req)* 100) AS refundRate, prov_req, deprov_req, prov_succ, prov_error,     prov_timeout, decode(prov_req, 0, 0,((prov_req-prov_timeout)/ prov_req)* 100) AS provTransactionRate, decode(     prov_req, 0, 0,(prov_succ / prov_req)* 100) AS provRate, deprov_succ, deprov_err, decode(deprov_req, 0, 0,(     deprov_succ / deprov_req)* 100) AS deactiveRate, revenue, avg_svc_latency FROM perf_dashboard WHERE     CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD') ORDER BY calling_date DESC")
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
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


# In[ ]:


# B Transaction
def B_Transaction():
    cursor = database.cursor()
    query = ("SELECT calling_date, sum(CASE WHEN calling_type IN (1) THEN trans_count ELSE 0 END) PA_ATTEMPT,     sum(CASE WHEN calling_type IN (1) AND RESULT IN (0, 805) THEN trans_count ELSE 0 END) PA_SUCCESS,     sum(CASE WHEN calling_type IN (1) AND RESULT IN (0, 805) THEN revenue ELSE 0 END) PA_REVENUE,     sum(CASE WHEN calling_type IN (2, 7, 77) THEN trans_count ELSE 0 END) PM_ATTEMPT,     sum(CASE WHEN calling_type IN (2, 7, 77) AND RESULT IN (0, 805) THEN trans_count ELSE 0 END) PM_SUCCESS,     sum(CASE WHEN calling_type IN (2, 7, 77) AND RESULT IN (0, 805) THEN revenue ELSE 0 END) PM_REVENUE,     sum(CASE WHEN calling_type IN (64, 93, 95, 96, 99, 100) THEN trans_count ELSE 0 END) RA_ATTEMPT,     sum(CASE WHEN calling_type IN (64, 93, 95, 96, 99, 100) AND RESULT IN (0, 805) THEN trans_count ELSE 0 END) RA_SUCCESS,     sum(CASE WHEN calling_type IN (64, 93, 95, 96, 99, 100) AND RESULT IN (0, 805) THEN revenue ELSE 0 END) RA_REVENUE,     sum(CASE WHEN calling_type IN (78, 65) THEN trans_count ELSE 0 END) RM_ATTEMPT ,     sum(CASE WHEN calling_type IN (78, 65) AND RESULT IN (0, 805) THEN trans_count ELSE 0 END) RM_SUCCESS,     sum(CASE WHEN calling_type IN (78, 65) AND RESULT IN (0, 805) THEN revenue ELSE 0 END) RM_REVENUE,     sum(CASE WHEN calling_type IN (1, 2, 7, 77, 78, 64, 93, 65, 95, 96, 99, 100) THEN trans_count ELSE 0 END) TOTAL_ATTEMPT,     sum(CASE WHEN calling_type IN (1, 2, 7, 77, 78, 64, 93, 65, 95, 96, 99, 100) AND RESULT IN (0, 805) THEN trans_count ELSE 0 END) TOTAL_SUCCESS,     sum(CASE WHEN calling_type IN (1, 2, 7, 77, 78, 64, 93, 65, 95, 96, 99, 100) AND RESULT IN (0, 805) THEN revenue ELSE 0 END) TOTAL_REVENUE     FROM cdr_report_goofree WHERE calling_type IN (1, 2, 7, 77, 78, 64, 93, 65, 95, 96, 99, 100) AND calling_date = TO_CHAR(SYSDATE-1, 'YYYYMMDD')     GROUP BY calling_date ORDER BY calling_date")
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
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


# In[ ]:


# C_Subscriber
def C_Subscriber():
    cursor = database.cursor()
    query = ("SELECT a.calling_date, a. total TOTAL_SUBSCRIBER, b.t_req_act TOT_REQ_ACTIVATION, b.t_req_act_succ TOT_SUCC_ACTIVATION,     b.t_req_deact TOT_REQ_DEACTIVATION, b.t_req_deact_succ TOT_SUCC_DEACTIVATION, b.s_a SUSPEND_TO_ACTIVE, b.a_s ACTIVE_TO_SUSPEND FROM subs_summary_region a, (     SELECT calling_date, sum(case when calling_type = 17 then 1 else 0 end) t_req_act, sum(case when calling_type = 17 and result in (0,805) then 1 else 0 end) t_req_act_succ,     sum(case when calling_type = 72 then 1 else 0 end) t_req_deact, sum(case when calling_type = 72 and result in (0,805) then 1 else 0 end) t_req_deact_succ,     sum(case when prev_subs_status = 'S' and next_subs_status ='A'  then 1 else 0 end) s_a, sum(case when prev_subs_status = 'A' and next_subs_status ='S'  then 1 else 0 end) a_s     FROM call_detailed_record WHERE CALLING_DATE = TO_CHAR(SYSDATE-2,'YYYYMMDD') group by calling_date) b WHERE a.calling_date = b.calling_date")
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            summary.write(r,8, row[7], content)
            
    cursor.close()


# In[ ]:


# C_Subscriber Free Paid
def C_FreePaid():
    cursor = database.cursor()
    query = ("SELECT sum(case when b.initial_amount =0 then 1 else 0 end), sum(case when b.initial_amount !=0 then 1 else 0 end)     from user_monthly_rating a, V_SUBS_STATUS b where a.content_id=b.content_id and a.rating_id=b.rating_id and a.play_order = '000'     and a.effective_dt_to like '9999%'")
    
    # Exec query summary
    cursor.execute(query)
    result = cursor.fetchall()
    
    for r, row in enumerate(result, start=18):
        for c, col in enumerate(row):
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            summary.write(r,9, row[0], content)
            summary.write(r,10, row[1], content)
            
    cursor.close()


# In[ ]:


# C_SUBS_WITHDRAWAL
def C_SUBS_WITHDRAWAL():
    cursor = database.cursor()
    query = ("SELECT calling_date, sum(case when calling_type = 18 and result = 0 then 1 else 0 end) SUBS_WITHDRAWAL,     sum(case when calling_type = 4 and result = 0 then 1 else 0 end) DELETE_CONTENT, sum(case when calling_type = 67     and result = 0 then 1 else 0 end) RENEWAL_CONTENT_DELETION, sum(case when calling_type in (64,93,95,96,99,100) and     result = 666 then 1 else 0 end) RENEWAL_CONTENT_INSUF_BALANCE, sum(case when calling_type in (64,93,95,96,99,100)     and result = 502 then 1 else 0 end) RENEWAL_CONTENT_ERROR_CHARGING, count(*) GRAND_TOTAL FROM call_detailed_record     WHERE prev_subs_status = 'A' AND next_subs_status ='S' AND CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD') AND     calling_type in (4,18,67,64,93,95,96,99,100) AND result in (0,502,666) GROUP BY calling_date")
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
            summary.write(r,1, row[0], date_format)
            summary.write(r,2, row[1], content)
            summary.write(r,3, row[2], content)
            summary.write(r,4, row[3], content)
            summary.write(r,5, row[4], content)
            summary.write(r,6, row[5], content)
            summary.write(r,7, row[6], content)
            
    cursor.close()


# In[ ]:


# C_CHURNED_USER
def C_CHURNED_USER():
    cursor = database.cursor()
    query = ("select calling_date, 'Subs Withdrawal' Churned_User, sum(case when region='Sumbagut' then 1 else 0 end)  Sumbagut,     sum(case when region='Sumbagteng' then 1 else 0 end) Sumbagteng, sum(case when region='Sumbagsel' then 1 else 0 end)  Sumbagsel,     sum(case when region='Jabotabek' then 1 else 0 end) Jabotabek, sum(case when region='Jabar' then 1 else 0 end) Jabar,     sum(case when region='Jateng' then 1 else 0 end) Jateng, sum(case when region='Jatim' then 1 else 0 end) Jatim,     sum(case when region='Balinusra' then 1 else 0 end) Balinusra, sum(case when region='Kalimantan' then 1 else 0 end) Kalimantan,     sum(case when region='Sulawesi' then 1 else 0 end)   Sulawesi, sum(case when region='Papua' then 1 else 0 end)  Papua,     sum(case when region='UNKNOWN' then 1 else 0 end)  Other, count(*)  total from CALL_DETAILED_RECORD a, prefix_region b     WHERE prev_subs_status = 'A' AND next_subs_status ='S' and substr(a.calling_mdn,1,7) = b.prefix AND CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD')     AND calling_type=18 AND result=0 GROUP BY calling_date union all select calling_date,'Delete Content' Churned_User, sum(case when region='Sumbagut' then 1 else 0 end)  Sumbagut,     sum(case when region='Sumbagteng' then 1 else 0 end) Sumbagteng, sum(case when region='Sumbagsel' then 1 else 0 end)  Sumbagsel,     sum(case when region='Jabotabek' then 1 else 0 end) Jabotabek, sum(case when region='Jabar' then 1 else 0 end) Jabar,     sum(case when region='Jateng' then 1 else 0 end) Jateng, sum(case when region='Jatim' then 1 else 0 end) Jatim,     sum(case when region='Balinusra' then 1 else 0 end) Balinusra, sum(case when region='Kalimantan' then 1 else 0 end) Kalimantan,     sum(case when region='Sulawesi' then 1 else 0 end)   Sulawesi, sum(case when region='Papua' then 1 else 0 end)  Papua,     sum(case when region='UNKNOWN' then 1 else 0 end)  Other, count(*)  total from CALL_DETAILED_RECORD a, prefix_region b     WHERE prev_subs_status = 'A' AND next_subs_status ='S' and substr(a.calling_mdn,1,7) = b.prefix AND CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD')     AND calling_type=4 AND result=0 GROUP BY calling_date union all select calling_date,'Renewal Content Deletion' Churned_User,     sum(case when region='Sumbagut' then 1 else 0 end)  Sumbagut, sum(case when region='Sumbagteng' then 1 else 0 end) Sumbagteng,     sum(case when region='Sumbagsel' then 1 else 0 end)  Sumbagsel, sum(case when region='Jabotabek' then 1 else 0 end) Jabotabek,     sum(case when region='Jabar' then 1 else 0 end) Jabar, sum(case when region='Jateng' then 1 else 0 end) Jateng,     sum(case when region='Jatim' then 1 else 0 end) Jatim, sum(case when region='Balinusra' then 1 else 0 end) Balinusra,     sum(case when region='Kalimantan' then 1 else 0 end) Kalimantan, sum(case when region='Sulawesi' then 1 else 0 end)   Sulawesi,     sum(case when region='Papua' then 1 else 0 end)  Papua, sum(case when region='UNKNOWN' then 1 else 0 end)  Other,     count(*)  total from CALL_DETAILED_RECORD a, prefix_region b WHERE prev_subs_status = 'A' AND next_subs_status ='S'     and substr(a.calling_mdn,1,7) = b.prefix AND CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD') AND calling_type=67     AND result=0 GROUP BY calling_date union all select calling_date,'Renewal Content Insuf Balance' Churned_User,     sum(case when region='Sumbagut' then 1 else 0 end)  Sumbagut, sum(case when region='Sumbagteng' then 1 else 0 end) Sumbagteng,     sum(case when region='Sumbagsel' then 1 else 0 end)  Sumbagsel, sum(case when region='Jabotabek' then 1 else 0 end) Jabotabek,     sum(case when region='Jabar' then 1 else 0 end) Jabar, sum(case when region='Jateng' then 1 else 0 end) Jateng,     sum(case when region='Jatim' then 1 else 0 end) Jatim, sum(case when region='Balinusra' then 1 else 0 end) Balinusra,     sum(case when region='Kalimantan' then 1 else 0 end) Kalimantan, sum(case when region='Sulawesi' then 1 else 0 end)   Sulawesi,     sum(case when region='Papua' then 1 else 0 end)  Papua, sum(case when region='UNKNOWN' then 1 else 0 end)  Other,     count(*)  total from CALL_DETAILED_RECORD a, prefix_region b WHERE prev_subs_status = 'A' AND next_subs_status ='S'     and substr(a.calling_mdn,1,7) = b.prefix AND CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD') AND calling_type in (64,93,95,96,99,100)     AND result=666 GROUP BY calling_date union all select calling_date,'Renewal Content Error Charging' Churned_User,     sum(case when region='Sumbagut' then 1 else 0 end)  Sumbagut, sum(case when region='Sumbagteng' then 1 else 0 end) Sumbagteng,     sum(case when region='Sumbagsel' then 1 else 0 end)  Sumbagsel, sum(case when region='Jabotabek' then 1 else 0 end) Jabotabek,     sum(case when region='Jabar' then 1 else 0 end) Jabar, sum(case when region='Jateng' then 1 else 0 end) Jateng,     sum(case when region='Jatim' then 1 else 0 end) Jatim, sum(case when region='Balinusra' then 1 else 0 end) Balinusra,     sum(case when region='Kalimantan' then 1 else 0 end) Kalimantan, sum(case when region='Sulawesi' then 1 else 0 end)   Sulawesi,     sum(case when region='Papua' then 1 else 0 end)  Papua, sum(case when region='UNKNOWN' then 1 else 0 end)  Other,     count(*)  total from CALL_DETAILED_RECORD a, prefix_region b WHERE prev_subs_status = 'A' AND next_subs_status ='S'     and substr(a.calling_mdn,1,7) = b.prefix AND CALLING_DATE = TO_CHAR(SYSDATE-1,'YYYYMMDD') AND calling_type in (64,93,95,96,99,100)     AND result=502 GROUP BY calling_date")
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
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


# In[ ]:


# A. NEW SUBSCRIPTION TOTAL
def A_NEW_SUBSCRIPTION_TOTAL():
    cursor = database.cursor()
    query = ("select calling_date, sum(trans_count), sum(case when result in ('0','805','661','666') then trans_count else 0 end),     round(100*sum(case when result in ('0','805','661','666') or (result = 502 and charge_result_cde = '3:27') then trans_count else 0 end)/(sum(trans_count)),2) || '%' ,     sum(case when result in ('0','805') then trans_count else 0 end), round(100*sum(case when result in ('0','805') or (result = 502 and charge_result_cde = '3:27') then trans_count else 0 end)/(sum(trans_count)),2) || '%' ,     sum(case when result = '0' then trans_count else 0 end), sum(case when result = '805' then trans_count else 0 end),     sum(case when result = '661' then trans_count else 0 end), sum(case when result = '666' then trans_count else 0 end)     from cdr_report_goofree where calling_type in (1,2,7,64,93,78,77,65,95,96,99,100)     and calling_date = to_char(sysdate-1,'YYYYMMDD') group by calling_date order by calling_date")
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
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


# In[ ]:


# B. NEW SUBSCRIPTION CHANNEL
def B_NEW_SUBSCRIPTION_CHANNEL():    
    
    # A_SUCCESS_RATE Purchasing
    def A_SUCCESS_RATE():
        cursor = database.cursor()
        query = ("SELECT CALLING_DATE,NOTE,SMS, UMB,HTTP,DTMF,OPENAPI,WEB,IVR,WAP,CSWEB,WCC,OTHERS FROM (select calling_date,         'ATTEMPT (A)' note, sum(case when substr(media_code,2,1)='0' then trans_count else 0 end) SMS, sum(case when substr(media_code,2,1)='U' then trans_count else 0 end) UMB,         sum(case when substr(media_code,2,1)='S' then trans_count else 0 end) HTTP, sum(case when substr(media_code,2,1)='D' then trans_count else 0 end) DTMF,         sum(case when substr(media_code,2,1)='O' then trans_count else 0 end) OPENAPI, sum(case when substr(media_code,2,1)='B' then trans_count else 0 end) WEB,         sum(case when substr(media_code,2,1)='I' then trans_count else 0 end) IVR, sum(case when substr(media_code,2,1)='P' then trans_count else 0 end) WAP,         sum(case when substr(media_code,2,1)='C' then trans_count else 0 end) CSWEB, sum(case when substr(media_code,2,1)='W' then trans_count else 0 end) WCC,         sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (1) group by calling_date union select calling_date, 'SUCCESS (A)' NOTE,         sum(case when substr(media_code,2,1)='0' and result in (0,805) then trans_count else 0 end) SMS, sum(case when substr(media_code,2,1)='U' and result in (0,805) then trans_count else 0 end) UMB,         sum(case when substr(media_code,2,1)='S' and result in (0,805) then trans_count else 0 end) HTTP, sum(case when substr(media_code,2,1)='D' and result in (0,805) then trans_count else 0 end) DTMF,         sum(case when substr(media_code,2,1)='O' and result in (0,805) then trans_count else 0 end) OPENAPI, sum(case when substr(media_code,2,1)='B' and result in (0,805) then trans_count else 0 end) WEB,         sum(case when substr(media_code,2,1)='I' and result in (0,805)then trans_count else 0 end) IVR, sum(case when substr(media_code,2,1)='P' and result in (0,805) then trans_count else 0 end) WAP,         sum(case when substr(media_code,2,1)='C' and result in (0,805) then trans_count else 0 end) CSWEB, sum(case when substr(media_code,2,1)='W' and result in (0,805) then trans_count else 0 end) WCC,         sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') and result in (0,805) then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (1) group by calling_date) where  calling_date = to_char(sysdate-1,'YYYYMMDD') order by calling_date")

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
                content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                content.set_border()
                date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                date_format.set_border()
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
        query = ("SELECT CALLING_DATE,NOTE,SMS, UMB,HTTP,DTMF,OPENAPI,WEB,IVR,WAP,CSWEB,WCC,OTHERS FROM (select calling_date,'ATTEMPT (B)' note,         sum(case when substr(media_code,2,1)='0' then trans_count else 0 end) SMS, sum(case when substr(media_code,2,1)='U' then trans_count else 0 end) UMB,         sum(case when substr(media_code,2,1)='S' then trans_count else 0 end) HTTP, sum(case when substr(media_code,2,1)='D' then trans_count else 0 end) DTMF,         sum(case when substr(media_code,2,1)='O' then trans_count else 0 end) OPENAPI, sum(case when substr(media_code,2,1)='B' then trans_count else 0 end) WEB,         sum(case when substr(media_code,2,1)='I' then trans_count else 0 end) IVR, sum(case when substr(media_code,2,1)='P' then trans_count else 0 end) WAP,         sum(case when substr(media_code,2,1)='C' then trans_count else 0 end) CSWEB, sum(case when substr(media_code,2,1)='W' then trans_count else 0 end) WCC,         sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (2,7,77) group by calling_date union select calling_date, 'SUCCESS (B)' NOTE,         sum(case when substr(media_code,2,1)='0' and result in (0,805) then trans_count else 0 end) SMS, sum(case when substr(media_code,2,1)='U' and result in (0,805) then trans_count else 0 end) UMB,         sum(case when substr(media_code,2,1)='S' and result in (0,805) then trans_count else 0 end) HTTP, sum(case when substr(media_code,2,1)='D' and result in (0,805) then trans_count else 0 end) DTMF,         sum(case when substr(media_code,2,1)='O' and result in (0,805) then trans_count else 0 end) OPENAPI, sum(case when substr(media_code,2,1)='B' and result in (0,805) then trans_count else 0 end) WEB,         sum(case when substr(media_code,2,1)='I' and result in (0,805)then trans_count else 0 end) IVR, sum(case when substr(media_code,2,1)='P' and result in (0,805) then trans_count else 0 end) WAP,         sum(case when substr(media_code,2,1)='C' and result in (0,805) then trans_count else 0 end) CSWEB, sum(case when substr(media_code,2,1)='W' and result in (0,805) then trans_count else 0 end) WCC,         sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') and result in (0,805) then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (2,7,77) group by calling_date) where  calling_date = to_char(sysdate-1,'YYYYMMDD') order by calling_date")

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
                content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                content.set_border()
                date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                date_format.set_border()
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
        query = ("SELECT CALLING_DATE,NOTE,RENEWAL,SMS, UMB,HTTP,DTMF,OPENAPI,WEB,IVR,WAP,CSWEB,WCC,OTHERS FROM (select calling_date,         'ATTEMPT (A)' note, sum(case when substr(media_code,2,1)='N' then trans_count else 0 end) RENEWAL, sum(case when substr(media_code,2,1)='0' then trans_count else 0 end) SMS,         sum(case when substr(media_code,2,1)='U' then trans_count else 0 end) UMB, sum(case when substr(media_code,2,1)='S' then trans_count else 0 end) HTTP,         sum(case when substr(media_code,2,1)='D' then trans_count else 0 end) DTMF, sum(case when substr(media_code,2,1)='O' then trans_count else 0 end) OPENAPI,         sum(case when substr(media_code,2,1)='B' then trans_count else 0 end) WEB, sum(case when substr(media_code,2,1)='I' then trans_count else 0 end) IVR,         sum(case when substr(media_code,2,1)='P' then trans_count else 0 end) WAP, sum(case when substr(media_code,2,1)='C' then trans_count else 0 end) CSWEB,         sum(case when substr(media_code,2,1)='W' then trans_count else 0 end) WCC, sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (64,93,95,96,99,100) group by calling_date union select calling_date, 'SUCCESS (A)' NOTE,         sum(case when substr(media_code,2,1)='N' and result in (0,805) then trans_count else 0 end) RENEWAL, sum(case when substr(media_code,2,1)='0' and result in (0,805) then trans_count else 0 end) SMS,         sum(case when substr(media_code,2,1)='U' and result in (0,805) then trans_count else 0 end) UMB, sum(case when substr(media_code,2,1)='S' and result in (0,805) then trans_count else 0 end) HTTP,         sum(case when substr(media_code,2,1)='D' and result in (0,805) then trans_count else 0 end) DTMF, sum(case when substr(media_code,2,1)='O' and result in (0,805) then trans_count else 0 end) OPENAPI,         sum(case when substr(media_code,2,1)='B' and result in (0,805) then trans_count else 0 end) WEB, sum(case when substr(media_code,2,1)='I' and result in (0,805)then trans_count else 0 end) IVR,         sum(case when substr(media_code,2,1)='P' and result in (0,805) then trans_count else 0 end) WAP, sum(case when substr(media_code,2,1)='C' and result in (0,805) then trans_count else 0 end) CSWEB,         sum(case when substr(media_code,2,1)='W' and result in (0,805) then trans_count else 0 end) WCC, sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') and result in (0,805) then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (64,93,95,96,99,100) group by calling_date) where  calling_date = to_char(sysdate-1,'YYYYMMDD') order by calling_date")

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
                content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                content.set_border()
                date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                date_format.set_border()
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
        query = ("SELECT CALLING_DATE,NOTE,RENEWAL,SMS,UMB,HTTP,DTMF,OPENAPI,WEB,IVR,WAP,CSWEB,WCC,OTHERS FROM (select calling_date,         'ATTEMPT (B)' note, sum(case when substr(media_code,2,1)='N' then trans_count else 0 end) RENEWAL, sum(case when substr(media_code,2,1)='0' then trans_count else 0 end) SMS,         sum(case when substr(media_code,2,1)='U' then trans_count else 0 end) UMB, sum(case when substr(media_code,2,1)='S' then trans_count else 0 end) HTTP,         sum(case when substr(media_code,2,1)='D' then trans_count else 0 end) DTMF, sum(case when substr(media_code,2,1)='O' then trans_count else 0 end) OPENAPI,         sum(case when substr(media_code,2,1)='B' then trans_count else 0 end) WEB, sum(case when substr(media_code,2,1)='I' then trans_count else 0 end) IVR,         sum(case when substr(media_code,2,1)='P' then trans_count else 0 end) WAP, sum(case when substr(media_code,2,1)='C' then trans_count else 0 end) CSWEB,         sum(case when substr(media_code,2,1)='W' then trans_count else 0 end) WCC, sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') then trans_count else 0 end) OTHERS         from cdr_report_goofree where calling_type in (78,65) group by calling_date union select calling_date, 'SUCCESS (B)' NOTE, sum(case when substr(media_code,2,1)='N' and result in (0,805) then trans_count else 0 end) RENEWAL,         sum(case when substr(media_code,2,1)='0' and result in (0,805) then trans_count else 0 end) SMS, sum(case when substr(media_code,2,1)='U' and result in (0,805) then trans_count else 0 end) UMB,         sum(case when substr(media_code,2,1)='S' and result in (0,805) then trans_count else 0 end) HTTP, sum(case when substr(media_code,2,1)='D' and result in (0,805) then trans_count else 0 end) DTMF,         sum(case when substr(media_code,2,1)='O' and result in (0,805) then trans_count else 0 end) OPENAPI, sum(case when substr(media_code,2,1)='B' and result in (0,805) then trans_count else 0 end) WEB,         sum(case when substr(media_code,2,1)='I' and result in (0,805)then trans_count else 0 end) IVR, sum(case when substr(media_code,2,1)='P' and result in (0,805) then trans_count else 0 end) WAP,         sum(case when substr(media_code,2,1)='C' and result in (0,805) then trans_count else 0 end) CSWEB, sum(case when substr(media_code,2,1)='W' and result in (0,805) then trans_count else 0 end) WCC,         sum(case when substr(media_code,2,1) not in('0','U','S','D','B','P','I','C','W','N','O') and result in (0,805) then trans_count else 0 end) OTHERS from cdr_report_goofree where calling_type in (78,65)         group by calling_date) where calling_date = to_char(sysdate-1,'YYYYMMDD') order by calling_date")

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
                content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                content.set_border()
                date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
                date_format.set_border()
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


# In[ ]:


# C. NEW SUBSCRIPTION HOUR
def C_NEW_SUBSCRIPTION_HOUR():
    cursor = database.cursor()
    query = ("SELECT calling_date,substr(calling_time, 1, 2), count(*), sum(CASE WHEN RESULT IN ('0', '805', '661', '666') OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END),     round(100 * sum(CASE WHEN RESULT IN ('0', '805', '661', '666') OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END)/ count(*), 2) || '%',     sum(CASE WHEN RESULT IN ('0', '805') OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END), round(100 * sum(CASE WHEN RESULT IN ('0', '805') OR (RESULT = 502 AND charge_result_cde = '3:27')     THEN 1 ELSE 0 END)/ count(*), 2) || '%' FROM call_detailed_record cdr WHERE	calling_type IN (1, 2, 7, 64, 93, 78, 77, 65, 95, 96, 99, 100) AND calling_date = to_char(sysdate-1, 'YYYYMMDD')     GROUP BY calling_date, substr(calling_time, 1, 2) ORDER BY calling_date, substr(calling_time, 1, 2)")
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
            new_subs.write(r,1, row[0], date_format)
            new_subs.write(r,2, row[1], content)
            new_subs.write(r,3, row[2], content)
            new_subs.write(r,4, row[3], content)
            new_subs.write(r,5, row[4], content)
            new_subs.write(r,6, row[5], content)
            new_subs.write(r,7, row[6], content)
            
    cursor.close()


# In[ ]:


# D. NEW SUBSCRIPTION HOUR PER CHANNEL
def D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL():
    cursor = database.cursor()
    query = ("SELECT calling_date,substr(calling_time, 1, 2), CASE WHEN substr(media_code, 2, 1) = 'U' THEN 'UMB'     WHEN substr(media_code, 2, 1) = 'I' THEN 'IVR' WHEN substr(media_code, 2, 1) = 'O' THEN 'OpenAPI'     WHEN substr(media_code, 2, 1) = '0' THEN 'SMS' WHEN substr(media_code, 2, 1) = 'B' THEN 'Web'     WHEN substr(media_code, 2, 1) = 'D' THEN 'DTMF' WHEN substr(media_code, 2, 1) = 'S' THEN 'HTTPGW'     WHEN substr(media_code, 2, 1) = 'C' THEN 'CSWeb' WHEN substr(media_code, 2, 1) = 'W'     AND lower(DVC_ID) != 'smsbroadcast' THEN 'WCC' WHEN substr(media_code, 2, 1) = 'W'     AND lower(DVC_ID) = 'smsbroadcast' THEN 'Sms Broadcast' WHEN substr(media_code, 2, 1) = 'M' THEN 'MKios'     WHEN substr(media_code, 2, 1) = 'Z' THEN 'Point' WHEN substr(media_code, 2, 1) = 'N' THEN 'Renewal'     WHEN substr(media_code, 2, 1) = 'J' THEN 'Recomendation' WHEN substr(media_code, 2, 1) = 'L' THEN 'Referal'     WHEN substr(media_code, 2, 1) = 'P' THEN 'WAP' WHEN substr(media_code, 2, 1) = 'R' THEN 'MyNSP'     WHEN substr(media_code, 2, 1) = 'V' THEN 'My Telkomsel' WHEN substr(media_code, 2, 1) = 'Q' THEN 'Music Max'     ELSE 'Promo' END, count(*), sum(CASE WHEN RESULT IN ('0', '805', '661', '666') OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END),     round(100 * sum(CASE WHEN RESULT IN ('0', '805', '661', '666') OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END)/ count(*), 2) || '%' ,     sum(CASE WHEN RESULT IN (0, 805) OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END),     round(100 * sum(CASE WHEN RESULT IN (0, 805) OR (RESULT = 502 AND charge_result_cde = '3:27') THEN 1 ELSE 0 END)/ count(*), 2) || '%'     FROM call_detailed_record WHERE calling_date = to_char(sysdate-1, 'yyyymmdd') AND calling_type IN (1, 2, 7, 64, 93, 77, 78, 65, 95, 96, 99, 100)     GROUP BY calling_date, substr(calling_time, 1, 2), CASE WHEN substr(media_code, 2, 1) = 'U' THEN 'UMB'     WHEN substr(media_code, 2, 1) = 'I' THEN 'IVR' WHEN substr(media_code, 2, 1) = 'O' THEN 'OpenAPI'     WHEN substr(media_code, 2, 1) = '0' THEN 'SMS' WHEN substr(media_code, 2, 1) = 'B' THEN 'Web'     WHEN substr(media_code, 2, 1) = 'D' THEN 'DTMF' WHEN substr(media_code, 2, 1) = 'S' THEN 'HTTPGW'     WHEN substr(media_code, 2, 1) = 'C' THEN 'CSWeb' WHEN substr(media_code, 2, 1) = 'W'     AND lower(DVC_ID) != 'smsbroadcast' THEN 'WCC' WHEN substr(media_code, 2, 1) = 'W'     AND lower(DVC_ID) = 'smsbroadcast' THEN 'Sms Broadcast' WHEN substr(media_code, 2, 1) = 'M' THEN 'MKios'     WHEN substr(media_code, 2, 1) = 'Z' THEN 'Point' WHEN substr(media_code, 2, 1) = 'N' THEN 'Renewal'     WHEN substr(media_code, 2, 1) = 'J' THEN 'Recomendation' WHEN substr(media_code, 2, 1) = 'L' THEN 'Referal'     WHEN substr(media_code, 2, 1) = 'P' THEN 'WAP' WHEN substr(media_code, 2, 1) = 'R' THEN 'MyNSP'     WHEN substr(media_code, 2, 1) = 'V' THEN 'My Telkomsel' WHEN substr(media_code, 2, 1) = 'Q' THEN 'Music Max'     ELSE 'Promo' END ORDER BY calling_date, substr(calling_time, 1, 2)")
    
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
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
            new_subs.write(r,1, row[0], date_format)
            new_subs.write(r,2, row[1], content)
            new_subs.write(r,3, row[2], content)
            new_subs.write(r,4, row[3], content)
            new_subs.write(r,5, row[4], content)
            new_subs.write(r,6, row[5], content)
            new_subs.write(r,7, row[6], content)
            new_subs.write(r,8, row[7], content)
            
    cursor.close()


# In[ ]:


# A. PROV TOTAL :
def A_PROV_TOTAL():
    cursor = database.cursor()
    query = ("select calling_date,count(*),sum(case when prov_result_cde=0 then 1 else 0 end),round(100*sum(case when prov_result_cde=0 then 1 else 0 end)/count(*),2) || '%'     from call_detailed_record where calling_type in (1,2,7,64,93,78,77,65,95,96,99,100) and calling_date = to_char(sysdate-1,'YYYYMMDD') and trim(prov_tr_id) is not null     group by calling_date order by calling_date")
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    prov.write('A1','A. PROV TOTAL',title)
    prov.write('B3','CALLING_DATE',bold)
    prov.set_column(1,4, 15)
    prov.write('C3','ATT',bold)
    prov.write('D3','SUCC',bold)
    prov.write('E3','SR',bold)
    
    for r, row in enumerate(result, start=3):
        for c, col in enumerate(row):
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
            prov.write(r,1, row[0], date_format)
            prov.write(r,2, row[1], content)
            prov.write(r,3, row[2], content)
            prov.write(r,4, row[3], content)
            
    cursor.close()


# In[ ]:


# B. PROV CHANNEL :
def B_PROV_CHANNEL():
    cursor = database.cursor()
    query = ("SELECT calling_date, channel,attempt_PA,success_PA ,attempt_PB, success_PB FROM (select  calling_date , nvl(media_name,'renewal') as channel,     sum(case when calling_type in (1) then 1 else 0 end) as attempt_PA, sum(case when calling_type in (1) and prov_result_cde=0 then 1 else 0 end) as success_PA,     sum(case when calling_type in (2,7,77) then 1 else 0 end) as attempt_PB, sum(case when calling_type in (2,7,77) and prov_result_cde=0 then 1 else 0 end) as success_PB     from call_detailed_record cdr left join (select distinct media_code, media_name from media_info) mi on substr(mi.media_code,2,1) = substr(cdr.media_code,2,1)     where calling_type in (1,2,7,77) and calling_date = to_char(sysdate-1,'yyyymmdd') and trim(prov_tr_id) is not null group by calling_date, nvl(media_name,'renewal')     order by calling_date, nvl(media_name,'renewal'))")
    
    # Exec query new_subs
    cursor.execute(query)
    result = cursor.fetchall()
    prov.write('A6','B. PROV CHANNEL',title)
    prov.merge_range('B8:B9','CALLING_DATE',bold)
    prov.merge_range('C8:C9','CHANNEL',bold)
    prov.set_column('C:C', 18)
    prov.merge_range('D8:E8','PURCHASE AUTORENEWAL ON (A)',bold)
    prov.write('D9','ATTEMTPT',bold)
    prov.write('E9','SUCCESS',bold)
    prov.merge_range('F8:G8','PURCHASE AUTORENEWAL OFF (B)',bold)
    prov.write('F9','ATTEMTPT',bold)
    prov.write('G9','SUCCESS',bold)
    prov.set_column(3,8, 16)
    prov.merge_range('H8:H9','SR (A)',bold)
    prov.merge_range('I8:I9','SR (B)',bold)

    for line in ['B','C','D','E','F','G','H','I']:
        prov.write(line+'18','',neat)
        prov.write(line+'19','',neat)
        prov.write(line+'20','',neat)
        
    for row_H in range(11):
        SRA_Source = 'D{}/E{}'.format(*([row_H + 10] * 11))
        SRA_Cetak = "H{}".format(*([row_H + 10] * 11))
        prov.write_formula(SRA_Cetak,'=IFERROR('+SRA_Source+',"-")',percentage)
        
    for row_I in range(11):
        SRB_Source = 'F{}/G{}'.format(*([row_I + 10] * 11))
        SRB_Cetak = "I{}".format(*([row_I + 10] * 11))
        prov.write_formula(SRB_Cetak,'=IFERROR('+SRB_Source+',"-")',percentage)
    
    for r, row in enumerate(result, start=9):
        for c, col in enumerate(row):
            content = workbook.add_format({'font_color': 'black', 'num_format': '#,##0', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            content.set_border()
            date_format = workbook.add_format({'font_color': 'black', 'num_format': 'yyyymmdd', 'align': 'center', 'valign': 'vcenter', 'border': 1})
            date_format.set_border()
            prov.write(r,1, row[0], date_format)
            prov.write(r,2, row[1], content)
            prov.write(r,3, row[2], content)
            prov.write(r,4, row[3], content)
            prov.write(r,5, row[4], content)
            prov.write(r,6, row[5], content)
            
    cursor.close()


# In[ ]:


# A_KPI()
# B_Transaction()
# C_Subscriber()
# C_FreePaid()
# C_SUBS_WITHDRAWAL()
# C_CHURNED_USER()
# A_NEW_SUBSCRIPTION_TOTAL()
B_NEW_SUBSCRIPTION_CHANNEL()
# C_NEW_SUBSCRIPTION_HOUR()
# D_NEW_SUBSCRIPTION_HOUR_PER_CHANNEL()
A_PROV_TOTAL()
B_PROV_CHANNEL()
workbook.close()
database.close()
SendEmail()
