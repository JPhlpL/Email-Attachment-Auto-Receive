# Improvements
# check if have attached file
# check if not xlsx format
# check if not fit on the format (xlsx)
# button for stop and exit

import win32com.client
import os
import mysql.connector
import pandas as pd
import datetime
import warnings
import time
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
# Email configuration
smtp_server = 'smtpserv'
smtp_port = 25
#for current timestamp
current_timestamp = datetime.now()
cur_month = current_timestamp.strftime("%B")
cur_month_num = current_timestamp.strftime("%m")
cur_date = current_timestamp.strftime("%d")
cur_year = current_timestamp.strftime("%Y")
#allocation
save_folder = "C:/Attachments/"
#database configuration
db_config = {
    "host": "localhost",
    "user": "prod",
    "password": "MTA4ODYuRGVuc28=",
    "database": "autosend_survey"
}

conn = mysql.connector.connect(**db_config)
cursor = conn.cursor()


def get_current_datetime():
    current_datetime = datetime.now()
    get_current_datetime = current_datetime.strftime("%Y-%m-%d %H:%M:%S")
    return get_current_datetime

def insert_log(log):
    print(log)
    insert_log = "INSERT INTO tblexecute_logs (EXECUTE_LOGS) VALUES (%s)"
    values = (log,)
    cursor.execute(insert_log, values)
    conn.commit()
    f.write(log + "\n")
    f.flush()

def import_num():
    import_year = cur_year
    import_month = cur_month_num
    import_initial = '{:06d}'.format(1)
    base_control = 'IMPORT'

    # add hash code for each transaction
    select_query = "SELECT MAX(RESPONSE_IMPORT_NUM) AS importNum FROM tblsupplier_response"
    cursor.execute(select_query)
    control_number = cursor.fetchone()[0]

    if not control_number or control_number is None :
        import_number = base_control + "-" + import_year + "-" + import_month + "-" + import_initial

    else:
        import_string = control_number.split('-')
        increment_num = int(import_string[4]) + 1
        final_num = '{:06d}'.format(increment_num)
        import_number = base_control + "-" + import_year + "-" + import_month + "-" + final_num

    return import_number

with open('execute_logs.txt', 'a') as f:

    def save_attachment(mail_item, response_import_number, sender):
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()

        #import number -> once per import excel transact
        import_number = response_import_number

        for attachment in mail_item.Attachments:
            #print (attachment.FileName)
            attachment.SaveAsFile(os.path.join(save_folder, attachment.FileName))

            if attachment.FileName.endswith(".xlsx"):
                excel_file = os.path.join(save_folder, attachment.FileName)
                df = pd.read_excel(excel_file, skiprows=1, na_filter=False)  # Skip the first row

                for index, row in df.iterrows():
                    filename = attachment.FileName
                    transact_num = row.iloc[0]  # Extract value from column 'A'
                    part_name = row.iloc[1]  # Extract value from column 'B'
                    part_num = row.iloc[2]  # Extract value from column 'C'
                    supplier = row.iloc[3]  # Extract value from column 'D'
                    onhand_stocks = row.iloc[4]  # Extract value from column 'E'
                    incoming_del = row.iloc[5]  # Extract value from column 'F'
                    shipment_sched = row.iloc[6]  # Extract value from column 'G'
                    qty_alloted_yes = row.iloc[7]  # Extract value from column 'H'
                    qty_alloted_no = row.iloc[8]  # Extract value from column 'I'

                    select_query = "SELECT COUNT(*) FROM tblsupplier_response " \
                                   "WHERE RESPONSE_EXCEL_NAME = %s AND " \
                                   "RESPONSE_TRANSACT_NUM = %s AND " \
                                   "RESPONSE_ITEM_DESC = %s AND " \
                                   "RESPONSE_ITEM_PART_NUM = %s AND "\
                                   "RESPONSE_SUPPLIER = %s AND " \
                                   "RESPONSE_ONHAND = %s AND " \
                                   "RESPONSE_INCOMING_DELIVER = %s AND " \
                                   "RESPONSE_SHIPMENT_SCHED = %s AND " \
                                   "RESPONSE_YES = %s AND " \
                                   "RESPONSE_NO_QTY = %s"

                    cursor.execute(select_query, (
                    filename, transact_num, part_name, part_num, supplier, onhand_stocks, incoming_del, shipment_sched,
                    qty_alloted_yes, qty_alloted_no))
                    count = cursor.fetchone()[0]

                    if count == 0:

                        insert_query = "INSERT INTO tblsupplier_response (RESPONSE_EXCEL_NAME, " \
                                       "RESPONSE_IMPORT_NUM," \
                                       "RESPONSE_TRANSACT_NUM," \
                                       "RESPONSE_ITEM_DESC, " \
                                       "RESPONSE_ITEM_PART_NUM, " \
                                       "RESPONSE_SUPPLIER, " \
                                       "RESPONSE_ONHAND, " \
                                       "RESPONSE_INCOMING_DELIVER, " \
                                       "RESPONSE_SHIPMENT_SCHED, " \
                                       "RESPONSE_YES, " \
                                       "RESPONSE_NO_QTY, " \
                                       "RESPONSE_METHOD) " \
                                       "VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
                        values = (filename, import_number, transact_num, part_name, part_num, supplier, onhand_stocks, incoming_del,
                                  shipment_sched, qty_alloted_yes, qty_alloted_no, "RESPONDED")
                        cursor.execute(insert_query, values)
                        conn.commit()

                        fetch_masteritem_sql = "SELECT COUNT(*) FROM tblmasteritem WHERE MASTERITEM_ITEM_DESC = %s AND MASTERITEM_PART_NUM = %s"
                        cursor.execute(fetch_masteritem_sql, (part_name, part_num))
                        masteritem_count = cursor.fetchone()[0]

                        if masteritem_count == 1:
                            current_modified = datetime.now()
                            update_query = "UPDATE tblmasteritem SET MASTERITEM_ONHAND = %s, MASTERITEM_INCOMING_DELIVER = %s, MASTERITEM_SHIPMENT_SCHED = %s, MASTERITEM_YES = %s, MASTERITEM_NO_QTY = %s, MASTERITEM_TIMESTAMP_MODIFIED = %s,MASTERITEM_SUPPLIER_SENT_FILE_COUNT = 1 WHERE MASTERITEM_PART_NUM = %s"
                            update_values = (
                            onhand_stocks, incoming_del, shipment_sched, qty_alloted_yes, qty_alloted_no,
                            current_modified, part_num)
                            cursor.execute(update_query, update_values)
                            conn.commit()

                            #######
                            log = get_current_datetime() + "->" + f"Sender: {sender}"
                            insert_log(log)
                            #######

                            #######
                            log = get_current_datetime() + "->" + f"Data updated on masterlist: {part_num}"
                            insert_log(log)
                            #######

                        #if master_count == 0:
                    #if count == 0:
                #end of loop

            # if not xlsx
            else:

                #######
                #log = get_current_datetime() + "->" + f"Sender: {sender}"
                #insert_log(log)
                #######

                #######
                #log = get_current_datetime() + "->" + f"Incorrect Data File"
                #insert_log(log)
                #######

                os.remove(os.path.join(save_folder, attachment.FileName))
                pass

        #end of saving attachment
        time.sleep(1)
        send_emails(response_import_number)
        # Create or put function for sending email

        cursor.close()
        conn.close()

    def send_emails(response_import_num):
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        fetch_masteritem_sql = ("SELECT DISTINCT(a.SUPPLIER_EMAIL), b.RESPONSE_IMPORT_NUM, b.RESPONSE_SUPPLIER FROM tblsupplier a JOIN tblsupplier_response b ON a.SUPPLIER_NAME = b.RESPONSE_SUPPLIER WHERE RESPONSE_IMPORT_NUM LIKE '%%%s%%'" % response_import_num)
        cursor.execute(fetch_masteritem_sql)
        supp_res = cursor.fetchall()
        for row in supp_res:
            supplier_email = row[0]
            imp_num = row[1]
            supplier_name = row[2]
            #
            log = get_current_datetime() + "->" + f"Email : {supplier_email}"
            insert_log(log)
            #
            log = get_current_datetime() + "->" + f"Import Number : {imp_num}"
            insert_log(log)
            #
            log = get_current_datetime() + "->" + f"Supplier Name : {supplier_name}"
            insert_log(log)
            #

            # initializing email setup
            to_recipient = supplier_email  # user // Must Change for deployment
            # working cc
            cc_emails = ['verna.magnayi.a4b@ap.denso.com',to_email]  # user
            sender_email = "autosendsurveyproc-noreply@ap.denso.com"  # user

            # Create the email message
            subject = "[Procurement - Autosend Survey] New transaction added as of " + str(
                cur_month) + " " + str(
                cur_date) + " " + str(cur_year)

            # query to check all record based on import number
            rec_query = ("SELECT RESPONSE_ITEM_DESC, RESPONSE_ONHAND, RESPONSE_INCOMING_DELIVER, RESPONSE_SHIPMENT_SCHED, RESPONSE_YES, RESPONSE_NO_QTY FROM tblsupplier_response WHERE RESPONSE_IMPORT_NUM LIKE '%%%s%%'" % imp_num)
            cursor.execute(rec_query)
            response = cursor.fetchall()

            message = """\Message"""


            for rec in response:
                RESPONSE_ITEM_DESC = rec[0]
                RESPONSE_ONHAND = rec[1]
                RESPONSE_INCOMING_DELIVER = rec[2]
                RESPONSE_SHIPMENT_SCHED = rec[3]
                RESPONSE_YES = rec[4]
                RESPONSE_NO_QTY = rec[5]

                log = get_current_datetime() + "->" + f" Data record successfully <br> Item Description -> {RESPONSE_ITEM_DESC}" \
                                                      f"<br> Onhand Qty -> {RESPONSE_ONHAND}" \
                                                      f"<br> Incoming Delivery -> {RESPONSE_ITEM_DESC}" \
                                                      f"<br> Item Description -> {RESPONSE_ITEM_DESC}" \
                                                      f"<br> Shipment Schedule -> {RESPONSE_ITEM_DESC}" \
                                                      f"<br> Response (If Yes) -> {RESPONSE_ITEM_DESC}" \
                                                      f"<br> Response (If No) -> {RESPONSE_ITEM_DESC}"
                insert_log(log)

                
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = to_recipient
            msg['Cc'] = ', '.join(cc_emails)
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'html'))

            # Connect to the SMTP server and send the email
            # Must Check
            try:
                server = smtplib.SMTP(smtp_server, smtp_port)

                # user
                to_recipient_list = to_recipient.split(';')
                to_emails = to_recipient_list + cc_emails
                server.sendmail(sender_email, to_emails, msg.as_string())

                server.quit()

                #######
                log = get_current_datetime() + "->" + f"New record added successfully"
                insert_log(log)
                #######

            except Exception as e:

                #######
                log = get_current_datetime() + "->" + f"Email sending failed: {str(e)}"
                insert_log(log)
                #######

        time.sleep(1)

    def check_emails():

        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)  # 6 represents the inbox folder
        emails = inbox.Items
        required_emails = emails.Restrict("@SQL=urn:schemas:httpmail:subject LIKE '%Supplier Inventory and Stock Condition%'")

        for email in required_emails:
            if email.Class == 43:  # 43 corresponds to a MailItem
                if email.SenderEmailType == "EX":
                    sender = email.Sender.GetExchangeUser().PrimarySmtpAddress
                else:
                    sender = email.SenderEmailAddress
                # import number -> once per import excel transact
                response_import_number = import_num()
                save_attachment(email, response_import_number, sender)

    if __name__ == "__main__":

        welcome_screen = """\

             
     _____       _              _____                 _           
    |  _  | _ _ | |_  ___  ___ | __  | ___  ___  ___ |_| _ _  ___ 
    |     || | ||  _|| . ||___||    -|| -_||  _|| -_|| || | || -_|
    |__|__||___||_|  |___|     |__|__||___||___||___||_| \_/ |___|
                                                                                                      
     _____  _              _      _____            _       _   
    |   __|| |_  ___  ___ | |_   |   __| ___  ___ |_| ___ | |_ 
    |__   ||  _|| . ||  _|| '_|  |__   ||  _||  _|| || . ||  _|
    |_____||_|  |___||___||_,_|  |_____||___||_|  |_||  _||_|  
                                                     |_|       
                                                     
    NOTE: KINDLY OPEN YOUR MICROSOFT OUTLOOK APPLICATION ALWAYS. THANK YOU!
         """

        print (welcome_screen)

        to_email = input("Enter your email: ")

        #######
        log = get_current_datetime() + "->" + f"Greetings {to_email}!"
        insert_log(log)
        time.sleep(1)
        #######

        #######
        log = get_current_datetime() + "->" + "Execute Received"
        insert_log(log)
        time.sleep(1)
        #######

        #######
        log = get_current_datetime() + "->" + "Initializing..."
        insert_log(log)
        time.sleep(1)
        #######

        #######
        log = get_current_datetime() + "->" + "Waiting for a file..."
        insert_log(log)
        time.sleep(1)
        #######

        if not os.path.exists(save_folder):
            os.makedirs(save_folder)
            #######
            log = get_current_datetime() + "->" + "Created a Folder"
            insert_log(log)
            time.sleep(1)
            #######

        while True:
            warnings.simplefilter("ignore")
            check_emails()
            # f.flush()






