from email import policy
from email.parser import BytesParser
from email.utils import parsedate_tz, mktime_tz
import datetime
import re
import time
import os
import logging
import mailbox
import glob
from multiprocessing.pool import ThreadPool
from login import create_connection
import mbox
import base64


try:
    if os.path.exists('logs/info.log'):
        os.remove('logs/info.log')
    if os.path.exists('logs/warning.log'):
        os.remove('logs/warning.log')
    if os.path.exists('logs/ticket_errors.log'):
        os.remove('logs/ticket_errors.log')
    if os.path.exists('logs/missing_tickets.log'):
        os.remove('logs/missing_tickets.log')
except Exception as e:
    print('Unable to delete', e)
# regex = '''([a-zA-Z0-9.!#$%&'*+-/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*)'''
regex = re.compile(
    r'''([a-zA-Z0-9.!#$%&'*+-/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)*)''', re.MULTILINE)
email_force_tld_regex = re.compile(
    r'''([a-zA-Z0-9.!#$%&'*+-/=?^_`{|}~-]+@[a-zA-Z0-9-]+(?:\.[a-zA-Z0-9-]+)+)''', re.MULTILINE)
sizeRegex = re.compile(r'''\bsize=[0-9]+\b''', re.MULTILINE)

items = 0
#### Loggers###############
logger = logging.getLogger('main')
missing_ticket = logging.getLogger('missing')
missing_sub_logger = logging.getLogger('info')
error = logging.getLogger('error')
############# SET LOG LEVEL ##################
missing_sub_logger.setLevel(logging.WARNING)
logger.setLevel(logging.INFO)
missing_ticket.setLevel(logging.DEBUG)
error.setLevel(logging.ERROR)
############  FORMATTER #######################
formatter = logging.Formatter('%(asctime)s:%(levelname)s:%(message)s')

############# FILE handlers ########################
file_handler = logging.FileHandler('logs/info.log')
missing_ticket_file_handler = logging.FileHandler('logs/missing_tickets.log')
missing_sub_logger_file_handler = logging.FileHandler('logs/warning.log')
error_file_handler = logging.FileHandler('logs/ticket_errors.log')
################## set formatter ################################
missing_sub_logger_file_handler.setFormatter(formatter)
missing_ticket_file_handler.setFormatter(formatter)
file_handler.setFormatter(formatter)
error_file_handler.setFormatter(formatter)

#################### Add handler ############################
logger.addHandler(file_handler)
missing_ticket.addHandler(missing_ticket_file_handler)
missing_sub_logger.addHandler(missing_sub_logger_file_handler)
error.addHandler(error_file_handler)


def get_timestamp(string, gmail=True):
    try:
        if gmail:
            # https://stackoverflow.com/questions/62092529/convert-date-string-from-gmail-to-timestamp-python
            return mktime_tz(parsedate_tz(string))
#             return datetime.datetime.strptime(string,"%a, %d %b %Y %H:%M:%S %z").timestamp()
        else:
            return datetime.datetime.strptime(string, "%m/%d/%y, %I:%M %p").timestamp()
    except Exception as e:
        return string


def get_mbox(path, progress, loading, progress_bar):
    global items
    emails = []
    items = 0
    mboxs = [mbox for mbox in glob.glob(f"{path}/**/*.mbox", recursive=True)]
    if len(mboxs) >= 1:
        for idx, m in enumerate(mboxs):
            mbox_name = ""
            try:
                mbox_name = os.path.dirname(m).split(
                    "Takeout")[0].split("/")[-1].split("\\")[0] + ".zip "
            except Exception as e:
                error.exception(e)
            loading.emit((1, f'Loading Mbox {idx+1} into memory'))
            # print(f'Loading Mbox {idx+1} into memory')
            mboxObj = mailbox.mbox(m)

            total_emails = len(mboxObj)
            # print("Emails Detected in Mbox {} : {}".format(idx+1,total_emails))
            loading.emit(
                (2, f"Emails Detected in Mbox {idx+1} : {total_emails}", total_emails))
            emails.extend(parse_mbox(mbox_name, mboxObj,
                          total_emails, progress, progress_bar))
        return emails
    else:
        return None


def process_emls(path, progress, loading, progress_bar, _error):
    emls = [eml for eml in glob.glob(f"{path}/**/*.eml", recursive=True)]
    emails = []
    if len(emls) > 0:
        loading.emit(
            (2, f"Total Emls files detected : {len(emls)}", len(emls)))
        futures = ThreadPool(5).imap(parse_eml, emls)
        idx = 0
        conn = create_connection('ticket.db')
        if conn is not None:
            with conn:
                cur = conn.cursor()
                mbox_query = '''SELECT * FROM mbox ORDER BY ID DESC LIMIT 1'''
                cur.execute(mbox_query)
                row = cur.fetchone()
                insert_query = '''INSERT INTO mbox(DATE,ITEMS,START,END)
                    VALUES(?,?,?,?)'''
                t = time.localtime()
                current_time = time.strftime("""%d-%m-%Y %I:%M:%S %p""", t)
                if row is None:
                    value = (current_time, len(emls), 1, len(emls))
                    cur.execute(insert_query, value)
                else:
                    value = (current_time, len(emls), int(
                        row[4])+1, int(row[4])+1+len(emls))
                    cur.execute(insert_query, value)
                cur.connection.commit()
        for future in futures:
            if future is not None:
                idx += 1
                incr = round((idx/len(emls))*100)
                progress_bar.emit((incr, idx))
                tmp = future
                if isinstance(future[3], Exception):
                    pass
                else:
                    emails.append(tmp)
                tmp[0] = idx
                progress.emit(tmp)

        return emails
    else:
        return None


def parse_eml(eml):
    try:
        with open(eml, 'rb') as fp:
            eml_msg = BytesParser(policy=policy.default).parse(fp)
            time_format = """%a, %b %d, %I:%M %p"""
            timestamp = get_timestamp(eml_msg['Date'])
            t = time.localtime(timestamp)
            local_time = time.strftime(time_format, t)
            if "support@idrive.com" in regex.findall(eml_msg['From']) or \
                "support@ibackup.com" in regex.findall(eml_msg['From']) or \
                "sales@idrive.com" in regex.findall(eml_msg['From']) or \
                "sales@remotepc.com" in regex.findall(eml_msg['From']) or \
                "sales@ibackup.com" in regex.findall(eml_msg['From']) or \
                "no-reply@idrive.com" in regex.findall(eml_msg['From']) or \
                "noreply@idrivee2.com" in regex.findall(eml_msg['From']) or \
                "privacy@idrive.com" in regex.findall(eml_msg['From']) or \
                "info@idrive.com" in regex.findall(eml_msg['From']) or \
                "info@remotepc.com" in regex.findall(eml_msg['From']) or \
                "privacy@remotepc.com" in regex.findall(eml_msg['From']) or \
                "support@idrivemirror.com" in regex.findall(eml_msg['From']) or \
                "support@send.idrive.com" in regex.findall(eml_msg['From']) or \
                "support@send.ibackup.com" in regex.findall(eml_msg['From']) or \
                "support@send.remotepc.com" in regex.findall(eml_msg['From']) or \
                "privacy@ibackup.com" in regex.findall(eml_msg['From']) or \
                    "support@remotepc.com" in regex.findall(eml_msg['From']):
                text = eml_msg.get_body(preferencelist=('plain')).get_content()
                emails = regex.findall(text)
                try:
                    missing_sub_logger.warning('Subject: {} : From: {} : Emails in body : {}'.format(eml_msg['Subject'],
                                                                                                     eml_msg['X-Original-Sender'], emails[:3]))
                except Exception as e:
                    error.exception(
                        f"Logging Error: {e} - {eml_msg['X-Original-Sender']} - {eml_msg['Subject']}")
                return [1, emails[:3], eml_msg['Subject'], local_time]

            elif "supportgroup@idrive.com" in regex.findall(eml_msg['From']) or \
                 "group1@remotepc.com" in regex.findall(eml_msg['From']):
                try:
                    logger.info('Emails: {} : Subject: {} : TimeStamp : {}'.format(
                        regex.findall(eml_msg['X-Original-Sender']), eml_msg['Subject'], local_time))
                except Exception as e:
                    error.exception(
                        f"Logging Error : {e} - {eml_msg['X-Original-Sender']} - {eml_msg['Subject']}")
                return [1, regex.findall(eml_msg['X-Original-Sender']), eml_msg['Subject'], local_time]

            else:
                try:
                    logger.info('Emails: {} : Subject: {} : TimeStamp : {}'.format(
                        regex.findall(eml_msg['From']), eml_msg['Subject'], local_time))
                except Exception as e:
                    error.exception(f"Logging Error : {e} - {eml_msg['From']}")
                return [1, regex.findall(eml_msg['From']), eml_msg['Subject'], local_time]

    except Exception as e:
        error.exception(f"File Parse Error: Name: {eml} : {e}")
        return [1, eml, "File inaccessible or too long", e]

# def find_attachment(msg):
#     attachments = []
#     if msg.get_content_maintype() == 'multipart':
#         # Iterate through the message's payload (i.e., the attachments)
#         for part in msg.get_payload():
#         # Check if the part is an attachment
#             if part.get_content_disposition() == 'attachment':
#                 try:
#                     #  Extract the size of the attachment from the "Content-Length" header field
#                     attachment_data = part.get_payload(decode=True)
#                     size = len(base64.b64decode(attachment_data))
#                     attachments.append(size)
#                 except Exception as e:
#                     error.exception(e)
#     return attachments


def get_attachment_size(part, attachments):
    if part.get_content_disposition() == 'attachment':
        try:
            #  Extract the size of the attachment from the "Content-Diposition" header field
            attachment = sizeRegex.findall(part['Content-Disposition'])
            if len(attachment) > 0:
                attachments.append(attachment[0].split("=")[-1])
            else:
                #  Extract the size of the attachment by reading the base64 payload
                attachment_data = part.get_payload(decode=True)
                size = len(base64.b64decode(attachment_data))
                attachments.append(size)
        except Exception as e:
            error.exception(e)


def find_attachment(message):
    attachments = []
    if message.is_multipart():
        for part in message.get_payload():
            get_attachment_size(part, attachments)
    else:
        get_attachment_size(message, attachments)
    return attachments


def parse_mbox(mbox_name, mboxObj, total_emails, progress, progress_bar):
    global items
    time_format = """%a, %b %d, %I:%M %p"""
    from_email = []
    conn = create_connection('ticket.db')
    if conn is not None:
        with conn:
            cur = conn.cursor()
            query = '''INSERT INTO emails(EMAIL,SUBJECT,DATE,ATTACHMENT)
                        VALUES(?,?,?,?)'''
            mbox_query = '''SELECT * FROM mbox ORDER BY ID DESC LIMIT 1'''
            cur.execute(mbox_query)
            row = cur.fetchone()
            insert_query = '''INSERT INTO mbox(DATE,ITEMS,START,END)
                    VALUES(?,?,?,?)'''
            t = time.localtime()
            current_time = mbox_name + \
                "[" + time.strftime("""%d-%m-%Y %I:%M:%S %p""", t) + "]"
            if row is None:
                value = (current_time, total_emails, 1, total_emails)
                cur.execute(insert_query, value)
            else:
                value = (current_time, total_emails, int(
                    row[4])+1, int(row[4])+1+total_emails)
                cur.execute(insert_query, value)

            for idx, msg in enumerate(mboxObj):
                items += 1
                per = round((idx/total_emails)*100)

                attachments = find_attachment(msg)
                progress_bar.emit((per, idx+1, msg['From']))

                if "support@idrive.com" in regex.findall(msg['From']) or \
                    "no-reply@idrive.com" in regex.findall(msg['From']) or \
                    "sales@idrive.com" in regex.findall(msg['From']) or \
                    "sales@remotepc.com" in regex.findall(msg['From']) or \
                    "sales@ibackup.com" in regex.findall(msg['From']) or\
                    "noreply@idrivee2.com" in regex.findall(msg['From']) or \
                    "support@remotepc.com" in regex.findall(msg['From']) or \
                    "privacy@idrive.com" in regex.findall(msg['From']) or \
                    "info@idrive.com" in regex.findall(msg['From']) or \
                    "support@idrivemirror.com" in regex.findall(msg['From']) or \
                    "info@remotepc.com" in regex.findall(msg['From']) or \
                    "privacy@remotepc.com" in regex.findall(msg['From']) or \
                    "support@send.idrive.com" in regex.findall(msg['From']) or \
                    "support@send.ibackup.com" in regex.findall(msg['From']) or \
                    "support@send.remotepc.com" in regex.findall(msg['From']) or \
                    "privacy@ibackup.com" in regex.findall(msg['From']) or \
                        "support@ibackup.com" in regex.findall(msg['From']):
                    email_data = mbox.GmailMboxMessage(msg)
                    emails = []
                    for content in email_data.read_email_payload():
                        try:
                            emails.extend(regex.findall(content[2]))
                        except Exception as e:
                            error.exception(
                                f"ParseMBox: {e} : {msg['X-Original-Sender']} - {msg['Subject']}")
                    try:
                        missing_sub_logger.warning('Subject: {} : From: {} : Emails in body : {}'.format(msg['Subject'],
                                                                                                         msg['X-Original-Sender'], emails[:3]))
                        logger.info('Emails: {} : Subject: {} : TimeStamp : {}'.format(
                            regex.findall(msg['X-Original-Sender']), msg['Subject'], msg['Date']))

                        timestamp = get_timestamp(msg['Date'])
                        t = time.localtime(timestamp)
                        local_time = time.strftime(time_format, t)
                        value = (
                            f'''{emails[:3]}''', msg['Subject'], local_time, f'''{attachments}''')
                        from_email.append(
                            [items, emails[:3], msg['Subject'], local_time, attachments, email_data])
                        progress.emit(
                            [items, emails[:3], msg['Subject'], local_time, attachments, email_data])
                        cur.execute(query, value)
                    except Exception as e:
                        error.exception(
                            f"Logging Error: {e} - {msg['X-Original-Sender']} - {msg['Subject']}")

                elif "supportgroup@idrive.com" in regex.findall(msg['From']) or \
                        "group1@remotepc.com" in regex.findall(msg['From']):
                    try:
                        logger.info('Emails: {} : Subject: {} : TimeStamp : {}'.format(
                            regex.findall(msg['X-Original-Sender']), msg['Subject'], msg['Date'], attachments))

                        timestamp = get_timestamp(msg['Date'])
                        t = time.localtime(timestamp)
                        local_time = time.strftime(time_format, t)
                        value = (f'''{regex.findall(msg['X-Original-Sender'])}''',
                                 msg['Subject'], local_time, f'''{attachments}''')
                        body_content = mbox.GmailMboxMessage(msg)
                        progress.emit([items, regex.findall(
                            msg['X-Original-Sender']), msg['Subject'], local_time, attachments, body_content])
                        from_email.append([items, regex.findall(
                            msg['X-Original-Sender']), msg['Subject'], local_time, attachments, body_content])
                        cur.execute(query, value)
                    except Exception as e:
                        error.exception(
                            f"Logging Error : {e} - {msg['X-Original-Sender']}")

                else:
                    try:
                        logger.info('Emails: {} : Subject: {} : TimeStamp : {}'.format(
                            regex.findall(msg['From']), msg['Subject'], msg['Date'], attachments))

                        timestamp = get_timestamp(msg['Date'])
                        t = time.localtime(timestamp)
                        local_time = time.strftime(time_format, t)
                        value = (
                            f'''{regex.findall(msg['From'])}''', msg['Subject'], local_time, f'''{attachments}''')
                        body_content = mbox.GmailMboxMessage(msg)
                        progress.emit([items, regex.findall(
                            msg['From']), msg['Subject'], local_time, attachments, body_content])
                        from_email.append([items, regex.findall(
                            msg['From']), msg['Subject'], local_time, attachments, body_content])
                        cur.execute(query, value)
                    except Exception as e:
                        error.exception(f"Logging Error : {e} - {msg['From']}")
            cur.connection.commit()
    else:
        # print("Error", "Something went wrong. Try again")
        pass
    return from_email


if __name__ == "__main__":
    pass
