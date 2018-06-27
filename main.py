#!/usr/local/lib/python3.6
# Use Dictionary instead of List
# line 57-61: reduce code lines to one line

from simple_salesforce import Salesforce
from slackclient import SlackClient
import time, datetime
import imaplib, email, sys
import re, os, socket
import credential, traceback

# Enable Debug Logging to Debug.log file
def debug():
    return False

# Function to print date
def date():
    Time = time.strftime("%I:%M %p", time.localtime())
    return Time

# Check Internet connectivity
def is_connected():
  try:
    host = socket.gethostbyname("www.google.com")
    s = socket.create_connection((host, 80), 2)
    return True
  except:
    log('Error: Internet Connection down, Retrying after 60 seconds\n')
    time.sleep(60)
    return False

# SalesForce Case list dump
def sf():
    global CASES_DATA
    dir_path = os.path.dirname(os.path.realpath(__file__))
    users = os.path.join(dir_path, "users.ini")
    with open(users) as f:
        CASES_DATA = [x.strip().split(':') for x in f.readlines()]
    try:
        sf = Salesforce(username=credential.username, password=credential.password, security_token=credential.security_token)
    except:
        log('Fatal: Invalid Credentials/Token; Exitting\n')
        time.sleep(60)
        return
    j = 0
    for user,ownerid in CASES_DATA:
        lst = []
        query = "SELECT CaseNumber FROM Case WHERE OwnerId = '%s'" %ownerid
        try:
            prt = sf.query_all(query)
        except simple_salesforce.exceptions.SalesforceGeneralError:
            log('Salesforce Error, retrying in 60 seconds\n')
            time.sleep(60)
            pass
        for i in range(len(prt['records'])):
            lst.append(str(prt['records'][i]['CaseNumber']))
        if CASES_DATA[j][0]==user:
            CASES_DATA[j].append(lst)
        j += 1

# Function to get & save latest MessID
def id():
    try:
        mail.select('inbox', readonly=True)
        type, data = mail.search(None, '(ALL)')
        mail_ids = data[0]
        id_list = mail_ids.split()
        global b
        b = int(id_list[-1])
    except imaplib.IMAP4.abort:
        log('Imaplib.IMAP4.abort Error: Retrying in 60 seconds\n')
        time.sleep(60)
        mailbox()
        log('Re-connected to the Mail Server!\n')
        return
    except TimeoutError:
        log('TimeoutError: Retrying in 60 seconds\n')
        time.sleep(60)
        return
    except OSError:
        log('OSError: Retrying in 60 seconds\n')
        time.sleep(60)
        return
    except BrokenPipeError:
        log('BrokenPipeError: Retrying in 60 seconds\n')
        time.sleep(60)
        return
    except:
        log(traceback.print_exc())
        sys.exit()

# Current Mail Fetch Details
def write_b():
    if b > a:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        curr = os.path.join(dir_path, "curr.ini")
        f = open(curr,'w')
        f.write(str(b))
        if debug():
            debug_log('Write b='+str(b)+'\n')         # Debugging
        f.close()
    else:
        log('Error: Rollback event occured, ignored\n')

# Fetch Last Seen MessID
def read_a():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    curr = os.path.join(dir_path, "curr.ini")
    f = open(curr,'r')
    l = f.readlines()
    global a
    a = int(l[0])
    f.close()
    log('Last processed MessID: '+str(a)+'\n')

# Fetch new mails function
def fetchmail():
    for i in range(a+1,b+1):
        if debug():
            debug_log('Mail Function: a='+str(a)+' b='+str(b)+' i='+str(i)+'\n')             # Debugging
        typ, data = mail.fetch(str(i), '(RFC822)')
        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                em_subject = msg['subject'].replace("\r\n","")
                email_subject = em_subject.split("- ref:")[0]
                email_from = msg['from']
                email_to = msg['to']
                if debug():
                    debug_log("Subject: "+email_subject+'\n'+"To: "+email_to+'\n')         # Debugging
                date_tuple = email.utils.parsedate_tz(msg['Date'])
                if date_tuple:
                    localdate = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                    local_date = str(localdate.strftime("%d %b %H:%M")+" (IST)")
                RAW_SUB = re.search(r'\s\d{4}', email_subject)
                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        body = part.get_payload(decode=True)
                        try:
                            string = body.decode('utf-8')[0:200]
                        except UnicodeDecodeError as u:
                            string = body.decode('ISO-8859-1')[0:200]
                        except:
                            log('Exception: '+str(u)+'\n')
                            log('part.get_content_type = '+part.get_content_type()+'\n')
                            string = "0"
                    else:
                        string = "0"
                if debug():
                    debug_log("Content type: "+str(part.get_content_type())+'\n')         # Debugging
                    debug_log("String: "+str(string.encode('utf-8').strip())+'\n')        # Debugging
                #print(string)
                #print(str(RAW_SUB[0]))
                #print(str(RAW_SUB[0].strip()))
                #IGNORE_EMAILS = ['internal-case-support@avinetworks.com',
                #        'noreply@salesforce.com',
                #        'avinetworks-support@avinetworks.com',
                #        'support-notification@avinetworks.com']
                try:
                    c = str(RAW_SUB[0].strip())
                    #print(c)
                    #################### Case Update Notification #######################
                    for j in range(len(CASES_DATA)):
                        if c in CASES_DATA[j][2]:
                            #print(j, CASES_DATA[j][0], CASES_DATA[j][2])
                            if any(re.findall('|'.join(credential.IGNORE_EMAILS),email_from)):
                                if 'internal-case-support@avinetworks.com' in email_from:
                                    if 'Case Attachment - ' in email_subject:
                                        log(local_date+'['+str(i)+'] ALERT:'+ CASES_DATA[j][0] +' Case:'+c+' From   : '+email_from+' Subject: '+email_subject+'\n')
                                        slack(local_date+' - '+email_subject+' ('+email_from+')', CASES_DATA[j][0], ':scroll:')
                                    else:
                                        log(local_date+'['+str(i)+'] IGNORED ALERT: '+ CASES_DATA[j][0] +' Case:'+c+' From   : '+email_from+' Subject: '+email_subject+'\n')
                                else:
                                    log(local_date+'['+str(i)+'] IGNORED ALERT: '+ CASES_DATA[j][0] +' Case:'+c+' From   : '+email_from+' Subject: '+email_subject+'\n')
                            else:
                                if "Sandbox: Case - " in email_subject:
                                    log(local_date+'['+str(i)+'] IGNORED ALERT: Sandbox Case: ' + c+'\n')
                                else:
                                    log(local_date+'['+str(i)+'] ALERT:'+ CASES_DATA[j][0] +' Case:'+c+' From   : '+email_from+' Subject: '+email_subject+'\n')
                                    slack(local_date+' - '+email_subject+' ('+email_from+')', CASES_DATA[j][0], ':scroll:')
                    #################### New Case Notification #######################
                    if "New Case Assigned:" in email_subject:
                        #if re.search(EMAIL_ID, email_to):
                            #sf()
                            log(local_date+'['+str(i)+'] '+ 'My Queue Case: '+c+' Subject: '+email_subject+'\n')
                    elif 'New Case:' in string:
                        if "Sandbox: Case - " in email_subject:
                            log(local_date+'['+str(i)+'] IGNORED ALERT: Sandbox Case: ' + c+'\n')
                        else:
                            log(local_date+'['+str(i)+'] New Case in QUEUE - Case: ' + c+'\n')
                            slack(local_date+' New '+email_subject, credential.channel, ':briefcase:')
                    else:
                        #print('Case: ' + c)
                        log(local_date+ '['+str(i)+'] Case:' + c+'\n')
                except TypeError:
                    #print('No Case ID')
                    log(local_date+'['+str(i)+'] No Case ID\n')

# Slack Alerts
def slack(message, channel, icon):
    sc = SlackClient(credential.token)
    sc.api_call('chat.postMessage', channel=channel, text=message, username='SREBot', icon_emoji=icon)

# Debug log file
def log(text):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    log_file = os.path.join(dir_path, "log_file.log")
    f = open(log_file,'a')
    f.write(str(date())+': '+str(text))
    f.close()

# Debug log file
def debug_log(text):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    debug_file = os.path.join(dir_path, "debug.log")
    f = open(debug_file,'a')
    f.write(str(date())+': '+str(text))
    f.close()

# Email Credentails:
ORG_EMAIL   = credential.ORG_EMAIL
EMAIL_ID    = credential.FROM_EMAIL
FROM_PWD    = credential.FROM_PWD
SMTP_SERVER = "imap.gmail.com"

# Email Box connect function:
def mailbox():
    global mail
    mail = imaplib.IMAP4_SSL(SMTP_SERVER)
    try:
        mail.login(EMAIL_ID,FROM_PWD)
    except:
        log('Unable to connect! Check Credentials\n')
        sys.exit()

# Connect to SF:
while True:
    if is_connected():
        sf()
        #log('Fetched results from SalesForce\n'+'Total Cases: '+str(len(CASES_LIST))+' => ' +str(CASES_LIST).strip('[]')+'\n')
        log('Fetched results from SalesForce\n')
        mailbox()
        log('Connected to the Mail Server!\n')
        break
    else:
        is_connected()

read_a()
d = 0

for i in range(len(CASES_DATA)):
    log(CASES_DATA[i][0] + ': ' + str(len(CASES_DATA[i][2]))+'\n')

# Main Loop
while True:
    d += 1
    if d == 10:
        sf()
        #log('SF info retreived: Total Cases: '+str(len(CASES_LIST))+'\n')
        log('SF info retreived.\n')
        d = 0
    try:
        if is_connected():
            id()
            if a != b:
                if debug():
                    debug_log('Loop Values: a='+str(a)+' b='+str(b)+'\n')     # Debugging
                fetchmail()
                write_b()
                a = b
            time.sleep(30)
        else:
            log('Error: Internet Connection down, Retrying after 60 seconds\n')
            continue
            time.sleep(60)
    except KeyboardInterrupt:
        mail.logout()
        log('Process Killed by Keyboard\n')
        sys.exit()
    except:
        traceback.print_exc()
        sys.exit()
