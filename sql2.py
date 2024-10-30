import datetime
import smtplib
from email.mime.text import MIMEText
import json
import traceback
import time
import logging
import config2
import requests

logger = logging.getLogger(__name__)


def is_time_between(begin_time, end_time, check_time=None):
    check_time = check_time or datetime.datetime.now().time()
    if begin_time < end_time:
        return begin_time <= check_time <= end_time
    else:
        return check_time >= begin_time or check_time <= end_time


def handle_task(task_func, error_message):
    global email_text
    try:
        task_func()
    except Exception as e:
        tb = ''.join(traceback.format_tb(e.__traceback__))
        try:
            error = json.loads(str(e))['result']['message']
        except:
            error = e.with_traceback(e.__traceback__)
        email_text += f"{error_message}:\n\n{tb}{error}\n\n"


def send_email(subject, body, recipients, cc=[]):
    gmail_user = 'mmarcotte.grt@gmail.com'
    gmail_password = 'fvqm sdft qnao nsbl'

    msg = MIMEText(body)
    msg['Subject'] = subject
    msg['From'] = 'Granite Automation'
    msg['To'] = ', '.join(recipients)
    msg['Cc'] = ', '.join(cc)

    try:
        server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        server.ehlo()
        server.login(gmail_user, gmail_password)
        server.sendmail(gmail_user, recipients + cc, msg.as_string())
        server.close()
        print('Email sent:', body)
    except Exception as e:
        print('Unable to send email:', str(e))


def run_main_tasks_before_630pm():
    current_time = datetime.datetime.now().time()
    if current_time < datetime.time(18, 30):
        # These are the 5 main tasks that should run only before 6:30 PM
        handle_task(config2.MergedVoice, "MergedVoice Automation Failure")
        handle_task(config2.Edgeboot, "Edgeboot Automation Failure")
        handle_task(config2.SerialColumn, "SerialColumn Automation Failure")
        handle_task(config2.DIAAutoTemplate, "DIAAutoTemplate Automation Failure")
        handle_task(config2.UpdateTicketData, "UpdateTicketData Automation Failure")


def run_tracking_tasks():
    # Updated intervals to 15 minutes, removing any intervals after 6:30 PM
    if is_time_between(datetime.time(7, 29), datetime.time(7, 44)) or is_time_between(datetime.time(12, 29),
                                                                                      datetime.time(12, 44)):
        handle_task(lambda: config2.TrackingUpdate(
            sheet_id=5036994589577092, ticket_column='Equipment Ticket', tracking_column='Tracking Number',
            tracking_status_column='Shipping Status', del_date_column='Delivery Date'),
                    "TrackingUpdate Automation Failure")

    if is_time_between(datetime.time(9, 59), datetime.time(10, 14)) or is_time_between(datetime.time(14, 59),
                                                                                       datetime.time(15, 14)):
        handle_task(lambda: config2.TrackingUpdate(
            sheet_id=4591122715568004, ticket_column='Equipment Ticket/PWO', tracking_column='Tracking Number v2',
            tracking_status_column='UPS Status v2', del_date_column='Delivery Date v2'),
                    "EPIK Tracking Automation Failure")


def run_final_tracking_tasks_at_730pm():
    if is_time_between(datetime.time(19, 29), datetime.time(19, 44)):
        # Final run of the tracking tasks at 7:30 PM
        handle_task(lambda: config2.TrackingUpdate(
            sheet_id=5036994589577092, ticket_column='Equipment Ticket', tracking_column='Tracking Number',
            tracking_status_column='Shipping Status', del_date_column='Delivery Date'),
                    "TrackingUpdate Automation Failure")

        handle_task(lambda: config2.TrackingUpdate(
            sheet_id=4591122715568004, ticket_column='Equipment Ticket/PWO', tracking_column='Tracking Number v2',
            tracking_status_column='UPS Status v2', del_date_column='Delivery Date v2'),
                    "EPIK Tracking Automation Failure")


email_text = ''

# Run the main tasks before 6:30 PM
run_main_tasks_before_630pm()

# Run tracking tasks at the designated times
run_tracking_tasks()

# Run final tracking tasks at 7:30 PM
run_final_tracking_tasks_at_730pm()

if email_text:
    send_email('Automation Failure', email_text, ['mmarcotte@granitenet.com'], cc=['vmurray@granitenet.com'])

# Display affirmation message
try:
    print((lambda s: s[0].upper() + s[1:] + ('.' if not s.endswith('.') else ''))(
        requests.get("https://www.affirmations.dev/").json()['affirmation']))
except Exception as e:
    logger.error(f"Failed to retrieve affirmation: {str(e)}")

time.sleep(10)
