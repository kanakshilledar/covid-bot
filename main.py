import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tweepy
import time
import logging
import pywhatkit
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# starting logging procedure
logging.basicConfig(filename = 'main.log', filemode = 'w', format='%(asctime)s - %(message)s', level=logging.INFO)
logger = logging.getLogger()

# Google Sheets API authorization
SCOPE = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive',]
logger.info('Added SCOPE')
# your json file
creds = ServiceAccountCredentials.from_json_keyfile_name('CREDENTIAL.json', SCOPE)
client = gspread.authorize(creds)
logger.info('Google API Authorization Done')

# Twitter authorization
# add your keys
CONSUMER_KEY = 'CONSUMER API'
CONSUMER_SECRET = 'CONSUMER SECRET'
ACCESS_KEY = 'ACCESS KEY'
ACCESS_SECRET = 'ACCESS SECRET'

auth = tweepy.OAuthHandler(CONSUMER_KEY, CONSUMER_SECRET)
auth.set_access_token(ACCESS_KEY, ACCESS_SECRET)
api = tweepy.API(auth)
try:
    api.verify_credentials()
    logger.info('Twitter API Authorization Done')
except Exception as e:
    logger.info('Error During Authentication'.format(e))

# gmail authorization
# sender details
SENDER_ADDRESS = 'EMAIL ADDRESS'
SENDER_PASS = 'EMAIL PASSWORD'

def mailer(name, contact, requirement, location, email, link):
    message_content = '''
    Hello!
    Thank You for filling the form. Here\'s what we have received from you:
    \tName: {}
    \tContact: {}
    \tRequirement: {}
    \tLocation: {}
    \tEmail: {}
    This is your twitter link {}. Please keep monitoring the the replies to this tweet you may find something useful.
    
    Note: This is an auto generated email so please dont reply to this email. We can't guarantee you help but will try our best.
    Thank You for using our service
    Get Well Soon!
    Team Covid Bot
    '''

    # setup MIME
    message = MIMEMultipart()
    message['From'] = SENDER_ADDRESS
    message['To'] = email
    message['Subject'] = 'Link for your tweet.'
    message.attach(MIMEText(message_content.format(name, contact, requirement, location, email, link), 'plain'))
    # create smtp session
    logger.info('SMTP Session Created')
    session = smtplib.SMTP('smtp.gmail.com', 587)
    session.starttls()
    session.login(SENDER_ADDRESS, SENDER_PASS)
    text = message.as_string()
    session.sendmail(SENDER_ADDRESS, email, text)
    logger.info('Mail Sent')
    session.quit()
    curtime = list(time.localtime())
    # sending whatsapp message
    # pywhatkit.sendwhatmsg('+91' + str(contact), message_content.format(name, contact, requirement, location, email, link), time_hour=curtime[3], time_min=curtime[4] + 1, browser='firefox')
    logger.info('Message Sent')

# sending tweet
def tweeter(name, requirement, location):
    message = '''Name: {}
    Requirement: {}
    Location: {}'''
    api.update_status(message.format(name, requirement, location))
    logger.info('Tweet Sent')
    logger.info('Getting Tweet ID')
    posts = api.user_timeline()
    id = posts[0].id_str
    global link
    link = 'https://twitter.com/Covid19_helper/status/' + id
    logger.info('Link Generated - ' + link)


def responseData(response):
    # extracting data from the rows
    name = response[1]
    contact = response[2]
    requirement = response[3]
    location = response[4]
    email = response[5]
    if len(response) != 7:
        print(name, requirement, location)
        tweeter(name, requirement, location)
        sheet.update('G' + str(i), 'done')
        print('[+] tweeted')
        mailer(name, contact, requirement, location, email, link)
        # waiting 2 seconds
        time.sleep(2)
    
# Finding Workbook to read from
sheet = client.open("User Details (Responses)").sheet1
logger.info('Successfully Opened Responses Worksheet')

# variable to keep check of rows
i = 2

# working on the responses
logger.info('Started Working on Responses')
while True:
    try:
        response = sheet.row_values(i)
        print(response)
        responseData(response)
        i += 1
        logger.info('Response taken')
    
    except IndexError:
        logger.info('No new data available')
        logger.info('Halt for 60 seconds before retrying')
        time.sleep(60)
