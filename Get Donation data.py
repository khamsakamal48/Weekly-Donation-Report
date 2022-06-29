#!/usr/bin/env python3

import requests, os, json, glob, csv, sys, smtplib, ssl, imaplib, time, email, re, fuzzywuzzy, itertools, datetime, logging, locale, pretty_html_table, calendar
import pandas as pd
from smtplib import SMTP
from datetime import datetime
from datetime import date
from datetime import timedelta
from requests.adapters import HTTPAdapter
from urllib3 import Retry
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from jinja2 import Environment
from pretty_html_table import build_table

# Printing the output to file for debugging
sys.stdout = open('Process.log', 'w')

# API Request strategy
print("Setting API Request strategy")
retry_strategy = Retry(
total=3,
status_forcelist=[429, 500, 502, 503, 504],
allowed_methods=["HEAD", "GET", "OPTIONS"],
backoff_factor=10
)
adapter = HTTPAdapter(max_retries=retry_strategy)
http = requests.Session()
http.mount("https://", adapter)
http.mount("http://", adapter)

# Set current directory
print("Setting current directory")
os.chdir(os.getcwd())

# Setting Locale
print("Setting Locale")
locale.setlocale(locale.LC_ALL, 'en_IN.UTF-8')
# print(locale.setlocale(locale.LC_ALL, 'en_IN.ISO8859-1'))
# locale.setlocale(locale.LC_ALL, 'en_IN.ISO8859-1')

from dotenv import load_dotenv

load_dotenv()

# Retrieve contents from .env file
RE_API_KEY = os.getenv("RE_API_KEY")
MAIL_USERN = os.getenv("MAIL_USERN")
MAIL_PASSWORD = os.getenv("MAIL_PASSWORD")
IMAP_URL = os.getenv("IMAP_URL")
IMAP_PORT = os.getenv("IMAP_PORT")
SMTP_URL = os.getenv("SMTP_URL")
SMTP_PORT = os.getenv("SMTP_PORT")
SEND_TO  = os.getenv("SEND_TO")

# Retrieve access_token from file
print("Retrieve token from API connections")
with open('access_token_output.json') as access_token_output:
    data = json.load(access_token_output)
    access_token = data["access_token"]

def housekeeping():
    # Housekeeping
    print("Deleting old JSON files.....if present")
    fileList = glob.glob('Gift_List_in_RE_*.json')

    # Iterate over the list of filepaths & remove each file.
    for filePath in fileList:
        try:
            os.remove(filePath)
        except:
            pass
    
def get_request_re():
    print("Running GET Request from RE function")
    time.sleep(5)
    # Request Headers for Blackbaud API request
    headers = {
    # Request headers
    'Bb-Api-Subscription-Key': RE_API_KEY,
    'Authorization': 'Bearer ' + access_token,
    }
    
    global re_api_response
    re_api_response = http.get(url, params=params, headers=headers).json()
    
    check_for_errors()
    
def check_for_errors():
    print("Checking for errors")
    error_keywords = ["invalid", "error", "bad", "Unauthorized", "Forbidden", "Not Found", "Unsupported Media Type", "Too Many Requests", "Internal Server Error", "Service Unavailable", "Unexpected", "error_code", "400"]
    
    if any(x in re_api_response for x in error_keywords):
        # Send emails
        print ("Will send email now")
        send_error_emails()
        
def attach_file_to_email(message, filename):
    # Open the attachment file for reading in binary mode, and make it a MIMEApplication class
    with open(filename, "rb") as f:
        file_attachment = MIMEApplication(f.read())
    # Add header/name to the attachments    
    file_attachment.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    # Attach the file to the message
    message.attach(file_attachment)
    
def send_error_emails():
    print("Sending email for an error")
    
    # Close writing to Process.log
    sys.stdout.close()
    
    message = MIMEMultipart()
    message["Subject"] = subject
    message["From"] = MAIL_USERN
    message["To"] = SEND_TO

    # Adding Reply-to header
    message.add_header('reply-to', MAIL_USERN)
        
    TEMPLATE="""
    <table style="background-color: #ffffff; border-color: #ffffff; width: auto; margin-left: auto; margin-right: auto;">
    <tbody>
    <tr style="height: 127px;">
    <td style="background-color: #363636; width: 100%; text-align: center; vertical-align: middle; height: 127px;">&nbsp;
    <h1><span style="color: #ffffff;">&nbsp;Raiser's Edge Automation: {{job_name}} Failed</span>&nbsp;</h1>
    </td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; background-color: #ffffff; border-color: #ffffff;">&nbsp;</td>
    </tr>
    <tr style="height: 18px;">
    <td style="width: 100%; height: 18px; background-color: #ffffff; border-color: #ffffff; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #455362;">This is to notify you that execution of Auto-updating Alumni records has failed.</span>&nbsp;</td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; background-color: #ffffff; border-color: #ffffff;">&nbsp;</td>
    </tr>
    <tr style="height: 61px;">
    <td style="width: 100%; background-color: #2f2f2f; height: 61px; text-align: center; vertical-align: middle;">
    <h2><span style="color: #ffffff;">Job details:</span></h2>
    </td>
    </tr>
    <tr style="height: 52px;">
    <td style="height: 52px;">
    <table style="background-color: #2f2f2f; width: 100%; margin-left: auto; margin-right: auto; height: 42px;">
    <tbody>
    <tr>
    <td style="width: 50%; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #ffffff;">Job :</span>&nbsp;</td>
    <td style="background-color: #ff8e2d; width: 50%; text-align: center; vertical-align: middle;">&nbsp;{{job_name}}&nbsp;</td>
    </tr>
    <tr>
    <td style="width: 50%; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #ffffff;">Failed on :</span>&nbsp;</td>
    <td style="background-color: #ff8e2d; width: 50%; text-align: center; vertical-align: middle;">&nbsp;{{current_time}}&nbsp;</td>
    </tr>
    </tbody>
    </table>
    </td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; background-color: #ffffff;">&nbsp;</td>
    </tr>
    <tr style="height: 18px;">
    <td style="height: 18px; width: 100%; background-color: #ffffff; text-align: center; vertical-align: middle;">Below is the detailed error log,</td>
    </tr>
    <tr style="height: 217.34375px;">
    <td style="height: 217.34375px; background-color: #f8f9f9; width: 100%; text-align: left; vertical-align: middle;">{{error_log_message}}</td>
    </tr>
    </tbody>
    </table>
    """
    
    # Create a text/html message from a rendered template
    emailbody = MIMEText(
        Environment().from_string(TEMPLATE).render(
            job_name = "Syncing Raisers Edge and AlmaBase",
            current_time=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            error_log_message = Argument
        ), "html"
    )
    
    # Add HTML parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(emailbody)
    attach_file_to_email(message, 'Process.log')
    emailcontent = message.as_string()
    
    # Create secure connection with server and send email
    context = ssl._create_unverified_context()
    with smtplib.SMTP_SSL(SMTP_URL, SMTP_PORT, context=context) as server:
        server.login(MAIL_USERN, MAIL_PASSWORD)
        server.sendmail(
            MAIL_USERN, SEND_TO, emailcontent
        )

    # Save copy of the sent email to sent items folder
    with imaplib.IMAP4_SSL(IMAP_URL, IMAP_PORT) as imap:
        imap.login(MAIL_USERN, MAIL_PASSWORD)
        imap.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), emailcontent.encode('utf8'))
        imap.logout()
        
    exit()
    
def print_json(d):
    print(json.dumps(d, indent=4))
    
def pagination_api_request():
    # Pagination request to retreive list
    global url
    while url:
        # Blackbaud API GET request
        get_request_re()

        # Incremental File name
        i = 1
        while os.path.exists("Gift_List_in_RE_%s.json" % i):
            i += 1
        with open("Gift_List_in_RE_%s.json" % i, "w") as list_output:
            json.dump(re_api_response, list_output,ensure_ascii=False, sort_keys=True, indent=4)
        
        # Check if a variable is present in file
        with open("Gift_List_in_RE_%s.json" % i) as list_output_last:
            if 'next_link' in list_output_last.read():
                url = re_api_response["next_link"]
            else:
                break
    
    # Get a list of all the file paths that ends with wildcard from in specified directory
    global fileList
    fileList = glob.glob('Gift_List_in_RE_*.json')

def get_ytd_donation():
    print("Getting YTD Gift list from Raisers Edge")
    
    current_month = datetime.now().strftime("%m")
    print("Current month: " + str(current_month))
    
    global current_year
    current_year = datetime.now().strftime("%Y")
    print("Current year: " + str(current_year))
    
    # Determining the current Financial Year to determine Start date
    print("Determining the current Financial Year to determine Start date")
    
    global financial_year
    if int(current_month) <= 4:
        financial_year = int(datetime.now().strftime("%Y")) - 1
    else:
        financial_year = current_year
    
    print("Financial Year: " + str(financial_year))
    
    start_gift_date = "04-01-" + str(financial_year)
    print("Start Gift date: " + start_gift_date)
    
    global url, params
    url = "https://api.sky.blackbaud.com/gift/v1/gifts?gift_type=Donation&start_gift_date=%s&gift_type=MatchingGiftPayment&gift_type=PledgePayment&gift_type=RecurringGiftPayment" % start_gift_date
    
    params = {}
    
    pagination_api_request()
    
    # Getting YTD donation amount from JSON
    print("Getting YTD donation amount from JSON")

    ytd_donation_amount_list = []
    
    for each_file in fileList:
        with open(each_file) as json_file:
            data = json.load(json_file)
            for each_value in data['value']:
                amount = each_value['amount']['value']
                ytd_donation_amount_list.append(int(amount))
                
    print("Donation list in RE: " + str(ytd_donation_amount_list))
    
    ytd_donation_amount = sum(ytd_donation_amount_list)

    ytd_donation_amount_in_inr = locale.currency(ytd_donation_amount, grouping=True)
    
    print("YTD Donation in RE: " + str(ytd_donation_amount_in_inr))

def get_weekly_donation():
    print("Getting Weekly Gift list from Raisers Edge")
    
    end_gift_date = date.today().strftime("%m-%d-%Y")
    
    today = date.today()
    
    start_gift_date = (today - timedelta(days=7)).strftime("%m-%d-%Y")
    
    print(start_gift_date)
    print(end_gift_date)
    
    global url, params
    url = "https://api.sky.blackbaud.com/gift/v1/gifts?gift_type=Donation&start_gift_date=%s&end_gift_date=%s&gift_type=MatchingGiftPayment&gift_type=PledgePayment&gift_type=RecurringGiftPayment&sort=date" % (start_gift_date, end_gift_date)
    
    pagination_api_request()
    
    # Getting YTD donation amount from JSON
    print("Getting weekly donation amount from JSON")
    
    weekly_donation_date_list = []
    weekly_donation_amount_list = []
    weekly_donation_donor_list = []
    weekly_donation_campaign_list = []
    
    for each_file in fileList:
        with open(each_file) as json_file:
            data = json.load(json_file)
            for each_value in data['value']:
                
                amount = each_value['amount']['value']
                weekly_donation_amount_list.append(int(amount))
                
                constituent_id = each_value['constituent_id']
                weekly_donation_donor_list.append(int(constituent_id))
                
                donation_d = each_value['date'].replace("T00:00:00", "")
                donation_date = datetime.strptime(donation_d, '%Y-%m-%d')
                weekly_donation_date_list.append(donation_date.strftime("%d-%m-%Y"))
                
                campaign_id = each_value['gift_splits'][0]['campaign_id']
                weekly_donation_campaign_list.append(int(campaign_id))
                
    print(weekly_donation_date_list)
    print(weekly_donation_amount_list)
    print(weekly_donation_donor_list)
    print(weekly_donation_campaign_list)
    
    weekly_donation_donor_name_list_individual = []
    weekly_donation_donor_name_list_organisation = []
    
    for each_id in weekly_donation_donor_list:
        url = "https://api.sky.blackbaud.com/constituent/v1/constituents/%s" % each_id
        params = {}
        
        get_request_re()
        
        if re_api_response['type'] == "Individual":
            name = re_api_response['first'] + " " + re_api_response['last']
            weekly_donation_donor_name_list_individual.append(name)
            weekly_donation_donor_name_list_organisation.append(" ")
        else:
            name = re_api_response['name']
            weekly_donation_donor_name_list_organisation.append(name)
            weekly_donation_donor_name_list_individual.append(" ")
            
    print(weekly_donation_donor_name_list_individual)
    print(weekly_donation_donor_name_list_organisation)
    
    weekly_donation_campaign_name = []
    
    for each_id in weekly_donation_campaign_list:
        url = "https://api.sky.blackbaud.com/fundraising/v1/campaigns/%s" % each_id
        params = {}
        
        get_request_re()
        
        weekly_donation_campaign_name.append(re_api_response['description'])
        
    print(weekly_donation_campaign_name)
    
    global re_donation
    re_donation = {
        'Date of Credit': weekly_donation_date_list,
        'Amount': weekly_donation_amount_list,
        'Name of Donor': weekly_donation_donor_name_list_individual,
        'Name of Company': weekly_donation_donor_name_list_organisation,
        'Purpose/ Project Description': weekly_donation_campaign_name
    }
    
    print_json(re_donation)
    
def prepare_weekly_report_table():
    """
    DONATION data
    :return:
    """
    data = pd.DataFrame(re_donation)
    return data

def send_email():
    print("Sending email...")
    
    message = MIMEMultipart()
    message["Subject"] = "Donation Summary"
    message["From"] = MAIL_USERN
    message["To"] = SEND_TO

    # Adding Reply-to header
    message.add_header('reply-to', MAIL_USERN)
        
    start = """<!DOCTYPE html>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta http-equiv="X-UA-Compatible" content="IE=edge" />
<style type="text/css">
/* CLIENT-SPECIFIC STYLES */
body, table, td, a { -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }
table, td { mso-table-lspace: 0pt; mso-table-rspace: 0pt; }
img { -ms-interpolation-mode: bicubic; }

/* RESET STYLES */
img { border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; }
table { border-collapse: collapse !important; }
body { height: 100% !important; margin: 0 !important; padding: 0 !important; width: 100% !important; }

/* iOS BLUE LINKS */
a[x-apple-data-detectors] {
    color: inherit !important;
    text-decoration: none !important;
    font-size: inherit !important;
    font-family: inherit !important;
    font-weight: inherit !important;
    line-height: inherit !important;
}

/* MEDIA QUERIES */
@media screen and (max-width: 480px) {
    .mobile-hide {
        display: none !important;
    }
    .mobile-center {
        text-align: center !important;
    }
}

/* ANDROID CENTER FIX */
div[style*="margin: 16px 0;"] { margin: 0 !important; }
</style>
<body style="margin: 0 !important; padding: 0 !important; background-color: #eeeeee;" bgcolor="#eeeeee">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
        <td align="center" style="background-color: #eeeeee;" bgcolor="#eeeeee">
        <!--[if (gte mso 9)|(IE)]>
        <table align="center" border="0" cellspacing="0" cellpadding="0" width="1200">
        <tr>
        <td align="center" valign="top" width="1200">
        <![endif]-->
        <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:1200px;">
            <tr>
                <td align="center" valign="center" style="font-size:0; padding: 10px;" bgcolor="#00B2EA">
                <!--[if (gte mso 9)|(IE)]>
                <table align="center" border="0" cellspacing="0" cellpadding="0" width="1200">
                <tr>
                <td align="left" valign="top" width="300">
                <![endif]-->
                <!--[if (gte mso 9)|(IE)]>
                </td>
                <td align="right" width="300">
                <![endif]-->
                <div style="display:inline-block; max-width: auto; min-width: auto; vertical-align:top; width:auto;">
                    <table align="center" border="0" cellpadding="10" cellspacing="0" width="100%" style="max-width:300px;">
                        <tr>
                            <td style="font-family: Open Sans, Helvetica, Arial, sans-serif; font-size: 18px; font-weight: 400; line-height: 24px;">
                                <a href="http://www.iitb.ac.in" target="_blank" style="color: #ffffff; text-decoration: none;"><img src="https://i.ibb.co/fk6J37P/iitblogowhite.png" width="85" height="85" style="display: block; border: 0px;"/></a>
                            </td>
                        </tr>
                    </table>
                </div>
                <!--[if (gte mso 9)|(IE)]>
                </td>
                </tr>
                </table>
                <![endif]-->
                </td>
            </tr>"""
    end = """<tr>
                <td align="center" style="padding: 5px; background-color: #ffffff;" bgcolor="#ffffff">
                <!--[if (gte mso 9)|(IE)]>
                <table align="center" border="0" cellspacing="0" cellpadding="0" width="1200">
                <tr>
                <td align="center" valign="top" width="600">
                <![endif]-->
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:1200px;">
                    <tr>
                        <td align="center" style="font-family: Open Sans, Helvetica, Arial, sans-serif; font-size: 14px; font-weight: 400; line-height: 24px; padding: 5px 0 10px 0;">
                            <p style="font-size: 14px; font-weight: 800; line-height: 18px; color: #333333;">
                                Indian Institute of Technology Bombay<br>
                                Powai, Mumbai 400 076, Maharashtra, India
                            </p>
                        </td>
                    </tr>
                    <tr>
                    </tr>
                </table>
                <!--[if (gte mso 9)|(IE)]>
                </td>
                </tr>
                </table>
                <![endif]-->
                </td>
            </tr>
        </table>
        <!--[if (gte mso 9)|(IE)]>
        </td>
        </tr>
        </table>
        <![endif]-->
        </td>
    </tr>
</table>
</body>
</html>"""
            
    emailbody = start + weekly_report_output + end
    
    print(weekly_report_output)
    
    # Add HTML parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(MIMEText(emailbody, "html"))
    emailcontent = message.as_string()
    
    # Create secure connection with server and send email
    context = ssl._create_unverified_context()
    with smtplib.SMTP_SSL(SMTP_URL, SMTP_PORT, context=context) as server:
        server.login(MAIL_USERN, MAIL_PASSWORD)
        server.sendmail(
            MAIL_USERN, SEND_TO, emailcontent
        )

    # Save copy of the sent email to sent items folder
    with imaplib.IMAP4_SSL(IMAP_URL, IMAP_PORT) as imap:
        imap.login(MAIL_USERN, MAIL_PASSWORD)
        imap.append('Sent', '\\Seen', imaplib.Time2Internaldate(time.time()), emailcontent.encode('utf8'))
        imap.logout()

def prepare_report():
    weekly_report = prepare_weekly_report_table()
    
    global weekly_report_output
    weekly_report_output = (build_table(weekly_report, 'blue_dark', font_family='Open Sans, Helvetica, Arial, sans-serif', even_color='black', padding='10px')).replace("background-color: #D9E1F2;font-family: Open Sans", "background-color: #D9E1F2; color: black;font-family: Open Sans")
    send_email()

def get_monthly_donation():
    print("Getting Monthly Gift list from Raisers Edge")
    
    # April
    apr_start_date = "04-01-" + str(financial_year)
    apr_end_date = "04-30-" + str(financial_year)
    
    # May
    may_start_date = "05-01-" + str(financial_year)
    may_end_date = "05-31-" + str(financial_year)
    
    # June
    jun_start_date = "06-01-" + str(financial_year)
    jun_end_date = "06-30-" + str(financial_year)
    
    # July
    jul_start_date = "07-01-" + str(financial_year)
    jul_end_date = "07-31-" + str(financial_year)
    
    # August
    aug_start_date = "08-01-" + str(financial_year)
    aug_end_date = "08-31-" + str(financial_year)
    
    # September
    sep_start_date = "09-01-" + str(financial_year)
    sep_end_date = "09-30-" + str(financial_year)
    
    # October
    oct_start_date = "10-01-" + str(financial_year)
    oct_end_date = "10-31-" + str(financial_year)
    
    # November
    nov_start_date = "11-01-" + str(financial_year)
    nov_end_date = "11-30-" + str(financial_year)
    
    # December
    dec_start_date = "12-01-" + str(financial_year)
    dec_end_date = "12-31-" + str(financial_year)
    
    next_year = int(financial_year) + 1
    
    # January
    jan_start_date = "01-01-" + str(next_year)
    jan_end_date = "01-31-" + str(next_year)
    
    # February
    feb_start_date = "02-01-" + str(next_year)
    
    if ((next_year % 400 == 0) or (next_year % 100 != 0) and (next_year % 4 == 0)):
        # next year is a leap year
        feb_end_date = "02-29-" + str(next_year)
    else:
        # next year is not a leap year
        feb_end_date = "02-28-" + str(next_year)
    
    # March
    mar_start_date = "03-01-" + str(next_year)
    mar_end_date = "03-31-" + str(next_year)
    
    start_dates = [apr_start_date, may_start_date, jun_start_date, jul_start_date, aug_start_date, sep_start_date, oct_start_date, nov_start_date, dec_start_date, jan_start_date, feb_start_date, mar_start_date]
    end_dates = [apr_end_date, may_end_date, jun_end_date, jul_end_date, aug_end_date, sep_end_date, oct_end_date, nov_end_date, dec_end_date, jan_end_date, feb_end_date, mar_end_date]
    
    print(start_dates)
    print(end_dates)
    
    monthly_donation_amount_list = []
    donation_amount = 0
    # for start_gift_date, end_gift_date in itertools.product(start_dates, end_dates):
        
    #     start_ = datetime.strptime(start_gift_date, '%m-%d-%Y')
    #     start_date = datetime.date(start_)
    #     print(start_date)
    #     print(date.today())
        
    #     if start_date > date.today():
    #         break
    #     else:
    #         global url, params
    #         url = "https://api.sky.blackbaud.com/gift/v1/gifts?gift_type=Donation&start_gift_date=%s&end_gift_date=%s&gift_type=MatchingGiftPayment&gift_type=PledgePayment&gift_type=RecurringGiftPayment" % (start_gift_date, end_gift_date)
    #         print(url)
    #         params = {}
    #         pagination_api_request()
            
    #         donation = []
    #         for each_file in fileList:
    #             with open(each_file) as json_file:
    #                 data = json.load(json_file)
    #                 for each_value in data['value']:
    #                     amount = each_value['amount']['value']
    #                     donation.append(int(amount))
            
    #         donation_amount = sum(donation)
    
    for each_start_date in start_dates:
        start_date = datetime.date(datetime.strptime(each_start_date, '%m-%d-%Y'))
        print(start_date)
        print(date.today())
        
        if start_date > date.today():
            break
        else:
            for each_end_date in end_dates:
                global url, params
                url = "https://api.sky.blackbaud.com/gift/v1/gifts?gift_type=Donation&start_gift_date=%s&end_gift_date=%s&gift_type=MatchingGiftPayment&gift_type=PledgePayment&gift_type=RecurringGiftPayment" % (each_start_date, each_end_date)
                print(url)
                params = {}
                pagination_api_request()
                
                donation = []
                for each_file in fileList:
                    with open(each_file) as json_file:
                        data = json.load(json_file)
                        for each_value in data['value']:
                            amount = each_value['amount']['value']
                            donation.append(int(amount))
                            
                end_dates.remove(each_end_date)
                break
            
            donation_amount = sum(donation)
        
        monthly_donation_amount_list.append(donation_amount)
    
        # Get month name
        print("Getting month name")
        month = calendar.month_name[monthinteger]
        monthly_donation_month_list.append(month)
        
        # Incrementing the month
        monthinteger += 1
        
        if monthinteger > 12:
            monthinteger_new = monthinteger - 1
            monthinteger = monthinteger_new
    print(monthly_donation_amount_list)
            
    
    
try:
    housekeeping()
    
    get_ytd_donation()
    housekeeping()
    
    get_weekly_donation()
    housekeeping()
    
    get_monthly_donation()
    
    prepare_report()
    
    # Close writing to Process.log
    sys.stdout.close()
    
    # Exit
    sys.exit()
    
except Exception as Argument:
    print("Error while getting YTD Donation from Raisers Edge")
    subject = "Error while getting YTD Donation from Raisers Edge"
    send_error_emails()