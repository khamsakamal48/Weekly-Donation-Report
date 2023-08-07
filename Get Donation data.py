import numpy as np
import requests
import os
import json
import glob
import locale
import msal
import base64
import logging
import pandas as pd

from datetime import datetime
from datetime import timedelta
from requests.adapters import HTTPAdapter
from urllib3 import Retry
from pretty_html_table import build_table
from dotenv import load_dotenv

def set_current_directory():
    os.chdir(os.getcwd())


def start_logging():
    global process_name

    # Get File Name of existing script
    process_name = os.path.basename(__file__).replace('.py', '').replace(' ', '_')

    logging.basicConfig(filename=f'Logs/{process_name}.log', format='%(asctime)s %(message)s', filemode='w',
                        level=logging.DEBUG)

    # Printing the output to file for debugging
    logging.info('Starting the Script')


def stop_logging():
    logging.info('Stopping the Script')


def housekeeping():
    logging.info('Doing Housekeeping')

    # Housekeeping
    multiple_files = glob.glob('*_RE_*.json')

    # Iterate over the list of filepaths & remove each file.
    logging.info('Removing old JSON files')
    for each_file in multiple_files:
        try:
            os.remove(each_file)
        except:
            pass

def set_api_request_strategy():
    logging.info('Setting API Request strategy')

    global http

    # API Request strategy
    logging.info('Setting API Request Strategy')

    retry_strategy = Retry(
        total=3,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=['HEAD', 'GET', 'OPTIONS'],
        backoff_factor=10
    )

    adapter = HTTPAdapter(max_retries=retry_strategy)
    http = requests.Session()
    http.mount('https://', adapter)
    http.mount('http://', adapter)


def get_env_variables():
    logging.info('Setting Environment variables')

    global RE_API_KEY, CLIENT_ID, O_CLIENT_ID, CLIENT_SECRET, TENANT_ID, FROM, CC_TO, SEND_TO, ERROR_EMAILS_TO

    load_dotenv()

    RE_API_KEY = os.getenv('RE_API_KEY')
    CLIENT_ID = os.getenv('CLIENT_ID')
    O_CLIENT_ID = os.getenv('O_CLIENT_ID')
    CLIENT_SECRET = os.getenv('CLIENT_SECRET')
    TENANT_ID = os.getenv('TENANT_ID')
    FROM = os.getenv('FROM')
    SEND_TO = eval(os.getenv('SEND_TO'))
    CC_TO = eval(os.getenv('CC_TO'))
    ERROR_EMAILS_TO = os.getenv('ERROR_EMAILS_TO')

def send_error_emails(subject, arg):
    logging.info('Sending email for an error')

    authority = f'https://login.microsoftonline.com/{TENANT_ID}'

    app = msal.ConfidentialClientApplication(
        client_id=O_CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority
    )

    scopes = ["https://graph.microsoft.com/.default"]

    result = None
    result = app.acquire_token_silent(scopes, account=None)

    if not result:
        result = app.acquire_token_for_client(scopes=scopes)

        TEMPLATE = """
           <table style="background-color: #ffffff; border-color: #ffffff; width: auto; margin-left: auto; margin-right: auto;">
           <tbody>
           <tr style="height: 127px;">
           <td style="background-color: #363636; width: 100%; text-align: center; vertical-align: middle; height: 127px;">&nbsp;
           <h1><span style="color: #ffffff;">&nbsp;Raiser's Edge Automation: {job_name} Failed</span>&nbsp;</h1>
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
           <td style="background-color: #ff8e2d; width: 50%; text-align: center; vertical-align: middle;">&nbsp;{job_name}&nbsp;</td>
           </tr>
           <tr>
           <td style="width: 50%; text-align: center; vertical-align: middle;">&nbsp;<span style="color: #ffffff;">Failed on :</span>&nbsp;</td>
           <td style="background-color: #ff8e2d; width: 50%; text-align: center; vertical-align: middle;">&nbsp;{current_time}&nbsp;</td>
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
           <td style="height: 217.34375px; background-color: #f8f9f9; width: 100%; text-align: left; vertical-align: middle;">{error_log_message}</td>
           </tr>
           </tbody>
           </table>
           """

        # Create a text/html message from a rendered template
        emailbody = TEMPLATE.format(
            job_name="Getting Donation Summary from Raisers Edge",
            current_time=datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            error_log_message=arg
        )

        # Set up attachment data
        with open('Process.log', 'rb') as f:
            attachment_content = f.read()
        attachment_content = base64.b64encode(attachment_content).decode('utf-8')

        if "access_token" in result:

            endpoint = f'https://graph.microsoft.com/v1.0/users/{FROM}/sendMail'

            email_msg = {
                'Message': {
                    'Subject': subject,
                    'Body': {
                        'ContentType': 'HTML',
                        'Content': emailbody
                    },
                    'ToRecipients': [
                        {
                            'EmailAddress': {
                                'Address': ERROR_EMAILS_TO
                            }
                        }
                    ],
                    'Attachments': [
                        {
                            '@odata.type': '#microsoft.graph.fileAttachment',
                            'name': 'Process.log',
                            'contentBytes': attachment_content
                        }
                    ]
                },
                'SaveToSentItems': 'true'
            }

            requests.post(
                endpoint,
                headers={
                    'Authorization': 'Bearer ' + result['access_token']
                },
                json=email_msg
            )

        else:
            logging.info(result.get('error'))
            logging.info(result.get('error_description'))
            logging.info(result.get('correlation_id'))

def print_json(d):
    print(json.dumps(d, indent=4))


def retrieve_token():
    logging.info('Retrieve token for API connections')

    with open('access_token_output.json') as access_token_output:
        data = json.load(access_token_output)
        access_token = data['access_token']

    return access_token

def get_request_re(url, params):
    logging.info('Running GET Request from RE function')

    global re_api_response

    # Retrieve access_token from file
    access_token = retrieve_token()

    # Request headers
    headers = {
        'Bb-Api-Subscription-Key': RE_API_KEY,
        'Authorization': 'Bearer ' + access_token,
    }

    re_api_response = http.get(url, params=params, headers=headers).json()

    return re_api_response

def post_request_re(url, params):
    logging.info('Running POST Request to RE function')

    # Retrieve access_token from file
    access_token = retrieve_token()

    # Request headers
    headers = {
        'Bb-Api-Subscription-Key': RE_API_KEY,
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json',
    }

    re_api_response = http.post(url, params=params, headers=headers, json=params).json()

    return re_api_response

def patch_request_re(url, params):
    logging.info('Running PATCH Request to RE function')

    # Retrieve access_token from file
    access_token = retrieve_token()

    # Request headers
    headers = {
        'Bb-Api-Subscription-Key': RE_API_KEY,
        'Authorization': 'Bearer ' + access_token,
        'Content-Type': 'application/json'
    }

    re_api_response = http.patch(url, headers=headers, data=json.dumps(params))

    return re_api_response


def load_from_json_to_parquet():
    logging.info('Loading from JSON to Parquet file')

    # Get a list of all the file paths that ends with wildcard from in specified directory
    fileList = glob.glob('API_Response_RE_*.json')

    df = pd.DataFrame()

    for each_file in fileList:
        # Open Each JSON File
        with open(each_file, 'r', encoding='utf8') as json_file:
            # Load JSON File
            json_content = json.load(json_file)

            # Convert non-string values to strings
            json_content = convert_to_strings(json_content)

            # Load to a dataframe
            df_ = api_to_df(json_content)

            df = pd.concat([df, df_])

    housekeeping()

    return df

def convert_to_strings(obj):
    """
    Recursively convert non-string values to strings in a dictionary
    """
    if isinstance(obj, dict):
        return {k: convert_to_strings(v) for k, v in obj.items()}
    elif isinstance(obj, list):
        return [convert_to_strings(item) for item in obj]
    else:
        return str(obj)


def api_to_df(response):
    logging.info('Loading API response to a DataFrame')

    # Load from JSON to pandas
    try:
        api_response = pd.json_normalize(response['value'])
    except:
        try:
            api_response = pd.json_normalize(response)
        except:
            api_response = pd.json_normalize(response, 'value')

    # Load to a dataframe
    df = pd.DataFrame(data=api_response)

    return df

def pagination_api_request(url, params):
    # Pagination request to retreive list
    while url:
        # Blackbaud API GET request
        get_request_re(url, params)

        # Incremental File name
        i = 1
        while os.path.exists(f'API_Response_RE_{process_name}_{i}.json'):
            i += 1

        with open(f'API_Response_RE_{process_name}_{i}.json', 'w', encoding='utf-8') as list_output:
            json.dump(re_api_response, list_output, ensure_ascii=True, sort_keys=True, indent=4)

        # Check if a variable is present in file
        with open(f'API_Response_RE_{process_name}_{i}.json') as list_output_last:

            if 'next_link' in list_output_last.read():
                url = re_api_response['next_link']

            else:
                break

def set_locale():
    logging.info('Setting Locale')

    try:
        # Setting Locale
        locale.setlocale(locale.LC_ALL, 'en_IN.UTF-8')

    except:
        pass

def prepare_report(data):
    """
    DONATION data
    :return:
    """
    data = pd.DataFrame(data)

    report_output = (
        build_table(data, 'blue_dark', font_family='Open Sans, Helvetica, Arial, sans-serif', even_color='black',
                    padding='10px', width='700px', font_size='16px')).replace(
        "background-color: #D9E1F2;font-family: Open Sans",
        "background-color: #D9E1F2; color: black;font-family: Open Sans")

    return report_output

def get_timeline():
    logging.info('Identifying Current Year and Financial Year')

    current_month = datetime.now().strftime("%m")

    current_year = datetime.now().strftime("%Y")

    current_date = datetime.now().strftime("%d")

    # Determining the current Financial Year to determine Start date
    if int(current_month) <= 4:

        if int(current_month) == 4 and int(current_date) > 7:
            financial_year = current_year

        else:
            financial_year = int(datetime.now().strftime("%Y")) - 1
    else:
        financial_year = current_year

    start_gift_date = '01-Apr-' + str(financial_year)

    return current_date, current_month, current_year, financial_year, start_gift_date

def get_donation():
    logging.info('Getting all Gifts from Raisers Edge')

    url = f'https://api.sky.blackbaud.com/gift/v1/gifts?gift_type=Donation&gift_type=MatchingGiftPayment&gift_type=PledgePayment&gift_type=RecurringGiftPayment&gift_type=GiftInKind'
    params = {}

    pagination_api_request(url, params)

    data = load_from_json_to_parquet().copy()

    data.to_parquet('Database/Data.parquet', index=False)

    # Pre-process data
    data = process_data(data).copy()

    return data

def process_data(data):
    logging.info('Pre-process Donation data')

    # Amount
    data['amount.value'] = pd.to_numeric(data['amount.value'])

    # Dates
    data['date'] = pd.to_datetime(data['date'])
    data['date_added'] = pd.to_datetime(data['date_added'], format='ISO8601')
    data['receipt_date'] = data['receipts'].apply(lambda x: get_dates(x))
    data['receipt_date'] = data['receipt_date'].fillna(data['date_added'].astype(str) + 'T00:00:00')
    data['receipt_date'] = data['receipt_date'].apply(lambda x: str(x).split(' ')[0])
    data['receipt_date'] = pd.to_datetime(data['receipt_date'], format='mixed', dayfirst=False)

    # Campaign ID
    data['campaign_id'] = data['gift_splits'].apply(lambda x: x[0]['campaign_id'])

    return data

def get_dates(value):
    try:
        date = value[0]['date']
    except:
        date = np.NaN

    return date

def get_ytd_donation():
    logging.info('Getting YTD Gifts from Raisers Edge')

    # amount = re_donation[
    #     (re_donation['date'] >= start_gift_date) &
    #     (re_donation['receipt_date'] >= start_gift_date)
    # ]['amount.value'].sum()

    amount = re_donation[
        (re_donation['receipt_date'] >= start_gift_date)
        ]['amount.value'].sum()

    if len(str(round(amount))) >= 10:
        amount = locale.currency(round(amount), grouping=True)[:-3]
        amount = amount.replace(',', '', 1)
    else:
        amount = locale.currency(round(amount), grouping=True)[:-3]

    amount = {
        'Financial Year': "F.Y. " + str(financial_year) + ' - ' + str(int(str(financial_year)[-2:]) + 1),
        'Amount': [amount]
    }

    amount = prepare_report(amount)

    return amount

def get_previous_year_donations():
    logging.info('Getting YTD Gifts donated in the previous financial years from Raisers Edge')

    amount = re_donation[
        (re_donation['receipt_date'] >= start_gift_date) &
        (re_donation['date'] < start_gift_date)
    ]['amount.value'].sum()

    if len(str(round(amount))) >= 10:
        amount = locale.currency(round(amount), grouping=True)[:-3]
        amount = amount.replace(',', '', 1)
    else:
        amount = locale.currency(round(amount), grouping=True)[:-3]

    amount = {
        'Financial Year': 'Previous Financial Years',
        'Amount': [amount]
    }

    amount = prepare_report(amount)

    return amount

def get_monthly_donation():
    logging.info('Getting Monthly Gifts from Raisers Edge')

    # data = re_donation[
    #     (re_donation['receipt_date'] >= start_gift_date) &
    #     (re_donation['date'] >= start_gift_date)
    #     ][['date', 'amount.value']].reset_index(drop=True).copy()

    data = re_donation[
        (re_donation['receipt_date'] >= start_gift_date)
        ][['receipt_date', 'amount.value']].reset_index(drop=True).copy()

    data['receipt_date'] = pd.to_datetime(data['receipt_date'], format='%d-%b-%Y')
    data['receipt_date'] = data['receipt_date'].dt.strftime('%B')

    data = data.rename(columns={
        'receipt_date': 'Month',
        'amount.value': 'Amount'
    }).copy()

    months = [
        'January',
        'February',
        'March',
        'April',
        'May',
        'June',
        'July',
        'August',
        'September',
        'October',
        'November',
        'December'
    ]

    data['Month'] = pd.Categorical(data['Month'], categories=months, ordered=True)

    data = data.groupby('Month').agg({'Amount': 'sum'}).dropna().reset_index().copy()

    data = data[data['Amount'] > 0].reset_index(drop=True).copy()

    data['Amount'] = data['Amount'].apply(lambda x: locale.currency(round(x), grouping=True)[:-3])

    data = prepare_report(data)

    return data

def get_weekly_donation():
    logging.info('Getting Weekly Gift list from Raisers Edge')

    # data = re_donation[
    #     (re_donation['receipt_date'] >= (pd.to_datetime('today') - timedelta(days=7)).strftime('%d-%b-%Y')) &
    #     (re_donation['receipt_date'] >= start_gift_date) &
    #     (re_donation['date'] >= start_gift_date)
    # ][[
    #     'date', 'amount.value', 'constituent_id', 'campaign_id'
    # ]].reset_index(drop=True).copy()

    data = re_donation[
        (re_donation['receipt_date'] >= (pd.to_datetime('today') - timedelta(days=7)).strftime('%d-%b-%Y')) &
        (re_donation['receipt_date'] >= start_gift_date)
        ][[
        'receipt_date', 'amount.value', 'constituent_id', 'campaign_id'
    ]].reset_index(drop=True).copy()

    data = data.sort_values(['receipt_date'], ascending=False).copy()
    data['receipt_date'] = data['receipt_date'].dt.strftime('%d-%b-%Y')

    data['Name of Donor'] = data['constituent_id'].apply(lambda x: get_donor_name(x))
    data['Purpose/ Project Description'] = data['campaign_id'].apply(lambda x: get_project(x))

    data = data.rename(columns={
        'receipt_date': 'Date of Credit',
        'amount.value': 'Amount'
    }).copy()

    data = data.drop(columns=['constituent_id', 'campaign_id']).copy()

    data['Amount'] = data['Amount'].apply(lambda x: locale.currency(round(x), grouping=True)[:-3])

    data = prepare_report(data)

    return data

def get_donor_name(id):
    logging.info('Getting Donor Name')

    url = f'https://api.sky.blackbaud.com/constituent/v1/constituents/{id}'
    params = {}

    response = get_request_re(url, params)

    if response['type'] == 'Individual':
        name = response['first'] + ' ' + response['last']
    else:
        name = response['name']

    return name

def get_project(id):
    logging.info('Getting Project Name')

    url = f'https://api.sky.blackbaud.com/fundraising/v1/campaigns/{id}'
    params = {}

    response = get_request_re(url, params)

    return response['description']

def send_email():
    logging.info('Sending email')

    subject = 'Donation Summary | Raisers Edge'

    authority = f'https://login.microsoftonline.com/{TENANT_ID}'

    app = msal.ConfidentialClientApplication(
        client_id=O_CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=authority
    )

    scopes = ['https://graph.microsoft.com/.default']

    result = None
    result = app.acquire_token_silent(scopes, account=None)

    if not result:
        result = app.acquire_token_for_client(scopes=scopes)

    start = '''
    <!DOCTYPE html>
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

    <!-- HIDDEN PREHEADER TEXT -->
    <div style="display: none; font-size: 1px; color: #fefefe; line-height: 1px; font-family: Open Sans, Helvetica, Arial, sans-serif; max-height: 0px; max-width: 0px; opacity: 0; overflow: hidden;">
    Dear Team,

    Below are the details of the Donations as recorded in Raisers Edge.
    </div>

    <table border="0" cellpadding="0" cellspacing="0" width="100%">
        <tr>
            <td align="center" style="background-color: #eeeeee;" bgcolor="#eeeeee">
            <!--[if (gte mso 9)|(IE)]>
            <table align="center" border="0" cellspacing="0" cellpadding="0" width="700">
            <tr>
            <td align="center" valign="top" width="700">
            <![endif]-->
            <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:700px;">
                <tr>
                    <td align="center" style="padding:0;Margin:0">
                    <table class="es-header-body" cellspacing="0" cellpadding="0" bgcolor="#ffffff" align="center" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px;background-color:#FFFFFF;width:700px">
                    <tr>
                        <td align="left" bgcolor="#305496" style="padding:0;Margin:0;padding-top:20px;padding-left:20px;padding-right:20px;background-color:#305496">
                        <table cellpadding="0" cellspacing="0" width="100%" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                        <tr>
                            <td align="center" valign="top" style="padding:0;Margin:0;width:700px">
                            <table cellpadding="0" cellspacing="0" width="100%" role="presentation" style="mso-table-lspace:0pt;mso-table-rspace:0pt;border-collapse:collapse;border-spacing:0px">
                            <tr>
                                <td align="center" style="padding:0;Margin:0;padding-bottom:20px;font-size:0px"><img src="https://i.ibb.co/fk6J37P/iitblogowhite.png" alt style="display:block;border:0;outline:none;text-decoration:none;-ms-interpolation-mode:bicubic" width="85" height="85"></td>
                            </tr>
                            </table></td>
                        </tr>
                        </table></td>
                    </tr>
                    </table></td>
                </tr>
                <tr>
                    <td align="justify" style="padding: 35px; background-color: #ffffff;" bgcolor="#ffffff">
                    <!--[if (gte mso 9)|(IE)]>
                    <table align="justify" border="0" cellspacing="0" cellpadding="0" width="700">
                    <tr>
                    <td align="justify" valign="top" width="700">
                    <![endif]-->
                    <table align="justify" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:700px;">
                        <tr>
                            <td align="justify" style="font-family: Open Sans, Helvetica, Arial, sans-serif; font-size: 16px; font-weight: 400; line-height: 24px; padding-top: 0px; padding-bottom: 10px; solid #eeeeee; ">
                                <p align="justify" style="font-size: 16px; font-weight: 400; line-height: 24px; color: #333333;">
                                    Dear Team,<br><br>
                                    Below are the details of the Donations as recorded in Raisers Edge.
                                </p>
                                <p align="center" style="font-size: 32px; font-weight: 800; line-height: 24px; color: #333333; padding-top: 10px;">
                                    YTD Summary
                                </p>
                                <p align="center" style="font-size: 16px; font-weight: 400; line-height: 24px; color: #333333;">
                                    <b>Gifts received in current financial year</b>
                                    <br>
                                </p>
    '''

    tr_date = '''
    <p align="center" style="font-size: 16px; font-weight: 400; line-height: 24px; color: #333333;">
        <br>
        <b>Gifts received in current financial year with gift date in previous financial years</b>
        <br>
    </p>
    '''

    month = '''
    <p align="center" style="font-size: 32px; font-weight: 800; line-height: 24px; color: #333333; padding-top: 10px;">
                            <br>
                            Monthwise Summary
                        </p>
    '''

    weekly = '''
    <p align="center" style="font-size: 32px; font-weight: 800; line-height: 24px; color: #333333; padding-top: 10px;">
                            <br>
                            Weekly Summary
                        </p>
    '''

    end = '''
    <p align="left" style="font-size: 16px; font-weight: 200; line-height: 24px; color: #333333;">
                                <br>
                                The weekly summary is for a period of <b>start_date</b> to <b>end_date</b>.<br><br>
                            </p>
                        </td>
                    </tr>
                    </tr>
                </table>
                <!--[if (gte mso 9)|(IE)]>
                </td>
                </tr>
                </table>
                <![endif]-->
                </td>
            </tr>

            <tr>
                <td align="center" style=" padding: 10px; background-color: #305496;" bgcolor="#305496">
                <!--[if (gte mso 9)|(IE)]>
                <table align="center" border="0" cellspacing="0" cellpadding="0" width="700">
                <tr>
                <td align="center" valign="top" width="700">
                <![endif]-->
                <table align="center" border="0" cellpadding="10" cellspacing="10" width="100%" style="max-width:700px; padding-top: 10px;">
                <td align="center" style="font-family: Open Sans, Helvetica, Arial, sans-serif; font-size: 14px; font-weight: 400; line-height: 24px; padding: 0px 0px 0px 0px;">
                    <p style="font-size: 32px; font-weight: 600; line-height: 24px; color: #ffffff;">
                    ज्ञानम् परमम् ध्येयम्
                    </p>
                </td>
                </table>
                <!--[if (gte mso 9)|(IE)]>
                </td>
                </tr>
                </table>
                <![endif]-->
                </td>
            </tr>
            <tr>
                <td align="center" style="padding: 5px; background-color: #ffffff;" bgcolor="#ffffff">
                <!--[if (gte mso 9)|(IE)]>
                <table align="center" border="0" cellspacing="0" cellpadding="0" width="700">
                <tr>
                <td align="center" valign="top" width="700">
                <![endif]-->
                <table align="center" border="0" cellpadding="0" cellspacing="0" width="100%" style="max-width:700px;">
                    <tr>
                        <td align="center" style="font-family: Open Sans, Helvetica, Arial, sans-serif; font-size: 14px; font-weight: 400; line-height: 24px; padding: 5px 0 10px 0;">
                            <p style="font-size: 14px; font-weight: 400; line-height: 18px; color: #808080;">
                                This is a system generated email
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
        </html>
    '''

    emailbody = start + ytd_report + tr_date + previous_ytd_report + month + monthly_donation + weekly + weekly_donation + end
    emailbody = emailbody.replace('start_date', (pd.to_datetime('today') - timedelta(days=7)).strftime('%d %b, %Y')).replace('end_date', pd.to_datetime('today').strftime('%d %b, %Y'))

    if "access_token" in result:
        endpoint = f'https://graph.microsoft.com/v1.0/users/{FROM}/sendMail'

        email_msg = {
            'Message': {
                'Subject': subject,
                'Body': {
                    'ContentType': 'HTML',
                    'Content': emailbody
                },
                'ToRecipients': get_recipients(SEND_TO),
                'ccRecipients': get_recipients(CC_TO),
            },
            'SaveToSentItems': 'true'
        }

        r = requests.post(
            endpoint,
            headers={
                'Authorization': 'Bearer ' + result['access_token']
            },
            json=email_msg
        )

        if r.ok:
            print('Sent email successfully')
        else:
            print(r.json())

    else:
        print(result.get('error'))
        print(result.get('error_description'))
        print(result.get('correlation_id'))

def get_recipients(email_list):
    value = []

    for email in email_list:
        email = {
            'emailAddress': {
                'address': email
            }
        }

        value.append(email)

    return value

try:
    # Set current directory
    set_current_directory()

    # Start Logging for Debugging
    start_logging()

    # Retrieve contents from .env file
    get_env_variables()

    # Housekeeping
    housekeeping()

    # Set API Request strategy
    set_api_request_strategy()

    # Set Locale
    set_locale()

    # Get Years
    current_date, current_month, current_year, financial_year, start_gift_date = get_timeline()

    # Get the complete Donation
    re_donation = get_donation().copy()
    re_donation.to_parquet('Database/RE Donations.parquet', index=False)

    # Format the dates
    re_donation['receipt_date'] = pd.to_datetime(re_donation['receipt_date'])
    re_donation['date'] = pd.to_datetime(re_donation['date'])

    # Get YTD Donation - based on receipt date
    ytd_report = get_ytd_donation()

    # Get YTD Donation - based on transaction date
    previous_ytd_report = get_previous_year_donations()

    # Get Monthly Donation
    monthly_donation = get_monthly_donation()

    # Get Weekly Donation
    weekly_donation = get_weekly_donation()

    # Send Email
    send_email()

except Exception as Argument:

    logging.error(Argument)

    send_error_emails('Error while getting YTD Donation from Raisers Edge', Argument)

finally:

    # Housekeeping
    housekeeping()

    # Stop Logging
    stop_logging()

    exit()