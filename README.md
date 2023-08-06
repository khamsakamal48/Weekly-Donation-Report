# Donation Report from Raisers Edge

### Pre-requisites
- Install below packages

```bash
sudo apt install python3-pip
sudo apt install git
pip install pretty_html_table
pip install pandas
pip install requests
pip install python-dotenv
pip install jinja2
pip install msal
pip install pyarrow
pip install fastparquet
pip install chardet
```

- Set up locale for Indian Currency

```bash
sudo apt install locales
sudo locale-gen en_IN.UTF-8
```

- Ensure and Set up proper system time on the server/machine

```bash
## Display Current Date and Time with timedatectl
timedatectl

## Sync Time to NIST Atomic Clock
sudo timedatectl set-ntp yes

## Set Timezone
sudo timedatectl set-timezone Asia/Kolkata

## Set Hardware Clock to Sync to Local Timezone
timedatectl set-local-rtc 1
```

- Create a **.env** file with below parameters. ***`Replace # ... with appropriate values`***

```bash
AUTH_CODE= # Raiser's Edge NXT Auth Code (encode Client 
REDIRECT_URL=# Redirect URL of application in Raiser's Edge NXT
CLIENT_ID=# Client ID of application in Raiser's Edge NXT
RE_API_KEY=# Raiser's Edge NXT Developer API Key
O_CLIENT_ID= # Client ID of the Application in O365 to send emails
CLIENT_SECRET= # Application Secret in O365
TENANT_ID=# Tenant ID of O365 Group
FROM=# From Email Address
SEND_TO='email_1, email_2' # Email ID of users who needs to receive the report
CC_TO='email_3, email_4' # Email ID of users who will be CC'd for the report
ERROR_EMAILS_TO=# Email ID of user who needs to receive error emails (if any)
```

### Installation
Clone the repository
```bash
git clone https://github.com/khamsakamal48/Weekly-Donation-Report.git
```

Run below command in Terminal
```bash
cd Weekly-Donation-Report
python3 'Request Access Token.py'
```
- Copy and paste the link in a browser to get the **TOKEN**
- Copy the **TOKEN** in the terminal and press ENTER

Run below command in Terminal
```bash
python3 'Refresh Access Token.py'
```

Set a CRON job to refresh token and start the program
```bash
*/42 * * * * cd Weekly-Donation-Report/ && python3 Refresh\ Access\ Token.py > /dev/null 2>&1

# At 10:31 every Monday
31 10 * * 1 cd Weekly-Donation-Report/ && python3 Get\ Donation\ data.py > /dev/null 2>&1
```