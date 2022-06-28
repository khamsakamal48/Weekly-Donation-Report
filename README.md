# Weekly-Doonation-Report

- Create a **.env** file with below parameters. ***`Replace # ... with appropriate values`***

```bash

AUTH_CODE= # Raiser's Edge NXT Auth Code (encode Client 
REDIRECT_URL=# Redirect URL of application in Raiser's Edge NXT
CLIENT_ID=# Client ID of application in Raiser's Edge NXT
RE_API_KEY=# Raiser's Edge NXT Developer API Key
MAIL_USERN= # Email Username
MAIL_PASSWORD= # Email password
IMAP_URL=# IMAP web address
IMAP_PORT=# IMAP Port
SMTP_URL=# SMTP web address
SMTP_PORT=# SMTP Port
SEND_TO=# Email ID of user who needs to receive error emails (if any)

```

 - Set up locale for Indian Currency

 ```bash

sudo apt install python3-pip
sudo apt install locales
sudo locale-gen en_IN.UTF-8

 ```

 ```bash

pip3 install pretty_html_table
pip3 install pandas

 ```