''' Goto https://myaccount.google.com/lesssecureapps
and switch 'ON' to 'Allow less secure apps:'
'''

# Sending an email from one account to one account 10 times.

import smtplib

from email.mime.text import MIMEText

gmail_user = "## senders email address ##"
gmail_appPassword = "## sender's password ##"

sent_from = ["## sender's email address ##"]
to = ["## recevier's email address ##"]

text = 'you owe me a million dollars, bro'

msg = MIMEText(text)

for i in range(10):
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(gmail_user, gmail_appPassword)
    server.sendmail(sent_from, to, msg.as_string())
    print("Email sent successfully.")

server.quit()
