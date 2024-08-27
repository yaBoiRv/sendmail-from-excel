# NOTE: Some of the current Anaconda software does not support the encryption done by the starttls command.
# Therefore, you must use conda install python=3.7.1=h33f27b4_4 to downgrade from 3.7.3-h8c8aaf0_0 in order to rectify this problem.

# Import necessary modules and packages
import pandas as pd
import smtplib
from smtplib import SMTP

while True:
    try:
        # Ask what email service they use
        mailService = input("Which email sending service do you want to use? (Reply \"gmail\" for Gmail, \"yahoo\" for Yahoo Mail, and \"outlook\" for Outlook/Hotmail).")

        # Make sure that the user input something valid
        if not mailService.lower() in ["gmail", "yahoo", "outlook"]:
            raise RuntimeError

    except RuntimeError:
        # Handle the error
        print("Please enter either \"gmail\", \"yahoo\", or \"outlook\".")
        continue

    else:
        # Configure email server string
        if mailService.lower() == "gmail":
            serverStr = "smtp.gmail.com"
        elif mailService.lower() == "yahoo":
            serverStr = "smtp.mail.yahoo.com"
        else:
            serverStr = "smtp-mail.outlook.com"

        # Break out of the loop
        break

# Read the Excel file (NOTE: The information given is not real)
emailList = pd.read_excel("email-list.xlsx")

# Grab all data from the Excel file
emails = emailList["Email"]
firstNames = emailList["Recipient First Name"]
lastNames = emailList["Recipient Last Name"]

# Configure SMTP
s = SMTP(serverStr, 587)
s.ehlo()
s.starttls()
s.ehlo()

# Get login information
senderEmail = input("What email address should the messages be sent from? ")
senderPassword = input("What is the password for the above email address? ")

# Login using the input information
s.login(senderEmail, senderPassword)

# Read the subject to be sent to all email addresses in the Excel file
sbjFile = open("subject.txt", "r", encoding="utf-8")
sbj = sbjFile.read()

# Read the message to be sent to all email addresses in the Excel file
msgFile = open("message.txt", "r", encoding="utf-8")
msg = msgFile.read()

# Begin sending messages
for i in range(len(emails)):
    # Edit the message with information from the Excel file
    msgWithFullName = msg.replace("fullName", f"{firstNames[i]} {lastNames[i]}")
    msgFull = msgWithFullName.replace("name", firstNames[i])

    # Edit the subject with information from the Excel file
    sbjWithFullName = sbj.replace("fullName", f"{firstNames[i]} {lastNames[i]}")
    sbjFull = sbjWithFullName.replace("name", firstNames[i])

    # Create the full email message with proper headers
    email_message = f"Subject: {sbjFull}\n"
    email_message += "Content-Type: text/plain; charset=utf-8\n"
    email_message += f"\n{msgFull}"

    # Send the email
    s.sendmail(senderEmail, emails[i], email_message.encode("utf-8"))
    print(f"Sent email to {firstNames[i]} {lastNames[i]} at {emails[i]}.")

# Terminate connection once email sending has completed and close files
s.quit()
msgFile.close()
sbjFile.close()
