import csv
import email, smtplib, ssl
from pathlib import Path
import time
from tqdm import tqdm
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



#contact file name without extension
contacts = "notification_contact_file"

#account credentials
sender_email = ""
password = ""


#count the number of rows in file
with open(f"{contacts}.csv", encoding='cp1252') as file:
    reader = csv.reader(file)
    next(reader)  # Skip header row
    rows = sum(1 for row in reader)
    print(f"Total contacts: {rows}")

#variable to be used in loop to increment the number of contacts
x = 0

#create a text file for logs and report
now = datetime.now()
myreport = open(f'report_{contacts}_{now}.txt', 'w')

#open contact file that will be used to send emails
with open(f"{contacts}.csv", encoding='cp1252') as file:
    reader = csv.reader(file)
    next(reader)  # Skip header row
    #loop through contacts to prepare email and send individually
    for firstname, lastname, email, info_manquant in tqdm(reader, total=rows):
        
        try:
            # Send email here
            subject = f"my subject"

            body = f"Hello {firstname} {lastname},\n\nThis is my message.\n\nThe rest of my message with personalized info from csv {info_manquant}."
            
            sender_email = sender_email

            receiver_email = f"{email}"

            # Create a multipart message and set headers
            message = MIMEMultipart()
            message["From"] = f"Sender name <{sender_email}>"
            message["To"] = receiver_email
            message["Subject"] = subject
            #message["Bcc"] = receiver_email  # Recommended for mass emails

            # Add body to email
            message.attach(MIMEText(body, "plain"))

            text = message.as_string()


            # Log in to server using secure context and send email
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(sender_email, password)
                server.sendmail(sender_email, receiver_email, text)

            #increment 
            x += 1

            #enter new line in report with info on contact that has been targeted
            myreport.write(f"{now} Email sent to: {firstname} {lastname}\n")

            #wait a few seconds before running loop again
            time.sleep(10)


        #for other error types
        except Exception as e:
            myreport.write(f"{now} Other Error & Email not sent to: {firstname} {lastname} - Error: {e}\n")
            pass

#print last error
print(f"\n\nTotal emails sent: {x}")

#write total line in report
myreport.write(f"\n\nTotal emails sent: {x}")
#close file
myreport.close()

