import openai
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import random

# Set the OpenAI API key
openai.api_key = "your Open Ai API key"

#import the email, name investor and company website frome excel 
import pandas as pd

path = "path to your excel file"
df = pd.read_excel(path)
data_vc_final = df[['Website (if available)', 'Partner Name', 'Partner Email']].dropna(axis= 0)

#server definition
server = None

sender_email = "your_email_adress"
sender_password = "your_email_password"
pdf_file = "path to your CV as a PDF file"
subject = "Summer internship"

for i in range(len(data_vc_final)):
    website = data_vc_final.loc[i]['Website (if available)'] #to adjust according to the name of the Excel column
    email = data_vc_final.loc[i]['Partner Email'] #to adjust according to the name of the Excel column
    name_partner = data_vc_final.loc[i]['Partner Name'] #to adjust according to the name of the Excel column
    
    # Use the OpenAI API to generate the email text
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=(f"Write an email to {name_partner} asking if they have a summer internship in VC. Personalize the email according to the website. The website is {website}."),
        temperature=0.5,
        max_tokens=1024
    )
    email_text = response["choices"][0]["text"]

    try:
        # This needs to be adjusted to your email provider server. FYI, Google and Microsoft emails don't work.
        server = smtplib.SMTP('mail.privateemail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)

         # Create the message
        message = MIMEMultipart()
        message["From"] = sender_email
        message["To"] = email
        message["Subject"] = subject
        message.attach(MIMEText(email_text))
    
        # Add the pdf file as an attachment
        with open(pdf_file, "rb") as f:
            attach = MIMEApplication(f.read(), _subtype = "pdf")
            attach.add_header("content-disposition", "attachment", filename = os.path.basename(pdf_file))
            message.attach(attach)
    
        server.sendmail(sender_email, email, message.as_string())
        print(f'Email sent to {email}')

        #generate a random number between 1 and 10
        delay = random.randint(1, 10)

        #sleep for the amount of seconds generated
        time.sleep(delay)

    except Exception as e:
        print(f'Error: {e}')
    finally:
        if server:
            server.quit()
