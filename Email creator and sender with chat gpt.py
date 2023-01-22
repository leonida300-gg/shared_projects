import openai
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import time
import random

# Set the OpenAI API key
openai.api_key = "add your Open Ai key"

#server definition
server = None

investor_name_list = ["Kuber Bansal", "WZL"]
email_list = ["kuberbansal@gmail.com", "info@wzl.io"]
website_list = ["https://crossculturevc.com", "shipwrecked.vc"]
subject = "Summer internship"

sender_email = "your_email_adress"
sender_password = "your_email_password"

#path to pdf
pdf_file = "C:\\Users\\caneparo\\Desktop\\CV_Paolo_Caneparo.pdf"

for i in range(len(email_list)):
    investor = investor_name_list[i]
    email = email_list[i]
    website = website_list[i]

    # Use the OpenAI API to generate the email text
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=(f"Write an email to {investor} asking if they have a summer internship position in VC open for Paolo Caneparo. Personalize the email according to the website and state that they can know more from my cv and they can contact me when they want for knowing more about me. The website is {website}."),
        temperature=0.5,
        max_tokens=1024
    )
    email_text = response["choices"][0]["text"]
        
    try:
        server = smtplib.SMTP("mail.privateemail.com", 587)
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

