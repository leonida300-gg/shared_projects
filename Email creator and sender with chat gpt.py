import openai
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

# Set the OpenAI API key
openai.api_key = "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"

#server definition
server = None

investor_name_list = ["Kuber Bansal"]
email_list = ["info@wzl.io"]
website_list = ["https://crossculturevc.com"]
subject = "Summer internship"

sender_email = "paolo@wzl.io"
sender_password = "PaoloWZL0601"

#path to pdf
pdf_file = "C:\\Users\\caneparo\\Desktop\\New cvs, cover and linkedin\\CV_Paolo_Caneparo"

for i in range(len(email_list)):
    investor = investor_name_list[i]
    email = email_list[i]
    website = website_list[i]

    # Use the OpenAI API to generate the email text
    response = openai.Completion.create(
        engine="text-davinci-002",
        prompt=(f"Write an email to {investor} asking if they have a summer internship position in VC open for Paolo Caneparo. Personalize the email according to the website. The website is {website}."),
        temperature=0.5,
        max_tokens=1024
    )
    email_text = response["choices"][0]["text"]
        
    try:
        server = smtplib.SMTP('smtp.open-xchange.com', 535)
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
    except Exception as e:
        print(f'Error: {e}')
    finally:
        if server:
            server.quit()

