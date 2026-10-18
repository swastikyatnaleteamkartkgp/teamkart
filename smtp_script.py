import os
import smtplib
import pandas as pd
import time
import random
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, make_msgid
from email.mime.base import MIMEBase
from email import encoders

# Email credentials
EMAIL = "swastikyatnale.teamkartkgp@gmail.com"
PASSWORD = os.environ.get("EMAIL_PASSWORD")

# Put in the correct csv file name 
data = pd.read_csv('Merged.csv', encoding = 'utf-8')
data.columns = data.columns.str.strip()

# Definitions
ATTACHMENT_PATH = "TeamKART_Sponsorship_Brochure.pdf"
BROCHURE_URL = "https://online.fliphtml5.com/vqomr/TeamKART-Sponsorship-Brochure/" 
YOUR_NAME = "Swastik Yatnale"
TK_LOGO_URL = "https://imgs.search.brave.com/sv9Okf6sV5Cmz8fLS-RwmJ4UnGHgVvUuETOSC-FziQQ/rs:fit:860:0:0:0/g:ce/aHR0cHM6Ly91Z2Mu/cHJvZHVjdGlvbi5s/aW5rdHIuZWUvZTYw/NTFhMTAtMWFiZC00/NWRhLWI4N2QtMzkz/ZDc5MmM5NjE2X3Rl/YW1rYXJ0LWVsZWN0/cmljLWxvZ28td2hp/dGUtc3EucG5nP2lv/PXRydWUmc2l6ZT1h/dmF0YXItdjNfMA"
YOUR_DEPARTMENT = "Department of Mechanical Engineering"
YOUR_YEAR = "Second"
YOUR_ROLE_TK = "Suspension Steering & Brakes Member"
YOUR_CONTACT = "+91 9890699650"
YOUR_LINKED_IN = "https://www.linkedin.com/in/swastikyatnale/"
YOUR_FACEBOOK = "https://www.facebook.com/TeamKART/"
CC_EMAILS = ["prajitpradeep.teamkartkgp@gmail.com","nihalsingh.teamkartkgp@gmail.com","shubhayubasak.teamkartkgp@gmail.com","mohamed.teamkartkgp@gmail.com"]
#SUBJECT = f"Follow Up: Potential partnership of TeamKART, IIT Kharagpur with {row["Company"]}"
HTML_HEAD = """
<!DOCTYPE html>
<html>
<head>
    <style>
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #1a1a1a;
            max-width: 600px;
            margin: 0 auto;
        }}
        .content {{
            padding: 0px;
            border: 0px solid #e0e0e0;
            border-radius: 0px;
        }}
        .highlight {{
            color: #E31E24;
            font-weight: bold;
        }}
        .links-section {{
            background-color: #f4f4f4;
            padding: 15px;
            border-left: 4px solid #E31E24;
            margin: 20px 0;
        }}
        .links-section a {{
            color: #E31E24;
            text-decoration: none;
            display: block;
            margin-bottom: 8px;
            font-size: 14px;
        }}
        .footer {{
            margin-top: 25px;
            padding-top: 15px;
            border-top: 1px solid #eeeeee;
        }}
    </style>
</head>"""

# Template for the body
HTML_BODY = """
<body>
<p>Dear <strong>{recipient_name}</strong>,</p>

    <p>I hope you are doing well. I wanted to follow up on my previous email on a potential collaboration between <strong>TeamKART, IIT Kharagpur</strong> and <strong>{company}</strong>.</p>
    <p>We are currently preparing our first Electric Vehicle for <strong>Formula Bharat 2027</strong>, with which we wish to begin a new chapter for our team. Partnership and support from a great organisation like {company} could aid us in our research and development of the project. </p>
    <p>I have attached our sponsorship brochure to this mail. I will be readily available to answer any queries regarding the team or the project.</p>
    <p>Thank you for your time. Looking forward to a positive response from your side.</p>
    <p>Regards,</p>


"""

"""
<p><strong>Kindly refer to:</strong></p>
        <div class="links-section">
            <a href="https://online.fliphtml5.com/vqomr/TeamKART-Sponsorship-Brochure/">Our Sponsorship Brochure</a>
            <a href="http://www.teamkart.org/">Our Team's Website</a>
            <a href="https://youtube.com/@teamkart3652">15 Years of TeamKART's Combustion Legacy</a>
            <a href="https://www.instagram.com/team.kart/">Our Instagram Handle</a>
            <a href="https://www.facebook.com/teamkart/">Facebook Page</a>
        </div>

"""

HTML_TAIL="""
        
        <div class="footer">
            <table style="border-collapse: collapse; font-family: Arial, sans-serif;">
                <tr>
                    <td style="padding-right: 15px;">
                        <img src="{tk_logo_url}" width="100" style="display: block;">
                    </td>
                    <td style="border-left: 2px solid #E31E24; padding: 0;"></td>
                    <td style="padding-left: 15px; line-height: 1.4; font-size: 10pt;">
                        <span style="font-weight: bold; font-size: 11pt;">Swastik Yatnale</span><br>
                        Second-Year Undergraduate Student<br>
                        Department of Mechanical Engineering<br>
                        Suspension Steering & Brakes Member, TeamKART<br>
                        Indian Institute of Technology, Kharagpur<br>
                        Contact: +91 9890699650<br>
                        <a href="https://www.linkedin.com/in/swastikyatnale/" style="color: #0044cc;">LinkedIn</a> | 
                        <a href="https://www.facebook.com/TeamKART/" style="color: #0044cc;">Facebook</a>
                    </td>
                </tr>
            </table>
        </div>
    </div>
</body>
</html>
"""

def send_emails():

    now = datetime.now()

    # Format it as DD-MM-YYYY
    today_date = now.strftime("%d-%m-%Y")
    
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL, PASSWORD)
        print(f"Successfully logged in. Sending emails for {today_date}...")
    except Exception as e:
        print(f"Login failed: {e}")
        return

    for index, row in data.iterrows():
        try:
            msg = MIMEMultipart("mixed")
            msg["From"] = formataddr((YOUR_NAME, EMAIL))
            msg["To"] = row["Email"]
            msg["Cc"] = ", ".join(CC_EMAILS)
            msg["Subject"] = f"Follow Up: Potential partnership of TeamKART, IIT Kharagpur with {row['Company']}"
            msg["Message-ID"] = make_msgid(domain="gmail.com")

            html_template = HTML_HEAD+HTML_BODY+HTML_TAIL
            
            html_content = html_template.format(
                recipient_name=row['Sal'],
                brochure_link = BROCHURE_URL,
                tk_logo_url = TK_LOGO_URL,
                your_name = YOUR_NAME,
                your_year = YOUR_YEAR,
                your_department = YOUR_DEPARTMENT,
                your_role = YOUR_ROLE_TK,
                your_contact = YOUR_CONTACT,
                your_linkedin = YOUR_LINKED_IN,
                your_facebook = YOUR_FACEBOOK,
                company = row['Company']
            )

            # Create a container for the email body
            body = MIMEMultipart("alternative")
            body.attach(MIMEText(html_content, "html"))
            msg.attach(body)
            
            # Attach the PDF file
            with open(ATTACHMENT_PATH, "rb") as attachment_file:
                attachment = MIMEBase("application", "pdf")
                attachment.set_payload(attachment_file.read())
            
            encoders.encode_base64(attachment)
            
            attachment.add_header(
                "Content-Disposition",
                "attachment",
                filename=os.path.basename(ATTACHMENT_PATH)
            )
            
            msg.attach(attachment)
            recipients = [row["Email"]] + CC_EMAILS
            server.sendmail(EMAIL, recipients, msg.as_string())
            ist_now = datetime.now() + timedelta(hours=5, minutes=30)
            print(f"Sent email to {row['Email']} at {ist_now.strftime('%H:%M:%S')} IST")
            
            time.sleep(random.randint(25, 55))

        except Exception as e:
            print(f"Error sending to {row['Email']}: {e}")

    server.quit()

if __name__ == "__main__":
    send_emails()

