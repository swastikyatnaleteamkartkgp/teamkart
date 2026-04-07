import os
import smtplib
import pandas as pd
import time
import random
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.utils import formataddr, make_msgid

# Email credentials
EMAIL = "swastikyatnale.teamkartkgp@gmail.com"
PASSWORD = os.environ.get("EMAIL_PASSWORD")

# Put in the correct csv file name 
data = pd.read_csv('PSU.csv', encoding = 'utf-8')
data.columns = data.columns.str.strip()

# Definitions
BROCHURE_URL = "https://online.fliphtml5.com/TeamKart/1-Qt2Y/" 
YOUR_NAME = "Swastik Yatnale"
TK_LOGO_URL = "https://imgs.search.brave.com/sv9Okf6sV5Cmz8fLS-RwmJ4UnGHgVvUuETOSC-FziQQ/rs:fit:860:0:0:0/g:ce/aHR0cHM6Ly91Z2Mu/cHJvZHVjdGlvbi5s/aW5rdHIuZWUvZTYw/NTFhMTAtMWFiZC00/NWRhLWI4N2QtMzkz/ZDc5MmM5NjE2X3Rl/YW1rYXJ0LWVsZWN0/cmljLWxvZ28td2hp/dGUtc3EucG5nP2lv/PXRydWUmc2l6ZT1h/dmF0YXItdjNfMA"
YOUR_DEPARTMENT = "Department of Mechanical Engineering"
YOUR_YEAR = "First"
YOUR_ROLE_TK = "Mechanical Subsystem Trainee"
YOUR_CONTACT = "+91 9890699650"
YOUR_LINKED_IN = "https://www.linkedin.com/in/swastikyatnale/"
YOUR_FACEBOOK = "https://www.facebook.com/TeamKART/"
CC_EMAILS = ["prajitpradeep.teamkartkgp@gmail.com"]
SUBJECT = "Greetings from Indian Institute of Technology Kharagpur."
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
            padding: 20px;
            border: 1px solid #e0e0e0;
            border-radius: 8px;
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
<body> <div class="content"> <p>Respected Sir/Ma’am,</p>
   <p>I hope this communication finds you well.</p>

    <p>My name is <strong>Swastik Yatnale</strong>. I am writing on behalf of <span class="highlight">TeamKART</span>, a student engineering initiative at the <strong>Indian Institute of Technology Kharagpur</strong>, functioning under the Department of Mechanical Engineering. Established in 2008, the initiative provides undergraduate students with structured, hands-on exposure to engineering design, manufacturing, and project execution through the complete development of Formula-style race cars.</p>

    <p>The project is centred on capacity building through experiential learning. Students work directly on real-world engineering challenges involving design decision-making, manufacturing processes, cost awareness, and system integration. All technical processes and learnings are formally documented to ensure continuity and long-term educational impact for future student cohorts.</p>

    <p>Over the years, the team has successfully manufactured eight combustion-engine vehicles and participated in multiple national and international competitions, receiving recognition for engineering and manufacturing excellence. Building on this foundation, TeamKART has recently undertaken its first electric vehicle project and is currently working on improving and refining it as part of its transition towards sustainable and green engineering practices.</p>

    <p>In this context, we seek to explore a <strong>CSR collaboration</strong> with public sector organisations whose mandates include promotion of technical education, skill development, and youth capacity building. Support from <strong>{company}</strong> would directly contribute to strengthening hands-on engineering education and creating measurable learning outcomes for students.</p>

    <p>We would be grateful for the opportunity to share further details regarding the initiative and discuss the possible scope of CSR engagement at your convenience.</p>

    <p>Thank you for your time and consideration.</p>
</div>

</body>

"""

HTML_TAIL="""
        <p><strong>Kindly refer to:</strong></p>
        <div class="links-section">
            <a href="https://online.fliphtml5.com/TeamKart/1-Qt2Y/#p=1">Our Sponsorship Brochure</a>
            <a href="http://www.teamkart.org/">Our Team's Website</a>
            <a href="https://youtube.com/@teamkart3652">15 Years of TeamKART's Combustion Legacy</a>
            <a href="https://www.instagram.com/team.kart/">Our Instagram Handle</a>
            <a href="https://www.facebook.com/teamkart/">Facebook Page</a>
        </div>

        <div class="footer">
            <p>Thank you for your time and consideration.</p>
            <p>Warm regards,</p>
            <table style="border-collapse: collapse; font-family: Arial, sans-serif;">
                <tr>
                    <td style="padding-right: 15px;">
                        <img src="{tk_logo_url}" width="100" style="display: block;">
                    </td>
                    <td style="border-left: 2px solid #E31E24; padding: 0;"></td>
                    <td style="padding-left: 15px; line-height: 1.4; font-size: 10pt;">
                        <span style="font-weight: bold; font-size: 11pt;">Swastik Yatnale</span><br>
                        First-Year Undergraduate Student<br>
                        Department of Mechanical Engineering<br>
                        Mechanical Subsystem Trainee, TeamKART<br>
                        IIT Kharagpur<br>
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
            msg = MIMEMultipart("alternative")
            msg["From"] = formataddr((YOUR_NAME, EMAIL))
            msg["To"] = row["Email"]
            msg["Cc"] = ", ".join(CC_EMAILS)
            msg["Subject"] = SUBJECT
            msg["Message-ID"] = make_msgid(domain="gmail.com")

            html_template = HTML_HEAD+HTML_BODY+HTML_TAIL
            
            html_content = html_template.format(
                recipient_name=row['Name'],
                # company=row['Company'],
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

            msg.attach(MIMEText(html_content, "html"))
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

