import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from email.mime.application import MIMEApplication
from src.UI import init_logging


developer_logger, user_logger = init_logging()

EMAIL_CORP_RELAY = "corpsmtp.walgreens.com"
DEFAULT_EMAIL_RELAY_PORT = 25


def send_email_consolidated(sender=None,
                            receiver=None,
                            cc=None,
                            subject=None,
                            body=None, currentDateTime=None):

    cwd = os.getcwd()

    output_dir = os.path.join(cwd, 'Output_File', 'Report_Files')
    output_dir1 = os.path.join(cwd, 'Output_File', 'Data_Files')



    mail_content = '''
<html>
<head>    <style>
    table { 
        margin-left: auto;
        margin-right: auto;
    }
    table, th, td {
        border: 1px solid black;
        border-collapse: collapse;
    }
    th, td {
        padding: 5px;
        text-align: center;
        font-family: Helvetica, Arial, sans-serif;
        font-size: 90%;
    }
    table tbody tr:hover {
        background-color: #dddddd;
    }
    .wide {
        width: 90%; 
    }
    </style><title></title></head>
<body>
''' f'''
Hello,<br><br> 
The Capital Projects Automation process is complete. Please find the updated file attached in this mail.<br><br>
<br>

<br>

Thanks and Regards, <br><br>
WBS Transformation Team.

<br><br>
<font color=red>This is system generated mail</font>
</body></html>
    '''
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = receiver
    total_list = receiver
    msg['Subject'] = f'Capital Projects-Automation of invoice pending approval'
    msg.attach(MIMEText(mail_content, 'html'))

    attachment_filename = [f for f in os.listdir(output_dir) if os.path.isfile(os.path.join(output_dir, f)) and f.endswith('.xlsx')]

    if attachment_filename:
        file_path = os.path.join(output_dir, attachment_filename[0])
        print(file_path)
        if os.path.exists(file_path):
            with open(file_path, "rb") as file:
                part = MIMEApplication(file.read(), Name=os.path.basename(file_path))
                part['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path)}"'
                msg.attach(part)
                developer_logger.info(f"Attached file: {file_path}")
        else:
            developer_logger.info(f"File not found: {file_path}")
            user_logger.warning(f"File not found for attachment: {file_path}")


    # Attach file from output_dir1
    attachment_filename1 = [f for f in os.listdir(output_dir1) if
                            os.path.isfile(os.path.join(output_dir1, f)) and f.endswith('.xlsx')]
    if attachment_filename1:
        file_path1 = os.path.join(output_dir1, attachment_filename1[0])
        if os.path.exists(file_path1):
            with open(file_path1, "rb") as file:
                part1 = MIMEApplication(file.read(), Name=os.path.basename(file_path1))
                part1['Content-Disposition'] = f'attachment; filename="{os.path.basename(file_path1)}"'
                msg.attach(part1)
                developer_logger.info(f"Attached file: {file_path1}")
        else:
            developer_logger.info(f"File not found: {file_path1}")
            user_logger.warning(f"File not found for attachment: {file_path1}")

    try:
        mail = smtplib.SMTP(EMAIL_CORP_RELAY, DEFAULT_EMAIL_RELAY_PORT)
        mail.sendmail(sender, total_list, msg.as_string())
        print(f"Email sent! {total_list}")
        developer_logger.info(f"Email sent to {total_list}")
        user_logger.debug(f"Email sent to {total_list}")
    except smtplib.SMTPException as err:
        print("Error: unable to send email")
        developer_logger.info(f"Unable to send Email {err}")
        user_logger.debug(f"Unable to send Email {err}")
