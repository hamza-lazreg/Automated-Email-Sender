import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import pandas as pd

def read_recipients_from_excel(file_path):
    try:
        # Read recipient data from Excel file into a DataFrame
        df = pd.read_excel(file_path)
        
        # Strip any whitespace from column headers
        df.columns = df.columns.str.strip()
        
        # Print the DataFrame to verify its structure
        print("Excel DataFrame:\n", df)
        
        # Ensure the columns are correctly named
        if 'Email' not in df.columns or 'name' not in df.columns:
            raise ValueError("Excel file must contain 'Email' and 'name' columns")
        
        # Drop rows with any NaN values in the specified columns
        df = df.dropna(subset=['Email', 'name'])
        
        # Convert DataFrame to a list of dictionaries
        recipients = df.to_dict(orient='records')
        return recipients
    except Exception as e:
        print(f"Error reading recipients from Excel: {e}")
        return []

def send_email(subject, html_content, to_email):
    from_email = 'enter your email'
    from_password = 'enter your password'

    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = to_email
    msg['Subject'] = subject

    msg.attach(MIMEText(html_content, 'html'))

    with open('logo.png', 'rb') as img:
        logo = MIMEImage(img.read())
        logo.add_header('Content-ID', '<logo>')
        msg.attach(logo)

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)
        server.sendmail(from_email, to_email, msg.as_string())
        print(f"Email sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")
    finally:
        server.quit()

if __name__ == '__main__':
    recipient_list = read_recipients_from_excel('recipients.xlsx')
    subject = "Special 20% Discount Just for You!"
    
    for recipient in recipient_list:
        try:
            email = recipient['Email']  # Ensure this matches the column name in your Excel file
            name = recipient['name']    # Ensure this matches the column name in your Excel file
            html_content = f"""
            <html>
            <body>
                <p>Dear {name},</p>
                <p>We are excited to announce a special <strong>20% discount</strong> on our products!</p>
                <p>Use the code <strong>DISCOUNT20</strong> at checkout to enjoy your discount.</p>
                <img src="cid:logo" alt="Company Logo">
                <p>Best regards,<br>EcoLogic</p>
            </body>
            </html>
            """
            send_email(subject, html_content, email)
        except KeyError as e:
            print(f"Missing expected column in Excel data: {e}")
