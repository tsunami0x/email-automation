import pandas as pd
import yagmail

print ("""
 ______                 _ _                    _                        _   _             
|  ____|               (_) |        /\        | |                      | | (_)            
| |__   _ __ ___   __ _ _| |______ /  \  _   _| |_ ___  _ __ ___   __ _| |_ _  ___  _ __  
|  __| | '_ ` _ \ / _` | | |______/ /\ \| | | | __/ _ \| '_ ` _ \ / _` | __| |/ _ \| '_ \ 
| |____| | | | | | (_| | | |     / ____ \ |_| | || (_) | | | | | | (_| | |_| | (_) | | | |
|______|_| |_| |_|\__,_|_|_|    /_/    \_\__,_|\__\___/|_| |_| |_|\__,_|\__|_|\___/|_| |_|

                                                        Codded By: @TSuNAmi_ORg

""")

def send_bulk_emails(df, email_subject, email_body_template, sender_email, sender_password):
    # Initialize the yagmail.SMTP object
    yag = yagmail.SMTP(user=sender_email, password=sender_password)

    # Function to send an email
    def send_email(to_email, name, subject, body_template):
        # Personalize the email body
        personalized_greeting = f"Hello, {name}"
        email_body = f"{personalized_greeting}\n\n{body_template}"
        
        # Send the email
        yag.send(
            to=to_email,
            subject=subject,
            contents=email_body
        )
        print(f'[+] Email sent to mail: {to_email} name: {name}')

    # Iterate over each row in the DataFrame
    for index, row in df.iterrows():
        send_email(row['email'], row['name'], email_subject, email_body_template)

# Input from user
excel_path = input("[!] Enter the path to the Excel file: ")
email_subject = input("[!] Enter the email subject: ")
email_body_template = input("[!] Enter the main email body: ")
sender_email = input("[!] Enter your email address: ")
sender_password = input("[!] Enter your email password (App Password if 2FA is enabled): ")

# Load the Excel sheet
df = pd.read_excel(excel_path)

# Call the function to send emails
send_bulk_emails(df, email_subject, email_body_template, sender_email, sender_password)

input("Press Enter to exit...")