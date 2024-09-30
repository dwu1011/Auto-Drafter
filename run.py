import os.path
import base64
import mimetypes
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.message import EmailMessage
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

resume_filename = "David_Wu_Resume.pdf"
your_email = "daviddwu22@gmail.com"
email_csv = "email.csv"                 

SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]


def csv_to_dataframe(csv_file):
    # Read the CSV file into a pandas DataFrame
    df = pd.read_csv(csv_file)

    # Display the DataFrame (optional)
    print("Data loaded from CSV!")

    # Return the DataFrame for further use
    return df

def get_credentials():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        # else:
        flow = InstalledAppFlow.from_client_secrets_file(
            "credentials.json", SCOPES
        )
        creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds

def make_drafts(service, creds, email_msg, email, company):
    service = build("gmail", "v1", credentials=creds)

    mime_message = EmailMessage()

    # headers
    mime_message["To"] = email
    mime_message["From"] = your_email                                      
    mime_message["Subject"] = f"{company} Software Engineering Role Continued Interest"     #TODO: Change this to the subject of your email

    # text
    mime_message.set_content(
        email_msg, subtype="html"
    )

    # attachment
    attachment_filename = resume_filename                                        

    type_subtype, _ = mimetypes.guess_type(attachment_filename)
    maintype, subtype = type_subtype.split("/")

    msg = MIMEBase(maintype, subtype)
    

    with open(attachment_filename, "rb") as fp:
        msg.set_payload(fp.read())
        encoders.encode_base64(msg)  # Encode the binary file content to base64

    filename = os.path.basename(attachment_filename)
    msg.add_header("Content-Disposition", "attachment", filename=filename)
    
    mime_message.add_attachment(msg)

    encoded_message = base64.urlsafe_b64encode(mime_message.as_bytes()).decode()

    create_draft_request_body = {"message": {"raw": encoded_message}}
    # pylint: disable=E1101
    draft = (
        service.users()
        .drafts()
        .create(userId="me", body=create_draft_request_body)
        .execute()
    )
    print("SENT")



def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = get_credentials()

    df = csv_to_dataframe(email_csv)
    input("Press Enter to continue...")

    try:
        # Call the Gmail API
        service = build("gmail", "v1", credentials=creds)
        
        # Loop through each column in the DataFrame
        for index, row in df.iterrows():
            cur_row = row.tolist()  
            
            print(cur_row)

            email_msg = f"""                                                                
<html>
    <body>
        <p>Hi {cur_row[1]}!</p>

        <p>I hope this email finds you well. My name is David Wu, and I'm reaching out to express my continued interest in the Software Engineering role at {cur_row[0]}, which I applied for on {cur_row[4]}.</p>

        <p>As a junior studying Computer Science and Math at the University of Michigan, I'm particularly drawn to {cur_row[0]} and the company's innovative approach to {cur_row[5]}.</p>

        <p>Furthermore, as I’m currently in the process for Stripe, Duolingo, and Palantir, I wanted to know if there is anything further I can do to enhance my candidacy for this opportunity. I’m super passionate about {cur_row[0]} and its mission and would love it if my process is expedited so that I have the opportunity to choose {cur_row[0]}.</p>

        <p>I'd also like to briefly share some relevant experiences that make me a strong candidate for this role. This past summer, I worked as a Software Engineering Intern at Endgrate, a startup where I led backend development for our core product. I've also gained valuable research experience as an intern at both the University of Michigan and Stony Brook University, where I applied machine learning and neural networks to predictive modeling projects. On campus, I serve as a board member of the Chinese Student Association, and I'm an active member of a technical consulting club!</p>

        <p>I’d love to chat more about the company, the role, and my fit if you have the time! I have attached my resume below for your convenience.</p>

        <p>Thank you so much for your time and consideration!</p>

        <p>Best,<br>David Wu</p>
    </body>
</html>
"""
            email = cur_row[3]
            company = cur_row[0]

            make_drafts(service, creds, email_msg, email, company)

    except HttpError as error:
        print(f"An error occurred: {error}")


if __name__ == "__main__":
    main()