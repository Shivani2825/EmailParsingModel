import imaplib
import email
import smtplib
import csv
import time
import os
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from spacy.matcher import Matcher
import spacy
from openai import OpenAI
import json

client = OpenAI(api_key = 'Your API-Key')

nlp = spacy.load("en_core_web_sm")

# Set up your email and password directly (replace with your actual credentials)

email_address = "yourmail@gmail.com"
email_password = "your password"


# Set up the IMAP server for Gmail
imap_server = "imap.gmail.com"

# Set up the SMTP server for sending acknowledgment emails
smtp_server = "smtp.gmail.com"
smtp_port = 587
smtp_username = "yourmail@gmail.com"
smtp_password = "your password"

# Global variable to store email body
email_body = ""

chatgpt_responses=[]

# CSV file path to store the extracted information
csv_file_path = "extracted_information.csv"

def write_to_csv(data):
    # Check if the CSV file exists
    is_file_exists = os.path.isfile(csv_file_path)

    # Write headers if the file does not exist
    with open(csv_file_path, mode='a', newline='') as file:
        writer = csv.writer(file)

        if not is_file_exists:
            headers = ["Sender Name", "Sender Email ID", "Dates", "Order Numbers", "Reasons for Issue"]
            writer.writerow(headers)

        writer.writerow(data)

def send_acknowledgment(to_email):
    subject = "Acknowledgment: Your Email Has Been Received"
    body = "Thank you for your email. We have received your message and will get back to you as soon as possible."

    # Set up the MIME
    message = MIMEMultipart()
    message["From"] = email_address
    message["To"] = to_email
    message["Subject"] = subject

    # Attach body to the email
    message.attach(MIMEText(body, "plain"))

    # Connect to the SMTP server
    with smtplib.SMTP(smtp_server, smtp_port) as server:
        server.starttls()
        server.login(smtp_username, smtp_password)

        # Send the email
        server.sendmail(email_address, to_email, message.as_string())

def chatgpt_generate_dictionary(prompt):
    # Set up the prompt and additional user messages
    user_messages = [{'role': 'system', 'content': 'You are a helpful assistant.'}, {'role': 'user', 'content': prompt}]
    
    try:
        # Generate a response from ChatGPT
        response = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=user_messages,
        )

        # Extract the assistant's reply
        assistant_reply = response.choices[0].message.content

        # Print or log the raw content of assistant_reply for debugging
        print("Email Information:", assistant_reply)

        # Parse the assistant's reply into a dictionary
        parsed_reply = json.loads(assistant_reply)
        print(parsed_reply)
        chatgpt_responses.append(assistant_reply)


        return parsed_reply

    except json.JSONDecodeError as json_error:
        print(f"Error decoding JSON: {json_error}")
        return {}  # Return an empty dictionary in case of JSON decoding error

    except Exception as e:
        print(f"Error generating dictionary from ChatGPT: {e}")
        return {}  # Return an empty dictionary for other exceptions


def extract_name_and_email(email_address):
    # Use a regular expression to extract the name and email from the sender's email address
    match = re.match(r"(.*) <(.+)>", email_address)
    if match:
        name, email_id = match.groups()
        return name.strip(), email_id.strip()
    else:
        return "", ""

def process_emails():
    global email_body  # Declare the global variable

    # Connect to the IMAP server
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_address, email_password)

    # Select the mailbox (e.g., "inbox")
    mail.select("inbox")

    # Search for all unseen (unread) emails
    status, messages = mail.search(None, "(UNSEEN)")

    # Fetch and parse emails
    for num in messages[0].split():
        status, msg_data = mail.fetch(num, "(RFC822)")
        msg = email.message_from_bytes(msg_data[0][1])
        print(f"Subject: {msg.get('Subject')}")

        # Get the sender's email address
        sender_email = msg.get('From')
        sender_name, sender_email_id = extract_name_and_email(sender_email)
        print(f"Sender Name: {sender_name}")
        print(f"Sender Email ID: {sender_email_id}")

        # Initialize a variable to store the entire formatted email body
        formatted_email_body = ""

        # Check if the payload is text
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/plain":
                    body = part.get_payload(decode=True)
                    if body:
                        # Assign the content of the body to the variable
                        formatted_email_body = body.decode('utf-8')
                        # Process the text with spaCy
                        doc = nlp(formatted_email_body)

                        # Extract and print named entities
                        extracted_info = {
#                             "dates": [ent.text for ent in doc.ents if ent.label_ == "DATE" and not re.match(r"\b\d{6,}\b", ent.text) and (re.match(r"\b(?:\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{1,2} (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2} \d{2,4})\b", ent.text, re.IGNORECASE))],
                            "dates":[],
                            "order_numbers": [],
                            "reasons_for_issue": set(),  # Use a set to store unique sentences
                        }
                                                # Replace the existing line with if-else statement for "dates"
                
                        for ent in doc.ents:
                            if ent.label_ == "DATE" and not re.match(r"\b\d{6,}\b", ent.text) and (re.match(r"\b(?:\d{1,2}[-/]\d{1,2}[-/]\d{2,4}|\d{1,2} (?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec) \d{1,2} \d{2,4})\b", ent.text, re.IGNORECASE)):
                                extracted_info["dates"].append(ent.text)
                            else:
                                # Handle the case where the date does not match your criteria
                                pass

# Now "dates" contains the filtered list of date texts


                        # Define a custom rule for matching sentences with specific phrases
                        matcher = Matcher(nlp.vocab)
                        pattern = [{"LOWER": {"in": ["working", "problem", "damaged", "fault", "error", "malfunctioning", "defective", "missing "]}}]
                        matcher.add("issue_phrases", [pattern])

                        # Iterate over the matches
                        for match_id, start, end in matcher(doc):
                            span = doc[start:end]
                            extracted_info["reasons_for_issue"].add(span.sent.text.strip())

                        regex = r"\d{6,}"

                        # Find all matches of the regular expression
                        matches = re.findall(regex, formatted_email_body)
                        for i in matches:
                            extracted_info["order_numbers"].append(i)
                    

#                         print("\nExtracted Information:")
#                         print("Dates:", ", ".join(extracted_info["dates"]))
#                         print("Order Numbers:", ", ".join(extracted_info["order_numbers"]))
#                         print("Reasons for Return:")
#                         for reason in extracted_info["reasons_for_issue"]:
#                             print("-", reason)
                       
                        # Create a list to store the extracted information
                        extracted_data = [sender_name, sender_email_id]

                        # Add dates, order numbers, and reasons for issue to the extracted_data list
                        extracted_data.extend(extracted_info["dates"] if extracted_info["dates"] else ['Not Mentioned'])
                        extracted_data.extend(extracted_info["order_numbers"] if extracted_info["order_numbers"] else ['Not Mentioned'])
                        extracted_data.extend(extracted_info["reasons_for_issue"] if extracted_info["reasons_for_issue"] else ['Not Mentioned'])
#                         print(extracted_data)
                        # Write the extracted data to the CSV file
                        write_to_csv(extracted_data)

                        # Generate a dictionary with ChatGPT
                        chatgpt_generated_dictionary = chatgpt_generate_dictionary(f"Create a dictionary with the list data and keys will Name, Email, Date,Order Number and Issue {extracted_data}")
#                         print(chatgpt_generated_dictionary)
                        print("-----------------------------------------------------------End---------------------------------------------------")
                        print("\n")
#                         print(chatgpt_responses)

                    else:
                        print("Body is not text.")
                    break
        else:
            body = msg.get_payload(decode=True)
            if body:
                # Assign the content of the body to the variable
                formatted_email_body = body.decode('utf-8')
                email_body = formatted_email_body  # Store the email body in the global variable
                print(f"Body:\n{formatted_email_body}")
            else:
                print("Body is nhttp://localhost:8888/notebooks/EmailTrigger.ipynb#ot text.")

        # Send acknowledgment email to the sender
#         send_acknowledgment(sender_email)

    # Logout from the email server
    mail.logout()

# Run the script in an infinite loop with a 5-second delay
while True:
    process_emails()
    time.sleep(5)
