import os
import email
from openpyxl import Workbook
import datetime
import re
import phonenumbers

# Set the path to the directory containing the emails
emails_dir = '/Users/reinhard/WarellesLeads/Emails'

# Set the path to the output Excel file
excel_file = '/Users/reinhard/WarellesLeads/Leads.xlsx'

# Create a new Excel workbook and sheet with the specified columns
wb = Workbook()
ws = wb.active
ws.append(['Name', 'Telephone Number', 'Email Address', 'Creation Date', 'Source', 'More Information', 'Zip Code'])

subject_type = ('brochure', 'plan', 'contact')

def NextLine(titel, content):
    # to loop over the text in the web leads where the information is stored on the next line
    lines = content.splitlines()
    for i, line in enumerate(lines):
        if line.startswith(titel):
            titel = lines[i+1].strip()
            break
    else:
        titel = None
    return titel


def format_phone_number(number):
    # Parse the phone number to the Belgium standard
    phone_number = phonenumbers.parse(number, "BE")
    formatted_number = phonenumbers.format_number(phone_number, phonenumbers.PhoneNumberFormat.E164)
    return formatted_number

def subject_is_item_in_list(subject, list):
    for item in list:
        if item in subject:
            return item
    else:
        return subject

def concatenate_strings(*args):
    # Join the non-None strings using the empty string as separator
    return "".join(str(arg) for arg in args if arg is not None)


# Loop through each email file in the directory
for filename in os.listdir(emails_dir):
    if filename.endswith('.eml'):
        # Open the email file and parse its contents
        with open(os.path.join(emails_dir, filename), 'rb') as f:
            msg = email.message_from_bytes(f.read())

        # Check if the email is about a sales lead, type1
        if msg['Subject'] == 'Warelles: er is een nieuwe lead!':
            # Extract the lead information from the email content
            content = msg.get_payload(decode=True).decode('utf-8')
            creation_date = content.split('Creatiedatum: ')[1].split('\n')[0]
            first_name = content.split('Voornaam: ')[1].split('\n')[0]
            last_name = content.split('Achternaam: ')[1].split('\n')[0]
            name = last_name + ' ' + first_name
            email_address = content.split('Mail: ')[1].split('\n')[0]
            unformatted_telephone_number = content.split('Telefoon: ')[1].split('\n')[0]
            telephone_number = format_phone_number(unformatted_telephone_number)
            zip_code = content.split('Postcode: ')[1].split('\n')[0]
            # Get the email subject and write the lead information to the Excel sheet
            subject = 'Socials'
            ws.append([name, telephone_number, email_address, creation_date, subject, None, zip_code])
        # Else it is a weblead, type2
        else:
            # the web lead emails contain multiple parts, formatting etc, we only need the text part
            for part in msg.walk():
                if part.get_content_type() == 'text/plain':
                    content = part.get_payload(decode=True).decode('utf-8')
            
            first_name = NextLine("Voornaam *", content)
            last_name = NextLine("Naam *", content)
            name = last_name + ' ' + first_name
            email_address = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', content).group()
            unformatted_telephone_number = NextLine("Telefoonnummer *", content)
            telephone_number = format_phone_number(unformatted_telephone_number)
            unformatted_creation_date = datetime.datetime.strptime(msg['Date'], '%a, %d %b %Y %H:%M:%S %z')
            creation_date = unformatted_creation_date.strftime(('%Y-%m-%d %H:%M:%S'))
            more_information =  concatenate_strings( (NextLine("Plan", content)), NextLine("Wenst u een vrijblijvend bezoek aan Warelles?", content), NextLine("Formuleer hier uw vraag", content))
            # Get the email subject and write the lead information to the Excel sheet
            subject_long = msg['Subject']
            subject = subject_is_item_in_list(subject_long, subject_type)
            ws.append([name, telephone_number, email_address, creation_date, subject, more_information, None])

# Save the Excel file
wb.save(excel_file)