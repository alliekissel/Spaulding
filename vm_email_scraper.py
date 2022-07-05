import win32com.client
import os
from datetime import datetime
import re
import speech_recognition as sr
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
import openpyxl
import pandas as pd

# MailItem Interface documention: https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?redirectedfrom=MSDN&view=outlook-pia#properties_
# create outlook application session
outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")

# go to Voicemails folder
parent_folder = outlook.GetDefaultFolder(6).Parent
voicemails = parent_folder.folders("Voicemails")

# Retrieve messages in Voicemails folder
voicemail_msgs = voicemails.Items

# do stuff with those
for message in voicemail_msgs:
    subject = message.Subject
    print(subject)


# first grab Caller ID from subject line, test with first email
first_vm = voicemail_msgs.GetFirst()
# delimit by hyphen, discard everything before
subject_list = first_vm.Subject.split(" - ")
print(subject_list)
caller_id = subject_list[-1]
re.sub(r"[\n\t\s]*", ' ', caller_id)
caller_id.strip()
print(caller_id)

# grab date, phone number, and recipient from email body
print(first_vm.Body)
#msg_split = first_vm.Body.split("\n")


### phone number
from_idx = first_vm.Body.find("From:")
phone_number_idx_end = first_vm.Body.find("\n", from_idx)

from_line = first_vm.Body[from_idx:phone_number_idx_end]
re.sub(r"[\n\t\s]*", '', from_line)
phone_number = from_line[-12:]
print(phone_number)


### recipient
to_idx = first_vm.Body.find("To:")
recipient_idx_end = first_vm.Body.find("\n", to_idx)

to_line = first_vm.Body[to_idx:recipient_idx_end]
re.sub(r"[\n\t\s]*", ' ', to_line)
to_line = to_line.replace('"','')
recipient = to_line[3:]
recipient = recipient.strip()
print(recipient)


### date
received_idx = first_vm.Body.find("Received:")
received_idx_end = first_vm.Body.find("\n", received_idx)

received_line = first_vm.Body[received_idx:received_idx_end]
re.sub(r"[\n\t\s]*", ' ', received_line)
received_line = received_line.replace('"','')
# find index of comma and remove everything before
comma_idx = received_line.find(",")
received_line = received_line[comma_idx+1:].strip()

datetime_list = received_line.split(" ")
day = datetime_list[0]
month = datetime.strptime(datetime_list[1], '%B').month
year = datetime_list[2]

date_str = datetime(int(year), int(month), int(day)).strftime("%x")
print(date_str)

# download audio attachment
attachments = first_vm.attachments
attachment = attachments.Item(1)
print(str(attachment))
attachment_name = str(attachment)
attachment.SaveASFile(os.getcwd() + '\\' + attachment_name)

# let's try voice-to-text

# initialize recognizer class, using Google speech recognition
r = sr.Recognizer()

# read audio file
with sr.AudioFile(os.getcwd() + '\\' + attachment_name) as source:
    audio_text = r.listen(source)
    
    # error if API is unreachable
    try:
        text = r.recognize_google(audio_text)
        print('Converting audio transcript to text...')
        print(text)
        
    except:
        print('API unreachable. Try again.')


# https://pandas.pydata.org/docs/reference/api/pandas.DataFrame.to_excel.html
dates_list = [date_str]
callers_list = [caller_id]
numbers_list = [phone_number]
messages_list = [text]
attachment_filename = os.getcwd() + '\\' + attachment_name
audio = '=HYPERLINK("{}", "{}")'.format(attachment_filename, "Link Name")
audio_list = [audio]
recipients_list = [recipient]
data_dict = {'Date': dates_list, 'Caller': callers_list, 'Number': numbers_list, 'Message text': messages_list, 'Audio file': audio_list, 'Recipient': recipients_list}

print(data_dict)


df = pd.DataFrame(data=data_dict)
print(df)

df2 = df.copy()
with pd.ExcelWriter("test_vm.xlsx") as writer:
    df.to_excel(writer, sheet_name='test1')
    df2.to_excel(writer, sheet_name='test2')

df3 = df.copy()
with pd.ExcelWriter("test_vm.xlsx", mode="a", engine="openpyxl") as writer:
    df3.to_excel(writer, sheet_name="test3")

def get_sharepoint_context_using_user():
 
    # Get sharepoint credentials
    sharepoint_url = 'https://spauldinggrpllc.sharepoint.com/sites/SpauldingGroup'

    # Initialize the client credentials
    user_credentials = UserCredential('<username>', '<password>')

    # create client context object
    ctx = ClientContext(sharepoint_url).with_credentials(user_credentials)

    return ctx

def create_sharepoint_directory(dir_name: str):
    """
    Creates a folder in the sharepoint directory.
    """
    if dir_name:

        ctx = get_sharepoint_context_using_user()

        result = ctx.web.folders.add(f'Shared Documents/Allie Projects/{dir_name}').execute_query()

        if result:
            # documents is titled as Shared Documents for relative URL in SP
            relative_url = f'Shared Documents/Allie Projects/{dir_name}'
            return relative_url


create_sharepoint_directory('testing_python_connection_to_sharepoint')