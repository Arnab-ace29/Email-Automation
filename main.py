# import libraries
import pandas as pd 
import os 
import base64 
import datetime
import pickle 
import os.path
import re 
import glob
import pytz 
import logging 
from typing import List 
from docx import Document 
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader 
from email.mime.text import MIMEText
from datetime import datetime
import openai 
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build 
from google.auth.transport.requests import Request

from llama_index import (
    GPTVectorStoreIndex,
    SimpleDirectoryReader,
    LLMPredictor,
    PromptHelper
)
from llama_index import ServiceContext
from langchain.chat_models import ChatOpenAI

pytesseract.pytesseract.tesseract_cmd = ('/usr/bin/tesseract') # -- > check


# openAI API env setup (Insert key Here)
os.environ["OPENAI_API_KEY"] = ""
openai.api_key = ""

# generate basic logging file
logging.basicConfig(filename='emailBot.log' , level=logging.INFO, format='%(message)s')
logging.getLogger('googleapiclient.discovery_cache').setLevel(logging.ERROR)
logging.getLogger('googleapiclient.discovery').setLevel(logging.WARNING)

# create american time format for logging
cst = pytz.timezone('America/Chicago')
datefmt='%Y-%m-%d %H:%M:%S'

# initialize a prompt helper
max_input_size = 4096
# set number of output tokens 1500
num_output = 1000
# set maximum chunk overlap
max_chunk_overlap = 400
# defining a prompt helper
prompt_helper = PromptHelper(max_input_size, num_output ,chunk_overlap_ratio =0.5)
# prompt_helper = PromptHelper(context_window=4096, num_output=256, chunk_overlap_ratio=0.1, chunk_size_limit=None)

# chunks logic
def qa_doc(q):
    '''
    llama index inplementation (with chunks) for QAing on files, mostly words, pdf.
    this function reads files from folder : doc_email and create vectors.
    Then this vector is used to querying based on our question
    '''
    docs = SimpleDirectoryReader('./doc_email/').load_data()
    llm_predictor_gpt4 = LLMPredictor(
            llm = ChatOpenAI(temperature = 0.5 , model_name = 'gpt-4')
    )
    service_context = ServiceContext.from_defaults(llm_predictor = llm_predictor_gpt4)
    try:
        index = GPTVectorStoreIndex.from_documents(docs ,
                                                  service_context = service_context,
                                                  prompt_helper = prompt_helper,
                                                  chunk_size_limit = 1024)
        query_engine = index.as_query_engine(service_context = service_context)
        response = query_engine.query(q)
        return response.response
    except Exception as e:
        # add logging of qa error in a file
        return 'Server is buys, please try again after sometime.'
        pass

# create a service 
def create_service(client_secret_file, api_name, api_version, *scope, prefix = ''):
    '''
    This function is used to autheticate the client secret file and token with the email account
    for communiication between the code and gmail API
    '''
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERISON = api_version
    SCOPES = [scope for scope in scope[0]]
    
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle','rb') as _token:
            creds = pickle.load(_token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CLIENT_SECRET_FILE , SCOPES
            )
            creds = flow.run_console()
        with open('token.pickle','wb') as _token:
            pickle.dump(creds , _token)
    try :
        service = build(API_SERVICE_NAME , API_VERSION , credentials = creds)
        return service
    except Exception as e :
        print(e)
        print(f'failed to create service {API_SERVICE_NAME}')
        return None


# search emails with attachments
def search_emails(query_string: str, label_ids: List=None):
    '''
    This function searched for unread mail which has attachments 
    
    '''
    try:
        message_list_response = service.users().messages().list(
            userId='me',
            labelIds=label_ids,
            q=query_string
        ).execute()

        message_items = [msg for msg in message_list_response.get('messages', []) 
                         if 'UNREAD' in service.users().messages().get(userId='me', id=msg['id']).execute()['labelIds']]

        next_page_token = message_list_response.get('nextPageToken')

        while next_page_token:
            message_list_response = service.users().messages().list(
                userId='me',
                labelIds=label_ids,
                q=query_string,
                pageToken=next_page_token
            ).execute()

            message_items.extend([msg for msg in message_list_response.get('messages', []) 
                                  if 'UNREAD' in service.users().messages().get(userId='me', id=msg['id']).execute()['labelIds']])
                                  
            next_page_token = message_list_response.get('nextPageToken')
        return message_items
    except Exception as e:
        raise NoEmailFound('No emails returned')
        
def get_file_data(message_id, attachment_id, file_name, save_location):
    '''
    This function get the attachments data 
    '''
    response = service.users().messages().attachments().get(
        userId='me',
        messageId=message_id,
        id=attachment_id
    ).execute()

    file_data = base64.urlsafe_b64decode(response.get('data').encode('UTF-8'))
    return file_data        

def get_message_detail(message_id, msg_format='metadata', metadata_headers: List=None):
    '''
    This function return the message details
    '''
    message_detail = service.users().messages().get(
        userId='me',
        id=message_id,
        format=msg_format,
        metadataHeaders=metadata_headers
    ).execute()
    return message_detail


# create new mail
def create_message(sender, to, subject, message_text, thread_id):
    """Create a message for an email."""
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    raw_message = base64.urlsafe_b64encode(message.as_string().encode("utf-8"))
    return {
        'raw': raw_message.decode("utf-8"),
        'threadId': thread_id
    }


# sends the mail with new message
def send_message(service, user_id, message):
    """Send an email message."""
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except Exception as e:
        print('An error occurred: %s' % e)
        return None

# text content from mail
def get_text_parts(part):
    """Utility function to get the text parts from the email body"""
    if part.get('parts'):
        return ''.join(get_text_parts(subpart) for subpart in part.get('parts'))
    if part.get('mimeType') == 'text/plain':
        return base64.urlsafe_b64decode(part.get('body').get('data')).decode()
    return ''

# defining a persona for gpt 
persona="""You are an email assistant. Your job is to generate responses to user's queries and requests. You will go through the 
email text and the attachment data and gererate a reply which will be best suited for the condition.
Always start the response with Thank you. Mention user name if you have it.
Do not add any signature at the end while generating drafts."""

# defining a portion of string to be excluded form text
string_to_remove = 'This message.'

# got 4 API call
def generate_response(prompt):
    """Generate a response using GPT-4."""
    completions = openai.ChatCompletion.create(
        model="gpt-4", 
        temperature=0.01,
        messages=[
            {"role": "system", "content": persona},
            {"role": "user", "content": prompt},
            {'role' : 'user' , 'content' : 'please dont use any salutation at the end.'}
        ]
    )
    message = completions.choices[0]["message"]["content"]
    return message

# helper function
def print_line():
    return '__________________________________________________________________'

# read IMG pdf
def read_pdf_img(file):
    '''
    this function get the data from pdf files which have images. This uses 
    pytesseract to extract data from the images
    '''
    text = ''
    pdfs = glob.glob(f"./doc_email/{file}")
    for pdf_path in pdfs:
        pages = convert_from_path(pdf_path, 500)
        for pageNum,imgBlob in enumerate(pages):
            text = pytesseract.image_to_string(imgBlob,lang='eng')
            with open(f'./doc_email/img/{pdf_path[:-4]}_page{pageNum}.txt', 'w') as the_file:
                the_file.write(text.strip())
                
            with open(f'./doc_email/img/{pdf_path[:-4]}_page{pageNum}.txt' , 'r') as f:
                contents = f.read().strip()
                text += contents
                os.remove(f'./doc_email/img/{pdf_path[:-4]}_page{pageNum}.txt')
    return text 

# reads pdf
def read_pdf(file):
    '''
    This function reads text data from PDF using PdfReader, if there are no text part in the pdf
    then it would use read_pdf_img to read data from the image data and finally return text
    '''
    try :
        text = ''
        reader = PdfReader(f'./doc_email/{file}')
        i = 0
        while i != len(reader.pages):
            page = reader.pages[i]
            text += page.extract_text()
            i+=1
        if len(text) == 0:
            text += read_pdf_img(file)
        return text
    except:
        pass

# read excel
def read_excel(file):
    '''
    this function is used to convert excel data to text
    '''
    df = pd.read_excel(f'./doc_email/{file}')
    text = df.to_string(index=False)
    return text

# read word
def read_docx(file):
    '''
    this function reads the data from word files using docx module and return text 
    '''
    try:
        document = Document(f'./doc_email/{file}')
        text = ' '.join([paragraph.text for paragraph in document.paragraphs])
        return text
    except Exception as e:
        pass

# helper function
def read_data(file_name , count):
    '''
    this function combines all three functions and generate text 
    '''
    text = ''
    if 'pdf' in file_name.lower():
        if read_pdf(file_name)!=None:
            text += f'Attachment : {count} \n'
            text+= read_pdf(file_name)
            text+= '\n\n'
    elif 'docx' in file_name.lower():
        if read_docx(file_name)!=None:
            text += f'Attachment : {count} \n'
            text += read_docx(file_name)
            text+= '\n\n'
    elif 'xlsx' in file_name.lower():
        if read_excel(file_name)!=None:
            text += f'Attachment : {count} \n'
            text += read_excel(file_name)
            text+= '\n\n'
    else:
        pass
    return text

# remove files from folder
def remove_files():
    '''
    this function is used to clear out attachments downloaded from a given folder
    '''
    files = glob.glob(f'{save_location}/*')
    for f in files:
        os.remove(f)

# download attach
def attachments(query_string):
    '''
    this function is used to download the attachments from the mail
    and return a text out of all the attachments 
    '''
    count = 1
    text = ''
    email_messages = search_emails(query_string)
    for email_message in email_messages:
        messageDetail = get_message_detail(email_message['id'], msg_format='full', metadata_headers=['parts'])
        messageDetailPayload = messageDetail.get('payload')
        if 'parts' in messageDetailPayload:
            for msgPayload in messageDetailPayload['parts']:
                file_name = msgPayload['filename']
                body = msgPayload['body']
                if 'attachmentId' in body:
                    attachment_id = body['attachmentId']
                    attachment_content = get_file_data(email_message['id'], attachment_id, file_name, save_location)
                    with open(os.path.join(save_location, file_name), 'wb') as _f:
                        _f.write(attachment_content)           
                    # read data
                    text+= read_data(file_name , count)
                    count+=1
        time.sleep(0.5)        
    return text


# main function
def main(service , Prompt, query_string):
    '''
    this is the main function which gets the unread mail, processes it, generate response from prompts,
    sends mail to the sender, remove files
    '''
    logging.getLogger('googleapicliet.discovery_cache').setLevel(logging.ERROR)
    # setup function create_service
    while True:
        results = service.users().messages().list(userId='me', labelIds=['INBOX'], q='is:unread', maxResults=1).execute()
        if not results:
            break
        messages = results.get('messages', [])
        if not messages:
            break
        else:
            attch_data = attachments(query_string)
            logging.info(f"start time: {datetime.now(cst).strftime(datefmt)}")
            message = messages[0]
            msg = service.users().messages().get(userId='me', id=message['id']).execute()
            service.users().messages().modify(userId='me', id=message['id'], body={'removeLabelIds': ['UNREAD']}).execute()
            thread_id = msg['threadId']
            email_data = msg['payload']['headers']
            from_name = ''
            subject = ''
            for values in email_data:
                name = values['name']
                if name == 'From':
                    from_name = values['value']
                if name == 'Subject':
                    subject = values['value']
            logging.info(f"from {from_name}: , subject : {subject}")
            for part in msg['payload']['parts']:
                try:
                    data_text_decoded = get_text_parts(part) # -- give the body of the mail 
                    text = re.sub(string_to_remove, "", data_text_decoded, flags=re.IGNORECASE)
                    if text.strip():
                        print(f'prompt : {text}')
                        new_text = '\n' + data_text_decoded
                        try:
                            response_text = ''
                            logging.info(f"response generations : success via openAI")
                            response_text += generate_response(text + attch_data)
                            is_response_obtained = True
                        except Exception as e:
                            is_response_obtained = False
                            pass
                        if not is_response_obtained:
                            try:
                                response_text = ''
                                logging.info(f"response generation : success via qa")
                                response_text += qa_doc(Prompt)
                            except Exception as e:
                                pass
  
                        if not response_text:
                            response_text = 'Server is down, please try again after sometime.'

                        logging.info(f"response generation : success")
                        sig = '(Mail generated by Chat GPT)'
                        response_text = response_text + '\n\n' + sig + '\n' + print_line() + '\n' + new_text
                        sender = "test@gmail.com"  # <- Change this to your sender email
                        to = from_name  # <- Change this to your recipient email
                        subject = subject  # <- Change this to your email subject
                        message = create_message(sender, to, subject, response_text, thread_id)
                        send_message(service, "me", message)
                        remove_files()
                        logging.info(f"files removed status : success")
                        logging.info(f"response status : success")
                        logging.info(f"end time: {datetime.now(cst).strftime(datefmt)}")
                        logging.info(print_line())
                except Exception as e:
                    print(e)
                    logging.info(f"exception occured")
                    remove_files()
                    logging.info(f"files removed status : success")
                    logging.info(f"response status : fail")
                    logging.info(f"end time: {datetime.now(cst).strftime(datefmt)}")
                    logging.info(print_line())
                    pass

# # frsit time versification -- run this when restarting the code
if __name__ == '__main__':
    CLIENT_FILE = 'key_V2.json'
    API_NAME = 'gmail'
    API_VERSION = 'v1'
    SCOPES = ['https://mail.google.com/']
    service = create_service(CLIENT_FILE, API_NAME, API_VERSION, SCOPES)
    save_location = './doc_email'
    query_string = 'has:attachment is:unread'

# necessary auth 
def run(Prompt):
    '''
    this function creats a service, and runs the main function
    '''
    CLIENT_FILE = 'key_V2.json'
    API_NAME = 'gmail'
    API_VERSION = 'v1'
    SCOPES = ['https://mail.google.com/']
    service = create_service(CLIENT_FILE, API_NAME, API_VERSION, SCOPES)
    save_location = './doc_email'
    query_string = 'has:attachment is:unread'
    main(service = service , Prompt = Prompt , query_string = query_string)

import time
def periodic_work(interval):
    
    '''
    this function helps in running the main function periodically
    '''
    while True:
        Prompt = 'Summarize the mail without losing any details in bullet points. Always start the response with Thank you. Do not add any signature while generating drafts.'
        # Prompt = ''
        run(Prompt) # Prompt is not used any where and is dummy
        #interval should be an integer, the number of seconds to wait
        time.sleep(interval)

# run all code 
periodic_work(10)




