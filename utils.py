import imaplib
import email
import PyPDF2
import zipfile
from io import BytesIO
import google.generativeai as genai
import json
import pandas as pd
from dotenv import load_dotenv
import time
import os


class MailSearch():

    def __init__(self) -> None:
        load_dotenv()
        self.mail = imaplib.IMAP4_SSL('outlook.office365.com')
        self.username = 'madhavraov1985@outlook.com'
        self.password = 'M!1system'
        self.email_ids = None
        self.df = pd.DataFrame({
            'Supplier Name': pd.Series(dtype='str'),
            'Invoice no.': pd.Series(dtype='str'),
            'Date': pd.Series(dtype='str'),
            'Total': pd.Series(dtype='str'),
            'Payment details': pd.Series(dtype='str'), 
            'Arrival Date': pd.Series(dtype='str') 
        })
        genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
        self.generation_config = {
            "temperature": 1,
            "top_p": 0.95,
            "top_k": 0,
            "max_output_tokens": 8192,
            }
        
        self.safety_settings = [
            {
                "category": "HARM_CATEGORY_HARASSMENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_HATE_SPEECH",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            ]
        
        self.model = genai.GenerativeModel(
            model_name="gemini-1.5-pro-latest",
            generation_config=self.generation_config,
            safety_settings=self.safety_settings)
        
    def mailLogin(self):
        self.mail.login(self.username, self.password)

    def selectInbox(self):
        self.mail.select('inbox')
        self.result, self.data = self.mail.search(None, 'ALL')

    def googleModel(self, invoice_data):

        prompt = f"""Extract the following values from {invoice_data}:
            - supplier_name 
            - invoice_no
            - date
            - total_amount
            - card_details
            - arrival_date

            My company name is Hotelbeds, so exclude it as a supplier name.

            Expected output: JSON format (without dollar signs and with underscores in keys)
            Example: {{"supplier_name": "Example Supplier", "invoice_no": "12345", ...}}
            """

        response = self.model.generate_content(prompt)

        return response
    
    def responseToJson(self, response):

        response_str = response._result.candidates[0].content.parts[0].text
        start_index = response_str.find('```json\n') + len('```json\n')
        end_index = response_str.find('\n```', start_index)
        json_str = response_str[start_index:end_index]
        json_str = json_str.replace("None", "null")
        json_str = json_str.replace("'", "\"")
        try:
            json_data = json.loads(json_str)

            column_mapping = {
                "supplier_name": "Supplier Name",
                "invoice_no": "Invoice no.",
                "date": "Date",
                "total_amount": "Total",
                "card_details": "Payment details",
                "arrival_date": "Arrival Date",
            }

            json_data = {column_mapping.get(k, k): v for k, v in json_data.items()} 
        except json.JSONDecodeError as e:
            print(f"Error decoding JSON: {e}")
            return None  # or {}

        return json_data
    
    def zipFileSearch(self, attachment_data):
        try:
            with zipfile.ZipFile(BytesIO(attachment_data)) as zip_file:
                for file_name in zip_file.namelist():
                    if file_name.endswith('.pdf'):
                        pdf_data = zip_file.read(file_name)
                        try:
                            pdf_reader = PyPDF2.PdfReader(BytesIO(pdf_data))
                            extracted_data = ""
                            for page_num in range(len(pdf_reader.pages)):
                                page = pdf_reader.pages[page_num]
                                extracted_data += page.extract_text()
                                time.sleep(5)
                                model_response = self.googleModel(extracted_data)
                                json_data = self.responseToJson(model_response)
                                new_row_df = pd.DataFrame([json_data])
                                self.df = pd.concat([self.df, new_row_df], ignore_index=True)

                        except Exception as e:
                            print(f'Error during extractin{e}')
        except zipfile.BadZipFile:
            print('Error: Not a valid zip file')

    def pdfFileSearch(self, attachment_data):
        try:
            pdf_reader = PyPDF2.PdfReader(BytesIO(attachment_data))
            extracted_data = ""
            for page_num in range(len(pdf_reader.pages)):
                page = pdf_reader.pages[page_num]
                extracted_data += page.extract_text()
                time.sleep(5)
                model_response = self.googleModel(extracted_data)
                json_data = self.responseToJson(model_response)
                
                new_row_df = pd.DataFrame([json_data])
                self.df = pd.concat([self.df, new_row_df], ignore_index=True)
        except Exception as e:
            print(f"Error during extraction: {e}")

    def searchEmailattachment(self):
        self.email_ids = self.data[0].split()
        for email_id in self.email_ids:
            self.result, self.data = self.mail.fetch(email_id, '(RFC822)')
            raw_email = self.data[0][1]
            email_message = email.message_from_bytes(raw_email)
            if email_message.get_content_maintype() == 'multipart':
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    filename = part.get_filename()
                    if filename:
                        if filename.lower().endswith('.zip'):
                            attachment_data = part.get_payload(decode=True)
                            self.zipFileSearch(attachment_data)
                        elif filename.lower().endswith('.pdf'):
                            attachment_data = part.get_payload(decode=True)
                            self.pdfFileSearch(attachment_data)
    

    def mailLogout(self):
        self.mail.logout()

    def getExtractedData(self):
        self.mailLogin()
        self.selectInbox()
        self.searchEmailattachment()
        self.mailLogout()
        return self.df
        


        
