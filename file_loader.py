import os
import requests
import io
import extract_msg
import numpy as np
import pandas as pd
import pdfplumber
import tabula
import textract

# !pip install utils --user
# !pip install tabula-py
# !pip install --user textract

pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)


def get_mail(file_path):
    msg = extract_msg.Message(file_path)
    return msg


def mail2df(msg):
    my_list = []
    my_list.append([msg.filename, msg.sender, msg.to, msg.date, msg.subject, msg.body,
                    msg.message_id])
    df = pd.DataFrame(my_list, columns=['File Name', 'From', 'To', 'Date', 'Subject', 'MailBody Text', 'Message ID'])
    msg.close()
    return df['MailBody Text']


def open_file(file_name, clean_up=True):
    df = pd.DataFrame()
    file_extension = file_name.split('.')[-1].lower()
    if file_extension == 'msg':
        msg = get_mail(file_name)
        df = mail2df(msg)
        msg.close()
    elif file_extension == 'csv':
        df = pd.read_csv(file_name)
    elif file_extension == 'txt':
        with open(file_name, 'r') as file:
            df = file.read().replace('\n', '')
    elif file_extension == 'xlsx' or file_extension == 'xls':
        df = pd.read_excel(file_name)
    elif file_extension == 'docx' or file_extension == 'doc':
        df = textract.process(file_name).decode()
    elif file_extension == 'pdf':
        df = tabula.read_pdf(file_name, pages='all')
        if len(df) == 0:
            with pdfplumber.open(file_name) as pdf:
                pages = pdf.pages
                text = [pages[i].extract_text() for i, pg in enumerate(pages) if
                        isinstance(pages[i].extract_text(), str)]
            df = ''.join(text)
            pdf.close()
    if clean_up == True:
        if os.path.isfile(file_name):
            os.remove(file_name)
    return (df)


def get_csv(fname, link):
    response = requests.get(link)
    df_csv_file = io.StringIO(response.content.decode('utf-8'))
    df = pd.read_csv(df_csv_file)
    df.to_csv(fname)
    return

def download_csv_file(file_name='currency_codes.csv'):
    get_csv(file_name, 'https://www.datahub.io/JohnSnowLabs/iso-4217-currency-codes/r/iso-4217-currency-codes-csv.csv')
    return

if __name__ == '__main__':
    file_name = 'currency_codes.csv'
    download_csv_file()
    df = open_file(file_name, clean_up=True)
    print(df.head())
