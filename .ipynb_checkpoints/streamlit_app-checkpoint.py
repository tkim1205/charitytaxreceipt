#import requests
import streamlit as st
from fpdf import FPDF
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import datetime
#from xhtml2pdf import pisa

smtp_server = "smtp.gmail.com"
smtp_port = 587

st.set_page_config(page_title="Church Tax Recipt", page_icon="⛪", layout="wide")

def main():
    st.title("Church Tax Recipt ⛪")

    st.subheader("Gmail Configuration")
    gmail_sender = st.text_input('Sender Gmail', 'address@gmail.com')
    gmail_app_pw = st.text_input('Gmail App Password', 'xjmq mhxs eqxd mvzh')

    st.subheader("HTML Templates")
    excel_email_list = st.file_uploader("Excel File - Email List", type='xlsx', accept_multiple_files=False, disabled=False, label_visibility="visible")
    email_template_file = st.file_uploader("HTML File - Email Template", type='html', accept_multiple_files=False, disabled=False, label_visibility="visible")

    st.subheader("Excel List")
    tax_receipt_template_file = st.file_uploader("HTML File - Tax Receipt Template", type='html', accept_multiple_files=False, disabled=False, label_visibility="visible")

    st.subheader("Actions")
    test_mode = st.toggle('Test Mode')

    if st.button("Preview"):
        try:
            if excel_email_list is not None and gmail_sender != '':
                with st.spinner('Running...'):
                    df = pd.read_excel(excel_email_list)
                    st.write(df)

                    email_template_html = file.read(email_template_file)
                    tax_receipt_template_html = file.read(tax_receipt_template_file)
                    
                    st.write(email_template_html)
                    st.write(tax_receipt_template_html)
                    
                st.success('Done!')

            elif gmail_sender == '':
                st.write("Please enter an Email Address")
            
            elif excel_email_list is None:
                st.write("Please choose a valid Excel file")
                
        except Exception as e:
            st.write("Error occurred 😓")
            st.exception(e)
    
if __name__ == '__main__': 
    main()