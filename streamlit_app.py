#import requests
import streamlit as st
from fpdf import FPDF
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import datetime
from io import StringIO
#from xhtml2pdf import pisa

smtp_server = "smtp.gmail.com"
smtp_port = 587

# Get the current date and year
current_date = datetime.datetime.now().strftime("%Y-%m-%d")
current_year = datetime.datetime.now().year

st.set_page_config(
    page_title="Church Tax Recipt", 
    page_icon="⛪"
)

def main():
    st.title("Church Tax Recipt ⛪")

    st.subheader("Gmail Configuration")
    gmail_sender = st.text_input('Sender Gmail', 'address@gmail.com')
    gmail_app_pw = st.text_input('Gmail App Password', 'abcd efgh ijkl mnop')

    st.subheader("Excel List")
    excel_email_list = st.file_uploader("Email List", type='xlsx', accept_multiple_files=False, disabled=False, label_visibility="visible")

    st.subheader("HTML Templates")
    email_template_file = st.file_uploader("Email Message", type='html', accept_multiple_files=False, disabled=False, label_visibility="visible")
    tax_receipt_template_file = st.file_uploader("Tax Receipt", type='html', accept_multiple_files=False, disabled=False, label_visibility="visible")


    st.subheader("Actions")
    test_mode = st.toggle('Test Mode')

    if st.button("Preview"):
        try:
            if excel_email_list is not None and email_template_file is not None and tax_receipt_template_file is not None:
                with st.spinner('Running...'):
                    df = pd.read_excel(excel_email_list)
                    st.write(df)

                    # Get email template to variable:
                    email_template_io = StringIO(email_template_file.getvalue().decode("utf-8"))
                    email_html_template = email_template_io.read()

                    # Get tax receipt template to variable:
                    tax_receipt_template_io = StringIO(tax_receipt_template_file.getvalue().decode("utf-8"))
                    tax_receipt_html_template = tax_receipt_template_io.read()

                    # Iterate through each row and send emails
                    for index, row in df.iterrows():
                        receipt_id = row['Offering Number']
                        lastname = row['Last Name']
                        firstname = row['First Name']
                        address = row['Address']
                        city = row['City']
                        province = row['Province']
                        postal = row['Postal Code']
                        email = row['Email']
                        amount = row['Donation Amount']
                        amount_float = float(amount)
                        amount_formatted = "{:,.2f}".format(amount_float)
                    
                    
                        # Email Subject
                        subject = "Toronto True Hope Church - {year} Tax Receipt - {id}".format(year = current_year, id = receipt_id)
                        
                        # Email Message
                        email_html = email_html_template.format(
                            firstname = firstname,
                            lastname = lastname,
                            year = current_year)
                        
                        # Tax Receipt
                        stage_tax_receipt_html = tax_receipt_html_template.format(
                            receipt_id = receipt_id,
                            lastname = lastname,
                            firstname = firstname,
                            address = address,
                            city = city,
                            province = province,
                            postal = postal,
                            email = email,
                            amount = amount_formatted,
                            date = current_date,
                            year = current_year)
                        
                        # Add line
                        html_line = """<br><hr style="margin-left: 75px; margin-right: 75px;"><br>"""
                        
                        # Final HTML Content for Tax Receipt
                        tax_receipt_html = stage_tax_receipt_html + html_line + stage_tax_receipt_html
                        
                        # Set output file name    
                        output_file_name = "TTHC - {year} Tax Receipt - {id}.pdf".format(year = current_year, id = receipt_id)

                        st.markdown(email_html, unsafe_allow_html=True)
                        st.markdown(tax_receipt_html, unsafe_allow_html=True)

                        break
                    
                #st.success('Done!')
            
            elif excel_email_list is None:
                st.write("Please choose a valid Excel file")

            elif email_template_file is None:
                st.write("Please choose a valid Email Template file")

            elif tax_receipt_template_file is None:
                st.write("Please choose a valid Tax Receipt file")
                
        except Exception as e:
            st.write("Error occurred 😓")
            st.exception(e)
    
if __name__ == '__main__': 
    main()