
from PIL import Image
import pytesseract
import streamlit as st
import pandas as pd
from io import BytesIO
import re
import os

# Specify the path to the Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Title of the app
st.title("Invoice Data Extraction with Tesseract")

# Initialize session state for storing extracted data
if 'data' not in st.session_state:
    st.session_state.data = []

# Upload an image
uploaded_image = st.file_uploader("Choose an image...", type=["jpg", "jpeg", "png"], key="image_uploader")

if uploaded_image is not None:
    # Open the uploaded image
    image = Image.open(uploaded_image)
    
    # Display the image
    st.image(image, caption="Uploaded Image", use_column_width=True)
    
    # Extract text from the image using Tesseract
    extracted_text = pytesseract.image_to_string(image)
    
    # Display the extracted text for debugging purposes
    st.write("Extracted Text:")
    st.write(extracted_text)

    # Define regex patterns to extract specific fields
    address_pattern = r'Address:\s*(.*)'  # Capture everything after "Address:"
    total_amount_pattern = r'\bTOTAL\s*[:\-]?\s*([\d,.]+)'  # Capture the total amount after "TOTAL:"
    tel_pattern = r'(Tel|Phone)[:.]?\s*([\+()\- \d]+)'  # Capture telephone number with "Tel" or "Phone"
    date_pattern = r'(Invoice Date|Date)[:.]?\s*([\d]{1,2}-[A-Za-z]{3}-[\d]{4})'  # Capture date in DD-MMM-YYYY format
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,7}\b'  # Email regex
    
    # Extracting the address using regex
    address_match = re.search(address_pattern, extracted_text, re.IGNORECASE)
    address = address_match.group(1).strip() if address_match else "Not found"
    
    # Extracting the total amount using regex
    total_amount_match = re.search(total_amount_pattern, extracted_text, re.IGNORECASE)
    total_amount = total_amount_match.group(1) if total_amount_match else "Not found"
    
    # Extracting the telephone number using regex
    tel_match = re.search(tel_pattern, extracted_text, re.IGNORECASE)
    telephone = tel_match.group(2).strip() if tel_match else "Not found"

    
    # Extracting the date using regex (DD-MMM-YYYY format)
    date_match = re.search(date_pattern, extracted_text, re.IGNORECASE)
    invoice_date = date_match.group(2) if date_match else "Not found"
    
    # Extracting the email using regex
    email_match = re.search(email_pattern, extracted_text)
    email = email_match.group(0).strip() if email_match else "Not found"

    # Debugging prints to check the matches
    st.write(f"Address Match: {address_match.group(0) if address_match else 'No match'}")
    st.write(f"Total Amount Match: {total_amount_match.group(0) if total_amount_match else 'No match'}")
    st.write(f"Telephone Match: {tel_match.group(0) if tel_match else 'No match'}")
    st.write(f"Date Match: {date_match.group(0) if date_match else 'No match'}")
    st.write(f"Email Match: {email_match.group(0) if email_match else 'No match'}")

    # Append the extracted data to session state
    st.session_state.data.append({
        'Invoice Date': invoice_date,
        'Address': address,
        'Email': email,
        'Telephone': telephone,
        'Total Amount': total_amount
    })

    # Create a DataFrame from the accumulated data
    df = pd.DataFrame(st.session_state.data)
    
    # Display the DataFrame in the Streamlit app
    st.write("Extracted Data:")
    st.write(df)

    # Save the DataFrame to an Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='InvoiceData')
        output.seek(0)  # Reset the buffer to the beginning
    
    # Create a download button for the Excel file
    st.download_button(
        label="Download Extracted Data as Excel",
        data=output,
        file_name="invoice_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
