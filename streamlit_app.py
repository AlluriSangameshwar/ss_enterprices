import streamlit as st
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO

# Function to generate Word bill
def generate_docx(customer_name, bill_to, bill_date, items):
    doc = Document()
    doc.add_heading('S. S. ENTERPRISES', level=1).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph(
        'Aluminium Interior Works\nPlot No. 651/A, East Kakatiyanagar, Neredmet, Malkajgiri, Secunderabad – 500056\n'
        'Cell: 9014462295, 7999110733'
    ).alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    doc.add_paragraph(f'Customer Name: {customer_name}')
    doc.add_paragraph(f'Bill to: {bill_to}')
    doc.add_paragraph(f'Date: {bill_date}')
    doc.add_paragraph('BILL', style='Heading 2').alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    headers = [
        "S.No.", "Item Name", "Sub Item Name", "Width in Sq.ft", "Height in Sq.ft", "Depth Sq.ft",
        "Total Sq.ft", "Price Sq.ft", "Per Sq.ft/Each", "Total Price"
