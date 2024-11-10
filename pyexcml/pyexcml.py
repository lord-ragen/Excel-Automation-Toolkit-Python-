import pandas as pd
from lxml import etree
from datetime import datetime

def transform_transaction_description(description):
    # Mapping the transaction description to XML-compliant values
    if "withdrawal" in description.lower():
        if "local currency" in description.lower():
            return "Cash Withdrawal in Local Currency"
        elif "fcy" in description.lower() or "foreign currency" in description.lower():
            return "Cash Withdrawal in FCY (Cheques)"
    elif "deposit" in description.lower():
        if "local currency" in description.lower():
            return "Cash deposit - local currency"
        elif "foreign currency" in description.lower():
            return "Cash Deposit - Foreign Currency"
    return description  # Default if no mapping found

def format_date(date_value):
    # Formats a date to 'YYYY-MM-DDTHH:MM:SS' format or returns empty if date is NaT or missing.
    return date_value.strftime('%Y-%m-%dT%H:%M:%S') if pd.notna(date_value) else ""

def safe_text(value):
    # Converts to string if value is not NaN; otherwise, returns an empty string
    return str(value) if pd.notna(value) else ""

def excel_to_xml(input_file_path, output_file_path="output.xml"):
    # Read Excel file
    df = pd.read_excel(input_file_path)
    
    # Define root element for XML
    root = etree.Element("TransactionsReport")

    # Iterate through each row in the DataFrame
    for _, row in df.iterrows():
        # Create a new XML element for each transaction
        transaction = etree.SubElement(root, "Transaction")
        
        # Account and Transaction Details
        account_number = safe_text(row.get('ACCOUNT NUMBER', ""))
        if account_number:
            etree.SubElement(transaction, "from_account").text = account_number
            etree.SubElement(transaction, "to_account").text = account_number
        
        account_type = safe_text(row.get('PERSONAL ACCOUNT TYPE', ""))
        if account_type:
            etree.SubElement(transaction, "account_type").text = account_type
        
        open_date = row.get('ACC OPEN DATE', "")
        if open_date:
            etree.SubElement(transaction, "opened").text = format_date(open_date)

        transaction_date = row.get('TRANSACTION DATE', "")
        if transaction_date:
            etree.SubElement(transaction, "date_transaction").text = format_date(transaction_date)
        
        trans_ref = safe_text(row.get('TRANS REFERENCE', ""))
        if trans_ref:
            etree.SubElement(transaction, "internal_ref_number").text = trans_ref
        
        sys_date = row.get('SYSDATE', "")
        if sys_date:
            etree.SubElement(transaction, "submission_date").text = format_date(sys_date)

        # Transacting Person Information
        person = safe_text(row.get('TRANSACTING PERSON', ""))
        if person:
            conductor = etree.SubElement(transaction, "t_conductor")
            etree.SubElement(conductor, "t_person").text = person
        
        gender = safe_text(row.get('SIGNATORY GENDER', ""))
        if gender:
            etree.SubElement(transaction, "gender").text = "M" if gender.upper() == "MALE" else "F"
        
        legal_id = safe_text(row.get('LEGAL ID', ""))
        if legal_id:
            identification = etree.SubElement(transaction, "t_person_identification")
            etree.SubElement(identification, "number").text = legal_id
        
        legal_doc_name = safe_text(row.get('LEGAL DOC NAME', ""))
        if legal_doc_name:
            etree.SubElement(identification, "type").text = legal_doc_name
        
        issue_date = row.get('ISSUE DATE', "")
        if issue_date:
            etree.SubElement(identification, "issue_date").text = format_date(issue_date)
        
        # Transaction Specifics
        trans_description = safe_text(row.get('TRANS DESCRIPTION', ""))
        if trans_description:
            etree.SubElement(transaction, "transaction_description").text = transform_transaction_description(trans_description)
        
        currency = safe_text(row.get('CURRENCY', ""))
        if currency:
            etree.SubElement(transaction, "currency_code_local").text = currency
        
        amount = row.get('AMOUNT', "")
        if pd.notna(amount):
            etree.SubElement(transaction, "amount_local").text = str(amount)
        
        amount_lcy = row.get('AMOUNT LCY', "")
        if pd.notna(amount_lcy):
            etree.SubElement(transaction, "amount_lcy").text = str(amount_lcy)
        
        # Additional Fields
        status_code = safe_text(row.get('STATUS CODE', ""))
        if status_code:
            etree.SubElement(transaction, "status_code").text = status_code
        
        institution_name = safe_text(row.get('INSTITUTION NAME', ""))
        if institution_name:
            etree.SubElement(transaction, "institution_name").text = institution_name

    # Save the XML tree to a file
    tree = etree.ElementTree(root)
    tree.write(output_file_path, pretty_print=True, xml_declaration=True, encoding="UTF-8")

