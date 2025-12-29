import json
import os
from pathlib import Path
import unicodedata
import pandas as pd
import fitz
import base64
import streamlit as st
import tempfile
from dateutil import parser
from dotenv import load_dotenv
import re
from openai import OpenAI

load_dotenv()


def sanitize_excel_sheet_name(name: str) -> str:
    """
    Sanitize an Excel sheet name by removing forbidden characters and limiting it to 31 characters.
    
    Forbidden characters are: `[:\/\\\?\*\[\]\,]`.
    The sanitized name is returned in uppercase.
    
    Args:
        name (str): The name to sanitize.
        
    Returns:
        str: The sanitized name.
    """
    forbidden_chars = r'[:\/\\\?\*\[\]\,]'
    clean_name = re.sub(forbidden_chars, "", name).upper()
    return clean_name[:31]


def normalize_excel_sheet_name(name: str) -> str:
    """
    Normalize an Excel sheet name by removing accents and converting to lowercase.

    This function is used to sanitize Excel sheet names before writing them to an Excel file.
    It uses the unicodedata.normalize() function to remove accents from the name, and then
    converts the name to lowercase and removes any whitespace characters.

    Args:
        name (str): The name to normalize.

    Returns:
        str: The normalized name.
    """
    name = unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode()
    name = name.lower()
    name = re.sub(r"[^a-z0-9]", "", name)
    return name


def to_french_date(date_str: str) -> str:
    """
    Convert a date string to a French date string.

    Args:
        date_str (str): The date string to convert.

    Returns:
        str: The French date string.
    """
    try:
        dt = parser.parse(date_str, dayfirst=False) 
        return dt.strftime("%d/%m/%Y")
    except Exception as e:
        print(f"Error converting {date_str}: {e}")
        return date_str 


def invoice_to_image(invoice):
    """
    Convert a PDF document to an image and a text.

    Args:
        path (str): The path to the PDF document.
        
    Returns:
        dict or Exception: The text and image, or an Exception if the API call fails.
    """
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp_file:
        tmp_file.write(invoice.read())
        path = tmp_file.name
    path = Path(path)
    doc = fitz.open(path)
    list_b64 = []
    matrix = fitz.Matrix(2, 2) 
    for page in doc:
        # Convert to image
        pix = page.get_pixmap(matrix=matrix)
        img_bytes = pix.tobytes("jpeg")
        try:
            base64_image = base64.b64encode(img_bytes).decode("utf-8")
        except Exception as e: 
            print(f"Error: {e}")
            return None
        url= f"data:image/jpeg;base64,{base64_image}"
        list_b64.append(url)
    return list_b64
    

def fill_excel_file(list_invoices_dict, csv_file, excel_name):
    """
    Fill an Excel file with the invoice data.
    
    Args:
        invoice_dict (dict): The invoice data.
        excel_path (str): The path to the Excel file.
    """
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_file:
        tmp_file.write(csv_file.read())
        excel_path = Path(tmp_file.name)
    try:
        sheets = pd.read_excel(excel_path, sheet_name=None)
    except Exception:
        sheets = {}
    csv_file.seek(0)
    normalized_names = {normalize_excel_sheet_name(sheet_name):sheet_name for sheet_name in sheets}
    COLUMNS = ["N° FACTURE", "REF", "Article", "Quantité facturée", "Prix unitaire", "Total payé HT", 
               "Stock entré en caisse", "Stock restant en caisse", "Boutique", "Casse ou échange"]
    for invoice in list_invoices_dict:
        company_name = invoice.get("company_name", "Unknown")
        number = invoice.get("invoice_reference", {}).get("number", "Unknown")  
        date = invoice.get("invoice_reference", {}).get("date", "Unknown")
        date = to_french_date(date)
        invoice_number = f"{number} du {date}"
        normalized_company_name = normalize_excel_sheet_name(company_name)
        if normalized_company_name in normalized_names:
            sheet_name = normalized_names[normalized_company_name]
            df_existing = sheets[sheet_name]
            df_existing.columns = COLUMNS
        else:
            sheet_name = sanitize_excel_sheet_name(company_name)
            normalized_names[normalize_excel_sheet_name(sheet_name)] = sheet_name
            df_existing = pd.DataFrame(columns=COLUMNS)
        rows = []
        for article in invoice.get("articles", []):
            row = {col: None for col in COLUMNS}
            row["N° FACTURE"] = invoice_number
            row["REF"] = article.get("reference", "")
            row["Article"] = article.get("designation", "")
            row["Quantité facturée"] = article.get("quantity")
            row["Prix unitaire"] = article.get("unit_price")
            row["Total payé HT"] = article.get("total_price")
            rows.append(row)
        df_new = pd.DataFrame(rows, columns=COLUMNS)
        empty_row = pd.DataFrame([[""] * len(df_existing.columns)], columns=df_existing.columns)
        df = pd.concat([df_existing, empty_row, df_new], ignore_index=True)
        sheets[sheet_name] = df
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    st.success(f"Le fichier Excel {csv_file.name} a été rempli avec les {len(invoices) } factures, vous pouvez le telecharger ci-dessous.😃🔥")
    warning_box = st.empty()
    warning_box.warning(f"⚠️ L'IA peut faire des erreurs, pensez à vérifier systématiquement le contenu du fichier Excel.")
    with open(excel_path, "rb") as f:
        st.download_button(label="Télécharger le fichier Excel", data=f, file_name=excel_name)


tools = [{
    "type":"function",
    "name": "extract_invoice_data",
    "description": "Extract structured data from an invoice",
    "strict": True,
    "parameters": {
        "type": "object",
        "properties": {
            "company_name": {
                "type": "string",
                "description": "Name of the company issuing the invoice"
            },
            "invoice_reference": {
                "type": "object",
                "description": "Invoice reference as written on the document",
                "properties": {
                    "number": {
                        "type": "string",
                        "description": "Invoice identifier extracted from the invoice header"
                    },
                    "date": {
                        "type": "string",
                        "description": "Date found near the invoice number or in the header, even if unlabeled or on a separate line (e.g. '14-08-2024')"
                    },
                },
                "required": ["number", "date"],
                "additionalProperties": False
            },
            "articles": {
                "type": "array",
                "description": "List of items listed on the invoice",
                "items": {
                    "type": "object",
                    "properties": {
                        "reference": {
                            "type": ["string", "null"],
                            "description": "Item reference or SKU. Null if not present."
                        },
                        "designation": {
                            "type": "string",
                            "description": "Item name or description"
                        },
                        "quantity": {
                            "type": "number",
                            "description": "Quantity of the item"
                        },
                        "unit_price": {
                            "type": "number",
                            "description": "Unit price before tax"
                        },
                        "total_price": {
                            "type": "number",
                            "description": "Total price for this item (quantity × unit price)"
                        }
                    },
                    "required": ["reference", "designation", "quantity", "unit_price", "total_price"],
                    "additionalProperties": False
                }
            },
        },
        "required": ["company_name", "invoice_reference", "articles"],
        "additionalProperties": False
    }
}]
st.title("Extraction factures")
invoices = st.file_uploader(
    "Factures",
    type=["pdf", "png", "jpg"],
    accept_multiple_files=True
)
csv_file = st.file_uploader("Fichier Excel", type=["csv", "xlsx"])
if st.button("Lancer le traitement"):
    if not invoices or not csv_file:
        st.error("Veuillez fournir des factures et un fichier Excel")
    else:
        list_invoices_dict = []
        for invoice in invoices:
            list_images = invoice_to_image(invoice)
            messages = [{"role":"user", "content": []}]
            for image_url in list_images:
                messages[0]["content"].append({"type":"input_image", "image_url":image_url})
            data = {
                    "model": "gpt-4.1",
                    "input": messages,
                    "tools": tools,
                    "tool_choice": {"type": "function", "name": "extract_invoice_data"},
                    "temperature": 0,
                    }
            for attempt in range(1,4):
                try:
                    print("api gpt call")
                    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
                    response = client.responses.create(**data)
                    final_response = response.output[0].to_json()
                    if isinstance(final_response, dict):
                        invoice_dict = final_response.get("arguments", None)
                        invoice_dict = json.loads(invoice_dict)
                    elif isinstance(final_response, str):
                        invoice_dict = json.loads(final_response)
                        invoice_dict = invoice_dict.get("arguments", None)
                        invoice_dict = json.loads(invoice_dict)
                    else:
                        raise TypeError(f"final_response is not a string or a dict, it's a {type(final_response)}")
                    list_invoices_dict.append(invoice_dict)
                    break
                except Exception as e:
                    print(f"Attempt {attempt}/3 \n API call failed with error: {e} - retrying...")
        fill_excel_file(list_invoices_dict, csv_file, csv_file.name)
  
  







        
