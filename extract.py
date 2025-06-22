import fitz  # PyMuPDF
import os
import pandas as pd

field_label_map = {
    "4UiwKT2uB4kI": "Financial Instituation Rep Name (Banking)",
    "6fUUXeadQ4sU": "Postal Code (Banking)",
    "6moGC7iUbe8P": "Change of Address, No (Vendor)",
    "8i3798Kq5tw8": "Address (Banking)",
    "8tIixNAU38MS": "Branch Name (Banking)",
    "9aHnAfzer3UZ": "Entered By (For Office Use)",
    "d3pDUH7wbo8O": "Enrol(New) (State)",
    "dFt4w0PRfoQ0": "Province/State (Vendor)",
    "E8E76S921VE1": "Contact Phone Number (Vendor)",
    "EXUPWLtTAYYW": "City (Vendor)",
    "hPEabhrPsXoH": "Change (State)",
    "Hrp6X2DaCyoF": "Legal Name (Vendor)",
    "k7SqxwdWe9c6": "Change of Address, Yes (Vendor)",
    "LJVBsEILfOgH": "City (Banking)",
    "M3oZvCpYwAcP": "Authorizing Name (Terms)",
    "Oa9sivAfVHIJ": "GST Registrant, Yes (GST)",
    "oncZK9FmdbA2": "Branch ID (Banking)",
    "orwOXjUM4F84": "Account Number (Banking)",
    "RIA9DYCaCLEY": "Checked By (For Office Use)",
    "RmPWjCpnKEsR": "Province/State (Banking)",
    "S8Jc9R2pSKk0": "Cancel (State)",
    "tHcfvVccTGM4": "Signer Title (Terms)",
    "ThIg7LGzSR01": "Postal Code (Vendor)",
    "U1qLkwYL4lkD": "Vendor Name (Banking)",
    "uD0ztjXVdGwX": "Vendor Name (For Office Use)",
    "UMDjYLHSyhgY": "GST Registrant, No (GST)",
    "V23jEPMIRaAR": "Business/GST (GST)",
    "VJzXA4Rgv8s8": "Email (Vendor)",
    "vRLME1Mtc2wL": "Bank ID (Banking)",
    "xaYLekBpKiwV": "Financial Institution Rep Phone (Banking)",
    "xhGsItERhWQQ": "Mailing Address (Vendor)",
    "XsJH7Bp9O1cC": "Vendor ID (For Office Use)",
    "YZZfAjI1JLEL": "Name of Financial Institution (Banking)",
    "zrFE5FkUMcsX": "Date of Entered By (For Office Use)",
    "fEMxk53ivbM4_b8QVeywqMcsR": "Date (Terms)",
    "fEMxk53ivbM4": "Signature"
}





def get_sorted_fields(field_dict):
    correct_order = [
        "Cancel (State)",
        "Change (State)",
        "Enrol(New) (State)",
        "Province/State (Vendor)",
        "Mailing Address (Vendor)",
        "Email (Vendor)",
        "Change of Address, Yes (Vendor)",
        "Change of Address, No (Vendor)",
        "Postal Code (Vendor)",
        "City (Vendor)",
        "Legal Name (Vendor)",
        "Contact Phone Number (Vendor)",
        "GST Registrant, Yes (GST)",
        "GST Registrant, No (GST)",
        "Business/GST (GST)",
        "Authorizing Name (Terms)",
        "Date (Terms)",
        "Signer Title (Terms)",
        "Vendor Name (Banking)",
        "Province/State (Banking)",
        "Address (Banking)",
        "Account Number (Banking)",
        "Branch Name (Banking)",
        "Name of Financial Institution (Banking)",
        "Branch ID (Banking)",
        "Bank ID (Banking)",
        "Postal Code (Banking)",
        "Financial Institution Rep Phone (Banking)",
        "City (Banking)",
        "Financial Instituation Rep Name (Banking)",
        "Vendor ID (For Office Use)",
        "Vendor Name (For Office Use)",
        "Entered By (For Office Use)",
        "Checked By (For Office Use)",
        "Date of Entered By (For Office Use)",
        "Signature"
    ]
    
    sorted_dict = {}
    for field in correct_order:
        if field in field_dict:
            sorted_dict[field] = field_dict[field]
    for field, value in field_dict.items():
        if field not in sorted_dict:
            sorted_dict[field] = value
    return sorted_dict










def extract_fields_from_pdf(filepath):
    doc = fitz.open(filepath)
    extracted = {}
    for page in doc:
        for widget in page.widgets():
            field_id = widget.field_name
            field_value = widget.field_value
            field_label = field_label_map.get(field_id, f"Unlabeled Field ({field_id})")
            extracted[field_label] = field_value
    return get_sorted_fields(extracted)













pdf_folder = "data"
excel_path = "info.xlsx"

data_columns = {}
all_labels = set()

for filename in os.listdir(pdf_folder):
    if filename.lower().endswith(".pdf"):
        path = os.path.join(pdf_folder, filename)
        print(f"processing")
        field_data = extract_fields_from_pdf(path)
        all_labels.update(field_data.keys())
        data_columns[filename] = field_data


all_labels = list(get_sorted_fields({label: "" for label in all_labels}).keys())
df = pd.DataFrame.from_dict(data_columns, orient='columns')
df = df.reindex(all_labels)


df.to_excel(excel_path)
print("saved")





