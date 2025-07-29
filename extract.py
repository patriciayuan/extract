import fitz  # PyMuPDF
import os
import pandas as pd

# field_label_map = {
#     "4UiwKT2uB4kI": "Financial Instituation Rep Name (Banking)",
#     "6fUUXeadQ4sU": "Postal Code (Banking)",
#     "6moGC7iUbe8P": "Change of Address, No (Vendor)",
#     "8i3798Kq5tw8": "Address (Banking)",
#     "8tIixNAU38MS": "Branch Name (Banking)",
#     "9aHnAfzer3UZ": "Entered By (For Office Use)",
#     "d3pDUH7wbo8O": "Enrol(New) (State)",
#     "dFt4w0PRfoQ0": "Province/State (Vendor)",
#     "E8E76S921VE1": "Contact Phone Number (Vendor)",
#     "EXUPWLtTAYYW": "City (Vendor)",
#     "hPEabhrPsXoH": "Change (State)",
#     "Hrp6X2DaCyoF": "Legal Name (Vendor)",
#     "k7SqxwdWe9c6": "Change of Address, Yes (Vendor)",
#     "LJVBsEILfOgH": "City (Banking)",
#     "M3oZvCpYwAcP": "Authorizing Name (Terms)",
#     "Oa9sivAfVHIJ": "GST Registrant, Yes (GST)",
#     "oncZK9FmdbA2": "Branch ID (Banking)",
#     "orwOXjUM4F84": "Account Number (Banking)",
#     "RIA9DYCaCLEY": "Checked By (For Office Use)",
#     "RmPWjCpnKEsR": "Province/State (Banking)",
#     "S8Jc9R2pSKk0": "Cancel (State)",
#     "tHcfvVccTGM4": "Signer Title (Terms)",
#     "ThIg7LGzSR01": "Postal Code (Vendor)",
#     "U1qLkwYL4lkD": "Vendor Name (Banking)",
#     "uD0ztjXVdGwX": "Vendor Name (For Office Use)",
#     "UMDjYLHSyhgY": "GST Registrant, No (GST)",
#     "V23jEPMIRaAR": "Business/GST (GST)",
#     "VJzXA4Rgv8s8": "Email (Vendor)",
#     "vRLME1Mtc2wL": "Bank ID (Banking)",
#     "xaYLekBpKiwV": "Financial Institution Rep Phone (Banking)",
#     "xhGsItERhWQQ": "Mailing Address (Vendor)",
#     "XsJH7Bp9O1cC": "Vendor ID (For Office Use)",
#     "YZZfAjI1JLEL": "Name of Financial Institution (Banking)",
#     "zrFE5FkUMcsX": "Date of Entered By (For Office Use)",
#     "fEMxk53ivbM4_b8QVeywqMcsR": "Date (Terms)",
#     "fEMxk53ivbM4": "Signature",
    
# }
field_label_map = {
    "CiR0MNaKu5s4": "Financial Instituation Rep Name (Banking)",
    "7Mt93Lhi6HMY": "Postal Code (Banking)",
    "8ZqnnZVTFXcT": "Change of Address, No (Vendor)",
    "pfAYIwTkTj03": "Address (Banking)",
    "8tIixNAU38MS": "Branch Name (Banking)",
    "MQqXyARK9Sw3": "Entered By (For Office Use)",
    "UrCnKA0OoUAA": "Enrol(New) (State)",
    "GnvrrDtuIwQ3": "Province/State (Vendor)",
    "TbJZU7a6i7s8": "Contact Phone Number (Vendor)",
    "IU3oSVQAwboQ": "City (Vendor)",
    "MEzv7Y1RY5Q0": "Change (State)",
    "n6DCtjpgXdUC": "Vendor Name (Vendor)",
    "NRnIb0L5kH8P": "Change of Address, Yes (Vendor)",
    "rexyEoPoY3A2": "City (Banking)",
    "YYMaqD7F2Ik7": "Authorizing Name (Terms)",
    "z6jNTzH63k42": "GST Registrant, Yes (GST)",
    "CZb5cePB42U5": "Branch ID (Banking)",
    "Trm2Tk4wLYsO": "Account Number (Banking)",
    "usvB8V4STss8_KZx38mnJjowT": "Checked By (For Office Use)",
    "wrnLCH1fQ681": "Province/State (Banking)",
    "S8Jc9R2pSKk0": "Cancel (State)",
    "F3GreQWbN60O": "Signer Title (Terms)",
    "YQrOeFoAwBQI": "Postal Code (Vendor)",
    "U1qLkwYL4lkD": "Vendor Name (Banking)",
    "OKiQNl6HyrYJ": "Vendor Name (For Office Use)",
    "RQOQPDNSZ54N": "GST Registrant, No (GST)",
    "VcbbWpgCdJQV": "Business/GST (GST)",
    "Kgj39JXeq7M7": "Email (Vendor)",
    "aa14kLe6jygR": "Bank ID (Banking)",
    "c7f4Ic3ICEQY": "Financial Institution Rep Phone (Banking)",
    "CNxZJO0lQ4gN": "Mailing Address (Vendor)",
    "7zxb9KP1t4YS": "Vendor ID (For Office Use)",
    "YZZfAjI1JLEL": "Name of Financial Institution (Banking)",
    "MQqXyARK9Sw3_WJgVGB4Jtp8T": "Date of Entered By (For Office Use)",
    "rkUXtg5zt6oP_7L6cunIVEHcA": "Date (Terms)",
    "fEMxk53ivbM4": "Signature",
    "26OzTCbIOj0U": "Void Check (Checklist)",
    "BMyd6UZoLNAN": "Date (Banking)",
    "g22JGSaqujE2": "Signature of Authorized Rep (Checklist)",
    "g4ye4vaTUmsP": "Legal Name (Vendor)",
    "tEZYFiVOYuUU": "Contant Name (Vendor)",
    "W6nXiED1dUUL": "Email Address (Checklist)",
    "rkUXtg5zt6oP": "Authorizing Signature (Terms)",
    "kCQY4FVGAi4K": "Name of Financial Institution (Banking)"

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
        print(f"📄 Processing: {filename}")
        field_data = extract_fields_from_pdf(path)
        all_labels.update(field_data.keys())
        data_columns[filename] = field_data


all_labels = list(get_sorted_fields({label: "" for label in all_labels}).keys())
df = pd.DataFrame.from_dict(data_columns, orient='columns')
df = df.reindex(all_labels)


df.to_excel(excel_path)
print("✅ Saved in transposed form to info.xlsx")





