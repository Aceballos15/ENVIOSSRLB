import os 
import requests
import pandas as pd
import openpyxl
import time 

script_folder = os.path.dirname(os.path.abspath(__file__))

archivo_excel = os.path.join(script_folder, "SRLAB.xlsx")
df = pd.read_excel(archivo_excel)

filtered_df = df[df["IMPUESTOS"] == 0].reset_index(drop=True)
print(filtered_df)

for index, row in filtered_df.iterrows():
    number = row["NUMERO_ENVIO"]
    nit = row["NIT"]
    verification_digit = row["DV"]
    doc_type = row["Tipon Documento"]
    name = row["NOMBRE"]
    email = row["EMAIL"]
    id_town_dian = row["ID_MUNICIPIO_DIAN"]
    organization_type = row["Tipo organnizacion "]
    invoice_date = row["invoice_date"]
    invoice_date_due = row["invoice_date_due"]
    product_code = row["CODIGO_PRODUCTO"]
    reference = row["REFERENCIA"]
    quantity = row["CANTIDAD"]
    unit_price = row["PRECIO_ANTES_IMPUESTOS"]
    taxes = row["IMPUESTOS"]
    percentage = row["PORCENTAJE"]
    subtotal = row["SUBTOTAL"]
    total = row["TOTAL"]
    number_invoice = ["NUMBER"]
    validate = row["VALIDATE"]
    cufe = row["CUFE"]
    xml = row["XML"]
    
    data = {
        "discrepancy_response": {
            "correction_concept_id": 5,
            "description": reference
        },
        "number": number,
        "sync": True,
        "type_document_id": 5,
        "type_operation_id": 14,
        "invoice_period": {
            "start_date": invoice_date,
            "end_date": invoice_date_due
        },
        "customer": {
            "identification_number": nit,
            "name": name,
            "type_document_identification_id": doc_type,
            "type_organization_id": organization_type
        },
        "legal_monetary_totals": {
            "line_extension_amount": subtotal,
            "tax_exclusive_amount": subtotal,
            "tax_inclusive_amount": taxes,
            "payable_amount": total
        },
        "credit_note_lines": [
            {
                "unit_measure_id": 70,
                "invoiced_quantity": quantity,
                "line_extension_amount": subtotal,
                "tax_totals": [
                    {
                        "tax_id": 1,
                        "tax_amount": taxes,
                        "taxable_amount": subtotal,
                        "percent": 0
                    }
                ],
                "description": reference,
                "code": "1",
                "type_item_identification_id": 1,
                "price_amount": total,
                "base_quantity": quantity
            }
        ]
    }   
