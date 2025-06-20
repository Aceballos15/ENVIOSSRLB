import os 
import requests
import pandas as pd
import openpyxl
import time 
import json 
from datetime import datetime

script_folder = os.path.dirname(os.path.abspath(__file__))

archivo_excel = os.path.join(script_folder, "SRLAB.xlsx")
df = pd.read_excel(archivo_excel)

filtered_df = df[df["NUMERO_ENVIO"] == "NXC00208"].reset_index(drop=True)
print(filtered_df)

for index, row in filtered_df.iterrows():
    number = row["NUMERO_ENVIO"]
    nit = row["NIT"]
    verification_digit = row["DV"]
    doc_type = row["Tipo Documento"]
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
            "description": "Decuento comercial"
        },
        "number": 208,
        "sync": True,
        "type_document_id": 5,
        "type_operation_id": 14,
        "invoice_period": {
            "start_date": invoice_date.strftime('%Y-%m-%d'),
            "end_date":invoice_date_due.strftime('%Y-%m-%d')
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
            "tax_inclusive_amount": total,
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
                "price_amount": unit_price,
                "base_quantity": quantity
            }
        ]
    }
    print(json.dumps(data))
    # tratar de enviar el documento a la DIAN 
    url_nd = "https://slsas.apifacturacionelectronica.com/api/ubl2.1/credit-note/"
    headers = {
        "Authorization": "Bearer a02dekuqix1xlIVPdzn7Ufa1dSQwW0ErkX5pG1EiXXS8k0AggboxQuxXC2ZImojfA56ULsy0hKbgStWl", 
        "Content-Type": "application/json"
    }
    response = requests.post(url_nd, headers=headers, json=data)
  
    if(response.text): 
        try:
            data_response = response.json()  
            if(data_response["is_valid"]): 
                print(f"NOTA CREDITO ENVIADA CORRECTAMENTE: {cufe}")
                
                filtered_df.at[index, "NUMBER"] = str(data_response["number"])
                filtered_df.at[index, "VALIDATE"] = str(data_response["is_valid"])
                filtered_df.at[index, "CUFE"] = str(data_response["uuid"])
                filtered_df.at[index, "XML"] = str(data_response["attached_document_base64_bytes"])
            else:
                 filtered_df.at[index, "VALIDATE"] = str("false")
        except requests.exceptions.JSONDecodeError:
            filtered_df.at[index, "VALIDATE"] = str("false")
            print("Error al decodificar la respuesta JSON. La respuesta podría no estar formateada correctamente.")  
    
    # Guardar al final
    filtered_df.to_excel(archivo_excel, index=False)
    time.sleep(3)
