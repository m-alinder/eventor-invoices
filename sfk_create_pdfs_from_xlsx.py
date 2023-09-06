import os, argparse
from datetime import date, datetime, timedelta
from pythonlib.SFKInvoice import SFKInvoice
import pandas as pd
import time 
from time import strftime


"""Create PDF files for all members

Usage:
python3 create_pdfs_from_xlsx.py files/Fakturor_2022_v3_paw.xlsx files/pdfs_v3_paw/ namn telefon email

"""

#today = date.today()
os.environ["TZ"] = "Europe/Stockholm"
time.tzset()

parser = argparse.ArgumentParser()
parser.add_argument("input_file", type=str, help="File with calculated invoice data")
parser.add_argument("export_directory", type=str, help="Directory to save result to")
parser.add_argument("info_name", type=str, help="Contact name on pdf")
parser.add_argument("info_phone", type=str, help="Contact phone on pdf")
parser.add_argument("info_email", type=str, help="Contact e-mail on pdf")
args = parser.parse_args()

def shorten_text(text:str):
    """Make short text is not too long"""
    if len(text) > 95:
        return text[0:92] + "... "
    return text

# Action
# Remeber start time
start_time = time.time()
print(" Start ".center(80, "-"))

xls = pd.ExcelFile(args.input_file)
df1 = pd.read_excel(xls, 'Aktivitetsöversikt')
df2 = pd.read_excel(xls, 'Fakturaöversikt')
#df1 Index(['id', 'Text', 'BatchId', 'E-id', 'E-mail', 'E-invoiceNo', 'Person',
#       'Event', 'Klass', 'amount', 'fee', 'lateFee', 'status', 'Ålder', 'OK?',
#       '%', 'Subvention', 'Att betala', 'Justering', 'Notering'],
#      dtype='object')
#df2 Index(['Fakturanummer', 'Person', 'Subvention (kr)', 'Justering (kr)',
#       'Totalt att betala (kr)', 'Fakturanamn', 'E-post', 'Faktura skickad',
#       'Faktura betald', 'Notering'],
#      dtype='object')

df = pd.merge(df1,df2,on="Person", suffixes=("_a","_f"))
g = df.groupby("Person")

data = []

# Loop all rows
for index, rows in g:
    person = {
        "invoice_no": int(rows["Fakturanummer"].iloc[0]),
        "name": rows["Person"].iloc[0],
        "e-mail": rows["E-mail"].iloc[0],
        "rows": [],
        "total_discount": rows["Subvention (kr)"].iloc[0],
        "total_adjustment": rows["Justering (kr)"].iloc[0],
        "total_amount": rows["Totalt att betala (kr)"].iloc[0],
        "note": rows["Notering_f"].iloc[0]
    }

    for idx, row in rows.iterrows():
        #print(row["id"])
        if pd.isna(row["Tjänst"]):
            text = str(row["Tävling"]) + " - " + str(row["Klass"])
        else:
            text = row["Tjänst"]

        r = {"id":row["Id"], 
            "text": shorten_text(text), 
            "amount": row["Belopp"], 
            "late_fee": row["Efteranmälningsavgift"], 
            "status": row["Status"], 
            "%": row["Subvention %"], 
            "discount": row["Subvention"], 
            "to_pay": row["Att betala"],
            "adjustment": row["Justering"],
            "note": row["Notering_a"]}
        person["rows"].append(r)
    data.append(person)
    #print(person)
    #for row in rows:
    #    print(row)

#print(data)
idx = 0
for invoice in data:
    idx += 1
    if idx < 10000:
        inv = SFKInvoice(args.export_directory, data=invoice, name=args.info_name, phone=args.info_phone, email=args.info_email)
        #print(invoice["invoice_no"], invoice["name"], invoice["total_amount"])
        
print ("Tidsåtgång: " + str(round((time.time() - start_time),1)) + " s")
print((" Klart (" + strftime("%Y-%m-%d %H:%M") + ") ").center(80, "-"))