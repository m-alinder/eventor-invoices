import os
import re
import sys
import pandas as pd
import numpy as np
import requests
import xml.etree.ElementTree as et

#pip3 install xlrd


# declare global variables
discount_rules = 2023
cache_dir = "eventor-cache/"
dfDiscounts = None

def parse_person_results_xml(xml_file):
    df_cols = ["Type", "EventId", "Name", "Date", "Class", "Family", "Given", "CompetitorStatus"]
    xtree = et.parse(xml_file)
    xroot = xtree.getroot()
    rows = []
    
    for resultlist in xroot.findall('ResultList'): 
        res = []
        event = resultlist.find('Event')
        if event is not None:
            try:
                eventType = event.attrib.get("eventForm")
                # eventForm is not always present. Assume individual when not available.
                if (eventType is None):
                    eventType = "IndSingleDay"
                eventId = event.find("EventId").text
                eventName = event.find("Name").text
                eventDate = event.find("EventRace").find("RaceDate").find("Date").text
                res.append(eventType)
                res.append(eventId)
                res.append(eventName)
                res.append(eventDate)

                if (eventType == "IndSingleDay"):
                    for classResult in resultlist.findall("ClassResult"):
                        eventClass = classResult.find("EventClass").find("Name").text
                        familyName = classResult.find("PersonResult").find("Person").find("PersonName").find("Family").text
                        givenName = classResult.find("PersonResult").find("Person").find("PersonName").find("Given").text
                        status = classResult.find("PersonResult").find("Result").find("CompetitorStatus").attrib.get("value")
                        res.append(eventClass)
                        res.append(familyName)
                        res.append(givenName)
                        res.append(status)

                        rows.append({df_cols[i]: res[i] 
                            for i, _ in enumerate(df_cols)})
                        
                        res = []
                        res.append(eventType)
                        res.append(eventId)
                        res.append(eventName)
                        res.append(eventDate)

                elif (eventType == "RelaySingleDay"):
                    eventClass = resultlist.find("ClassResult").find("EventClass").find("Name").text
                    familyName = resultlist.find("ClassResult").find("TeamResult").find("TeamMemberResult").find("Person").find("PersonName").find("Family").text
                    givenName = resultlist.find("ClassResult").find("TeamResult").find("TeamMemberResult").find("Person").find("PersonName").find("Given").text
                    status = resultlist.find("ClassResult").find("TeamResult").find("TeamMemberResult").find("CompetitorStatus").attrib.get("value")
                    res.append(eventClass)
                    res.append(familyName)
                    res.append(givenName)
                    res.append(status)

                    rows.append({df_cols[i]: res[i] 
                        for i, _ in enumerate(df_cols)})
                    
                elif (eventType == "IndMultiDay"):
                    for classResult in resultlist.findall("ClassResult"):
                        for personResult in classResult.findall("PersonResult"):
                            eventClass = classResult.find("EventClass").find("Name").text
                            familyName = personResult.find("Person").find("PersonName").find("Family").text
                            givenName = personResult.find("Person").find("PersonName").find("Given").text
                            status = personResult.find("RaceResult").find("Result").find("CompetitorStatus").attrib.get("value")
                            res.append(eventClass)
                            res.append(familyName)
                            res.append(givenName)
                            res.append(status)

                            rows.append({df_cols[i]: res[i] 
                                for i, _ in enumerate(df_cols)})
                            
                            res = []
                            res.append(eventType)
                            res.append(eventId)
                            res.append(eventName)
                            res.append(eventDate)
            except:
                print("Unknown structure: " + eventType + ", " + eventId + ', ' + eventName)
        
    out_df = pd.DataFrame(rows, columns=df_cols)
    return out_df

def get_person_results(apikey, personId, fromDate, toDate):
    # First, try get status from local file
    fileName = cache_dir + fromDate + "_" + toDate + "_" + str(personId) + ".xml"

    try:
        res = parse_person_results_xml(fileName)
        return res
    except:
        print(fileName + " not found. Requesting personal results from Eventor.")

    # Add the Authorization header
    headers = {'ApiKey': apikey}

    base_url = 'https://eventor.orientering.se'

    # Get results from eventor
    response = requests.get(f'{base_url}/api/results/person?personId={str(personId)}&fromDate={fromDate}&toDate={toDate}', headers=headers)
    if (response.status_code == 200):
        text_file = open(fileName, "w")
        text_file.write(response.content.decode("utf-8") )
        text_file.close()
        res = parse_person_results_xml(fileName)
        return res
    else:
        exit(-1)

def parse_club_info_xml(xml_file):
    xtree = et.parse(xml_file)
    xroot = xtree.getroot()

    df = pd.DataFrame(columns=('id', 'email'))
    for list in xroot.findall('Person'): 
        id = list.find('PersonId')
        tele = list.find('Tele')
        if id is not None and tele is not None:
            email = tele.attrib.get("mailAddress")
            df.loc[len(df)] = [id.text, email if email is not None else ""]
    return df

def get_club_info(apikey, clubId, fromDate, toDate):
    # First, try get status from local file
    fileName = cache_dir + fromDate + "_" + toDate + "_info_" + str(clubId) + ".xml"

    try:
        res = parse_club_info_xml(fileName)
        return res
    except:
        print(fileName + " not found. Requesting personal info from Eventor.")

    # Add the Authorization header
    headers = {'ApiKey': apikey}

    base_url = 'https://eventor.orientering.se'

    # Get results from eventor
    response = requests.get(f'{base_url}/api/persons/organisations/{str(clubId)}?includeContactDetails=true', headers=headers)
    if (response.status_code == 200):
        text_file = open(fileName, "w")
        text_file.write(response.content.decode("utf-8") )
        text_file.close()
        res = parse_club_info_xml(fileName)
        return res
    else:
        exit(-1)

def translate_to_swe(word):
    eng_to_swe = {
        "IndSingleDay"          : "Individuell",
        "IndMultiDay"           : "Multi",
        "RelaySingleDay"        : "Stafett",
        "DidNotStart"           : "Ej Start",
        "Cancelled"             : "Ej Start",
        "MisPunch"              : "OK",
        "DidNotFinish"          : "OK",
        "Disqualified"          : "OK",
        "OK"                    : "OK",
        "RentalPunchingCard"    : "Hyrbricka"
    }

    ord = eng_to_swe.get(word)
    if (ord is not None):
        return ord
    return word

def get_age(birthDate,eventDate):
    return eventDate.year - birthDate.year

def paid_cash(competition:str):
    # Usually paid cash
    return re.search(r"veteran|skogsflicks|motionsorientering", competition, re.IGNORECASE) != None

def is_relay(competition_type:str):
    return re.search(r'stafett', competition_type, re.IGNORECASE) != None

def check_ok(status, competition, competition_type):
    nok = status == '' or status == 'Ej Start' or status == 'Tjänst'

    # Competitions payed cash and Relays are always ok
    return not nok or paid_cash(competition) or is_relay(competition_type)

assert check_ok('', '', '') == False
assert check_ok('Ej Start', '', '') == False
assert check_ok('OK', '', '') == True

def normalize_amount(amount):
    return int(float(amount.replace(",","."))) if type(amount) == str else int(amount)

assert normalize_amount(10) == 10
assert normalize_amount("10") == 10
assert normalize_amount(500.40) == 500
assert normalize_amount("500,40") == 500
assert normalize_amount(250.20) == 250
assert normalize_amount("250,20") == 250

def normalize_fee(late_fee):
    return 0 if late_fee == "" else int(float(late_fee.replace(",","."))) if type(late_fee) == str else int(float(late_fee))

assert normalize_fee("250,20") == 250
assert normalize_fee("25") == 25

def calculate_discount(valid:bool, competition:str, competition_type:str, age) -> int:
    """Calculate discount in precentage"""

    # 100% discount for relay or paid with cash
    if (paid_cash(competition) or is_relay(competition_type)):
        return 100

    # Not valid entries give 0% discount (did not start, card rental)
    if not valid: 
        return 0

    column = "Vuxen"
    if age <= 16:
        column = "Barn"

    for row in dfDiscounts.itertuples():
        if row.Tävling == competition:
            return dfDiscounts.at[row.Index, column]

    #print("Standardsubvention för: " + competition)
    return 40


assert calculate_discount(False, "", "", 10) == 0
assert calculate_discount(False, "", "", 40) == 0

def calculate_discount_amount(amount, lateFee, competition:str, competition_type:str, age, valid, discount, person):
    """Calculate only discounted amount
    
    In certain cases (e.g., relays) also lateFee might be in subject for discount
    """

    amount = normalize_amount(amount) 
    lateFee = normalize_fee(lateFee)

    # Pay nothing for relay or paid with cash
    if (paid_cash(competition) or is_relay(competition_type)):
        return amount

    #print(f"[calculate_discount_amount] amount: '{amount}' latefee: '{lateFee}' '{competition}' age: '{age}' valid: '{valid}' d: '{discount}' p: '{person}'")
    if not valid:
        return 0

    if (amount == lateFee):
        # Assume normal fee when aount equals lateFee
        discount_amount = round((amount) * (discount / 100))
    else:
        # Amount contains total fee, lateFee included.
        # We only discount for the standard fee
        discount_amount = round((amount-lateFee) * (discount / 100))

    return discount_amount

# Amount is total fee to pay, lateFee already part of amount.
assert calculate_discount_amount("100", "0", "En tävling", "Single", 25, False, 40, "Person") == 0
assert calculate_discount_amount("100", "0", "En tävling", "Single", 25, True, 40, "Person") == 40
assert calculate_discount_amount("100", "0", "En tävling", "Single", 25, False, 100, "Person") == 0
assert calculate_discount_amount("100", "0", "En tävling", "Single", 25, True, 100, "Person") == 100
assert calculate_discount_amount("120", "20", "Sjövalla FK 1 i DM, stafett, Göteborg + Västergötland", "Stafett", 16, True, 100, "Person") == 120
assert calculate_discount_amount("120", "20", "Sjövalla FK 1 i DM, stafett, Göteborg + Västergötland", "Stafett", 25, True, 100, "Person") == 120
assert calculate_discount_amount("150", "50", "En tävling", "Single", 16, True, 40, "Person") == 40
assert calculate_discount_amount("150", "50", "En stafett-tävling", "Stafett", 16, True, 100, "Person") == 150
assert calculate_discount_amount("150", "50", "En tävling", "Single", 25, True, 40, "Person") == 40
assert calculate_discount_amount("150", "50", "En stafett-tävling", "Stafett", 25, True, 100, "Person") == 150
assert calculate_discount_amount("750,40", "250,20", "Sjövalla FK 1 i 25manna", "Stafett", 25, True, 100, "Person") == 750
assert calculate_discount_amount("100", "100", "En tävling", "Single", 25, True, 40, "Person") == 40

# 2022-12-13 'amount' and 'fee' is in very few cases different, but seems like we can ignore 'fee'
def calculate_amount_to_pay(amount, late_fee, competition:str, age, valid:bool, discount, person, discount_amount, adjustment) -> int:
    """Total amount to pay, including potential discount
    
    Requirements:
    - Entry must be valid (in scope of discount)
    - Late fee is never part of discount
    """
    amount = normalize_amount(amount)
    late_fee = normalize_fee(late_fee)

    # Person will pay full amount 
    if not valid:
        return amount + adjustment

    # if discount == 100:
    #     return adjustment

    return amount - discount_amount + adjustment

# def calculate_amount_to_pay(amount, late_fee, competition:str, age, valid:bool, discount, person, discount_amount, adjustment) -> int:
assert calculate_amount_to_pay(130, 30, "En tävling", 16, False, 40, "Person", 0, 0) == 130
assert calculate_amount_to_pay(100, 0,  "En tävling", 16, False, 40, "Person", 0, 0) == 100
assert calculate_amount_to_pay(50,   50, "En tävling", 16, False, 100, "Person", 0, 0) == 50
assert calculate_amount_to_pay(100, 0,  "En tävling", 16, True, 40, "Person", 40, 0) == 60
assert calculate_amount_to_pay(150, 50, "En tävling", 16, True, 40, "Person", 40, 0) == 110
assert calculate_amount_to_pay(100, 0,  "En tävling", 16, True, 100, "Person", 100, 0) == 0
assert calculate_amount_to_pay(150, 50, "En tävling", 16, True, 100, "Person", 100, 0) == 50
assert calculate_amount_to_pay(100, 0,  "En stafett-tävling", 16, True, 100, "Person", 100, 0) == 0
assert calculate_amount_to_pay(100, 50, "En stafett-tävling", 16, True, 100, "Person", 100, 0) == 0
assert calculate_amount_to_pay(100, 0,  "En stafett-tävling", 25, True, 100, "Person", 100, 0) == 0
assert calculate_amount_to_pay(100, 50, "En stafett-tävling", 25, True, 100, "Person", 100, 0) == 0
assert calculate_amount_to_pay(500.40, 250.20, "Sjövalla FK 1 i 25manna", 25, True, 100, "Person", 500, 0) == 0
assert calculate_amount_to_pay(100, 0,  "En tävling med justering +123 kr", 16, False, 40, "Person", 0, 123) == 223
assert calculate_amount_to_pay(100, 0,  "En tävling med justering -80 kr", 16, False, 40, "Person", 0, -80) == 20

def save_excel(df:pd.DataFrame, dfRemoved:pd.DataFrame, invoiceData, filename:str):

    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object. We also turn off the
    # index column at the left of the output dataframe.
    df.to_excel(writer, sheet_name='Aktivitetsöversikt', index=False)
    pd.DataFrame({}).to_excel(writer, sheet_name='Fakturaöversikt', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Aktivitetsöversikt']
    worksheet2 = writer.sheets['Fakturaöversikt']

    # Get the dimensions of the dataframe.
    (max_row, max_col) = df.shape

    # Make the columns wider for clarity.
    worksheet.set_column(0,  max_col - 1, 12)

    # Set the autofilter
    worksheet.autofilter(0, 0, max_row, max_col - 1)

    # Formats
    SEK = workbook.add_format({'locked': True, 'num_format': '# ##0 kr', 'bg_color': '#F1F1F1'})
    #bold_format = workbook.add_format({'bold': True})
    valid_format = workbook.add_format({'bg_color': '#C6EFCE'})
    invalid_format = workbook.add_format({'bg_color': '#FFC7CE'})
    eventor_format = workbook.add_format({'bg_color': '#F5D301'})
    eventor_mix_format = workbook.add_format({'bg_color': '#F8E201'})
    sfk_format = workbook.add_format({'bg_color': '#3F9049', 'font_color':'#FFFFFF'})
    manual_format = workbook.add_format({'bg_color': '#C5D9F1', 'font_color':'#000000'})
    date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})

    # For row cells
    dont_change_format = workbook.add_format({'locked': True, 'bg_color': '#F1F1F1'})
    editable_format = workbook.add_format({'locked': False, 'bg_color': '#FFFFFF'})

    # Set column colors / formats
    worksheet.write('A1', 'Id', eventor_format)
    worksheet.set_column('A:A', 8)
    worksheet.write('B1', 'Person_id', eventor_format)
    worksheet.set_column('B:B', 8)
    worksheet.write('C1', 'E-mail', eventor_format)
    worksheet.set_column('C:C', 27)
    worksheet.write('D1', 'Person', eventor_format)
    worksheet.set_column('D:D', 20)
    worksheet.write('E1', 'Ålder', eventor_mix_format)
    worksheet.set_column('E:E', 4)
    worksheet.write('F1', 'Datum', eventor_format)
    worksheet.set_column('F:F', 20)
    worksheet.write('G1', 'Tävling', eventor_format)
    worksheet.set_column('G:G', 30)
    worksheet.write('H1', 'Klass', eventor_format)
    worksheet.set_column('H:H', 20)
    worksheet.write('I1', 'Tjänst', eventor_format)
    worksheet.set_column('I:I', 30)
    worksheet.write('J1', 'EventTyp', eventor_format)
    worksheet.set_column('J:J', 10)
    worksheet.write('K1', 'Status', eventor_format)
    worksheet.set_column('K:K', 10)
    worksheet.write('L1', 'OK', sfk_format)
    worksheet.set_column('L:L', 10)
    worksheet.write('M1', 'Belopp', eventor_format)
    worksheet.set_column('M:M', 11)
    worksheet.write('N1', 'Efteranmälningsavgift', eventor_format)
    worksheet.set_column('N:N', 11)
    worksheet.write('O1', 'Subvention %', sfk_format)
    worksheet.set_column('O:O', 11)
    worksheet.write('P1', 'Subvention', sfk_format)
    worksheet.set_column('P:P', 11)
    worksheet.write('Q1', 'Att betala', sfk_format)
    worksheet.set_column('Q:Q', 11)
    worksheet.write('R1', 'Justering', editable_format)
    worksheet.set_column('R:R', 11)
    worksheet.write('S1', 'Notering', editable_format)
    worksheet.set_column('S:S', 50)

    # Add comments
    worksheet.write_comment('M1', 'Avgift - enligt Eventor. Avrundat till hela kronor')
    worksheet.write_comment('N1', 'Avgift för efteranmälan (subventioneras inte). Avrundat till hela kronor')
    worksheet.write_comment('K1', 'Om "Ej start" eller inte')
    worksheet.write_comment('L1', 'Om posten ska subventioneras eller inte givet SFK regler')
    worksheet.write_comment('O1', 'Subvention i procent per post enligt SFK regler')
    worksheet.write_comment('P1', 'Eventuell subvention i kr per post enligt SFK regler')
    worksheet.write_comment('Q1', 'Att betala för post efter eventuell subvention enligt SFK regler')
    worksheet.write_comment('R1', 'Ange justering i kr (minus för avdrag från "Att betala", plus för tillägg)')
    worksheet.write_comment('S1', 'Notera eventuella justeringar (inklusive ditt namn)')

    # Add backrground to cells that should not be altered
    #worksheet.set_column(f"A2:Q{max_row}", None, dont_change_format)
    #worksheet.set_column(f"R2:S{max_row}", None, editable_format)

    # If discount is ok or not
    worksheet.conditional_format(f"L2:L{max_row}", {'type':     'cell',
                                        'criteria': '==',
                                        'value':    'True',
                                        'format':   valid_format})
    worksheet.conditional_format(f"L2:L{max_row}", {'type':     'cell',
                                        'criteria': '==',
                                        'value':    'False',
                                        'format':   invalid_format})

    # Setup sheet 2
    worksheet2.write('A1', 'Fakturanummer', sfk_format)
    worksheet2.write('B1', 'Person', sfk_format)
    worksheet2.write('C1', 'Subvention (kr)', sfk_format)
    worksheet2.write('D1', 'Justering (kr)', sfk_format)
    worksheet2.write('E1', 'Totalt att betala (kr)', sfk_format)
    worksheet2.write('F1', 'Fakturanamn', sfk_format)
    worksheet2.write('G1', 'E-post', manual_format)
    worksheet2.write('H1', 'Faktura skickad', manual_format)
    worksheet2.write('I1', 'Faktura betald', manual_format)
    worksheet2.write('J1', 'Notering', manual_format)
    worksheet2.write_comment('C1', 'Total subvention (som information) för aktuell person')
    worksheet2.write_comment('D1', 'Summa för eventuella justeringar för aktuell person')
    worksheet2.write_comment('E1', 'Total, subventionerad, summa att betala för aktuell person.\n\nOBS! Om justering är utförd måste nästa skript köras innan denna summa uppdateras.', {'width': 200, 'height': 100})
    worksheet2.write_comment('H1', 'Ange när faktura mejlats ut')
    worksheet2.write_comment('I1', 'Ange när faktura har betalats in')
    worksheet2.write_comment('J1', 'Om något behöver noteras')
    worksheet2.set_column(0, 0, 12)
    worksheet2.set_column(1, 1, 22)
    worksheet2.set_column(2, 2, 12)
    worksheet2.set_column(3, 3, 12)
    worksheet2.set_column(4, 4, 16)
    worksheet2.set_column(5, 5, 15)
    worksheet2.set_column(6, 6, 32)
    worksheet2.set_column(7, 7, 12)
    worksheet2.set_column(8, 8, 12)
    worksheet2.set_column(9, 9, 50)
    idx = 0
    for name in invoiceData:
        idx = idx + 1
        worksheet2.write(idx, 0, invoiceData[name]['invoiceNo'],dont_change_format)
        worksheet2.write(idx, 1, invoiceData[name]['name'],dont_change_format)
        worksheet2.write_formula(f"C{idx+1}", f"=SUMIF(Aktivitetsöversikt!$D$2:$D${max_row+1},\"={invoiceData[name]['name']}\",Aktivitetsöversikt!$P$2:$P${max_row+1})", SEK)
        worksheet2.write_formula(f"D{idx+1}", f"=SUMIF(Aktivitetsöversikt!$D$2:$D${max_row+1},\"={invoiceData[name]['name']}\",Aktivitetsöversikt!$R$2:$R${max_row+1})", SEK)
        worksheet2.write_formula(f"E{idx+1}", f"=SUMIF(Aktivitetsöversikt!$D$2:$D${max_row+1},\"={invoiceData[name]['name']}\",Aktivitetsöversikt!$Q$2:$Q${max_row+1})+D{idx+1}", SEK)
        #worksheet2.write(idx, 2, invoiceData[name]['discount'], SEK)
        #worksheet2.write(idx, 3, invoiceData[name]['adjustment'], SEK)
        #worksheet2.write(idx, 4, invoiceData[name]['total_amount'], SEK)
        worksheet2.write(idx, 5, invoiceData[name]['invoiceName'],dont_change_format)
        worksheet2.write(idx, 6, invoiceData[name]['email'], editable_format)
        worksheet2.write(idx, 7, '', editable_format)
        worksheet2.write(idx, 8, '', editable_format)
        worksheet2.write(idx, 9, '', editable_format)

    # Set the autofilter
    worksheet2.autofilter(0, 0, idx, 7)

    protect_options = {
        'objects':               False,
        'scenarios':             False,
        'format_cells':          True,
        'format_columns':        True,
        'format_rows':           True,
        'insert_columns':        False,
        'insert_rows':           False,
        'insert_hyperlinks':     False,
        'delete_columns':        False,
        'delete_rows':           False,
        'select_locked_cells':   True,
        'sort':                  False,
        'autofilter':            True,
        'pivot_tables':          False,
        'select_unlocked_cells': True,
    }
    worksheet.protect('',protect_options)
    worksheet2.protect('',protect_options)


    if dfRemoved is not None:
        dfRemoved.to_excel(writer, sheet_name='Raderade Tävlingar', index=False)
        worksheet3 = writer.sheets['Raderade Tävlingar']
        worksheet3.set_column(0,  0, 25)
        worksheet3.set_column(1,  1, 50)

    # Close the Pandas Excel writer and output the Excel file.
    #writer.save()
    writer.close()

def create_discount(competition:str,  competition_type:str, age) -> int:
    """Calculate discount in precentage"""

    if (discount_rules == 2023):
        # Regler för subventioner - 2023

        # Relays give 100% discount
        if is_relay(competition_type):
            return 100

        # 100% when payed cash
        if paid_cash(competition):
            return 100

        # Members of age <= 16 get 100% discount for "Vårserie" and "Ungdomsnatt"
        if age <= 16 and re.search(r"vårserie|ungdomsnatt", competition, re.IGNORECASE) != None:
            return 100

        # Standard discount
        return 40

    else:
        # Tidigare regler för subventioner

        # Relays give 100% discount
        if is_relay(competition_type):
            return 100

        # 100% when payed cash
        if paid_cash(competition):
            return 100

        # Members of age < 21 get 100% discount for "Vårserie" and "DM"
        if age < 21 and re.search(r"vårserie|dm.*göteborg", competition, re.IGNORECASE) != None:
            return 100

        # Publiktävling give standard discount even if an SM    
        if re.search(r"publiktävling", competition, re.IGNORECASE) != None:
            return 40

        # SM events give 100% discount
        if re.search(r"sm, |sm sprint", competition, re.IGNORECASE) != None:
            return 100

        # Standard discount
        return 40

assert create_discount("Landehofs hösttävling", "", 10) == 40
assert create_discount("Landehofs hösttävling", "", 50) == 40
assert create_discount("25mannamedeln + Swedish League, #9 (WRE)", "", 50) == 40
assert create_discount("Sjövalla FK 1 i 25manna", "Stafett", 50) == 100
assert create_discount("Vårserien, #3", "", 10) == 100
assert create_discount("Vårserien, #3", "", 50) == 40
if (discount_rules == 2023):
    assert create_discount("SM, sprint (Swedish League, #6 + WRE)", "", 10) == 40
    assert create_discount("SM, sprint (Swedish League, #6 + WRE)", "", 50) == 40
    assert create_discount("DM, sprint, Göteborg", "", 10) == 40
    assert create_discount("DM, sprint, Göteborg", "", 50) == 40
    assert create_discount("Ungdomsnatt etapp 1", "", 10) == 100
else:
    assert create_discount("SM, sprint (Swedish League, #6 + WRE)", "", 10) == 100
    assert create_discount("SM, sprint (Swedish League, #6 + WRE)", "", 50) == 100
    assert create_discount("DM, sprint, Göteborg", "", 10) == 100
    assert create_discount("DM, sprint, Göteborg", "", 50) == 40
    
assert create_discount("DM, sprint, Halland", "", 10) == 40
assert create_discount("DM, sprint, Halland", "", 50) == 40
assert create_discount("Veteran-OL Göteborg", "", 50) == 100
assert create_discount("Skogsflicksmatch", "", 50) == 100

def create_discounts(df):
    dfDiscounts = df.drop_duplicates(subset=['Tävling', 'Datum'], keep='first')
    dfDiscounts = dfDiscounts[['Datum', 'Tävling', 'EventTyp']]
    dfDiscounts = dfDiscounts.dropna()
    dfDiscounts['Barn'] = np.vectorize(create_discount)(dfDiscounts['Tävling'], dfDiscounts['EventTyp'], 10)
    dfDiscounts['Vuxen'] = np.vectorize(create_discount)(dfDiscounts['Tävling'], dfDiscounts['EventTyp'], 45)
    return dfDiscounts

def save_discounts_xlsx(df:pd.DataFrame, filename:str):
    print("Sparar subventioner till: " + filename)
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object. We also turn off the
    # index column at the left of the output dataframe.
    df.to_excel(writer, sheet_name='Discounts', index=False)

    # Get the xlsxwriter worksheet objects.
    worksheet = writer.sheets['Discounts']

    # Make the columns wider for clarity.
    worksheet.set_column(0,  0, 20)
    worksheet.set_column(1,  1, 50)
    worksheet.set_column(2,  2, 12)
    writer.close()


def main(): 
    if len(sys.argv) <= 3:
        print(sys.argv[0] + ' <apikey> <clubid> <invoicefile.xls>')
        exit()

    apikey = sys.argv[1]
    clubid = sys.argv[2]
    filnamn = sys.argv[3]

    # Subventionsfil baserad på SFKs regler
    discountfile = filnamn.rsplit('.', 1)[0] + ' - discounts.xlsx'

    # Ladda fakturafilen som laddats hem från eventor
    print("Öppnar filen: " + filnamn)
    try:
        dfInvoices = pd.read_excel(filnamn, 'Invoices', skiprows=1)
    except:
        print("Filen innehåller inte arket 'Invoices'.")
        exit(-1)

    # Försök att ladda subventionsfilen om den finns, annars skapas den senare.
    print("Läser in subventioner från: " + discountfile)
    global dfDiscounts
    try:
        dfDiscounts = pd.read_excel(discountfile, 'Discounts', skiprows=0)
    except:
        print("Kan inte ladda subventioner, skapar nya...")

    # Skapa cache dir för eventor cache.
    try:
        print("Använder " + cache_dir + " för eventor cache")
        os.makedirs(cache_dir)
    except FileExistsError:
        pass

    # Beräkna start och slutdatum för tävlingarna i listan
    startDate = dfInvoices['Datum'].min().strftime('%Y-%m-%d')
    endDate = dfInvoices['Datum'].max() + pd.DateOffset(days=1)
    endDate = endDate.strftime('%Y-%m-%d')

    dfClubInfo = get_club_info(apikey, clubid, startDate, endDate)

    # Add additional services
    servicefile = "Tjänster.xlsx"
    print("Läser in tjänster från: " + servicefile)
    try:
        dfServices = pd.read_excel(servicefile, 'Tjänster', skiprows=0)
        dfInvoices = pd.concat([dfInvoices, dfServices], ignore_index=True)
        dfInvoices = dfInvoices.sort_values(by=['Efternamn', 'Förnamn', 'Person-id', 'Datum'])
    except:
        print("Inga extra tjänster tillgängliga.")

    # Remove O-Ringen
    dfInvoices = dfInvoices.fillna("")
    dfInvoices = dfInvoices[~dfInvoices['Tävling'].str.contains('O-Ringen')]

    # Remove entries if discounts contains an x
    dfRemoved = None
    if dfDiscounts is not None:
        dfRemoved = pd.DataFrame(columns = ['Datum', 'Tävling'])

        for row in dfDiscounts.itertuples():
            if (str(row.Barn).lower() == 'x' or str(row.Vuxen).lower() == 'x'):
                print("Tar bort tävling: " + row.Tävling)
                dfRemoved = pd.concat([dfRemoved, pd.DataFrame([[row.Datum, row.Tävling]], columns=['Datum', 'Tävling'])], ignore_index=True)
                dfInvoices = dfInvoices[~dfInvoices['Tävling'].str.contains(row.Tävling)]

    # Setup 
    dfInvoices['Ålder'] = dfInvoices.apply(lambda row: get_age(row['Födelsedatum'], row['Datum']), axis=1)
    dfInvoices['Tjänst'] = dfInvoices.apply(lambda row: translate_to_swe(row['Tjänst']), axis=1)

    if 'Type' not in dfInvoices.columns:
        dfInvoices['EventTyp'] = ""
    if 'CompetitorStatus' not in dfInvoices.columns:
        dfInvoices['Status'] = dfInvoices["Tjänst"].map(lambda x: r"Tjänst" if pd.notna(x) else r"")
    if 'E-mail' not in dfInvoices.columns:
        dfInvoices['E-mail'] = ""

    dfInvoices.columns = dfInvoices.columns.str.replace('-', '_')

    # Lägg till email, tävlingstyp och tävlingsstatus till tabellen
    personId = -1
    for row in dfInvoices.itertuples():
        if (row.Person_id != personId):
            personId = row.Person_id
            print("Start processing " + row.Förnamn + " " + row.Efternamn + " (" + str(personId) + ")")
            dfPersonStatus = get_person_results(apikey, personId, startDate, endDate)
            personInfo = dfClubInfo.loc[dfClubInfo['id'] == str(personId)].email
            if len(personInfo) == 1:
                email = personInfo.item()
            else:
                email = ""
                print("No email available for " + row.Förnamn + " " + row.Efternamn + " (" + str(personId) + ")")

        dfInvoices.at[row.Index, 'E-mail'] = email

        if (len(row.Tävling) > 0):
            resultFound = False
            for result in dfPersonStatus.itertuples():
                if row.Tävling == result.Name and row.Klass == result.Class:
                    dfInvoices.at[row.Index, 'EventTyp'] = translate_to_swe(result.Type)
                    dfInvoices.at[row.Index, 'Status'] = translate_to_swe(result.CompetitorStatus)

                    # Remove used entry to handle multi events with same name as we otherwise would
                    # match first entry again.
                    dfPersonStatus.drop(result.Index, inplace=True)
                    resultFound = True
                    break

            if (not resultFound):
                print(row.Tävling + ", " + row.Klass + " saknar resultat")

    # Kontrollera om subvention skall ges
    dfInvoices['OK'] = dfInvoices.apply(lambda row: check_ok(row['Status'], row['Tävling'], row['EventTyp']), axis=1)

    dfInvoices['Justering'] = 0
    dfInvoices['Notering'] = " "

    # Create discounts if not available now when the event type has been populated
    if dfDiscounts is None:
        dfDiscounts = create_discounts(dfInvoices)
        save_discounts_xlsx(dfDiscounts, discountfile)

    # Beräkna subventioner utifrån tävling, ålder och avgifter
    dfInvoices['Subvention %'] = np.vectorize(calculate_discount)(dfInvoices['OK'],dfInvoices['Tävling'],dfInvoices['EventTyp'],dfInvoices['Ålder'])
    dfInvoices['Subvention'] = np.vectorize(calculate_discount_amount)(dfInvoices['Belopp'],dfInvoices['Efteranmälningsavgift'],dfInvoices['Tävling'],dfInvoices['EventTyp'],dfInvoices['Ålder'],dfInvoices['OK'],dfInvoices['Subvention %'],dfInvoices['Förnamn'])
    dfInvoices['Att betala'] = np.vectorize(calculate_amount_to_pay)(dfInvoices['Belopp'],dfInvoices['Efteranmälningsavgift'],dfInvoices['Tävling'],dfInvoices['Ålder'],dfInvoices['OK'],dfInvoices['Subvention %'],dfInvoices['Förnamn'],dfInvoices['Subvention'], dfInvoices['Justering'])

    # Slå ihop för- och efternamn till en column
    dfInvoices['Person'] = dfInvoices['Förnamn'] + ' ' + dfInvoices['Efternamn']

    # Radera överflödiga kolumner
    dfInvoices.drop(['Fakturanummer', 'Text', 'Förnamn','Efternamn', 'Födelsedatum', 'Valuta', 'Arrangörer'], axis=1, inplace=True)

    #old_cols = dfInvoices.columns.values
    new_cols = ['Id', 'Person_id', 'E-mail', 'Person', 'Ålder', 'Datum', 'Tävling',
                'Klass', 'Tjänst', 'EventTyp', 'Status', 'OK', 'Belopp', 'Efteranmälningsavgift',
                'Subvention %', 'Subvention', 'Att betala', 'Justering', 'Notering']
    
    dfInvoices = dfInvoices.reindex(columns=new_cols)

    # Group by person
    grp = dfInvoices[['Person','Subvention','Att betala', 'Justering']].groupby(["Person"])

    subvention = grp['Subvention'].aggregate('sum')
    att_betala = grp['Att betala'].aggregate('sum')
    justering = grp['Justering'].aggregate('sum')

    # Object for storing money to pay per person
    invoices_data = {}

    # Loopa över all personer
    for idx, name in enumerate(dfInvoices["Person"].unique()):
        invoices_data[name] = {"name":name, "discount": subvention.loc[name], 
            "total_amount": att_betala.loc[name],
            "invoiceNo": idx+1,
            "invoiceName": f"Faktura-{idx+1}.pdf",
            "adjustment": justering.loc[name],
            "email": dfInvoices.loc[dfInvoices['Person'] == name, ['E-mail']].values[0][0]}

    # Spara resultatet
    utfil = filnamn.rsplit('.', 1)[0] + ' - result.xlsx'
    print("Sparar beräknade resultat till: " + utfil)
    save_excel(dfInvoices, dfRemoved, invoices_data, utfil)

if __name__=="__main__":
    main()
