import argparse
import pdfplumber
import pandas as pd
import re

def read_pdf_files_in_directory(directory, outfile):
    """
    Läser in alla PDF-filer i en given katalog och skriver ut deras innehåll rad för rad.

    Args:
        directory (str): Sökvägen till katalogen.
        outfile (str): Filnamn för resultat.
    """

    import os

    dfError = pd.DataFrame(columns = ['Filnamn'])
    df = pd.DataFrame(columns = ['Datum', 'Tävling', 'Namn', 'Klass', 'Tjänst', 'Avgift'])

    for filename in os.listdir(directory):
        if filename.endswith('.pdf'):
            #print(f"Processing file: {filename}")
            file_path = os.path.join(directory, filename)
            
            try:
                with pdfplumber.open(file_path) as pdf:
                    event = ""
                    eventDate = ""

                    for page in pdf.pages:

                        #print(page.extract_text(x_tolerance=2))
                        text = page.extract_text(x_tolerance=2)
                        # Skriv ut texten rad för rad
                        for line in text.split("\n"):

                            # Anmälan för <namn> i klass ...
                            pattern = r"^.* f.r (.+) in? (.+) \d+,?\d* \d+,?\d* SEK (\d+,?\d*) SEK$"
                            match = re.match(pattern, line)
                            if match:
                                name, clazz, cost = match.groups()
                                df = pd.concat([df, pd.DataFrame([[eventDate, event, name, clazz, "", cost]], columns=['Datum', 'Tävling', 'Namn', 'Klass', 'Tjänst', 'Avgift'])], ignore_index=True)
                                #print(f"{eventDate}, {event}, {name}, {clazz}, {cost}, anmälan")
                                continue

                            # Hyrbricka SportIdent för <namn>
                            pattern = r"^(.+) f.r (.+) \d+,?\d* \d+,?\d* SEK (\d+,?\d*) SEK$"
                            match = re.match(pattern, line)

                            if match:
                                subject, name, cost = match.groups()
                                df = pd.concat([df, pd.DataFrame([[eventDate, event, name, "", subject, cost]], columns=['Datum', 'Tävling', 'Namn', 'Klass', 'Tjänst', 'Avgift'])], ignore_index=True)
                                #print(f"{eventDate}, {event}, {name}, {subject}, {cost}, tjänst")
                                continue

                            # Tävling <tävlingsnamn>
                            pattern = r"^Tävling +(.+)$"
                            match = re.match(pattern, line)
                            if match:
                                #print(line)
                                event = match.group(1)
                                continue

                            # Tävlingsdatum <datum>
                            pattern = r"^Tävlingsdatum (.+)$"
                            match = re.match(pattern, line)
                            if match:
                                #print(line)
                                eventDate = match.group(1)
                                if (event == ""):
                                    event = filename[11:-4]
                                continue
                    
                    if (event == "" or eventDate == ""):
                        print(f"Kunde inte analysera: {filename}")
                        dfError = pd.concat([dfError, pd.DataFrame([[filename]], columns=['Filnamn'])])

            except Exception as e:
                print(f"Error processing {filename}: {e}")

    writer = pd.ExcelWriter(outfile, engine='xlsxwriter')

    # Sort the table
    df = df.sort_values(by=['Datum', 'Tävling', 'Namn'])

    # Convert the dataframe to an XlsxWriter Excel object. We also turn off the
    # index column at the left of the output dataframe.
    df.to_excel(writer, sheet_name='Deltagare', index=False)

    # Get the xlsxwriter workbook and worksheet objects.
    #workbook  = writer.book
    worksheet = writer.sheets['Deltagare']
    worksheet.set_column(0,  0, 12)
    worksheet.set_column(1,  1, 50)
    worksheet.set_column(2,  2, 30)
    worksheet.set_column(3,  3, 10)
    worksheet.set_column(4,  4, 30)

    dfSummary = df.drop(df[df.Klass == ""].index).groupby(['Datum', 'Tävling']).size().reset_index(name='Deltagare')
    dfSummary.to_excel(writer, sheet_name='Tävlingar', index=False)
    worksheet = writer.sheets['Tävlingar']
    worksheet.set_column(0,  0, 12)
    worksheet.set_column(1,  1, 50)
    worksheet.set_column(2,  2, 10)


    dfAmount = df.drop(df[df.Klass == ""].index).groupby(['Datum', 'Tävling', 'Avgift']).size().reset_index(name='Deltagare')
    dfAmount['Nummer'] = dfAmount['Avgift'].str.replace(',', '.')
    dfAmount['Nummer'] = dfAmount['Nummer'].astype(float)
    dfAmount = dfAmount.sort_values(by=['Datum', 'Tävling', 'Nummer'])
    dfAmount = dfAmount.drop('Nummer', axis=1)
    dfAmount.to_excel(writer, sheet_name='Avgifter', index=False)
    worksheet = writer.sheets['Avgifter']
    worksheet.set_column(0,  0, 12)
    worksheet.set_column(1,  1, 50)
    worksheet.set_column(2,  2, 10)
    worksheet.set_column(3,  3, 10)

    dfService = df.drop(df[df.Tjänst == ""].index).groupby(['Datum', 'Tävling', 'Tjänst', 'Avgift']).size().reset_index(name='Antal')
    dfService.to_excel(writer, sheet_name='Tjänster', index=False)
    worksheet = writer.sheets['Tjänster']
    worksheet.set_column(0,  0, 12)
    worksheet.set_column(1,  1, 50)
    worksheet.set_column(2,  2, 30)
    worksheet.set_column(3,  4, 10)

    # If result file was given. Add sheet with missing entries
    if (args.result_file):
        xls = pd.ExcelFile(args.result_file)
        dfRes = pd.read_excel(xls, 'Aktivitetsöversikt')

        dfEventorOnly = dfRes[['Datum', 'Person','Tävling']]
        dfEventorOnly = dfEventorOnly.rename(columns={"Person": "Namn"})

        dfRes = dfRes[['Person','Tävling', 'Belopp']]
        dfRes = dfRes.rename(columns={"Person": "Namn", "Belopp": "Eventor"})
        dfRes["Eventor"] = dfRes["Eventor"].astype(str).str.replace(',', '.')
        dfRes["Eventor"] = pd.to_numeric(dfRes["Eventor"])
        dfRes = dfRes.sort_values(['Namn', "Tävling"])

        dfPdfs = df.sort_values(['Namn', "Tävling"])
        dfPdfs = dfPdfs[dfPdfs['Tjänst'].str.len().lt(1)]
        dfPdfs["Avgift"] = dfPdfs["Avgift"].astype(str).str.replace(',', '.')
        dfPdfs["Avgift"] = pd.to_numeric(dfPdfs["Avgift"])

        dfMissing = dfPdfs.merge(dfRes.drop_duplicates(), on=['Tävling', 'Namn'], 
                    how='left', indicator=True)
        
        dfMissing['AvgiftOk'] = dfMissing['Avgift'] == dfMissing['Eventor']
        #dfMissing = dfMissing[dfMissing['_merge'] == 'left_only']
        dfMissing = dfMissing[dfMissing['AvgiftOk'] == False]
        dfMissing = dfMissing.sort_values(['Datum'])
        dfMissing.reset_index(drop=True, inplace=True)
        dfMissing.drop(columns=['_merge', 'Tjänst', 'AvgiftOk'], inplace=True)

        dfMissing.to_excel(writer, sheet_name='AvgifterAttGranska', index=False)

        # Get the xlsxwriter workbook and worksheet objects.
        #workbook  = writer.book
        worksheet = writer.sheets['AvgifterAttGranska']
        worksheet.set_column(0,  0, 12)
        worksheet.set_column(1,  1, 50)
        worksheet.set_column(2,  2, 30)
        worksheet.set_column(3,  3, 15)
        worksheet.set_column(4,  4, 15)
        worksheet.set_column(5,  5, 15)

        dfEventorOnly = dfEventorOnly.merge(dfPdfs.drop_duplicates(), on=['Tävling', 'Namn'], 
                    how='left', indicator=True)
        dfEventorOnly = dfEventorOnly[dfEventorOnly['_merge'] == 'left_only']
        dfEventorOnly = dfEventorOnly.rename(columns={"Datum_x": "Datum"})
        dfEventorOnly = dfEventorOnly[['Datum', 'Namn', 'Tävling']]
        dfEventorOnly = dfEventorOnly.sort_values(['Datum'])
        dfEventorOnly = dfEventorOnly[dfEventorOnly['Tävling'].str.len().gt(1)]

        dfEventorOnly.to_excel(writer, sheet_name='BaraIEventor', index=False)

        # Get the xlsxwriter workbook and worksheet objects.
        #workbook  = writer.book
        worksheet = writer.sheets['BaraIEventor']
        worksheet.set_column(0,  0, 24)
        worksheet.set_column(1,  1, 30)
        worksheet.set_column(2,  2, 50)


    dfError.to_excel(writer, sheet_name='Fel Format', index=False)
    worksheet = writer.sheets['Fel Format']
    worksheet.set_column(0,  0, 60)

    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
    return df

parser = argparse.ArgumentParser()
parser.add_argument("input_directory", type=str, help="Directory with received eventor invoices")
parser.add_argument("output_file", type=str, help="Xlsx file to save result to")
parser.add_argument('-r', '--result-file', type=str, help='Result file to compare pdf content with')
args = parser.parse_args()

directory_path = "eventor-pdfs"  # Byt ut med din faktiska sökväg
dfPdfs = read_pdf_files_in_directory(args.input_directory, args.output_file)
