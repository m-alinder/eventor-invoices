# Fakturor för Sjövalla FK

## Generera underlag

- Exportera aktuell medlemsfil från Eventor https://eventor.orientering.se/OrganisationAdmin/Members?organisationId=321https://eventor.orientering.se/OrganisationAdmin/ExportMembersToExcel/321?year=2023
Öppna filen i Excel och spara sedan ner som 97-2004 format. Annars är det bara en XML-fil som koden just nu inte hanterar.
Sparad som: files/members_sjovalla_fk_2023_converted.xls

- Skapa fakturaperiod i Eventor till exempel mellan 1/1-1/8 (Eventor tar inte med sista dagen, dvs 31/7 effektivt). Id för denna fakturaperiod: 1103

- Exporetera fakturorna till excel genom att välja redigera och sedan "Exportera till Excel" (Exporten kan ta lite tid).

- Kör skriptet för att generera fakturaunderlag: *python3 sfk_update_costs_from_eventor_xls.py \<api-key> \<clubid> \<fakturafil.xls>*

- Skriptet hämtar tävlingsinformation från eventor och skapar två filer:
1. "fakturafil - discounts.xlsx"
1. "fakturafil - result.xlsx"

- Vill man justera subventionerna så updatera "fakturafil - discounts.xlsx" och kör skriptet igen.

- Vill man lägga till manuella poster så kan man göra detta i filen "Tjänster.xlsx" och kör scriptet igen.

## Gå igenom avgifter
Kontrollera resultatfilen.

## Generera fakturor

Kör scriptet: *python3 sfk_create_pdfs_from_xlsx.py "fakturafil - result.xlsx" fakturor/ "Förnamn Efternam" telefonnummer email@olklubb.se*

## Skicka ut fakturor
