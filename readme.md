# Fakturor för Sjövalla FK

## Skapa en fakturaexport från Eventor

- Skapa fakturaperiod i Eventor till exempel mellan 1/1-1/8 (Eventor tar inte med sista dagen, dvs 31/7 effektivt).
- Exporetera fakturorna till excel genom att välja redigera och sedan "Exportera till Excel" (Exporten kan ta lite tid).

## Skapa ett fakturaunderlag med subventioner

För att kunna skapa ett subventionerat fakturaunderlag så måste man ha följande:
- Fakturaexport från Eventor
- Eventor API nyckel
- Eventor klubbid (SFK = 321)
- Senaste versionen av beräkningsskriptet, https://github.com/m-alinder/eventor-invoices

```
sfk_update_costs_from_eventor_xls.py [-h] -i INFILE -a APIKEY -c CLUBID [-o OUTFILE] [-e EXTRAS] [-d DISCOUNTS] [-n NUMBER] 
```
Skriptet tar ett antal arguement:
- -i, används för att ange infil. -i kan användas flera gånger om man har flera exporter från Eventor som man vill ha med i fakturorna.
- -a, klubbens Eventor api-nyckel så att tävlingsresultat kan extraheras från eventor.
- -c, klubbens eventorklubbid
- -o, exceloutput från skriptet. En xlsx fil med beräknade subventioner.
- -d, subventionsfil i xlsx format att använda. Skapas om den inte redan existerar.
- -n, numret på första fakturan


Kör skriptet för att generera fakturaunderlag:
``` 
python3 sfk_update_costs_from_eventor_xls.py -a <api-key> -c <clubid> -i <eventor-fakturafil.xls> -d <subventioner.xlsx> -o <resultat.xlsx>
```

Skriptet hämtar tävlings- och personinformation från eventor och skapar två filer:
- <subventioner.xlsx>
- <resultat.xlsx>

All data som hämtats från eventor sparas i mappen "eventor-cache" och återanvänds vid nästkommande körningar.

### Granska subventioner
I subventionsfilen listas alla tävlingar och man kan där justera subventionerna i procent för en specifik tävling. Man kan också markera subventionen med ett "x" om man vill att tävlingen inte skall komma med i fakturorna.

När subventionsfilen har updaterats kan man köra skriptet igen, med samma argument. Eftesom subvetionsfilen nu existerar så används värdena från denna till att skapa en ny resultatfil.

### Extraposter
Vill man lägga till extraposter i de slutliga fakturorna så kan man skapa en extrafil i samma format som resultatfilen i vilken man lägger in de poster som man vill ska komma med.

Dessa extraposter läggs till i resultatfilen om man kör skriptet ytterligare en gång och inkluderar "-e extrafil.xlsx".

## Granska resultatfilen
Kontrollera att kostnader och subventioner i resultatfilen blivit korrekta. Eventuella felaktiga kostnader från eventor bör justeras i fakturaexporten från eventor, eftersom man då enkelt kan köra skriptet igen och få med de nya kostnaderna.

När allt ser bra ut i resultatfilens "aktivitetsöversikt" så kontrollera att kostnaderna under fliken "fakturaöversikt" är updaterade. Kör en manuell omräkning av värden.


## Generera fakturor
Innan fakturorna genereras kontrollera att kostnaderna i "fakturaöverikten" är updaterade.

Kör scriptet:
```
python3 sfk_create_pdfs_from_xlsx.py "resultat.xlsx" fakturor/ "Förnamn Efternam" telefonnummer email@olklubb.se
```

Vill man ändra betalningsvillkoren till något annat än 30 dagar så måste det göras manuellt i filen pythonlib/SFKInvoice.py genom att ändra raden "betalningsvillkor = 30" till de antal dagar man vill ha.

När skriptet kört klart så ligger alla fakturorna i mappen "fakturor/"

## Skicka ut fakturor

TODO
