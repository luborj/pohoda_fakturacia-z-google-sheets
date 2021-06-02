import gspread, requests, xmltodict, config
from datetime import datetime

# Pohoda XML
# Stiahnutie adresara z Pohody
xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<dat:dataPack id="001" ico="{config.ico}" application="TestAD" version = "2.0" note="Export"
 xmlns:dat="http://www.stormware.cz/schema/version_2/data.xsd"
 xmlns:lAdb="http://www.stormware.cz/schema/version_2/list_addBook.xsd"
 xmlns:ftr="http://www.stormware.cz/schema/version_2/filter.xsd"
 xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd"
>
 <dat:dataPackItem id="01" version="2.0">
   <lAdb:listAddressBookRequest version="2.0" addressBookVersion="2.0">
            <lAdb:requestAddressBook>
            </lAdb:requestAddressBook>
        </lAdb:listAddressBookRequest>
    </dat:dataPackItem>
</dat:dataPack>"""
headers = {'Content-Type': 'application/xml; charset=UTF-8', 'STW-Authorization': config.stw_code}
response = requests.post(config.api_url, data=xml, headers=headers).content
response_parse = xmltodict.parse(response)

adresar = {}
address_book = response_parse['rsp:responsePack']['rsp:responsePackItem']['lAdb:listAddressBook']['lAdb:addressbook']

i = 0
while i < len(address_book):
  company_id = address_book[i]['adb:addressbookHeader']['adb:id']
  if address_book[i]['adb:addressbookHeader']['adb:identity']['typ:address']['typ:company'] != None:
      company_name = address_book[i]['adb:addressbookHeader']['adb:identity']['typ:address']['typ:company'] # Firma
      adresar.update({company_id : company_name})
      i += 1
  else:
      company_name = address_book[i]['adb:addressbookHeader']['adb:identity']['typ:address']['typ:name'] # Osoba
      adresar.update({company_id : company_name})
      i += 1

today_date = datetime.today().strftime('%Y-%m-%d') # Datum vystavenia faktury

# Google Sheets
# vsetky data o klientoch su ulozene v Google Sheets

gc = gspread.service_account(filename='service_account.json') #Autorizacny subor
sh = gc.open(config.file_name) #Nazov suboru
sh = sh.worksheet(config.sheet_name)
list_all = sh.get_all_values()
companies = sh.col_values(1)
companies = list(dict.fromkeys(companies[1:])) #Vynechat prvy riadok tabulky - nadpisy buniek tabulky
len_loop = len(sh.col_values(3)) #Vsetky texty fakturacie

### pohoda.xml
xml_file = open("pohoda.xml", "w")
xml_header = f"""<?xml version="1.0" encoding="UTF-8"?>
<dat:dataPack id="fa002" ico="{config.ico}" application="StwTest" version="2.0" note="Import FA" xmlns:dat="http://www.stormware.cz/schema/version_2/data.xsd" xmlns:inv="http://www.stormware.cz/schema/version_2/invoice.xsd" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd">"""

xml_file.write(xml_header)

# parovanie firiem, vyber poloziek na fakturu
for company in companies:
    for pohoda_uid, pohoda_company in adresar.items():
        if company == pohoda_company:
            x = 0
            while x < len_loop:
                if (company == list_all[x][0]) and (list_all[x][4] == 'Y'):
                    xml_invoice = f"""
                                    <!-- {company} !-->
                                    <dat:dataPackItem id="AD{pohoda_uid}" version="2.0">
                                        <inv:invoice version="2.0">
                                            <inv:invoiceHeader>
                                                <inv:invoiceType>issuedInvoice</inv:invoiceType>
                                                <inv:date>{today_date}</inv:date>
                                                <inv:accounting>
                                                    <typ:ids>2</typ:ids>
                                                </inv:accounting>
                                                <inv:text>Fakturujeme Vám:</inv:text>
                                                <inv:partnerIdentity>
                                                    <typ:id>{pohoda_uid}</typ:id>
                                                </inv:partnerIdentity>
                                                <inv:paymentType>
                                                    <typ:paymentType>draft</typ:paymentType>
                                                </inv:paymentType>
                                                <inv:classificationVAT>
                                                              <typ:classificationVATType>inland</typ:classificationVATType>
                                                </inv:classificationVAT>
                                                <inv:account>
                                                    <typ:ids>FIO</typ:ids>
                                                </inv:account>
                                            </inv:invoiceHeader>
	                                        <inv:invoiceDetail>"""
                    xml_file.write(xml_invoice)
                    n = 0
                    while n < len_loop:
                        if (company == list_all[n][0]) and (list_all[n][4] == 'Y'):
                            xml_items = f"""
                                                <inv:invoiceItem>
        			                                <inv:text>{list_all[n][1]}</inv:text>
        			                                <inv:quantity>{list_all[n][3]}</inv:quantity>
                                                    <inv:rateVAT>high</inv:rateVAT>
        			                                <inv:homeCurrency>
        			                                <typ:unitPrice>{list_all[n][2]}</typ:unitPrice>
        			                                </inv:homeCurrency>
        			                            </inv:invoiceItem>"""
                            xml_file.write(xml_items)
                            n += 1
                        else:
                            n += 1
                    xml_invoice_end = f"""
                                            </inv:invoiceDetail>
                                        </inv:invoice>
                                    </dat:dataPackItem>"""
                    xml_file.write(xml_invoice_end)
                    break
                else:
                    x += 1

xml_header_end = f"""
</dat:dataPack>"""
xml_file.write(xml_header_end)
xml_file.close()

# Posli vygenerovane XML na Pohoda API
xml_data = open('pohoda.xml', 'rb')
send_api_request = requests.post(config.api_url, data=xml_data, headers=headers).content
xml_data.close()

# Zapis odpoved z API do debug.log
api_response = open("debug.log", "wb")
api_response.write(send_api_request)
api_response.close()

