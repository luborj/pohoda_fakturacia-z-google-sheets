import gspread, requests, xmltodict, config
from datetime import datetime

class API():
    def api_request(self, data):
        request_header = {'Content-Type': 'application/xml; charset=UTF-8', 'STW-Authorization': config.stw_code}
        self.response = requests.post(config.api_url, data=data, headers=request_header).content

    def api_response(self):
        response = xmltodict.parse(self.response)
        return response

    def api_response_plain(self):
        return self.response

class XML():
    def get_addressbook(self):
        get_addressbook_var = f"""<?xml version="1.0" encoding="UTF-8"?>
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
        return get_addressbook_var

    def invoice_header(self):
        invoice_header_var = f"""<?xml version="1.0" encoding="UTF-8"?>
                            <dat:dataPack id="fa002" ico="{config.ico}" application="StwTest" version="2.0" note="Import FA" xmlns:dat="http://www.stormware.cz/schema/version_2/data.xsd" xmlns:inv="http://www.stormware.cz/schema/version_2/invoice.xsd" xmlns:typ="http://www.stormware.cz/schema/version_2/type.xsd">"""
        return invoice_header_var

    def invoice_detail(self):
        invoice_detail_var = f"""
                                <!-- {customer_name} !-->
                                    <dat:dataPackItem id="AD{pohoda_uid}" version="2.0">
                                        <inv:invoice version="2.0">
                                            <inv:invoiceHeader>
                                                <inv:invoiceType>issuedInvoice</inv:invoiceType>
                                                <inv:date>{today}</inv:date>
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
                                                    <typ:ids>CSOB</typ:ids>
                                                </inv:account>
                                            </inv:invoiceHeader>
                                            <inv:invoiceDetail>"""
        return invoice_detail_var

    def invoice_item(self):
        invoice_item_var = f"""
                                                <inv:invoiceItem>
                                                    <inv:text>{list_all[n][1]}</inv:text>
                                                    <inv:quantity>{list_all[n][3]}</inv:quantity>
                                                    <inv:rateVAT>high</inv:rateVAT>
                                                    <inv:homeCurrency>
                                                    <typ:unitPrice>{list_all[n][2]}</typ:unitPrice>
                                                    </inv:homeCurrency>
                                                </inv:invoiceItem>"""
        return invoice_item_var

    def invoice_close(self):
        invoice_close_var = f"""
                                            </inv:invoiceDetail>
                                        </inv:invoice>
                                    </dat:dataPackItem>"""
        return invoice_close_var

    def invoice_file_close(self):
        invoice_file_close_var = f"""
                            </dat:dataPack>"""
        return invoice_file_close_var

# Vyparsovanie adresara z Pohoda API
adresar = XML().get_addressbook()
api_address_book = API()
api_address_book.api_request(adresar)
api_address_book = api_address_book.api_response()

local_address_book = {}
address_book = api_address_book['rsp:responsePack']['rsp:responsePackItem']['lAdb:listAddressBook']['lAdb:addressbook']

i = 0
while i < len(address_book):
  customer_name_id = address_book[i]['adb:addressbookHeader']['adb:id']
  if address_book[i]['adb:addressbookHeader']['adb:identity']['typ:address']['typ:company'] != None:
      customer_name_name = address_book[i]['adb:addressbookHeader']['adb:identity']['typ:address']['typ:company'] # Firma
      local_address_book.update({customer_name_id : customer_name_name})
      i += 1
  else:
      customer_name_name = address_book[i]['adb:addressbookHeader']['adb:identity']['typ:address']['typ:name'] # Osoba
      local_address_book.update({customer_name_id : customer_name_name})
      i += 1

today = datetime.today().strftime('%Y-%m-%d') # Datum vystavenia faktury

# Google Sheets spojenie
gc = gspread.service_account(filename='service_account.json') # Autorizačný súbor pre prístup k Google
sh = gc.open(config.file_name) # Nazov suboru
sh = sh.worksheet(config.sheet_name)
list_all = sh.get_all_values()
customers = sh.col_values(1)
customers = list(dict.fromkeys(customers[1:])) # Prvý riadok vynechať, to sú nadpisy stĺpcov
len_loop = len(sh.col_values(3)) # Všetky položky faktúrácie

# Pohoda XML
invoice_file = open("pohoda.xml", "w")
invoice_header = XML().invoice_header()
invoice_file.write(invoice_header)

# Parsovanie klientov z adresára ktorý sme si stiahli v prvom kroku
# Následne ich spárujeme s Google Sheet
# Pridáme položky z Google Sheets, zapíšeme do Pohoda XML
for customer_name in customers:
    for pohoda_uid, pohoda_customer_name in local_address_book.items():
        if customer_name == pohoda_customer_name:
            x = 0
            while x < len_loop:
                if (customer_name == list_all[x][0]) and (list_all[x][4] == 'Y'):
                    invoice_file.write(XML().invoice_detail())
                    n = 0
                    while n < len_loop:
                        if (customer_name == list_all[n][0]) and (list_all[n][4] == 'Y'):
                            invoice_file.write(XML().invoice_item())
                            n += 1
                        else:
                            n += 1
                    invoice_file.write(XML().invoice_close())
                    break
                else:
                    x += 1
invoice_file.write(XML().invoice_file_close())
invoice_file.close()

# Odošli Pohoda XML na API
invoice_xml = open('pohoda.xml', 'rb')
send_invoice = API()
send_invoice.api_request(invoice_xml)
invoice_xml.close()

# Zapíš odpoveď z API do súboru debug.log
debug_log = open("debug.log", "wb")
debug_log.write(send_invoice.api_response_plain())
debug_log.close()