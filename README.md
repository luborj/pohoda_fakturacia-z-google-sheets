# Pohoda - fakturácia z Google Sheets

Tento skript umožňuje generovať faktúry z Google Sheets.

## Google Sheets

Pre prístup do Google Sheets je potrebné mať vytvorený [API prístup](https://console.cloud.google.com/apis/). Následne súbor `service_account.json` umiestnite do adresára so skriptom.

Štruktúra tabuľky Google Sheets je následovná:

```markdown
|      Firma     |    Text   | Cena za kus | Počet kusov | Fakturovať Y/N |
| Firma 1 s.r.o. | Položka 1 | 1           | 5           | Y              |
| Firma 1 s.r.o. | Položka 2 | 100         | 1           | Y              |
| Firma 2 s.r.o. | Položka 1 | 1           | 1           | Y              |
```
*Prvý riadok v tabuľke slúži len ako nadpis.*

## Pohoda API

Script komunikuje s Pohodou prostredníctvom služby [Pohoda mServer](https://www.stormware.sk/pohoda/xml/mserver/).

## config.py

`api_url` Adresa Pohoda mServer vo formáte http://IP:port/xml

`ico = "123"` IČO účtovnej jednotky

`authorization` V hlavičke HTTP požiadavky je potrebné uviesť parameter [STW-Authorization](https://www.stormware.sk/prirucka-pohoda-online/Datova_komunikacia/POHODA_mServer_/), ktorý slúži na autentizáciu do programu. Jedná sa o zakódovaný reťazec meno:heslo do Base64.

`file_name = 'filename'`

`sheet_name = "sheetname"`
