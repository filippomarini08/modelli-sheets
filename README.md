# Modelli Sheets

Modelli di fogli di calcolo in italiano da vendere su Etsy e Gumroad. Funzionano in Excel e in Google Sheets.

| File | Per chi | Cosa fa |
|---|---|---|
| `prodotti/forfettario-2026.xlsx` | Partite IVA in regime forfettario | Registro incassi, limite di 85.000 €, stima di contributi INPS e imposta sostitutiva, quanto mettere da parte |
| `prodotti/margini-ecommerce.xlsx` | Chi vende online | Registro ordini con commissioni di pagamento, IVA e costi, margine per ordine e per mese |

Le schede di vendita (titolo, descrizione, tag, prezzo) sono in `listing/`.

## Rigenerare i file

I modelli si generano da script, così ogni modifica è tracciata.

```bash
python3 build/forfettario.py prodotti/forfettario-2026.xlsx
python3 build/margini.py prodotti/margini-ecommerce.xlsx
```

Serve Python 3 con `openpyxl`. Dopo la generazione le formule vanno ricalcolate aprendo il file (o con LibreOffice) prima di controllarne i valori.

## Aggiornamento annuale

Ogni gennaio vanno aggiornati nel modello forfettario i parametri INPS (quote fisse, minimale, aliquote) dalle circolari INPS dell'anno, e il nome del file.

## Controlli fatti (versione 0.1.0)

- Zero errori di formula dopo il ricalcolo con LibreOffice.
- Forfettario: con 40.000 € di incassi, attività professionale (78%) e Gestione separata, il modello dà contributi 8.133,84 € e imposta 3.459,92 €, come il calcolo a mano.
- Margini: ordine da 65 € con carta Stripe, costo merce 22 €, spedizione 6,50 €, imballo 1,20 € dà commissione 1,23 €, IVA 11,72 € e margine 22,35 €.

## Licenza

Proprietario, tutti i diritti riservati. Chi compra un modello riceve una licenza d'uso personale, non di rivendita.
