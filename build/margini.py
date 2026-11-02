"""Genera il modello "Margini e-commerce: ordini, commissioni e guadagno reale" (xlsx, compatibile con Google Sheets).

Uso: python3 build/margini.py prodotti/margini-ecommerce.xlsx
"""
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

from common import FONT, BOLD, TITLE, SUB, MUTED, INPUT_FILL, THIN, EURO, PCT, header_row, set_widths, input_cell

ROWS = 1000
LAST = 4 + ROWS
wb = Workbook()

# ---------------------------------------------------------------- Istruzioni
ist = wb.active
ist.title = "Istruzioni"
ist.sheet_view.showGridLines = False
set_widths(ist, {"A": 3, "B": 100})
lines = [
    ("Margini e-commerce: quanto guadagni davvero su ogni ordine", TITLE),
    ("Per chi vende online (Shopify, WooCommerce, Etsy, Vinted Pro, Instagram). Funziona in Excel e in Google Sheets.", SUB),
    ("", None),
    ("Come si usa", BOLD),
    ("1. Nel foglio Parametri controlla le commissioni dei tuoi metodi di pagamento e scegli se applichi l'IVA o sei forfettario.", FONT),
    ("2. Nel foglio Ordini registra ogni vendita: data, incasso, metodo di pagamento e i costi (merce, spedizione, imballo). La riga 5 è un esempio: sovrascrivila o cancellala.", FONT),
    ("3. Il foglio Riepilogo mostra mese per mese incassi, costi, commissioni, IVA e margine.", FONT),
    ("", None),
    ("Cosa significano i colori", BOLD),
    ("Celle gialle: da compilare. Le altre contengono formule.", FONT),
    ("", None),
    ("Come fa i conti", BOLD),
    ("Commissione = incasso × percentuale + quota fissa del metodo di pagamento scelto.", FONT),
    ("IVA = parte dell'incasso che va versata allo Stato (zero se sei in regime forfettario).", FONT),
    ("Margine = incasso − commissione − IVA − costo merce − costo spedizione − altri costi.", FONT),
    ("", None),
    ("Avvertenze", BOLD),
    ("Le commissioni preimpostate sono indicative: controllale sul tuo contratto e modificale nel foglio Parametri. Il margine non considera tasse sul reddito, abbonamenti e pubblicità.", FONT),
    ("Commissione Stripe per carte europee standard (1,5% + 0,25 €): https://sellfy.com/it/strumenti/calcolatore-commissioni-stripe/", FONT),
]
for i, (txt, font) in enumerate(lines, start=2):
    c = ist.cell(row=i, column=2, value=txt)
    if font:
        c.font = font
    c.alignment = Alignment(wrap_text=True, vertical="top")

# ---------------------------------------------------------------- Parametri
par = wb.create_sheet("Parametri")
par.sheet_view.showGridLines = False
set_widths(par, {"A": 3, "B": 40, "C": 16, "D": 16, "E": 50})
par["B2"] = "Parametri"
par["B2"].font = TITLE
par["B3"] = "Compila o correggi le celle gialle."
par["B3"].font = SUB
par["B5"] = "Regime IVA"
par["C5"] = "IVA ordinaria"
input_cell(par["C5"])
par["E5"] = "IVA ordinaria oppure Forfettario (IVA non dovuta)"
par["B6"] = "Aliquota IVA sulle vendite"
par["C6"] = 0.22
par["C6"].number_format = PCT
input_cell(par["C6"])
par["E6"] = "Di solito 22%"
for r in (5, 6):
    par[f"B{r}"].font = FONT
    par[f"E{r}"].font = MUTED
dv_reg = DataValidation(type="list", formula1='"IVA ordinaria,Forfettario"', allow_blank=False)
par.add_data_validation(dv_reg)
dv_reg.add("C5")

par["B9"] = "Metodo di pagamento"
par["C9"] = "Commissione %"
par["D9"] = "Quota fissa (€)"
par["E9"] = "Note"
header_row(par, 9, 2, 5)
metodi = [
    ("Carta (Stripe, carte UE standard)", 0.015, 0.25, "Listino Stripe per l'Italia, carte standard SEE"),
    ("PayPal", 0.034, 0.35, "Indicativo (tariffa nazionale standard): controlla il tuo listino"),
    ("Shopify Payments", 0.0, 0.0, "Inserisci la tariffa del tuo piano"),
    ("Satispay", 0.0, 0.0, "Inserisci la tua tariffa"),
    ("Bonifico", 0.0, 0.0, ""),
    ("Contrassegno", 0.0, 0.0, "Aggiungi il costo del corriere in Altri costi"),
    ("Marketplace (Etsy, Amazon, altro)", 0.0, 0.0, "Inserisci commissione di vendita + pagamento"),
    ("Altro", 0.0, 0.0, ""),
]
for i, (m, pct, fix, note) in enumerate(metodi, start=10):
    par.cell(row=i, column=2, value=m).font = FONT
    c = par.cell(row=i, column=3, value=pct)
    c.number_format = "0.00%"
    input_cell(c)
    c = par.cell(row=i, column=4, value=fix)
    c.number_format = EURO
    input_cell(c)
    par.cell(row=i, column=5, value=note).font = MUTED
input_cell(par["B10"])
for i in range(10, 18):
    input_cell(par.cell(row=i, column=2))

# ---------------------------------------------------------------- Ordini
o = wb.create_sheet("Ordini")
o.freeze_panes = "C5"
heads = ["Data", "N. ordine", "Canale", "Prodotto", "Q.tà", "Incasso dal cliente (€)", "Metodo di pagamento",
         "Costo merce (€)", "Costo spedizione (€)", "Altri costi (€)", "Commissione (€)", "IVA (€)", "Margine (€)",
         "Margine %", "Mese"]
widths = [12, 12, 14, 30, 7, 16, 30, 14, 14, 13, 14, 12, 14, 11, 7]
for j, (h, w) in enumerate(zip(heads, widths), start=1):
    o.cell(row=4, column=j, value=h)
    o.column_dimensions[o.cell(row=4, column=j).column_letter].width = w
header_row(o, 4, 1, 15)
o.row_dimensions[4].height = 34
o["A1"] = "Ordini"
o["A1"].font = TITLE
o["A2"] = "Una riga per ordine. L'incasso è quanto ha pagato il cliente, spedizione e IVA comprese."
o["A2"].font = SUB
for r in range(5, LAST + 1):
    for j in range(1, 16):
        c = o.cell(row=r, column=j)
        c.font = FONT
        c.border = THIN
        if j <= 10:
            c.fill = INPUT_FILL
    o.cell(row=r, column=1).number_format = "dd/mm/yyyy"
    for j in (6, 8, 9, 10, 11, 12, 13):
        o.cell(row=r, column=j).number_format = EURO
    o.cell(row=r, column=14).number_format = PCT
    o.cell(row=r, column=11, value=(
        f'=IF(OR(F{r}="",G{r}=""),0,'
        f'IFERROR(F{r}*INDEX(Parametri!$C$10:$C$17,MATCH(G{r},Parametri!$B$10:$B$17,0))'
        f'+INDEX(Parametri!$D$10:$D$17,MATCH(G{r},Parametri!$B$10:$B$17,0)),0))'))
    o.cell(row=r, column=12, value=f'=IF(OR(F{r}="",Parametri!$C$5="Forfettario"),0,F{r}*Parametri!$C$6/(1+Parametri!$C$6))')
    o.cell(row=r, column=13, value=f'=IF(F{r}="","",F{r}-K{r}-L{r}-H{r}-I{r}-J{r})')
    o.cell(row=r, column=14, value=f'=IF(OR(F{r}="",F{r}=0),"",M{r}/F{r})')
    o.cell(row=r, column=15, value=f'=IF(A{r}="","",MONTH(A{r}))')
example = [date(2026, 11, 3), "#1001", "Sito", "Felpa grigia M", 1, 65, "Carta (Stripe, carte UE standard)", 22, 6.5, 1.2]
for j, v in enumerate(example, start=1):
    o.cell(row=5, column=j, value=v)
dv_met = DataValidation(type="list", formula1="=Parametri!$B$10:$B$17", allow_blank=True)
o.add_data_validation(dv_met)
dv_met.add(f"G5:G{LAST}")
o.conditional_formatting.add(f"M5:M{LAST}", CellIsRule(operator="lessThan", formula=["0"], font=Font(name="Arial", color="9B2C2C", bold=True)))

# ---------------------------------------------------------------- Riepilogo
s = wb.create_sheet("Riepilogo")
s.sheet_view.showGridLines = False
s["B2"] = "Riepilogo per mese"
s["B2"].font = TITLE
s["B3"] = "Anno"
s["C3"] = 2026
input_cell(s["C3"])
s["B3"].font = FONT
cols = ["Mese", "Ordini", "Incassi (€)", "Commissioni (€)", "IVA (€)", "Costi (€)", "Margine (€)", "Margine %"]
for j, h in enumerate(cols, start=2):
    s.cell(row=5, column=j, value=h)
header_row(s, 5, 2, 9)
set_widths(s, {"A": 3, "B": 14, "C": 10, "D": 15, "E": 16, "F": 13, "G": 14, "H": 15, "I": 12})
mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre",
        "Ottobre", "Novembre", "Dicembre"]
rng = lambda col: f"Ordini!{col}5:{col}{LAST}"
for k, m in enumerate(mesi, start=1):
    r = 5 + k
    cond = f'(Ordini!$A$5:$A${LAST}<>"")*(YEAR(Ordini!$A$5:$A${LAST})=$C$3)*(MONTH(Ordini!$A$5:$A${LAST})={k})'
    s.cell(row=r, column=2, value=m)
    s.cell(row=r, column=3, value=f"=SUMPRODUCT({cond})")
    s.cell(row=r, column=4, value=f"=SUMPRODUCT({cond}*{rng('F')})")
    s.cell(row=r, column=5, value=f"=SUMPRODUCT({cond}*{rng('K')})")
    s.cell(row=r, column=6, value=f"=SUMPRODUCT({cond}*{rng('L')})")
    s.cell(row=r, column=7, value=f"=SUMPRODUCT({cond}*({rng('H')}+{rng('I')}+{rng('J')}))")
    s.cell(row=r, column=8, value=f"=D{r}-E{r}-F{r}-G{r}")
    s.cell(row=r, column=9, value=f"=IF(D{r}=0,0,H{r}/D{r})")
tot = 18
s.cell(row=tot, column=2, value="Totale")
for j, col in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H")):
    s.cell(row=tot, column=j, value=f"=SUM({col}6:{col}17)")
s.cell(row=tot, column=9, value=f"=IF(D{tot}=0,0,H{tot}/D{tot})")
for r in range(6, tot + 1):
    for j in range(2, 10):
        c = s.cell(row=r, column=j)
        c.font = BOLD if r == tot else FONT
        c.border = THIN
        if j in (4, 5, 6, 7, 8):
            c.number_format = EURO
        if j == 9:
            c.number_format = PCT
s["B20"] = "Scontrino medio"
s["C20"] = f"=IF(C{tot}=0,0,D{tot}/C{tot})"
s["C20"].number_format = EURO
s["B21"] = "Margine medio per ordine"
s["C21"] = f"=IF(C{tot}=0,0,H{tot}/C{tot})"
s["C21"].number_format = EURO
for r in (20, 21):
    s[f"B{r}"].font = FONT
    s[f"C{r}"].font = BOLD
s["B23"] = "Il margine non comprende tasse sul reddito, abbonamenti, pubblicità e altri costi fissi."
s["B23"].font = MUTED

wb.active = 0
wb.save(sys.argv[1])
print("ok", sys.argv[1])
