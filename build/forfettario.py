"""Genera il modello "Forfettario 2026: registro incassi e stima tasse" (xlsx, compatibile con Google Sheets).

Uso: python3 build/forfettario.py prodotti/forfettario-2026.xlsx
"""
import sys
from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from common import (FONT, BOLD, TITLE, SUB, MUTED, INPUT_FILL, HEAD_FILL, BAND_FILL, THIN, EURO, PCT,
                    header_row, set_widths, input_cell)

ROWS = 500  # righe disponibili nel registro incassi

wb = Workbook()

# ---------------------------------------------------------------- Istruzioni
ist = wb.active
ist.title = "Istruzioni"
set_widths(ist, {"A": 3, "B": 100})
lines = [
    ("Forfettario 2026: registro incassi e stima di tasse e contributi", TITLE),
    ("Per chi ha la partita IVA in regime forfettario. Funziona in Excel e in Google Sheets.", SUB),
    ("", None),
    ("Come si usa", BOLD),
    ("1. Nel foglio Parametri compila le celle gialle: attività, gestione INPS, riduzione del 35% e aliquota dell'imposta.", FONT),
    ("2. Nel foglio Incassi registra ogni fattura incassata: data di incasso e importo sono obbligatori. La riga 5 è un esempio: sovrascrivila o cancellala.", FONT),
    ("3. Il foglio Riepilogo si aggiorna da solo: incassi per mese, quanto manca al limite di 85.000 €, stima di contributi e imposta e quanto mettere da parte.", FONT),
    ("", None),
    ("Cosa significano i colori", BOLD),
    ("Celle gialle: da compilare. Tutte le altre contengono formule: non serve toccarle.", FONT),
    ("", None),
    ("Come fa i conti", BOLD),
    ("Nel forfettario conta l'incassato, non il fatturato: il registro somma gli importi per data di incasso.", FONT),
    ("Reddito imponibile = incassi × coefficiente di redditività della tua attività.", FONT),
    ("Contributi: Gestione separata = reddito × aliquota; Artigiani/Commercianti = quota fissa + percentuale sulla parte oltre il minimale, eventualmente ridotti del 35%.", FONT),
    ("Imposta sostitutiva = (reddito imponibile − contributi) × 15% o 5%. Per semplicità si usano i contributi stimati dell'anno: nella realtà si deducono quelli effettivamente versati.", FONT),
    ("", None),
    ("Avvertenze", BOLD),
    ("È una stima per pianificare, non sostituisce il commercialista. Non considera acconti, saldi, rivalsa INPS, casse professionali private né il massimale contributivo.", FONT),
    ("I valori INPS cambiano ogni anno: quelli inseriti sono del 2026 e vanno aggiornati nel foglio Parametri.", FONT),
    ("", None),
    ("Fonti dei parametri", BOLD),
    ("Coefficienti di redditività e aliquote: Legge 190/2014, allegato 4.", FONT),
    ("Contributi INPS 2026 (quote fisse, minimale, Gestione separata 26,07%): https://quickfisco.it/blog/regime-forfettario/contributi-inps-regime-forfettario-2024-artigiani-commercianti-e-professionisti/", FONT),
]
for i, (txt, font) in enumerate(lines, start=2):
    c = ist.cell(row=i, column=2, value=txt)
    if font:
        c.font = font
    c.alignment = Alignment(wrap_text=True, vertical="top")
ist.sheet_view.showGridLines = False

# ---------------------------------------------------------------- Parametri
par = wb.create_sheet("Parametri")
par.sheet_view.showGridLines = False
set_widths(par, {"A": 3, "B": 46, "C": 56, "D": 60})
par["B2"] = "Parametri"
par["B2"].font = TITLE
par["B3"] = "Compila le celle gialle. Le altre si calcolano da sole."
par["B3"].font = SUB

rows = [
    ("Anno", 2026, None, "Anno del registro"),
    ("Attività (gruppo ATECO)", "Commercio all'ingrosso e al dettaglio", None, "Scegli dall'elenco il gruppo del tuo codice ATECO"),
    ("Coefficiente di redditività", "=INDEX($C$22:$C$30,MATCH(C6,$B$22:$B$30,0))", PCT, "Si ricava dall'attività"),
    ("Gestione INPS", "Commercianti", None, "Gestione separata, Artigiani o Commercianti"),
    ("Riduzione contributi del 35%", "No", None, "Solo Artigiani e Commercianti, se l'hai richiesta all'INPS"),
    ("Aliquota imposta sostitutiva", 0.15, PCT, "15%, oppure 5% per i primi 5 anni se hai i requisiti"),
    ("Limite di ricavi annui", 85000, EURO, "Oltre questa soglia si esce dal forfettario l'anno dopo"),
]
for i, (label, value, fmt, note) in enumerate(rows, start=5):
    par.cell(row=i, column=2, value=label).font = FONT
    c = par.cell(row=i, column=3, value=value)
    c.font = FONT
    if fmt:
        c.number_format = fmt
    if i in (5, 6, 8, 9, 10, 11):
        input_cell(c)
    par.cell(row=i, column=4, value=note).font = MUTED

dv_att = DataValidation(type="list", formula1="=$B$22:$B$30", allow_blank=False)
dv_gest = DataValidation(type="list", formula1='"Gestione separata,Artigiani,Commercianti"', allow_blank=False)
dv_sino = DataValidation(type="list", formula1='"Sì,No"', allow_blank=False)
dv_aliq = DataValidation(type="list", formula1='"0.15,0.05"', allow_blank=False)
for dv, ref in ((dv_att, "C6"), (dv_gest, "C8"), (dv_sino, "C9"), (dv_aliq, "C10")):
    par.add_data_validation(dv)
    dv.add(ref)

par["B13"] = "Contributi INPS 2026"
par["B13"].font = BOLD
inps = [
    ("Gestione separata: aliquota", 0.2607, "0.00%"),
    ("Artigiani: quota fissa annua", 4521.36, EURO),
    ("Commercianti: quota fissa annua", 4611.64, EURO),
    ("Reddito minimale", 18808, EURO),
    ("Artigiani: aliquota oltre il minimale", 0.24, PCT),
    ("Commercianti: aliquota oltre il minimale", 0.2448, "0.00%"),
]
for i, (label, value, fmt) in enumerate(inps, start=14):
    par.cell(row=i, column=2, value=label).font = FONT
    c = par.cell(row=i, column=3, value=value)
    c.number_format = fmt
    input_cell(c)
par["D14"] = "Valori 2026: aggiornali ogni anno con le circolari INPS"
par["D14"].font = MUTED

par["B21"] = "Attività"
par["C21"] = "Coefficiente"
header_row(par, 21, 2, 3)
coeff = [
    ("Industrie alimentari e delle bevande", 0.40),
    ("Commercio all'ingrosso e al dettaglio", 0.40),
    ("Commercio ambulante di prodotti alimentari e bevande", 0.40),
    ("Commercio ambulante di altri prodotti", 0.54),
    ("Costruzioni e attività immobiliari", 0.86),
    ("Intermediari del commercio", 0.62),
    ("Alloggio e ristorazione", 0.40),
    ("Attività professionali, scientifiche, tecniche, sanitarie, istruzione, servizi finanziari e assicurativi", 0.78),
    ("Altre attività economiche", 0.67),
]
for i, (label, v) in enumerate(coeff, start=22):
    par.cell(row=i, column=2, value=label).font = FONT
    par.cell(row=i, column=2).alignment = Alignment(wrap_text=True, vertical="top")
    c = par.cell(row=i, column=3, value=v)
    c.number_format = PCT
    c.font = FONT
par["B31"] = "Fonte: Legge 190/2014, allegato 4. Verifica con il commercialista il gruppo del tuo codice ATECO."
par["B31"].font = MUTED

# ---------------------------------------------------------------- Incassi
inc = wb.create_sheet("Incassi")
inc.freeze_panes = "A5"
set_widths(inc, {"A": 14, "B": 14, "C": 16, "D": 28, "E": 36, "F": 14, "G": 10, "H": 30})
inc["A1"] = "Registro incassi"
inc["A1"].font = TITLE
inc["A2"] = "Una riga per ogni incasso. Data di incasso e importo sono obbligatori."
inc["A2"].font = SUB
heads = ["Data incasso", "Data fattura", "N. fattura", "Cliente", "Descrizione", "Importo (€)", "Mese", "Note"]
for j, h in enumerate(heads, start=1):
    inc.cell(row=4, column=j, value=h)
header_row(inc, 4, 1, 8)
example = ["2026-01-15", "2026-01-10", "1/2026", "Esempio Srl", "Sostituisci o cancella questa riga di esempio", 1200, None, ""]
from datetime import date
for r in range(5, 5 + ROWS):
    for j in range(1, 9):
        c = inc.cell(row=r, column=j)
        c.font = FONT
        if j in (1, 2, 3, 4, 5, 6, 8):
            c.fill = INPUT_FILL
        c.border = THIN
    inc.cell(row=r, column=1).number_format = "dd/mm/yyyy"
    inc.cell(row=r, column=2).number_format = "dd/mm/yyyy"
    inc.cell(row=r, column=6).number_format = EURO
    inc.cell(row=r, column=7, value=f'=IF(A{r}="","",MONTH(A{r}))')
    inc.cell(row=r, column=7).alignment = Alignment(horizontal="center")
inc["A5"] = date(2026, 1, 15)
inc["B5"] = date(2026, 1, 10)
inc["C5"] = "1/2026"
inc["D5"] = "Esempio Srl"
inc["E5"] = "Sostituisci o cancella questa riga di esempio"
inc["F5"] = 1200
dv_date = DataValidation(type="date", operator="greaterThan", formula1="DATE(2000,1,1)", allow_blank=True,
                         error="Inserisci una data valida", errorTitle="Data non valida")
dv_amt = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        error="Inserisci un importo positivo", errorTitle="Importo non valido")
inc.add_data_validation(dv_date)
inc.add_data_validation(dv_amt)
dv_date.add(f"A5:B{4 + ROWS}")
dv_amt.add(f"F5:F{4 + ROWS}")

# ---------------------------------------------------------------- Riepilogo
rie = wb.create_sheet("Riepilogo")
rie.sheet_view.showGridLines = False
set_widths(rie, {"A": 3, "B": 40, "C": 18, "D": 4, "E": 16, "F": 18})
rie["B2"] = "Riepilogo dell'anno"
rie["B2"].font = TITLE
rie["B3"] = '="Stima basata sugli incassi registrati per il "&Parametri!C5'
rie["B3"].font = SUB

last = 4 + ROWS
summary = [
    ("Incassi dell'anno", f"=SUMPRODUCT((YEAR(Incassi!A5:A{last})=Parametri!C5)*(Incassi!A5:A{last}<>\"\")*Incassi!F5:F{last})", EURO),
    ("Limite del forfettario", "=Parametri!C11", EURO),
    ("Quota del limite usata", "=IF(C6=0,0,C5/C6)", PCT),
    ("Margine prima del limite", "=MAX(0,C6-C5)", EURO),
    ("Coefficiente di redditività", "=Parametri!C7", PCT),
    ("Reddito imponibile", "=C5*C9", EURO),
    ("Contributi INPS stimati", "=IF(Parametri!C8=\"Gestione separata\",C10*Parametri!C14,"
     "(IF(Parametri!C8=\"Artigiani\",Parametri!C15,Parametri!C16)"
     "+MAX(0,C10-Parametri!C17)*IF(Parametri!C8=\"Artigiani\",Parametri!C18,Parametri!C19))"
     "*IF(Parametri!C9=\"Sì\",0.65,1))", EURO),
    ("Imposta sostitutiva stimata", "=MAX(0,C10-C11)*Parametri!C10", EURO),
    ("Totale da pagare stimato", "=C11+C12", EURO),
    ("Netto stimato", "=C5-C13", EURO),
    ("Da mettere da parte su ogni incasso", "=IF(C5=0,0,C13/C5)", PCT),
]
for i, (label, formula, fmt) in enumerate(summary, start=5):
    rie.cell(row=i, column=2, value=label).font = FONT
    c = rie.cell(row=i, column=3, value=formula)
    c.number_format = fmt
    c.font = BOLD if i in (13, 14, 15) else FONT
    c.border = THIN
    rie.cell(row=i, column=2).border = THIN
rie["B17"] = "Con Artigiani e Commercianti la quota fissa è dovuta anche con incassi bassi: per questo all'inizio dell'anno la percentuale da mettere da parte risulta alta."
rie["B17"].font = MUTED
rie["B17"].alignment = Alignment(wrap_text=True, vertical="top")
rie.merge_cells("B17:C18")
rie.row_dimensions[17].height = 30

rie["E4"] = "Mese"
rie["F4"] = "Incassi (€)"
header_row(rie, 4, 5, 6)
mesi = ["Gennaio", "Febbraio", "Marzo", "Aprile", "Maggio", "Giugno", "Luglio", "Agosto", "Settembre",
        "Ottobre", "Novembre", "Dicembre"]
for k, m in enumerate(mesi, start=1):
    r = 4 + k
    rie.cell(row=r, column=5, value=m).font = FONT
    c = rie.cell(row=r, column=6, value=(
        f"=SUMPRODUCT((Incassi!A5:A{last}<>\"\")*(YEAR(Incassi!A5:A{last})=Parametri!C5)"
        f"*(MONTH(Incassi!A5:A{last})={k})*Incassi!F5:F{last})"))
    c.number_format = EURO
    c.font = FONT
    c.border = THIN
    rie.cell(row=r, column=5).border = THIN
rie.cell(row=17, column=5, value="Totale").font = BOLD
c = rie.cell(row=17, column=6, value="=SUM(F5:F16)")
c.number_format = EURO
c.font = BOLD

# Warn when close to the limit
rie.conditional_formatting.add("C7", FormulaRule(formula=["$C$7>=0.9"], font=Font(name="Arial", color="9B2C2C", bold=True)))

wb.active = wb.sheetnames.index("Istruzioni")
wb.save(sys.argv[1])
print("ok", sys.argv[1])
