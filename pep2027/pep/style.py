"""Formatierungs-Helfer (Arial, Farben der PEP-Planung)."""
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo, TableFormula

FONT = "Arial"
THIN = Side(style="thin", color="A6A6A6")
THIN_DARK = Side(style="thin", color="000000")
MED = Side(style="medium", color="000000")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOX_DARK = Border(left=THIN_DARK, right=THIN_DARK, top=THIN_DARK, bottom=THIN_DARK)
NOBORDER = Border()

INPUT_FILL = "FFF2CC"      # helles Gelb = Eingabefeld
CALC_FONT = "000000"
LINK_FONT = "0563C1"


def fill(hex_):
    return PatternFill("solid", start_color="FF" + hex_, end_color="FF" + hex_) if hex_ else PatternFill()


def font(size=10, bold=False, color="000000", italic=False, underline=None, name=FONT):
    return Font(name=name, size=size, bold=bold, color="FF" + color, italic=italic, underline=underline)


def put(ws, ref, value=None, *, size=10, bold=False, color="000000", bg=None, italic=False, h=None, v="center",
        wrap=False, rot=0, nf=None, border=None, underline=None, indent=0):
    c = ws[ref]
    if value is not None:
        c.value = value
    c.font = font(size, bold, color, italic, underline)
    if bg:
        c.fill = fill(bg)
    c.alignment = Alignment(horizontal=h, vertical=v, wrap_text=wrap, text_rotation=rot, indent=indent)
    if nf:
        c.number_format = nf
    if border is not None:
        c.border = border
    return c


def style_range(ws, rng, **kw):
    for row in ws[rng]:
        for c in row:
            put(ws, c.coordinate, **kw)


def title(ws, ref, text, size=16, bg=None, color="000000"):
    return put(ws, ref, text, size=size, bold=True, bg=bg, color=color, v="center")


def note(ws, ref, text, size=9, color="595959", italic=True, wrap=False):
    return put(ws, ref, text, size=size, color=color, italic=italic, wrap=wrap, v="top" if wrap else "center")


def link(ws, ref, target_sheet, text, cell="A1", size=10, bold=False):
    c = ws[ref]
    tgt = target_sheet.replace("'", "''")
    c.value = f'=HYPERLINK("#\'{tgt}\'!{cell}","{text}")'
    c.font = font(size, bold, LINK_FONT, underline="single")
    c.alignment = Alignment(vertical="center")
    return c


def dxf(bg=None, fg=None, bold=None, italic=None, underline=None, border=None):
    f = None
    if fg or bold is not None or italic is not None or underline:
        f = Font(color=("FF" + fg) if fg else None, bold=bold, italic=italic, underline=underline)
    fl = PatternFill(bgColor="FF" + bg, fill_type="solid") if bg else None
    return DifferentialStyle(font=f, fill=fl, border=border)


def cf(ws, rng, formula, style, stop=True):
    rule = FormulaRule(formula=[formula], stopIfTrue=stop)
    rule.dxf = style
    ws.conditional_formatting.add(rng, rule)


def add_table(ws, name, ref, style="TableStyleMedium2", formulas=None, stripes=True):
    """Excel-Tabelle anlegen. formulas: {Spaltenname: Formel ohne '='} -> berechnete Spalte."""
    t = Table(displayName=name, ref=ref)
    t.tableStyleInfo = TableStyleInfo(name=style, showRowStripes=stripes, showColumnStripes=False)
    t._initialise_columns()
    header = ws[ref.split(":")[0]].row
    first_col = ws[ref.split(":")[0]].column
    for i, tc in enumerate(t.tableColumns):
        v = ws.cell(header, first_col + i).value
        if not isinstance(v, str) or not v:
            raise ValueError(f"Tabellenkopf fehlt: {name} Spalte {i + 1}")
        tc.name = v
    if formulas:
        for tc in t.tableColumns:
            if tc.name in formulas:
                tc.calculatedColumnFormula = TableFormula(attr_text=formulas[tc.name])
    ws.add_table(t)
    return t


def dv_list(ws, rng, source, allow_blank=True, error=True, prompt=None, title_=None):
    dv = DataValidation(type="list", formula1=source, allow_blank=allow_blank, showDropDown=False)
    dv.showErrorMessage = error
    dv.errorStyle = "stop" if error else "information"
    if not error:
        dv.showErrorMessage = True
        dv.errorStyle = "information"
        dv.error = "Wert ist nicht in der Liste. Er wird trotzdem übernommen."
        dv.errorTitle = "Hinweis"
    else:
        dv.error = "Bitte einen Wert aus der Liste wählen."
        dv.errorTitle = "Ungültige Eingabe"
    if prompt:
        dv.showInputMessage = True
        dv.prompt = prompt
        dv.promptTitle = title_ or "Eingabe"
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv


def dv_date(ws, rng):
    dv = DataValidation(type="date", operator="between", formula1="DATE(2000,1,1)", formula2="DATE(2100,12,31)",
                        allow_blank=True)
    dv.showErrorMessage = True
    dv.error = "Bitte ein gültiges Datum eingeben (TT.MM.JJJJ)."
    dv.errorTitle = "Datum"
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv


def widths(ws, mapping):
    for k, w in mapping.items():
        ws.column_dimensions[k].width = w


def q(sheet):
    """Blattnamen fuer Formeln quoten."""
    return "'" + sheet.replace("'", "''") + "'"


def landscape(ws, paper="A4", fit_w=1, fit_h=1, area=None, title_rows=None, title_cols=None, margins=0.4):
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3 if paper == "A3" else ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = fit_w
    ws.page_setup.fitToHeight = fit_h
    ws.page_margins.left = ws.page_margins.right = margins
    ws.page_margins.top = ws.page_margins.bottom = margins + 0.1
    ws.page_margins.header = ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True
    if area:
        ws.print_area = area
    if title_rows:
        ws.print_title_rows = title_rows
    if title_cols:
        ws.print_title_cols = title_cols
    ws.oddFooter.left.text = "&A"
    ws.oddFooter.right.text = "Seite &P von &N – Stand &D"
    ws.oddFooter.left.size = ws.oddFooter.right.size = 8


def protect(ws, password=None, unlocked=(), autofilter=True):
    """Blattschutz: Formeln gegen versehentliches Ueberschreiben sichern, Eingabezellen bleiben frei."""
    from openpyxl.styles import Protection
    ws.protection.sheet = True
    if password:
        ws.protection.password = password
    ws.protection.formatColumns = False
    ws.protection.formatRows = False
    ws.protection.formatCells = True
    ws.protection.autoFilter = not autofilter
    ws.protection.sort = True
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    for rng in unlocked:
        for row in ws[rng] if ":" in rng else [[ws[rng]]]:
            for c in row:
                c.protection = Protection(locked=False)
