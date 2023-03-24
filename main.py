from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule

import openpyxl
import os.path


# create input.xlsx as interface
def create_input_file():
    # check if input file exists in program, if not create it
    path = './input.xlsx'
    check_file = os.path.isfile(path)

    if not check_file:
        wb = openpyxl.Workbook()
        sheet = wb['Sheet']
        sheet.title = 'Samples'
        wb.create_sheet('Codes')

        cell_value = PatternFill(start_color='00339966', end_color='00339966', fill_type='solid')
        assay_name = PatternFill(start_color='00C0C0C0', end_color='00C0C0C0', fill_type='solid')
        assay_code = PatternFill(start_color='00C0C0C0', end_color='00C0C0C0', fill_type='solid')

        samples_sheet = wb['Samples']
        codes_sheet = wb['Codes']

        # create headers for samples sheet
        samples_sheet.column_dimensions['A'].width = 20
        samples_sheet.column_dimensions['B'].width = 20
        samples_sheet.column_dimensions['C'].width = 20
        samples_sheet['A1'].fill = PatternFill(fgColor='00C0C0C0', fill_type='solid')
        samples_sheet['B1'].fill = PatternFill(fgColor='00C0C0C0', fill_type='solid')
        samples_sheet['C1'].fill = PatternFill(fgColor='00C0C0C0', fill_type='solid')
        samples_sheet['A1'] = "Extraction Group"
        samples_sheet['B1'] = "Sample ID"
        samples_sheet['C1'] = "Assay Code"
        samples_sheet['A2'] = "Paste Data Here"

        # create headers for codes sheet
        codes_sheet['A1'] = "Assay"
        codes_sheet.column_dimensions['A'].width = 20
        codes_sheet['A1'].fill = PatternFill(fgColor='00808080', fill_type='solid')

        codes_sheet['B1'] = "code1"
        codes_sheet['C1'] = "code2"
        codes_sheet['D1'] = "code3"

        codes_sheet['A2'] = "assay1"
        codes_sheet['A3'] = "assay2"
        codes_sheet['A4'] = "assay3"

        samples_sheet.conditional_formatting.add('A2:C1048576',
                                                 FormulaRule(formula=['NOT(ISBLANK(A2))'],
                                                             stopIfTrue=True,
                                                             fill=cell_value))

        codes_sheet.conditional_formatting.add('B1:ZZ1',
                                               FormulaRule(formula=['NOT(ISBLANK(B1))'],
                                                           stopIfTrue=True,
                                                           fill=assay_code))

        codes_sheet.conditional_formatting.add('A2:A5000',
                                               FormulaRule(formula=['NOT(ISBLANK(A2))'],
                                                           stopIfTrue=True,
                                                           fill=assay_name))

        last_cell_row = codes_sheet.max_row
        last_cell_col = codes_sheet.max_column

        last_cell = codes_sheet.cell(last_cell_row, last_cell_col).coordinate
        editing_area = "B2:" + last_cell

        codes_sheet.conditional_formatting.add(editing_area,
                                               FormulaRule(formula=['NOT(ISBLANK(B2))'],
                                                           stopIfTrue=True,
                                                           fill=assay_name))

        wb.save(filename='input.xlsx')


create_input_file()

# read input.xlsx
input_file = openpyxl.load_workbook('input.xlsx')

samples = input_file['Samples']
codes = input_file['Codes']

cell = samples['A3']

print(cell.value)
print(samples.max_row)


