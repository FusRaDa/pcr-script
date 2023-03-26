from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation, DataValidationList

import openpyxl
import os.path

# check if input file exists in program, if not create it
input_path = './input.xlsx'
input_xlsx = os.path.isfile(input_path)


# create input.xlsx as interface
def create_input_file():

    wb = openpyxl.Workbook()
    sheet = wb['Sheet']
    sheet.title = 'Samples'
    wb.create_sheet('Codes')

    # create default samples sheet
    cell_value = PatternFill(start_color='00339966', end_color='00339966', fill_type='solid')

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

    samples_sheet.conditional_formatting.add('A2:C5000',
                                             FormulaRule(formula=['NOT(ISBLANK(A2))'],
                                                         stopIfTrue=True,
                                                         fill=cell_value))

    # create default codes sheet
    assay_name = PatternFill(start_color='00C0C0C0', end_color='00C0C0C0', fill_type='solid')
    assay_code = PatternFill(start_color='00C0C0C0', end_color='00C0C0C0', fill_type='solid')

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

    for row in range(2, last_cell_row + 1):
        for col in range(2, last_cell_col + 1):
            codes_sheet.cell(row=row, column=col).value = "N"

    last_cell = codes_sheet.cell(last_cell_row, last_cell_col).coordinate
    editing_area = "B2:" + last_cell

    no = PatternFill(start_color='00FF0000', end_color='00FF0000', fill_type='solid')
    yes = PatternFill(start_color='0000FF00', end_color='0000FF00', fill_type='solid')

    codes_sheet.conditional_formatting.add(editing_area,
                                           FormulaRule(formula=['IF(B2="N", "True", "")'],
                                                       stopIfTrue=True,
                                                       fill=no))

    codes_sheet.conditional_formatting.add(editing_area,
                                           FormulaRule(formula=['IF(B2="Y", "True", "")'],
                                                       stopIfTrue=True,
                                                       fill=yes))

    dv = DataValidation(type='list', formula1='"Y,N"', allow_blank=False, showDropDown=False, showErrorMessage=True)
    dv.error = "Invalid Entry - Use Dropdown"

    dv.add(editing_area)

    codes_sheet.add_data_validation(dv)

    print("input file created")

    wb.save(input_path)


def update_codes_sheet():

    input_file = openpyxl.load_workbook(input_path)
    codes_sheet = input_file['Codes']

    max_row = codes_sheet.max_row
    max_col = codes_sheet.max_column

    for row in range(2, max_row + 1):
        for col in range(2, max_col + 1):
            if codes_sheet.cell(row=row, column=col).value is None:
                codes_sheet.cell(row=row, column=col).value = "N"

    last_cell = codes_sheet.cell(max_row, max_col).coordinate
    editing_area = "B2:" + last_cell

    no = PatternFill(start_color='00FF0000', end_color='00FF0000', fill_type='solid')
    yes = PatternFill(start_color='0000FF00', end_color='0000FF00', fill_type='solid')

    codes_sheet.conditional_formatting.add(editing_area,
                                           FormulaRule(formula=['IF(B2="N", "True", "")'],
                                                       stopIfTrue=True,
                                                       fill=no))

    codes_sheet.conditional_formatting.add(editing_area,
                                           FormulaRule(formula=['IF(B2="Y", "True", "")'],
                                                       stopIfTrue=True,
                                                       fill=yes))

    # clear all data validation
    codes_sheet.data_validations = DataValidationList()

    dv = DataValidation(type='list', formula1='"Y,N"', allow_blank=False, showDropDown=False, showErrorMessage=True)
    dv.error = "Invalid Entry - Use Dropdown"

    dv.add(editing_area)

    codes_sheet.add_data_validation(dv)

    print("codes sheet updated")

    input_file.save(input_path)


if not input_xlsx:
    create_input_file()

if input_xlsx:
    update_codes_sheet()

