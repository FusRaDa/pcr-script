import json
import openpyxl


def get_codes_sheet_json():
    input_file = openpyxl.load_workbook('input.xlsx')

    codes_sheet = input_file['Codes']

    max_col = codes_sheet.max_column
    max_row = codes_sheet.max_row

    full_data = []

    code_keys = []
    assay_keys = []

    for col in codes_sheet.iter_cols(min_row=1, min_col=2, max_col=max_col, max_row=1):
        for cell in col:
            code_keys.append(cell.value)

    for row in codes_sheet.iter_rows(min_row=2, min_col=1, max_col=1, max_row=max_row):
        for cell in row:
            assay_keys.append(cell.value)

    code_index = 0
    for col in codes_sheet.iter_cols(min_row=2, min_col=2, max_col=max_col, max_row=max_row):
        data_dict = {'code': code_keys[code_index]}
        code_index += 1

        assay_index = 0
        for cell in col:
            data_dict[assay_keys[assay_index]] = cell.value
            assay_index += 1

        full_data.append(data_dict)

    return full_data

    # file_name = 'assay_data.json'
    # with open(file_name, 'w') as empty_json:
    #     json.dump(full_data, empty_json, indent=4)
    #     empty_json.close()

print(get_codes_sheet_json())


def get_samples_list():
    # read input.xlsx
    input_file = openpyxl.load_workbook('input.xlsx')

    samples = input_file['Samples']

    ext_array = []
    sample_array = []
    assay_array = []

    samples_list = []

    for row in range(2, samples.max_row + 1):
        ext_array.append(samples.cell(row=row, column=1).value)

    for row in range(2, samples.max_row + 1):
        sample_array.append(samples.cell(row=row, column=2).value)

    for row in range(2, samples.max_row + 1):
        assay_array.append(samples.cell(row=row, column=3).value)

    for i in range(len(ext_array)):
        dict_obj = {'ext': ext_array[i], 'id': sample_array[i], 'code': assay_array[i]}
        samples_list.append(dict_obj)

    return samples_list