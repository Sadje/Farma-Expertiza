import pandas as pd
from docxtpl import DocxTemplate
import os
import openpyxl
from docx2pdf import convert
import check_data


def create_errorsFile(errors):
    data_errors = pd.DataFrame(list(errors.items()), columns=['Num', 'Error'])
    data_errors.to_excel('errors.xlsx', index=False)


def read_anketa(file) -> tuple:
    errors = {}
    flag = True
    etalon_true = ['Соответствует', 'Частично соответствует', 'Не соответствует']
    path = 'Raw' + r'\\' + file                             # Raw
    wb = openpyxl.load_workbook(path, data_only=True)
    sheet1 = wb['Форма']
    sheets = wb.sheetnames[1:]
    sheets_kt = wb.sheetnames[-9:-6]
    sheets_st = wb.sheetnames[-6:-3]
    context = {'num_project': sheet1['C2'].value,
               'name_project': sheet1['C3'].value,
               'code_project': sheet1['C4'].value,
               'org_project': sheet1['C5'].value}

    temp_counter, list_results = check_data.check_types(sheet1)
    list_true = check_data.check_true_result(sheet1)
    if temp_counter != len(list_results):
        errors[sheet1['C2'].value] = ['Проблемы в Видах и/или Типах результатов']
    elif temp_counter != len(list_true):
        errors[sheet1['C2'].value].append('Проблема в соответствии результатов')
    else:
        for idx_res in range(len(list_results)):
            context[f'name_result_{idx_res + 1}'] = f'{list_results[idx_res][0]}'
            context[f'type_result_{idx_res + 1}'] = f'{list_results[idx_res][1]}'
            if list_true[idx_res] not in etalon_true:
                errors[sheet1['C2'].value].append(f'Некорректный термин соответствия для результата {idx_res + 1}')
            else:
                context[f'true_result_{idx_res + 1}'] = f'{list_true[idx_res]}'
            list_all_problems = check_data.check_problems(wb[sheets[idx_res]])
            list_all_kts = check_data.check_kt(wb[sheets_kt[idx_res]])
            list_all_sts = check_data.check_st(wb[sheets_st[idx_res]])
            if len(list_all_problems) > 0:
                for count_problems, problems in enumerate(list_all_problems):
                    context[f'problem_{idx_res + 1}_{count_problems + 1}'] = problems[0]
                    context[f'effect_{idx_res + 1}_{count_problems + 1}'] = problems[1]
                    context[f'import_{idx_res + 1}_{count_problems + 1}'] = problems[2]
            else:
                errors[sheet1['C2'].value].append('Не выбрано ни одной проблемы')
            if len(list_all_kts) > 0:
                for count_kt, kt in enumerate(list_all_kts):
                    context[f'kt_{idx_res + 1}_{count_kt + 1}'] = kt
            else:
                errors[sheet1['C2'].value].append('Не выбрано ни одной критической технологии')
            if len(list_all_sts) > 0:
                for count_st, st in enumerate(list_all_sts):
                    context[f'st_{idx_res + 1}_{count_st + 1}'] = st
            else:
                errors[sheet1['C2'].value].append('Не выбрано ни одной сквозной технологии')


    if len(errors) > 0:
        create_errorsFile(errors)
    return len(list_results), context


def find_anketa(num_project):
    for address, dirs, files in os.walk('Raw'):             # Raw
        for file in files:
            if 'KPM' in file:
                new_file = file.replace('KPM', 'КПМ')
            elif 'Lab' in file:
                new_file = file.replace('Lab', 'Лаб')
            else:
                break
            file_name = new_file[7:-5]
            if num_project == file_name:
                return read_anketa(file)


def create_pdf(file):
    pdf_file = f'{file[:-4]}pdf'
    convert('Doc\\' + file, 'PDF\\' + pdf_file)


def create_context(data):

    for key, values in data.items():
        count_results, context = find_anketa(key)
        context['zadanie'] = values[0]
        context['expert_project'] = values[1]
        if count_results > 0:
            doc = DocxTemplate(f'Template_result_{count_results}.docx')
            doc.render(context)
            doc_file = f'Zakluychenie {key}.docx'
            doc.save('Doc\\' + doc_file)
            create_pdf(doc_file)
        break



def create_dict_zadaniya() -> dict:
    dict_zadaniya = {}
    df_zadaniya = pd.read_excel('zadaniya.xlsx', sheet_name='ФАРМА')
    for i in range(df_zadaniya['Number'].shape[0]):
        dict_zadaniya[df_zadaniya['Number'][i]] = [df_zadaniya['Zadanie'][i], df_zadaniya['Expert short'][i]]

    return dict_zadaniya

def main():
    create_context(create_dict_zadaniya())


if __name__ == '__main__':
    main()
