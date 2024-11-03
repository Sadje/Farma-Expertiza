import pandas as pd


def check_types(sheet1):
    all_result = []

    for row in sheet1['B8':'C10']:
        list_result = []
        for cell in row:
            if cell.value is not None:
                list_result.append(cell.value)
        if len(list_result) > 0:
            all_result.append(list_result)
    start_count = len(all_result)
    if len(all_result) > 0:
        df = pd.read_excel('type_results.xlsx')
        for l_res in all_result:
            string_res = ''.join(l_res)
            for i in range(df['Name'].shape[0]):
                if string_res == df['Name'][i] + df['Type'][i]:
                    break
            else:
                all_result.remove(l_res)

    return start_count, all_result


def check_problems(sheet_problems):
    list_all_problems = []
    for row in sheet_problems['B7':'D49']:
        list_problems = []
        for cell in row:
            if cell.value is not None:
                list_problems.append(cell.value)
        if len(list_problems) == 3:
            list_all_problems.append(list_problems)

    return list_all_problems


def check_kt(sheet_kts):
    list_kts = []
    for row in sheet_kts['D7': 'D10']:
        for cell in row:
            if cell.value is not None:
                if cell.value is False:
                    list_kts.append('Нет')
                elif cell.value is True:
                    list_kts.append('Да')
                else:
                    list_kts.append(cell.value)

    return list_kts


def check_st(sheet_sts):
    list_sts = []
    for row in sheet_sts['D7': 'D10']:
        for cell in row:
            if cell.value is not None:
                if cell.value is False:
                    list_sts.append('Нет')
                elif cell.value is True:
                    list_sts.append('Да')
                else:
                    list_sts.append(cell.value)

    return list_sts


def check_true_result(sheet_true) -> list:
    list_true = []
    for row in sheet_true['I8':'I10']:
        for cell in row:
            if cell.value is not None:
                list_true.append(cell.value)
    return list_true


def check_conc(sheet_conc):
    list_conc = []
    for row in sheet_conc['J8':'J10']:
        for cell in row:
            if cell.value is not None and cell.value not in list_conc:
                list_conc.append(cell.value)

    return ". ".join(list_conc)


def check_ugt(sheet_ugt):
    flag = False
    list_tasks = []
    for row in sheet_ugt['D7':'D61']:
        for cell in row:
            if cell.value is not None:
                list_tasks.append(cell.value)
    if list_tasks.count(True) > 0:
        flag = True
    return flag


def check_doc(sheet_doc):
    flag = False
    list_docs = []
    for row in sheet_doc['D7':'D61']:
        for cell in row:
            if cell.value is not None:
                list_docs.append(cell.value)
    if list_docs.count(True) > 0:
        flag = True
    return flag
