import pandas as pd


def check_types(sheet1):
    all_result = []

    for row in sheet1['B8':'C10']:
        list_result = []
        for cell in row:
            if cell.value != None:
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
            if cell.value != None:
                list_problems.append(cell.value)
        if len(list_problems) == 3:
            list_all_problems.append(list_problems)

    return list_all_problems


