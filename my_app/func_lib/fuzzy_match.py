from fuzzywuzzy import fuzz
from my_app.func_lib.open_wb import open_wb
from my_app.settings import app_cfg
import time


def fuzzy_match(name1, name2):
    match = False
    ratios = fuzz.ratio(name1, name2)
    if ratios >= 85 and ratios < 100 :
        print(name1, '/ \t', name2, ratios)
        match = True

    #time.sleep(1)

    return match


if __name__ == "__main__":
    list_a = ['jim']
    list_b = ['jim', 'gym', 'ang', 'jime']
    my_file = 'C:/Users/jpisano/ta_adoption_data/ta_data_updates/tmp_TA Customer List.xlsx'
    my_wb, my_ws = open_wb(my_file)

    sheet1 = my_wb.sheet_by_name('Sheet1')
    sheet2 = my_wb.sheet_by_name('Sheet2')

    unique_names = []
    jims_list = []
    duplicate_names = []
    aka_list = []
    aka = {}

    for my_row in range(1, sheet1.nrows):
        duplicate_names.append(sheet1.cell_value(my_row, 0))

    for my_row in range(0, sheet2.nrows):
        unique_names.append(sheet1.cell_value(my_row, 0))

    for name in unique_names:
        # for test_name in duplicate_names:
        aka_list.append(name)
        print (aka_list)

        aka.update(aka_list)
        for test_name in duplicate_names:
            if fuzzy_match(name, test_name):
                aka_list.append(test_name)
                aka.update(aka_list)

    print(len(unique_names))
    print (len(duplicate_names))
    print(aka)

    # fuzzy_match(list_a, list_b)
