from my_app.func_lib.get_list_from_ss import get_list_from_ss
from my_app.settings import app_cfg


#
# Pull cx data from it's SS and return a dict
# example: cx_dict = {'customer':[cx_contact],cx[status]}
#
def get_linked_sheet_update(my_map, my_tag, my_keys):
    my_sheet = app_cfg[my_tag]
    rows = get_list_from_ss(my_sheet)

    # Get the key column name to link on
    # Then look for the column number
    key_col_num = -1
    for col_data in my_keys:
        if col_data[0] == my_tag:
            # print('Link to: ', col_data[1], col_data[2])
            my_key = col_data[1]
            for row in rows:
                for idx, col_name in enumerate(row):
                    if col_name == my_key:
                        key_col_num = idx
                        break

    link_col_data = []
    for col_data in my_map:
        if col_data[1] == my_tag:
            link_col_data.append([col_data[0], col_data[2]])

    my_dict = {}
    for i, row in enumerate(rows):
        if i == 0:
            continue
        dict_key = row[key_col_num]

        row_data = []
        for column in link_col_data:
            row_data.append(row[column[1]])

        my_dict[dict_key] = row_data

    return my_dict
