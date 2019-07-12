from my_app.func_lib.open_wb import open_wb
from my_app.ss_lib.Ssheet_class import Ssheet
from my_app.settings import app_cfg


def build_sheet_map(file_name, my_map, tag, run_dir=app_cfg["UPDATES_DIR"]):
    print('MAPPING>>>>>>>>>> ', run_dir + '\\' + file_name)

    # First look and the tag and decide if we looking at
    # A local excel sheet or Smart Sheet
    if tag[:2] == 'SS':
        # Get the list of columns
        my_sheet = Ssheet(file_name, True)
        my_columns = my_sheet.columns

        # Loop across the Smart Sheet columns
        for ss_col in range(len(my_columns)):
            ss_col_num = my_sheet.columns[ss_col]['index']
            ss_col_name = my_sheet.columns[ss_col]['title']

            # Loop across the sheet map and look for a match
            for idx, val in enumerate(my_map):
                col_name = val[0]
                if col_name == ss_col_name and val[1] == tag:
                    # We have a match on the source col name and file tag
                    val[2] = ss_col_num

    elif tag[:3] == 'XLS':
        workbook, sheet = open_wb(file_name, run_dir)
        # Loop across all column headings in the bookings file and
        # Find the column number that matches the col_name in my_dict
        for wb_col_num in range(sheet.ncols):
            for idx, val in enumerate(my_map):
                col_name = val[0]
                if col_name == sheet.cell_value(0, wb_col_num) and val[1] == tag:
                    val[2] = wb_col_num
    else:
        print('Missing Map TAG')
        exit()

    return my_map


# if __name__ == "__main__":
#     # Populate the sheet map with column meta data
#     sheet_map = build_sheet_map(app['XLS_BOOKINGS'], sheet_map, 'XLS_BOOKINGS')
#     sheet_map = build_sheet_map(app['XLS_RENEWALS'], sheet_map, 'XLS_RENEWALS')
#     sheet_map = build_sheet_map(app['SS_COVERAGE'], sheet_map, 'SS_COVERAGE')
#     sheet_map = build_sheet_map(app['SS_AS'], sheet_map, 'SS_AS')
#     sheet_map = build_sheet_map(app['SS_CX'], sheet_map, 'SS_CX')
#     sheet_map = build_sheet_map(app['SS_SAAS'], sheet_map, 'SS_SAAS')
#     print(sheet_map)
