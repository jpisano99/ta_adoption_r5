import os
import json
from my_app.settings import app_cfg
from my_app.func_lib.push_xlrd_to_xls import push_xlrd_to_xls
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.build_sku_dict import build_sku_dict


def file_checks(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    update_dir = app_cfg['UPDATES_DIR']
    archive_dir = app_cfg['ARCHIVES_DIR']

    # Check that all key directories exist
    path_to_main_dir = (os.path.join(home, working_dir))
    if not os.path.exists(path_to_main_dir):
        print(path_to_main_dir, " does NOT Exist !")
        exit()

    path_to_run_dir = (os.path.join(home, working_dir, run_dir))

    if not os.path.exists(path_to_run_dir):
        print(path_to_run_dir, " does NOT Exist !")
        exit()

    path_to_updates = (os.path.join(home, working_dir, update_dir))
    if not os.path.exists(path_to_updates):
        print(path_to_updates, " does NOT Exist !")
        exit()

    path_to_archives = (os.path.join(home, working_dir, archive_dir))
    if not os.path.exists(path_to_archives):
        print(path_to_archives, " does NOT Exist !")
        exit()

    # OK directories are there any files ?
    if not os.listdir(path_to_run_dir):
        print('Directory', path_to_run_dir, 'contains NO files')
        exit()

    #  Get the required Files to begin processing from app_cfg (settings.py)
    files_needed = {}
    # Do we have RAW files to process ?
    for var in app_cfg:
        if var.find('RAW') != -1:
            # Look for any config var containing the word 'RAW' and assume they are "Missing'
            files_needed[app_cfg[var]] = 'Missing'

    # See if we have the files_needed are there and they have consistent dates (date_list)
    run_files = os.listdir(path_to_run_dir)
    date_list = []
    for file_needed, status in files_needed.items():
        for run_file in run_files:
            date_tag = run_file[-13:-13 + 8]  # Grab the date if any
            run_file = run_file[:len(run_file)-14]  # Grab the name without the date
            if run_file == file_needed:
                date_list.append(date_tag)  # Grab the date
                files_needed[file_needed] = 'Found'
                break

    # All time stamps the same ?
    base_date = date_list[0]
    for date_stamp in date_list:
        if date_stamp != base_date:
            print('ERROR: Inconsistent date stamp found')
            exit()

    # Do we have all the files we need ?
    for file_name, status in files_needed.items():
        if status != 'Found':
            print('ERROR: File ', file_name, 'is missing')
            exit()

    # Read the config_dict.json file
    # with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
    #     config_dict = json.load(json_input)
    # print(config_dict)

    # Since we have a consistent date then Create the json file for config_data.json.
    # Put the time_stamp in it
    config_dict = {'data_time_stamp': base_date,
                   'last_run_dir': path_to_run_dir}
    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE']), 'w') as json_output:
        json.dump(config_dict, json_output)

    # Delete all previous tmp_ files
    for file_name in run_files:
        if file_name[0:4] == 'tmp_':
            os.remove(os.path.join(path_to_run_dir, file_name))

    # Here is what we have - All things should be in place
    print('Our directories:')
    print('\tPath to Main Dir:', path_to_main_dir)
    print('\tPath to Updates Dir:', path_to_updates)
    print('\tPath to Archives Dir:', path_to_archives)
    print('\tPath to Run Dir:', path_to_run_dir)

    # Process the RAW data (Renewals and Bookings)
    # Clean up rows, combine multiple Bookings files, add custom table names
    processing_date = date_list[0]
    file_paths = []
    bookings = []
    renewals = []
    start_row = 0
    print()
    print('We are processing files:')

    for file_name in files_needed:
        file_path = file_name + ' ' + processing_date + '.xlsx'
        file_path = os.path.join(path_to_run_dir, file_path)

        file_paths.append(file_path)

        my_wb, my_ws = open_wb(file_name + ' ' + processing_date + '.xlsx', run_dir)
        # my_wb = xlrd.open_workbook(file_path)
        # my_ws = my_wb.sheet_by_index(0)
        print('\t\t', file_name + '', processing_date + '.xlsx', ' has ', my_ws.nrows,
              ' rows and ', my_ws.ncols, 'columns')

        if file_name.find('Bookings') != -1:
            if start_row == 0:
                # For the first workbook include the header row
                start_row = 2
            elif start_row == 2:
                # For subsequent workbooks skip the header
                start_row = 3
            for row in range(start_row, my_ws.nrows):
                bookings.append(my_ws.row_slice(row))

        elif file_name.find('Renewals') != -1:
            for row in range(2, my_ws.nrows):
                renewals.append(my_ws.row_slice(row))

    # Push the lists out to an Excel File
    push_xlrd_to_xls(bookings, app_cfg['XLS_BOOKINGS'], run_dir, 'ta_bookings')

    as_bookings = get_as_skus(bookings)
    push_xlrd_to_xls(as_bookings, app_cfg['XLS_AS_SKUS'], run_dir, 'as_bookings')

    push_xlrd_to_xls(renewals, app_cfg['XLS_RENEWALS'], run_dir, 'ta_renewals')

    print('We have ', len(bookings), 'bookings line items')
    print('We have ', len(as_bookings), 'Services line items')
    print('We have ', len(renewals), 'renewal line items')

    return


def get_as_skus(bookings):
    # Build a SKU dict as a filter
    tmp_dict = build_sku_dict()
    sku_dict = {}
    header_row = bookings[0]
    header_vals = []
    for my_cell in header_row:
        header_vals.append(my_cell.value)

    # Strip out all but Service sku's
    for sku_key, sku_val in tmp_dict.items():
        if sku_val[0] == 'Service':
            sku_dict[sku_key] = sku_val

    sku_col_header = 'Bundle Product ID'
    sku_col_num = 0
    as_bookings = [header_row]

    # Get the col number that has the SKU's
    for idx, val in enumerate(header_vals):
        if val == sku_col_header:
            sku_col_num = idx
            break

    # Gather all the rows with AS skus
    for booking in bookings:
        if booking[sku_col_num].value in sku_dict:
            as_bookings.append(booking)

    print('All AS SKUs have been extracted from the current data!')
    return as_bookings


if __name__ == "__main__" and __package__ is None:
    print('Package Name:', __package__)
    print('running check_update_files')
    file_checks(os.path.join(app_cfg['UPDATES_DIR']))
    # file_checks()
    # file_checks(os.path.join(app_cfg['ARCHIVES_DIR'], '04-04-19 Updates'))
