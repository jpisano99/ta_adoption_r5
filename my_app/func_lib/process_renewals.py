from datetime import datetime
import datetime
import xlrd
from my_app.settings import app_cfg
from my_app.func_lib.sheet_desc import sheet_map
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.build_sheet_map import build_sheet_map


def process_renewals(run_dir=app_cfg["UPDATES_DIR"]):
    print('MAPPING>>>>>>>>>> ', run_dir + '\\' + app_cfg['XLS_RENEWALS'])
    # Open up the renewals excel workbooks

    wb, sheet = open_wb(app_cfg['XLS_RENEWALS'], run_dir)

    # Get the renewal columns we are looking for
    my_map = build_sheet_map(app_cfg['XLS_RENEWALS'], sheet_map, 'XLS_RENEWALS', run_dir)

    print('sheet_map ', id(sheet_map))
    print('my map ', id(my_map))

    # List comprehension replacement for above
    # Strip out the columns from the sheet map that we don't need
    my_map = [x for x in my_map if x[1] == 'XLS_RENEWALS']

    # Create a simple column name dict
    col_nums = {sheet.cell_value(0, col_num): col_num for col_num in range(0, sheet.ncols)}

    # Loop over all of the renewal records
    # Build a dict of {customer:[next renewal date, next renewal revenue, upcoming renewals]}
    my_dict = {}
    for row_num in range(1, sheet.nrows):
        customer = sheet.cell_value(row_num, col_nums['End Customer'])
        if customer in my_dict:
            tmp_record = []
            tmp_records = my_dict[customer]
        else:
            tmp_record = []
            tmp_records = []

        # Loop over the my map gather the columns we need
        for col_map in my_map:
            my_cell = sheet.cell_value(row_num, col_map[2])

            # Is this cell a Date type (3) ?
            # If so format as a M/D/Y
            if sheet.cell_type(row_num, col_map[2]) == 3:
                my_cell = datetime.datetime(*xlrd.xldate_as_tuple(my_cell, wb.datemode))
                my_cell = my_cell.strftime('%m-%d-%Y')

            tmp_record.append(my_cell)

        tmp_records.append(tmp_record)
        my_dict[customer] = tmp_records

    #
    # Sort each customers renewal dates
    #
    sorted_dict = {}
    summarized_dict = {}
    summarized_rec = []

    for customer, renewals in my_dict.items():
        # Sort this customers renewal records by date order
        renewals.sort(key=lambda x: x[0])
        sorted_dict[customer] = renewals

        next_renewal_date = renewals[0][0]
        next_renewal_rev = 0
        next_renewal_qtr = renewals[0][2]
        for renewal_rec in renewals:
            if renewal_rec[0] == next_renewal_date:
                # Tally this renewal record and get the next
                next_renewal_rev = float(renewal_rec[1] + next_renewal_rev)
            elif renewal_rec[0] != next_renewal_date:
                # Record these summarized values
                summarized_rec.append([next_renewal_date, next_renewal_rev, next_renewal_qtr])
                # Reset these values and get the next renewal date for this customer
                next_renewal_date = renewal_rec[0]
                next_renewal_rev = renewal_rec[1]
                next_renewal_qtr = renewal_rec[2]

            # Check to see if this is the last renewal record
            # If so exit the loop
            if renewals.index(renewal_rec) == len(renewals)-1:
                break

        summarized_rec.append([next_renewal_date, next_renewal_rev, next_renewal_qtr])
        summarized_dict[customer] = summarized_rec
        summarized_rec = []

    # print(sorted_dict['FIRST NATIONAL BANK OF SOUTHERN AFRICA LTD'])
    # print(summarized_dict['SPECTRUM HEALTH SYSTEM'])
    # print (len(summarized_dict['SPECTRUM HEALTH SYSTEM']))
    return summarized_dict
