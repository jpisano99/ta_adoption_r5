import datetime
import xlrd
import os

from my_app.settings import app_cfg
from my_app.func_lib.sheet_desc import sheet_map as sm
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.create_customer_order_dict import create_customer_order_dict
from my_app.func_lib.get_linked_sheet_update import get_linked_sheet_update
from my_app.func_lib.build_sheet_map import build_sheet_map
from my_app.func_lib.sheet_desc import sheet_map, sheet_keys
from my_app.func_lib.build_sku_dict import build_sku_dict
from my_app.func_lib.push_xls_to_ss import push_xls_to_ss


def build_dashboard(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir))

    # from my_app.func_lib.sheet_desc import sheet_map
    #
    # Open the order summary
    #
    wb_orders, sheet_orders = open_wb(app_cfg['XLS_ORDER_SUMMARY'], run_dir)

    # wb_orders, sheet_orders = open_wb('tmp_TA Scrubbed Orders_as_of ' + app_cfg['PROD_DATE'])

    # Loop over the orders XLS worksheet
    # Create a simple list of orders with NO headers
    order_list = []
    for row_num in range(1, sheet_orders.nrows):  # Skip the header row start at 1
        tmp_record = []
        for col_num in range(sheet_orders.ncols):
            my_cell = sheet_orders.cell_value(row_num, col_num)

            # If we just read a date save it as a datetime
            if sheet_orders.cell_type(row_num, col_num) == 3:
                my_cell = datetime.datetime(*xlrd.xldate_as_tuple(my_cell, wb_orders.datemode))
            tmp_record.append(my_cell)
        order_list.append(tmp_record)

    # Create a dict of customer orders
    customer_order_dict = create_customer_order_dict(order_list)
    print()
    print('We have summarized ', len(order_list), ' of interesting line items into')
    print(len(customer_order_dict), ' unique customers')
    print()

    # Build Sheet Maps
    sheet_map = build_sheet_map(app_cfg['SS_CX'], sm, 'SS_CX')
    sheet_map = build_sheet_map(app_cfg['SS_AS'], sheet_map, 'SS_AS')
    sheet_map = build_sheet_map(app_cfg['SS_SAAS'], sheet_map, 'SS_SAAS')

    #
    # Get dict updates from linked sheets CX/AS/SAAS
    #
    cx_dict = get_linked_sheet_update(sheet_map, 'SS_CX', sheet_keys)
    as_dict = get_linked_sheet_update(sheet_map, 'SS_AS', sheet_keys)
    saas_dict = get_linked_sheet_update(sheet_map, 'SS_SAAS', sheet_keys)

    print()
    print('We have CX Updates: ', len(cx_dict))
    print('We have AS Updates: ', len(as_dict))
    print('We have SAAS Updates: ', len(saas_dict))
    print()

    # Create Platform dict for platform lookup
    tmp_dict = build_sku_dict()
    platform_dict = {}
    for key, val in tmp_dict.items():
        if val[0] == 'Product' or val[0] == 'SaaS':
            platform_dict[key] = val[1]

    #
    # Init Main Loop Variables
    #
    new_rows = []
    new_row = []
    bookings_col_num = -1
    sensor_col_num = -1
    svc_bookings_col_num = -1
    platform_type_col_num = -1
    sku_col_num = -1
    my_col_idx = {}

    # Create top row for the dashboard
    # also make a dict (my_col_idx) of {column names : column number}
    for col_idx, col in enumerate(sheet_map):
        new_row.append(col[0])
        my_col_idx[col[0]] = col_idx
    new_rows.append(new_row)

    #
    # Main loop
    #
    for customer, orders in customer_order_dict.items():
        new_row = []
        order = []
        orders_found = len(orders)

        # Default Values
        bookings_total = 0
        sensor_count = 0
        service_bookings = 0
        platform_type = 'Not Identified'

        saas_status = 'No Status'
        cx_contact = 'None assigned'
        cx_status = 'No Update'
        as_pm = ''
        as_cse1 = ''
        as_cse2 = ''
        as_complete = ''  # 'Project Status/PM Completion'
        as_comments = ''  # 'Delivery Comments'

        #
        # Get update from linked sheets (if any)
        #
        if customer in saas_dict:
            saas_status = saas_dict[customer][0]
            if saas_status is True:
                saas_status = 'Provision Complete'
            else:
                saas_status = 'Provision NOT Complete'
        else:
            saas_status = 'No Status'

        if customer in cx_dict:
            cx_contact = cx_dict[customer][0]
            cx_status = cx_dict[customer][1]
        else:
            cx_contact = 'None assigned'
            cx_status = 'No Update'

        if customer in as_dict:
            if as_dict[customer][0] == '':
                as_pm = 'None Assigned'
            else:
                as_pm = as_dict[customer][0]

            if as_dict[customer][1] == '':
                as_cse1 = 'None Assigned'
            else:
                as_cse1 = as_dict[customer][1]

            if as_dict[customer][2] == '':
                as_cse2 = 'None Assigned'
            else:
                as_cse2 = as_dict[customer][2]

            if as_dict[customer][3] == '':
                as_complete = 'No Update'
            else:
                # 'Project Status/PM Completion'
                as_complete = as_dict[customer][3]

            if as_dict[customer][4] == '':
                as_comments = 'No Comments'
            else:
                as_comments = as_dict[customer][4]

        #
        # Loop over this customers orders
        # Create one summary row for this customer
        # Total things
        # Build a list of things that may change order to order (ie Renewal Dates, Customer Names)
        #
        platform_count = 0
        for order_idx, order in enumerate(orders):
            # calculate totals in this loop (ie total_books, sensor count etc)
            bookings_total = bookings_total + order[my_col_idx['Total Bookings']]
            sensor_count = sensor_count + order[my_col_idx['Sensor Count']]

            if order[my_col_idx['Product Type']] == 'Service':
                service_bookings = service_bookings + order[my_col_idx['Total Bookings']]

            if order[my_col_idx['Bundle Product ID']] in platform_dict:
                platform_count += 1
                platform_type = platform_dict[order[my_col_idx['Bundle Product ID']]]
                if platform_count > 1:
                    platform_type = platform_type + ' plus ' + str(platform_count-1)

        #
        # Modify/Update this record as needed and then add to the new_rows
        #
        order[my_col_idx['Total Bookings']] = bookings_total
        order[my_col_idx['Sensor Count']] = sensor_count
        order[my_col_idx['Service Bookings']] = service_bookings

        order[my_col_idx['CSM']] = cx_contact
        order[my_col_idx['Comments']] = cx_status

        order[my_col_idx['Project Manager']] = as_pm
        order[my_col_idx['AS Engineer 1']] = as_cse1
        order[my_col_idx['AS Engineer 2']] = as_cse2
        order[my_col_idx['Project Status/PM Completion']] = as_complete  # 'Project Status/PM Completion'
        order[my_col_idx['Delivery Comments']] = as_comments

        order[my_col_idx['Provisioning completed']] = saas_status

        order[my_col_idx['Product Description']] = platform_type

        order[my_col_idx['Orders Found']] = orders_found

        new_rows.append(order)
    #
    # End of main loop
    #

    # Do some clean up and ready for output
    #
    # Rename the columns as per the sheet map
    cols_to_delete = []
    for idx, map_info in enumerate(sheet_map):
        if map_info[3] != '':
            if map_info[3] == '*DELETE*':
                # Put the columns to delete in a list
                cols_to_delete.append(idx)
            else:
                # Rename to the new column name
                new_rows[0][idx] = map_info[3]

    # Loop over the new_rows and
    # delete columns we don't need as per the sheet_map
    for col_idx in sorted(cols_to_delete, reverse=True):
        for row_idx, my_row in enumerate(new_rows):
            del new_rows[row_idx][col_idx]

    #
    # Write the Dashboard to an Excel File
    #
    push_list_to_xls(new_rows, app_cfg['XLS_DASHBOARD'], run_dir,'ta_dashboard')
    # push_xls_to_ss(app_cfg['XLS_DASHBOARD']+'_as_of_01_31_2019.xlsx', 'jims dash')

    return


if __name__ == "__main__" and __package__ is None:
    print(__package__)
    print('running process bookings')
    build_dashboard()
