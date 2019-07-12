# import datetime
from datetime import datetime
import os
import json
from my_app.settings import app_cfg
from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.build_coverage_dict import build_coverage_dict
from my_app.func_lib.build_sku_dict import build_sku_dict
from my_app.func_lib.find_team import find_team
from my_app.func_lib.sheet_desc import sheet_map
from my_app.func_lib.build_sheet_map import build_sheet_map
from my_app.func_lib.process_renewals import process_renewals
from my_app.func_lib.build_customer_list import build_customer_list
from my_app.func_lib.cleanup_orders import cleanup_orders
from my_app.func_lib.push_list_to_xls import push_list_to_xls


def process_bookings(run_dir=app_cfg['UPDATES_DIR']):
    home = app_cfg['HOME']
    working_dir = app_cfg['WORKING_DIR']
    path_to_run_dir = (os.path.join(home, working_dir, run_dir))

    bookings_path = os.path.join(path_to_run_dir, app_cfg['XLS_BOOKINGS'])
    renewals_path = os.path.join(path_to_run_dir, app_cfg['XLS_RENEWALS'])

    # Read the config_dict.json file
    with open(os.path.join(path_to_run_dir, app_cfg['META_DATA_FILE'])) as json_input:
        config_dict = json.load(json_input)
    data_time_stamp = datetime.strptime(config_dict['data_time_stamp'], '%m-%d-%y')
    last_run_dir = config_dict['last_run_dir']

    print("Run Date: ", data_time_stamp, type(data_time_stamp))
    print('Run Directory:', last_run_dir)
    print(bookings_path)
    print(renewals_path)

    # Go to Smartsheets and build these two dicts to use reference lookups
    # team_dict: {'sales_levels 1-6':[('PSS','TSA')]}
    # sku_dict: {sku : [sku_type, sku_description]}
    team_dict = build_coverage_dict()
    sku_dict = build_sku_dict()

    #
    # Open up the bookings excel workbooks
    #
    wb_bookings, sheet_bookings = open_wb(app_cfg['XLS_BOOKINGS'], run_dir)

    # From the current up to date bookings file build a simple list
    # that describes the format of the output file we are creating
    # and the columns we need to add (ie PSS, TSA, Renewal Dates)

    my_sheet_map = build_sheet_map(app_cfg['XLS_BOOKINGS'], sheet_map,
                                   'XLS_BOOKINGS', run_dir)

    # print('sheet_map ', id(sheet_map))
    # print('sheet_map ', sheet_map)
    # print('my_sheet_map ', id(my_sheet_map))
    # print('my_sheet_map ', my_sheet_map)
    # exit()
    #
    # init a bunch a variable we need for the main loop
    #
    order_header_row = []
    order_rows = []
    order_row = []
    trash_rows = []

    dest_col_nums = {}
    src_col_nums = {}

    # Build a dict of source sheet {'col_name' : src_col_num}
    # Build a dict of destination sheet {'col_name' : dest_col_num}
    # Build the header row for the output file
    for idx, val in enumerate(my_sheet_map):
        # Add to the col_num dict of col_names
        dest_col_nums[val[0]] = idx
        src_col_nums[val[0]] = val[2]
        order_header_row.append(val[0])

    # Initialize the order_row and trash_row lists
    order_rows.append(order_header_row)
    trash_rows.append(sheet_bookings.row_values(0))

    print('There are ', sheet_bookings.nrows, ' rows in Raw Bookings')

    #
    # Main loop of over raw bookings excel data
    #
    # This loop will build two lists:
    # 1. Interesting orders based on SKUs (order_rows)
    # 2. Trash orders SKUs we don't care about (trash_rows)
    # As determined by the sku_dict
    # We have also will assign team coverage to both rows
    #
    for i in range(1, sheet_bookings.nrows):

        # Is this SKU of interest ?
        sku = sheet_bookings.cell_value(i, src_col_nums['Bundle Product ID'])

        if sku in sku_dict:
            # Let's make a row for this order
            # Since it has an "interesting" sku
            customer = sheet_bookings.cell_value(i, src_col_nums['ERP End Customer Name'])
            order_row = []
            sales_level = ''
            sales_level_cntr = 0

            # Grab SKU data from the SKU dict
            sku_type = sku_dict[sku][0]
            sku_desc = sku_dict[sku][1]
            sku_sensor_cnt = sku_dict[sku][2]

            # Walk across the sheet_map columns
            # to build this output row cell by cell
            for val in my_sheet_map:
                col_name = val[0]  # Source Sheet Column Name
                col_idx = val[2]  # Source Sheet Column Number

                # If this is a 'Sales Level X' column then
                # Capture it's value until we get to level 6
                # then do a team lookup
                if col_name[:-2] == 'Sales Level':
                    sales_level = sales_level + sheet_bookings.cell_value(i, col_idx) + ','
                    sales_level_cntr += 1

                    if sales_level_cntr == 6:
                        # We have collected all 6 sales levels
                        # Now go to find_team to do the lookup
                        sales_level = sales_level[:-1]
                        sales_team = find_team(team_dict, sales_level)
                        pss = sales_team[0]
                        tsa = sales_team[1]
                        order_row[dest_col_nums['pss']] = pss
                        order_row[dest_col_nums['tsa']] = tsa

                if col_idx != -1:
                    # OK we have a cell that we need from the raw bookings
                    # sheet we need so grab it
                    order_row.append(sheet_bookings.cell_value(i, col_idx))
                elif col_name == 'Product Description':
                    # Add in the Product Description
                    order_row.append(sku_desc)
                elif col_name == 'Product Type':
                    # Add in the Product Type
                    order_row.append(sku_type)
                elif col_name == 'Sensor Count':
                    # Add in the Sensor Count
                    order_row.append(sku_sensor_cnt)
                else:
                    # this cell is assigned a -1 in the sheet_map
                    # so assign a blank as a placeholder for now
                    order_row.append('')

            # Done with all the columns in this row
            # Log this row for BOTH customer names and orders
            # Go to next row of the raw bookings data
            order_rows.append(order_row)

        else:
            # The SKU was not interesting so let's trash it
            trash_rows.append(sheet_bookings.row_values(i))

    print('Extracted ', len(order_rows), " rows of interesting SKU's' from Raw Bookings")
    print('Trashed ', len(trash_rows), " rows of trash SKU's' from Raw Bookings")
    #
    # End of main loop
    #

    #
    # Renewal Analysis
    #
    renewal_dict = process_renewals(run_dir)
    for order_row in order_rows[1:]:
        customer = order_row[dest_col_nums['ERP End Customer Name']]
        if customer in renewal_dict:
            next_renewal_date = datetime.strptime(renewal_dict[customer][0][0], '%m-%d-%Y')
            next_renewal_rev = renewal_dict[customer][0][1]
            next_renewal_qtr = renewal_dict[customer][0][2]

            order_row[dest_col_nums['Renewal Date']] = next_renewal_date
            order_row[dest_col_nums['Product Bookings']] = next_renewal_rev
            order_row[dest_col_nums['Fiscal Quarter ID']] = next_renewal_qtr

            if len(renewal_dict[customer]) > 1:
                renewal_comments = '+' + str(len(renewal_dict[customer])-1) + ' more renewal(s)'
                order_row[dest_col_nums['Renewal Comments']] = renewal_comments

    # Now we build a an order dict
    # Let's organize as this
    # order_dict: {cust_name:[[order1],[order2],[orderN]]}
    order_dict = {}
    orders = []
    order = []

    for idx, order_row in enumerate(order_rows):
        if idx == 0:
            continue
        customer = order_row[0]
        orders = []

        # Is this customer in the order dict ?
        if customer in order_dict:
            orders = order_dict[customer]
            orders.append(order_row)
            order_dict[customer] = orders
        else:
            orders.append(order_row)
            order_dict[customer] = orders

    # Create a simple customer_list
    # Contains a full set of unique sorted customer names
    # Example: customer_list = [[erp_customer_name,end_customer_ultimate], [CustA,CustA]]
    customer_list = build_customer_list(run_dir)
    print('There are ', len(customer_list), ' unique Customer Names')

    # Clean up order_dict to remove:
    # 1.  +/- zero sum orders
    # 2. zero revenue orders
    order_dict, customer_platforms = cleanup_orders(customer_list, order_dict, my_sheet_map)

    #
    # Create a summary order file out of the order_dict
    #
    summary_order_rows = [order_header_row]
    for key, val in order_dict.items():
        for my_row in val:
            summary_order_rows.append(my_row)
    print(len(summary_order_rows), ' of scrubbed rows after removing "noise"')

    #
    # Push our lists to an excel file
    #
    # push_list_to_xls(customer_platforms, 'jim ')
    print('order summary name ', app_cfg['XLS_ORDER_SUMMARY'])

    push_list_to_xls(summary_order_rows, app_cfg['XLS_ORDER_SUMMARY'],
                     run_dir, 'ta_summary_orders')
    push_list_to_xls(order_rows, app_cfg['XLS_ORDER_DETAIL'], run_dir, 'ta_order_detail')
    push_list_to_xls(customer_list, app_cfg['XLS_CUSTOMER'], run_dir, 'ta_customers')
    push_list_to_xls(trash_rows, app_cfg['XLS_BOOKINGS_TRASH'], run_dir, 'ta_trash_rows')

    # exit()
    #
    # Push our lists to a smart sheet
    #
    # push_xls_to_ss(wb_file, app_cfg['XLS_ORDER_SUMMARY'])
    # push_xls_to_ss(wb_file, app_cfg['XLS_ORDER_DETAIL'])
    # push_xls_to_ss(wb_file, app_cfg['XLS_CUSTOMER'])
    # exit()
    return


if __name__ == "__main__" and __package__ is None:
    print('Package Name:', __package__)
    print('running process bookings')
    # process_bookings(os.path.join(app_cfg['ARCHIVES_DIR'], '04-04-19 Updates'))
    process_bookings()
