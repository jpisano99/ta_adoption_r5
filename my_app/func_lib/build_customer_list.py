from my_app.settings import app_cfg
from my_app.func_lib.build_sheet_map import build_sheet_map
from my_app.func_lib.sheet_desc import sheet_map
from my_app.func_lib.open_wb import open_wb


def build_customer_list(run_dir):
    my_map = build_sheet_map(app_cfg['XLS_BOOKINGS'], sheet_map, 'XLS_BOOKINGS', run_dir)

    wb_bookings, sheet_bookings = open_wb(app_cfg['XLS_BOOKINGS'], run_dir)
    customer_list = []
    col_num_end_customer = -1
    col_num_erp_customer = -1

    #
    # First find the column numbers for these column names in the book
    #
    for val in my_map:
        if val[0] == 'ERP End Customer Name':
            col_num_erp_customer = val[2]
        elif val[0] == 'End Customer Global Ultimate Name':
            col_num_end_customer = val[2]

    #
    # Main loop of bookings excel data
    #
    top_row = []
    for i in range(sheet_bookings.nrows):
        if i == 0:
            # Grab these values to make column headings
            top_row_erp_val = sheet_bookings.cell_value(i, col_num_erp_customer)
            top_row_end_val = sheet_bookings.cell_value(i, col_num_end_customer)
            top_row = [top_row_erp_val, top_row_end_val]
            continue

        # Capture both of the Customer names
        customer_name_erp = sheet_bookings.cell_value(i, col_num_erp_customer)
        customer_name_end = sheet_bookings.cell_value(i, col_num_end_customer)
        customer_list.append((customer_name_erp, customer_name_end))

    # Create a simple customer_list list of tuples
    # Contains a full set of unique sorted customer names
    # customer_list = [(erp_customer_name,end_customer_ultimate), (CustA,CustA)]
    customer_list = set(customer_list)

    # Convert the SET to a LIST so we can sort it
    customer_list = list(customer_list)

    # Sort the LIST
    customer_list.sort(key=lambda tup: tup[0])

    # Convert the customer name tuples to a list
    tmp_list = []
    for customer in customer_list:
        tmp_list.append(list(customer))
    customer_list = tmp_list

    # Place column headings at the top
    customer_list.insert(0, top_row)

    return customer_list


if __name__ == "__main__":
    our_customers = build_customer_list()
    print('We have: ', len(our_customers), ' customers')
    print(our_customers)
