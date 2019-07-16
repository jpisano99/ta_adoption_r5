from my_app.func_lib.open_wb import open_wb
from my_app.settings import app_cfg
import time


as_wb, as_ws = open_wb(app_cfg['TESTING_TA_AS_FIXED_SKU_RAW'])
cust_wb, cust_ws = open_wb(app_cfg['TESTING_BOOKINGS_RAW_WITH_SO'])
sub_wb, sub_ws = open_wb(app_cfg['TESTING_RAW_SUBSCRIPTIONS'])

print("AS Fixed SKUs Rows:", as_ws.nrows)
print('Bookings Rows:', cust_ws.nrows)
print('Subscription Rows:', sub_ws.nrows)

cntr = 0
cust_db = {}
cust_name_db = {}
so_db = {}

for row_num in range(1, cust_ws.nrows):
    cust_so_dict = {}
    sku_list = []

    # Gather the fields we want
    cust_id = cust_ws.cell_value(row_num, 15)
    cust_erp_name = cust_ws.cell_value(row_num, 13)
    cust_ultimate_name = cust_ws.cell_value(row_num, 14)
    cust_so = cust_ws.cell_value(row_num, 11)
    cust_sku = cust_ws.cell_value(row_num, 19)

    # FORMAT: cust_db
    # {cust_id1: {so1: [sku1, sku2,..]
    #             so2: [sku1, sku2,..]
    #             }

    # FORMAT: cust_name_db
    # {cust_id1: [erp_name1, erp_name2]}
    #

    # Check to see if we already have this cust_id ?
    if cust_id not in cust_db:
        # Create a new cust_id and basic record
        so_db[cust_so] = [cust_sku]
        cust_db[cust_id] = so_db

        # Add this name to the name_db
        cust_name_db[cust_id] = [cust_erp_name]
    else:
        # Grab the SO dict from this existing customer id
        cust_so_dict = cust_db[cust_id]

        # If this SO is already in this cust_id just append this SKU
        if cust_so in cust_so_dict:
            # This SO is in our dict, add this SKU to the list
            sku_list = so_db[cust_so]
            sku_list.append(cust_sku)
            so_db[cust_so] = sku_list
            cust_db[cust_id] = so_db
        else:
            cust_db[cust_id] = so_db
            so_db[cust_so] = [cust_sku]
            cust_db[cust_id] = so_db

        #
        # Update the cust_name_db list (if needed)
        #
        names = cust_name_db[cust_id]
        add_it = True
        for name in names:
            if name == cust_erp_name:
                # We already got it don't add it
                add_it = False
                continue
        if add_it:
            names.append(cust_erp_name)
            cust_name_db[cust_id] = names

print("Customer IDs: ", len(cust_db))
print("Customer Names: ", len(cust_name_db))
print("Customer SOs: ", len(so_db))

# for my_id, names in cust_name_db.items():
#     if len(names)>1:
#         print('Customer ID', my_id, ' has the following aliases')
#         for name in names:
#             print('\t\t',name)


# {cust_id1: {name: [SO1,SO2,SO3], cust_id2: [name1,name2] }
# {cust_id1: [[order1,order2], [order1,order2]],


#
# Find AS SO numbers
#
hit = 0
miss = 0
as_db = {}
for row_num in range(1, as_ws.nrows):
    # Gather the fields we want
    as_pid = as_ws.cell_value(row_num, 0)
    as_cust_name = as_ws.cell_value(row_num, 2)
    as_so = as_ws.cell_value(row_num, 16)

    # if as_so not in as_db:
    #     as_db[as_so] = [as_pid, as_cust_name]
    # else:
    #     as_record = as_db[as_so]

    # print(as_pid, as_cust_name,as_so)
    if as_so in so_db:
        # print('\t\tFound ', as_so,as_cust_name)
        hit += 1
    else:
        # print('\t\tNOT Found', as_cust_name)
        miss += 1

    #time.sleep(1)
print('hits', hit)
print('miss', miss)

#
# Build a quick and dirty reference dict of {cust_name: cust_ids}
#
cust_id_db = {}
for my_id, names in cust_name_db.items():
    for name in names:
        if name not in cust_id_db:
            cust_id_db[name] = my_id



print('names db' , len(cust_id_db))


sub_hit = 0
sub_miss = 0
for row_num in range(1, sub_ws.nrows):
    # Gather the fields we want
    # sub_pid = as_ws.cell_value(row_num, 0)
    sub_cust_name = sub_ws.cell_value(row_num, 2)
    sub_id = sub_ws.cell_value(row_num, 4)
    sub_status = sub_ws.cell_value(row_num, 5)
    sub_start_date = sub_ws.cell_value(row_num, 6)
    jim = sub_ws.cell(row_num, 6)
    print(jim)

    if sub_cust_name in cust_id_db:
        sub_hit = + 1
    else:
        sub_miss = + 1
        print (sub_cust_name, type(sub_start_date), sub_start_date)
    time.sleep(1)

print('sub hits', hit)
print('sub miss', miss)

