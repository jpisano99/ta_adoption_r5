from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.push_xlrd_to_xls import push_xlrd_to_xls
from my_app.settings import app_cfg
import xlrd
from datetime import datetime
import time

as_wb, as_ws = open_wb(app_cfg['TESTING_TA_AS_FIXED_SKU_RAW'])
cust_wb, cust_ws = open_wb(app_cfg['TESTING_BOOKINGS_RAW_WITH_SO'])
sub_wb, sub_ws = open_wb(app_cfg['TESTING_RAW_SUBSCRIPTIONS'])

print("AS Fixed SKUs Rows:", as_ws.nrows)
print('Bookings Rows:', cust_ws.nrows)
print('Subscription Rows:', sub_ws.nrows)

#
# Process Main Bookings File
#
cntr = 0
cust_db = {}
cust_name_db = {}
so_db = {}

for row_num in range(1, cust_ws.nrows):
    my_so_dict = {}
    my_sku_list = []
    my_name_list = []
    my_so_info_list = []

    # Gather the fields we want
    cust_id = cust_ws.cell_value(row_num, 15)
    cust_erp_name = cust_ws.cell_value(row_num, 13)
    cust_ultimate_name = cust_ws.cell_value(row_num, 14)
    cust_so = cust_ws.cell_value(row_num, 11)
    cust_sku = cust_ws.cell_value(row_num, 19)

    if cust_id == '':
        cust_id = 'Unknown'

    #
    # Check cust_db
    # {cust_id1: {so1: [sku1, sku2,..]
    #             so2: [sku1, sku2,..]
    #
    # Is this new one ?
    if cust_id not in cust_db:
        # Create a new cust_id and basic record
        my_so_dict[cust_so] = [cust_sku]
        cust_db[cust_id] = my_so_dict
    else:
        # Grab the SO dict from this existing customer id
        my_so_dict = cust_db[cust_id]

        # If this SO is already in this cust_id just append this SKU
        if cust_so in my_so_dict:
            # This SO is in our dict, insert this SKU to this SO
            my_sku_list = my_so_dict[cust_so]
            my_sku_list.append(cust_sku)
            my_so_dict[cust_so] = my_sku_list
            cust_db[cust_id] = my_so_dict
        else:
            my_sku_list = [cust_sku]
            my_so_dict[cust_so] = my_sku_list
            cust_db[cust_id] = my_so_dict

    #
    # Check cust_name_db
    # {cust_id1: [erp_name1, erp_name2]}
    #
    # Is this new one ?
    if cust_id not in cust_name_db:
        my_name_list = [cust_erp_name]
        cust_name_db[cust_id] = my_name_list
    else:
        my_name_list = cust_name_db[cust_id]
        add_it = True
        for name in my_name_list:
            if name == cust_erp_name:
                add_it = False
                break
        if add_it:
            my_name_list.append(cust_erp_name)
            cust_name_db[cust_id] = my_name_list

    #
    # Check so_db
    # {so: [cust_id, cust_erp_name]}
    #
    # Is this new one ?
    if cust_so not in so_db:
        my_so_info_list = [(cust_id, cust_erp_name)]
        so_db[cust_so] = my_so_info_list
    else:
        my_so_info_list = so_db[cust_so]
        add_it = True
        for info in my_so_info_list:
            if info == (cust_id, cust_erp_name):
                add_it = False
                break
        if add_it:
            my_so_info_list.append((cust_id, cust_erp_name))
            so_db[cust_id] = my_so_info_list

    # print(so_db)
    # time.sleep(1)
print("Customer IDs: ", len(cust_db))
print("Customer Names: ", len(cust_name_db))
print("Customer SOs: ", len(so_db))

# # Display Customer IDs and Aliases
# for my_id, names in cust_name_db.items():
#     if len(names) > 1:
#         print('Customer ID', my_id, ' has the following aliases')
#         for name in names:
#             print('\t\t', name)
#             time.sleep(1)

# # Display Sales Order info
# for my_so, info in so_db.items():
#     if len(info) > 1:
#         print('Sales Order', my_so, ' has multiple Customer IDs and Names')
#         for my_info in info:
#             print('\t\t', my_info)
#             time.sleep(1)

# # Display Customer and Order Details
# for cust_id, my_orders in cust_db.items():
#     print()
#     print ("Customer ID ", cust_id, " has ", len(my_orders), ' orders')
#     print('Names ', cust_name_db[cust_id])
#     for order_num, skus in my_orders.items():
#         print ('\tOrder Number ', order_num, ' has ', len(skus),' skus')
#         for sku in skus:
#             print('\t\t', sku)
#             time.sleep(1)


# {cust_id1: {name: [SO1,SO2,SO3], cust_id2: [name1,name2] }
# {cust_id1: [[order1,order2], [order1,order2]],


#
# Process AS AS-F SKU File - find AS SO and PID numbers
#
hit = 0
miss = 0
as_db = {}
for row_num in range(1, as_ws.nrows):
    as_record = []
    # Gather the fields we want
    as_pid = as_ws.cell_value(row_num, 0)
    as_cust_name = as_ws.cell_value(row_num, 2)
    as_so = as_ws.cell_value(row_num, 16)

    if as_so not in as_db:
        as_record = [as_pid, as_cust_name]
        as_db[as_so] = as_record
    else:
        as_record = as_db[as_so]
        as_record.append([as_pid, as_cust_name])
        as_db[as_so] = as_record

    if as_so in so_db:
        # print('\t\tFound ', as_so,as_cust_name)
        hit += 1
    else:
        # print('\t\tNOT Found', as_cust_name)
        miss += 1

    print(as_record)
    print(as_db)
    time.sleep(1)
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
            if name == 'THE VANGUARD GROUP INC':
                print('my id',my_id)


#
# Process Subscriptions
#
today = datetime.today()
expired = []
thirty_days = []
sixty_days = []
ninety_days = []
ninety_plus = []
header_row = []

for row_num in range(0, sub_ws.nrows):
    # Gather the fields we want
    if row_num == 0 :
        header_row.append(sub_ws.cell_value(row_num, 2))
        header_row.append(sub_ws.cell_value(row_num, 4))
        header_row.append(sub_ws.cell_value(row_num, 5))
        header_row.append(sub_ws.cell_value(row_num, 6))
        header_row.append(sub_ws.cell_value(row_num, 8))
        print(header_row)
        continue
    else:
        sub_cust_name = sub_ws.cell_value(row_num, 2)
        sub_id = sub_ws.cell_value(row_num, 4)
        sub_status = sub_ws.cell_value(row_num, 5)
        sub_start_date = sub_ws.cell_value(row_num, 6)
        sub_renew_date = sub_ws.cell_value(row_num, 8)

    if sub_cust_name in cust_id_db:
        sub_cust_id = cust_id_db[sub_cust_name]
    else:
        sub_cust_id = 'Unknown'
        print(sub_cust_id,sub_cust_name)

    year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sub_start_date, sub_wb.datemode)
    sub_start_date = datetime(year, month, day)

    year, month, day, hour, minute, second = xlrd.xldate_as_tuple(sub_renew_date, sub_wb.datemode)
    sub_renew_date = datetime(year, month, day)

    days_to_renew = (sub_renew_date - today).days

    #
    # Bucket this customer renewal by age
    #
    if days_to_renew < 0:
        expired.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
    elif days_to_renew <= 30:
        thirty_days.append([sub_cust_id, sub_cust_name, sub_id, sub_renew_date, days_to_renew, sub_status])
    elif days_to_renew <= 60:
        sixty_days.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
    elif days_to_renew <= 90:
        ninety_days.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
    elif days_to_renew > 90:
        ninety_plus.append([sub_cust_id, sub_cust_name, sub_id, sub_status])
        # print(ninety_plus)
        # time.sleep(1)

subs_total = len(expired)+len(thirty_days)+len(sixty_days)+len(ninety_days)+len(ninety_plus)
print()
print('Total Subscriptions: ',subs_total)
print('\tExpired:', len(expired))
print('\t30 days:', len(thirty_days))
print('\t60 days:', len(sixty_days))
print('\t90 days:', len(ninety_days))
print('\t90+ days:', len(ninety_plus))
print()

print(header_row)
thirty_days.insert(0, header_row)

push_list_to_xls(thirty_days,'jim_subs.xlsx')
print('sub hits', hit)
print('sub miss', miss)


cust_id_db

