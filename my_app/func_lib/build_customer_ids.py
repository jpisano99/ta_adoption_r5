from my_app.func_lib.open_wb import open_wb
from my_app.func_lib.push_list_to_xls import push_list_to_xls
from my_app.func_lib.push_xlrd_to_xls import push_xlrd_to_xls
from my_app.func_lib.build_sku_dict import build_sku_dict
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
# Create a SKU Filter
#
# Options Are: Product / Software / Service / SaaS / *
sku_filter_val = '*'
tmp_dict = build_sku_dict()
sku_filter_dict = {}

for key, val in tmp_dict.items():
    if val[0] == sku_filter_val:
        sku_filter_dict[key] = val
    elif sku_filter_val == '*':
        # Selects ALL Interesting SKUs
        sku_filter_dict[key] = val

print('SKU Filter set to:', sku_filter_val)

#
# Build a xref dict of valid customer ids for lookup by SO and ERP Name
#
xref_cust_name = {}
xref_so = {}
for row_num in range(1, cust_ws.nrows):
    cust_id = cust_ws.cell_value(row_num, 15)
    cust_erp_name = cust_ws.cell_value(row_num, 13)
    cust_so = cust_ws.cell_value(row_num, 11)

    # Only add valid ID/Name Pairs to the reference
    if cust_id == '-999' or cust_id == '':
        continue

    if cust_erp_name not in xref_cust_name:
        xref_cust_name[cust_erp_name] = cust_id
    if cust_so not in xref_so:
        xref_so[cust_so] = cust_id

#
# Process Main Bookings File
#
cntr = 0
cust_db = {}
cust_name_db = {}
so_db = {}
# Main loop starts here
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

    # We have a missing or bad cust_id try to look one up
    if cust_id == '' or cust_id == '-999':
        if cust_erp_name in xref_cust_name:
            cust_id = xref_cust_name[cust_erp_name]

        if cust_so in xref_so:
            cust_id = xref_so[cust_so]

        # If id is still bad flag cust_id as UNKNOWN
        if cust_id == '' or cust_id == '-999':
            cust_id = 'UNKNOWN'

    #
    # Check cust_db
    # {cust_id1: {so1: [sku1, sku2,..]
    #             so2: [sku1, sku2,..]
    #
    # Is this new one ?
    if cust_id not in cust_db:
        # Create a new cust_id and basic record
        if cust_sku in sku_filter_dict:
            my_so_dict[cust_so] = [cust_sku]
            cust_db[cust_id] = my_so_dict
    else:
        # Grab the SO dict from this existing customer id
        my_so_dict = cust_db[cust_id]

        # If this SO is already in this cust_id just append this SKU
        if cust_so in my_so_dict:
            if cust_sku in sku_filter_dict:
                # This SO is in our dict, insert this SKU to this SO
                my_sku_list = my_so_dict[cust_so]
                my_sku_list.append(cust_sku)
                my_so_dict[cust_so] = my_sku_list
                cust_db[cust_id] = my_so_dict
        else:
            if cust_sku in sku_filter_dict:
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

#
# Build a quick and dirty reverse lookup of customer names {cust_name: cust_id}
#
cust_id_db = {}
for my_id, names in cust_name_db.items():
    for name in names:
        if name not in cust_id_db:
            cust_id_db[name] = my_id


print("Customer IDs with AS Services: ", len(cust_db))
print("Customer Unique Customer ID's: ", len(cust_name_db))
print("Customer Unique Customer Names: ", len(cust_id_db))
print("Customer Unique SOs: ", len(so_db))


# A quick check on customer ids
id_list = [['Customer ID', 'Customer Aliases']]
for cust_id, cust_aliases in cust_name_db.items():
    alias_list = []
    alias_str = ''
    for cust_alias in cust_aliases:
        alias_list.append(cust_alias)
        alias_str = alias_str + cust_alias + ' : '
    alias_str = alias_str[:-3]
    id_list.append([cust_id, alias_str])

push_list_to_xls(id_list, 'unique_cust_ids.xlsx')

# print(len(cust_id_db))
# for id, name in cust_id_db.items():
#     print(id,name)
#     time.sleep(1)
# exit()


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


#
# Process AS AS-F SKU File - find AS SO and PID numbers
#
hit = 0
miss = 0
as_db = {}
for row_num in range(1, as_ws.nrows):
    my_as_info_list = []
    # Gather the fields we want
    as_pid = as_ws.cell_value(row_num, 0)
    as_cust_name = as_ws.cell_value(row_num, 2)
    as_so = as_ws.cell_value(row_num, 16)

    if as_so not in as_db:
        my_as_info_list.append((as_pid, as_cust_name))
        as_db[as_so] = my_as_info_list
    else:
        my_as_info_list = as_db[as_so]
        add_it = True
        for info in my_as_info_list:
            if info == (as_pid, as_cust_name):
                add_it = False
                break
        if add_it:
            my_as_info_list.append((as_pid, as_cust_name))
            as_db[as_so] = my_as_info_list

    if as_so in so_db:
        # print('\t\tFound ', as_so,as_cust_name)
        hit += 1
    else:
        # print('\t\tNOT Found', as_cust_name)
        miss += 1

    # print(my_as_info_list)
    # print(as_db)
    # time.sleep(1)
print('hits', hit)
print('miss', miss)

# # Display AS Sales Order info
# for my_so, info in as_db.items():
#     if len(info) > 1:
#         print('AS Sales Order', my_so, ' has multiple PIDs')
#         for my_info in info:
#             print('\t\t', my_info)
#             time.sleep(1)
# exit()


#
# Build ths Subscription db (sub_db)
#
sub_db = {}
for row_num in range(1, sub_ws.nrows):
    # Gather the fields we want
    sub_cust_name = sub_ws.cell_value(row_num, 2)
    sub_id = sub_ws.cell_value(row_num, 4)

    my_sub_id_list = []
    # Is this new one ?
    if sub_cust_name not in sub_db:
        my_sub_id_list = [sub_id]
        sub_db[sub_cust_name] = my_sub_id_list
    else:
        my_sub_id_list = sub_db[sub_cust_name]
        add_it = True
        for id in my_sub_id_list:
            if id == sub_id:
                add_it = False
                break
        if add_it:
            my_sub_id_list.append(sub_id)
            sub_db[sub_cust_name] = my_sub_id_list

# # Display Customer Subscription IDs
# for sub_cust_name, sub_ids in sub_db.items():
#     if len(sub_ids) > 1:
#         print('Customer ', sub_cust_name, ' has multiple Subscription IDs')
#         for sub_id in sub_ids:
#             print('\t\t', sub_id)
#             time.sleep(1)
#
# exit()


#
# Databases we have constructed to link data
#
############
# Based on the Bookings File
#
# cust_db
# {cust_id: {so1: [sku1, sku2,..]
#             so2: [sku1, sku2,..]}
#
# cust_id_db
# {erp_name: cust_id}

# cust_name_db
# {cust_id: [erp_name1, erp_name2]}
#
# Check so_db
# {so: [cust_id, cust_erp_name]}
#
############
# Based on the AS-F Data
#
# as_db
# {so: [(as_pid1, as_cust_name1),(as_pid2, as_cust_name2)]}
#
############
# Based on the Subscription and Renewal Data
#
# sub_db
# {erp_name: [sub_id1,sub_id2]}
#


#
# Make the Magic List
#
magic_list = []
header_row = ['cust_id', 'SO', 'AS PID', 'AS Customer Name']
# header_row = ['cust_id', 'cust_name', 'sub id', 'renew date', 'renew date', 'sub status', 'so', 'as pid', 'as status']
magic_list.append(header_row)
print (magic_list)

for cust_id, so_dict in cust_db.items():
    magic_cust_id = cust_id
    magic_cust_name_list = cust_name_db[magic_cust_id]  # List of all Names / aliases for this customer id
    magic_as_pid_list = []
    magic_sub_id_list = []
    magic_so_list = []

    # Let's find all the AS PIDs in as_db for this Customer ID
    for so, sku_list in so_dict.items():
        # print('checking customer id', magic_cust_id,' SO num ', so, 'has', len(sku_list), ' skus')
        if so in as_db:
            # This SO has an AS PID associated with it
            magic_so_list.append(so)
            my_as_details = as_db[so]

            # OK Let's get all PIDS and associated AS data
            for as_detail in my_as_details:
                # Here is where we need to create a line item for each AS PID
                # as_db
                # {so: [(as_pid1, as_cust_name1),(as_pid2, as_cust_name2)]}
                magic_as_pid_list.append(as_detail[0])

                magic_list.append([magic_cust_id, so, as_detail[0], as_detail[1]])
                print([magic_cust_id, so, as_detail[0], as_detail[1]])
                # time.sleep(.5)

push_list_to_xls(magic_list,'magic.xlsx')













    #
    # # time.sleep(.1)
    # if magic_cust_id == '54028':
    #     if magic_as_pid_list != []:
    #         print()
    #         print('Customer ID', magic_cust_id, ' with Cust Aliases:', magic_cust_name_list)
    #         print('\tMatched AS SO List:', magic_so_list, '\tAS pid list:', magic_as_pid_list)

    # Let's find all Subscriptions for this Customer ID / Name
    # We need to search by ERP Customer Name
    # jim={}
    #
    # print('Customer ID: ', magic_cust_id , 'has ', len(magic_cust_name_list), ' aliases:', magic_cust_name_list)
    # for cust_name in magic_cust_name_list:
    #     if cust_name in sub_db:
    #         # Found subscription(s)
    #         print('\tName found in Subscription:', cust_name, ' subscriptions ', sub_db[cust_name])
    #         magic_sub_id_list.append(sub_db[cust_name])
    # if magic_sub_id_list == []:
    #     magic_sub_id_list.append("No Subscriptions Found")
    # # print('\t\t Found subscriptions: ', magic_sub_id_list)

exit()



    # sub_db
    # {erp_name: [sub_id1,sub_id2]}

    # cust_id_db
    # {erp_name: cust_id}


exit()


#
#   IGNORE BELOW
#






for row_num in range(1, sub_ws.nrows):
    # Gather the fields we want
    magic_sub_cust_name = sub_ws.cell_value(row_num, 2)
    magic_sub_id = sub_ws.cell_value(row_num, 4)
    magic_sub_status = sub_ws.cell_value(row_num, 5)
    magic_sub_start_date = sub_ws.cell_value(row_num, 6)
    magic_sub_renew_date = sub_ws.cell_value(row_num, 8)

    # Get the customer ID
    if magic_sub_cust_name in cust_id_db:
        print(len(magic_sub_cust_name))
        print('\t',magic_sub_cust_name)
        time.sleep(.5)

        magic_bookings_cust_id = cust_id_db[magic_sub_cust_name]
    else:
        pass
        # print("Customer ID Miss", magic_sub_cust_name)

    # Get

# Diagnostic
match = 0
misses = 0
for row_num in range(1, as_ws.nrows):
    found = False
    as_cust_name = as_ws.cell_value(row_num, 2)
    #print(as_cust_name)
    for row_num in range(1, sub_ws.nrows):
        sub_cust_name = sub_ws.cell_value(row_num, 2)
        #time.sleep(1)
        #print("\t",sub_cust_name)
        if sub_cust_name == as_cust_name:
            found=True
            match += 1
            print('\tMatched', as_cust_name)
            break
    if not found:
        misses += 1
        print(as_cust_name, 'not found in AS sheet')

    #time.sleep(1)

print('matches' ,match,misses)
exit()




today = datetime.today()
expired = []
thirty_days = []
sixty_days = []
ninety_days = []
ninety_plus = []


for row_num in range(1, sub_ws.nrows):
    # Gather the fields we want
    sub_cust_name = sub_ws.cell_value(row_num, 2)
    sub_id = sub_ws.cell_value(row_num, 4)
    sub_status = sub_ws.cell_value(row_num, 5)
    sub_start_date = sub_ws.cell_value(row_num, 6)
    sub_renew_date = sub_ws.cell_value(row_num, 8)

    if sub_cust_name in cust_id_db:
        # Get the cust_id that matches this subscription name
        sub_cust_id = cust_id_db[sub_cust_name]

        # Go get a list of SOs for this cust_id
        # Use this to find and AS engagements
        my_so_dict = cust_db[sub_cust_id]
        my_so_list = []
        for so, skus in my_so_dict.items():
            my_so_list.append(so)

        # Go get a list of AS PIDs for these SO's
        my_as_pids = []
        for so in my_so_list:
            if so in as_db:
                # Found an AS record
                as_info = as_db[so]
                as_pid = as_info[0][0]
                as_cust_name = as_info[0][1]
                my_as_pids.append(as_pid)
            # else:
            #     my_as_pids.append("NO AS Engagements Found !")
    else:
        # We can't find a match on this customer name
        # Maybe check aliases ?
        sub_cust_id = 'Unknown'
        my_as_pids = ''
        my_so_list = ''

    print(sub_cust_id, sub_cust_name, 'have ', len(my_as_pids), ' PIDS')
    print('\t\t',my_as_pids)
    # print('\tSOs',my_so_list)
    # print('\tAS PIDS', my_as_pids)
    # print()
    time.sleep(1)

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

