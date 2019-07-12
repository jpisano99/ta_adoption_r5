from my_app.func_lib.build_sku_dict import build_sku_dict


def cleanup_orders(customer_list, order_dict, col_map):
    sku_col_num = -1
    booking_col_num = -1
    for idx, val in enumerate(col_map):
        if val[0] == 'Bundle Product ID':
            sku_col_num = idx
        if val[0] == 'Total Bookings':
            booking_col_num = idx

    # Create Platform dict for platform lookup
    tmp_dict = build_sku_dict()
    platform_dict = {}
    for key, val in tmp_dict.items():
        if val[0] == 'Product' or val[0] == 'SaaS':
            platform_dict[key] = val[1]

    # Loop over each customer and see if they have any orders
    customer_platforms = []
    for customer_names in customer_list[1:]:  # skip the head row
        customer_name = customer_names[0]  # (erp_customer_name,end_customer_ult)

        # If this customer has orders then
        # Loop over the orders for this customer
        if customer_name in order_dict:
            dirty_orders = order_dict[customer_name]
            clean_orders = []
            platform_found = False

            for idx, order in enumerate(dirty_orders):
                sku = order[sku_col_num]
                booking = order[booking_col_num]
                valid_order = False

                if sku in platform_dict:
                    platform_found = True
                    customer_platforms.append([customer_name, sku, booking])

                if booking == 0:
                    # Drop this order
                    continue

                # Rescan dirty orders (starting at idx)
                # Look for same SKU and opposite booking amount
                for tmp_order in dirty_orders[idx:]:
                    tmp_sku = tmp_order[sku_col_num]
                    tmp_booking = tmp_order[booking_col_num]
                    if sku == tmp_sku and (booking * -1) == tmp_booking:
                        # Skip this order since it has a corresponding SKU
                        # for the opposite booking amount
                        # Exit this loop and get the next order
                        valid_order = False
                        break
                    else:
                        valid_order = True

                if valid_order:
                    clean_orders.append(order)

            if platform_found is False:
                customer_platforms.append([customer_name, 'None found', 0])

            order_dict[customer_name] = clean_orders

    return order_dict, customer_platforms


# Execute `main()` function
if __name__ == '__main__':
    pass
