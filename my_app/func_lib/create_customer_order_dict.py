def create_customer_order_dict(my_rows):
    #
    # Takes an unordered list of rows (order_rows)
    # Returns an order_dict: {customer_name:[[order1],[order2],[orderN]]}
    # If we wanted to sort on customer name my_rows.sort(key=lambda x: x[0])
    #
    order_dict = {}
    for my_row in my_rows:
        customer = my_row[0]
        orders = []
        # Is customer already in the order dict ?
        if customer in order_dict:
            orders = order_dict[customer]
            orders.append(my_row)
            order_dict[customer] = orders
        else:
            orders.append(my_row)
            order_dict[customer] = orders

    return order_dict
