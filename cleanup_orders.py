import time


def cleanup_orders(customer_list, master_dict, col_name_list):

    for col_num, col_name in enumerate(col_name_list):
        # print(col_num, col_name)
        if col_name == 'Bundle Product ID':
            sku_col_num = col_num
        if col_name == 'Total Bookings':
            booking_col_num = col_num

    for customer in customer_list:
        orders = master_dict[customer[0]]

        customer_revenue = 0
        orders_negative = []
        orders_positive = []
        orders_removed = []
        zero_order = 0

        # Let's make two lists of orders (+ & -)
        # and throw out zero dollar orders
        for order in orders:
            sku = order[sku_col_num]
            booking = order[booking_col_num]

            if booking > 0 :
                # This appears to be a good order
                orders_positive.append(order)
            elif booking < 0:
                # We found a negative booking
                orders_negative.append(order)
            elif booking == 0:
                # We found a zero booking
                # Let's not include it
                zero_order += 1
                pass

        # print('-------------------------------')
        # print('BEFORE:', customer)
        # print('pos orders: ', len(orders_positive), orders_positive)
        # print('neg orders: ', len(orders_negative), orders_negative)
        # print('zero orders: ', zero_order)
        # print('removed orders: ', len( orders_removed), orders_removed)
        # print ('TOTAL ORDERS: ',len(master_dict[customer]))
        # print('-------------------------------')



        # Ok for this customer we have orders org'd in two lists
        # loop over the positive  orders and look for any +/- duplicate transactions
        # We remove these since it is a net zero revenue amount
        while len(orders_positive) > 0:
            order_count = len(orders_positive)
            #print('order count',order_count)

            for i, order_pos in enumerate(orders_positive):
                i += 1
                sku = order_pos[sku_col_num]
                booking = order_pos[booking_col_num]
                #print(sku,booking)
                dupe = False
                #print ('i',i)

                # Look in this customers orders for minus revenue
                for order_neg in orders_negative:
                    if order_neg[sku_col_num] == sku and order_neg[booking_col_num] == (booking * -1):
                        # We have a dupe so remove both the + and - transaction
                        # Start all over then
                        dupe = True
                        orders_removed.append(order_pos)
                        orders_removed.append(order_neg)

                        orders_negative.remove(order_neg)
                        orders_positive.remove(order_pos)
                        break

                if dupe:
                    # break out and start over
                    break

            if i == order_count:
                # break out and start over
                # print (i,order_count)
                # time.sleep(3)
                break
            else:
                #print (i,order_count)
                #time.sleep(3)
                continue

        master_dict[customer[0]] = orders_positive
        # print('-------------------------------')
        # print('AFTER:', customer)
        # print('pos orders: ', len(orders_positive), orders_positive)
        # print('neg orders: ', len(orders_negative), orders_negative)
        # print('zero orders: ', zero_order)
        # print('removed orders: ', len( orders_removed), orders_removed)
        # print ('TOTAL ORDERS: ',zero_order+len(orders_positive)+len(orders_negative)+len(orders_removed))
        # print('-------------------------------')
        #time.sleep(5)
    return master_dict


# Execute `main()` function
if __name__ == '__main__':
    customer_list = ['ABU DHABI COMMERCIAL BANK']
    scrub_orders(customer_list, master_dict)

