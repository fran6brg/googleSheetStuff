# **************************************************************************** #
#                                                                              #
#                                                         :::      ::::::::    #
#    read.py                                            :+:      :+:    :+:    #
#                                                     +:+ +:+         +:+      #
#    By: francisberger <francisberger@student.42    +#+  +:+       +#+         #
#                                                 +#+#+#+#+#+   +#+            #
#    Created: 2019/12/31 21:05:06 by francisberg       #+#    #+#              #
#    Updated: 2020/01/01 05:31:15 by francisberg      ###   ########.fr        #
#                                                                              #
# **************************************************************************** #

# functions
from gspread_lib import *
from gspread_format_lib import *

if __name__ == '__main__':
    ss_name = 'compte courant'
    ws_name = 'transactions'
    ss = get_ss(ss_name)
    ws = get_ws_from_ss(ss, ws_name)
    update_data_with(ss, 'CA20200101_003947', ws)
    # data = get_records(ws)
    # bold_cell(ws, 'B1:B1')
    # update_cell_label(ws, 'A1', 'col A')
    # row = get_values_inside_row(ws, 3)
    # col = get_values_inside_col(ws, 3)