# **************************************************************************** #
#                                                                              #
#                                                         :::      ::::::::    #
#    gspread_format_lib.py                              :+:      :+:    :+:    #
#                                                     +:+ +:+         +:+      #
#    By: francisberger <francisberger@student.42    +#+  +:+       +#+         #
#                                                 +#+#+#+#+#+   +#+            #
#    Created: 2020/01/01 00:50:45 by francisberg       #+#    #+#              #
#    Updated: 2020/01/01 04:54:50 by francisberg      ###   ########.fr        #
#                                                                              #
# **************************************************************************** #

import gspread_formatting as gsf

def align_cells(ws, cells):
    print "align_cells range=", cells
    fmt = gsf.cellFormat(
        horizontalAlignment='CENTER'
    )
    gsf.format_cell_range(ws, cells, fmt)
    pass


def bold_cell(ws, cells):
    print "bold_cell range=", cells
    fmt = gsf.cellFormat(
        textFormat=gsf.textFormat(
            bold=True)
    )
    gsf.format_cell_range(ws, cells, fmt)
    pass