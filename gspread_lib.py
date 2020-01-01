# **************************************************************************** #
#                                                                              #
#                                                         :::      ::::::::    #
#    gspread_lib.py                                     :+:      :+:    :+:    #
#                                                     +:+ +:+         +:+      #
#    By: francisberger <francisberger@student.42    +#+  +:+       +#+         #
#                                                 +#+#+#+#+#+   +#+            #
#    Created: 2020/01/01 00:50:52 by francisberg       #+#    #+#              #
#    Updated: 2020/01/01 05:27:28 by francisberg      ###   ########.fr        #
#                                                                              #
# **************************************************************************** #

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import hashlib
import time

# scope
scope = ["https://spreadsheets.google.com/feeds",
        'https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"]

# get client
creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)
client = gspread.authorize(creds)

# UTILS

def pprint_nl(data):
    pprint(data)
    print ""
    pass


# GET

# get a specific worksheet
def get_ss(ss_name):
    try:
        return client.open(ss_name) 
    except:
        print "Error : this worksheet does not exist."

# get a specific spreadsheet
def get_ws_from_ss(ss, ws_name):
    print "get_ws_from_ss", ws_name, "of ss.id =", ss.id
    try:
        return ss.worksheet(ws_name) 
    except:
        print "Error : this sheet does not exist."
        pass

# get a list of all records
def get_records(ws):
    print "get_records of ws.id=", ws.id 
    try:
        data = ws.get_all_records()
        pprint_nl(data)
        return data 
    except:
        print "Error : function get_records()."

# get col values inside ws
def get_values_inside_col(ws, col_int):
    print "get_values_inside_col of ws.id=", ws.id, "| col=", col_int
    try:
        data = ws.col_values(col_int)
        # pprint_nl(data)
        return data 
    except:
        print "Error : function get_values_inside_col()."

# get row values inside ws
def get_values_inside_row(ws, row_int):
    print "get_values_inside_row of ws.id=", ws.id, "| row=", row_int
    try:
        data = ws.row_values(row_int)
        # pprint_nl(data)
        return data 
    except:
        print "Error : function get_values_inside_row()."

# get cell value inside ws
def get_cell_value(ws, col_int, row_int):
    # print "get_cell_value of col=", col_int, "row=", row_int 
    try:
        data = ws.cell(col_int, row_int).value
        # pprint_nl(data)
        return data 
    except:
        print "Error : function get_cell_value()."



# UPDATE

def update_cell_label(ws, label, value):
    print "update_cell_label label=", label
    try:
        ws.update_acell(label, value)
    except:
        print "Error : update_cell_label."

def update_cell_numeric(ws, row_int, col_int, value):
    print "update_cell_numeric col=", col_int, "row=", row_int
    try:
        ws.update_cell(row_int, col_int, value)
    except:
        print "Error : update_cell_numeric."

def add_other(labels_ws, other):
    nb_match = labels_ws.row_count
    tab = [other, "to classify"]
    labels_ws.insert_row(tab, index=nb_match)
    pass

def find_id(ids, id_):
    if id_ in ids:
        print "transaction already registered"
        return 1
    else:
        print "new transaction"
        return 0
    pass

def update_data_with(ss, releve, ws):
    tr_ss = get_ss(releve)
    tr_ws = get_ws_from_ss(tr_ss, 'Sheet0')
    i = tr_ws.find("Date").row + 1 + 160
    nb_tr = tr_ws.row_count - i
    labels_ws = ss.worksheet('with_labels')
    count = 0;
    ids = get_values_inside_col(ws, 1)
    matched_other = get_values_inside_col(labels_ws, 1)
    while i <= nb_tr:
        row_values = get_values_inside_row(tr_ws, i)
        date_compta = row_values[0]
        libelle = row_values[1]
        pprint(libelle)
        ope_type = libelle.splitlines()[0].strip(' ')
        if ope_type == "VIREMENT EN VOTRE FAVEUR" or ope_type == "PRELEVEMENT" or ope_type == "VIREMENT EMIS":
            other = "na"
            date_tr = date_compta
        else:
            other = libelle.splitlines()[1].strip(' ')[:-5].strip(' ')
            date_tr = libelle.splitlines()[1].strip(' ')[-5:].strip(' ')
        try:
            debit = row_values[2].encode('ascii', 'ignore').decode('ascii').strip()
        except:
            debit = ""
        if debit != "":
            debit = float(debit.replace(',','.'))
        else:  
            debit = 0.0
        try:
            credit = row_values[3].encode('ascii', 'ignore').decode('ascii').strip()
        except:
            credit = ""
        if credit != "":
            credit = float(credit.replace(',','.'))
        else:
            credit = 0.0
        id_ = str(int(hashlib.md5(date_compta + libelle + str(debit) + str(credit)).hexdigest(), 16))
        
        if find_id(ids, id_):
            i += 1
            time.sleep(5)
            continue

        if other not in matched_other:
            add_other(labels_ws, other)

        label = '=VLOOKUP(E' + str(2 + i) + ';with_labels!$A:$B;2;false)'
            
        tab = [id_, date_compta, date_tr, ope_type, other, debit, credit, label, libelle]
        # pprint(tab)
        # print ""
        # print "insert at row", 2
        ws.insert_row(tab, index=2)
        i += 1
        time.sleep(5)
    pass