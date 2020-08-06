'''
@File    :   handle_po_mgr.py
@Time    :   2020/07/31 15:50:12
@Author  :   Tony Tang
@Version :   1.0
@Contact :   wei.tang_ks@ht-tech.com
@License :   (C)Copyright 2020-2021
@Desc    :   customer po mgr
'''
import connect_db as conn
import time
import os
import send_email as se
import pandas as pd
import numpy as np
import json
import re
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl import load_workbook
from xlrd import open_workbook
from itertools import groupby


os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'


# None to ''
def xstr(s):
    return '' if s is None else str(s).strip()


# Get po data
def get_entry_data(po_query):
    json_data = []

    sql = f'''
        select t2.F_101 as 料号, t2.FName as 物料名称,t1.到货批号, t1.实入数量 as 总数量, t3.单位,t1.实入数量 / t3.单位 as 标签数量,'' as 已打印标签数量,
        t1.实入数量 / t3.单位 as 剩余打印数量, '' as 本次打印数量 from erpbase..TblToInSub t1
        inner join AIS20141114094336.dbo.t_ICItem  t2 on t2.FNumber = t1.物料编号
        inner join erpbase.dbo.unitlist t3 on t3.料号 = t2.F_101 
        where t1.入库单编号 = '{po_query['entry_number']}'
        '''
    results = conn.MssConn.query(sql)
    for row in results:
        result = {}
        result['part_no'] = xstr(row[0])
        result['part_name'] = xstr(row[1])
        result['lot_id'] = xstr(row[2])
        result['total_qty'] = xstr(row[3])
        result['unit_qty'] = xstr(row[4])
        result['lbl_qty'] = xstr(row[5])
        result['lbl_printed_qty'] = xstr(row[6])
        result['lbl_non_printed_qty'] = xstr(row[7])
        result['lbl_printing_qty'] = xstr(row[8])

        json_data.append(result)
    return json_data


# Print label
def print_label(sel_data):
    for row in sel_data:
        print(row)
