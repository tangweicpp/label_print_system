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


# Get entry number
def get_entry_no(po_query, ret_info):
    json_data = []

    sql = f"SELECT distinct 到货单编号 FROM erpbase..tblToRec where 收货日期 < '{po_query['end_date']}' and 收货日期 > '{po_query['start_date']}' order by 到货单编号 "
    results = conn.MssConn.query(sql)
    for row in results:
        result = {}
        result['value'] = xstr(row[0])
        result['entryNumber'] = xstr(row[0])

        json_data.append(result)

    ret_info['ret_desc'] = "success"
    ret_info['ret_code'] = 200
    return json_data


# Get entry data
def get_entry_data(po_query, ret_info):
    json_data = []

    # By 入库单
    # sql = f'''
    #     select t2.F_101 as 料号, t2.FName as 物料名称,t1.到货批号, t1.实入数量 as 总数量, t3.单位,t1.实入数量 / t3.单位 as 标签数量,'' as 已打印标签数量,
    #     t1.实入数量 / t3.单位 as 剩余打印数量, '' as 本次打印数量 from erpbase..TblToInSub t1
    #     inner join AIS20141114094336.dbo.t_ICItem  t2 on t2.FNumber = t1.物料编号
    #     inner join erpbase.dbo.unitlist t3 on t3.料号 = t2.F_101
    #     where t1.入库单编号 = '{po_query['entry_number']}'
    # '''

    # By 到货单
    sql = f'''
        select t2.F_101 as 料号, t2.FName as 物料名称,t1.到货批号, t1.到货数量 as 总数量, t3.单位,t1.到货数量 / t3.单位 as 标签数量,'' as 已打印标签数量,
        t1.到货数量 / t3.单位 as 剩余打印数量, '' as 本次打印数量 from erpbase..tblToRecEntry t1
        inner join AIS20141114094336.dbo.t_ICItem  t2 on t2.FNumber = t1.物料编号
        left join erpbase.dbo.unitlist t3 on t3.料号 = t2.F_101 
        where t1.到货单编号 = '{po_query['entry_number']}'
    '''

    results = conn.MssConn.query(sql)
    if not results:
        ret_info['ret_desc'] = '查询不到该入库单号，请确认输入的是否正确？'
        ret_info['ret_code'] = 201
        return False

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

        if not result['unit_qty']:
            ret_info['ret_desc'] = f"物料：{result['part_name']}  料号：{result['part_no']} 没有维护单位数量，请先维护好，否则无法打印标签"
            ret_info['ret_code'] = 201
            return False

        json_data.append(result)

    ret_info['ret_desc'] = "success"
    ret_info['ret_code'] = 200

    return json_data


# Print label
def print_handle(sel_data, ret_info):
    if not sel_data:
        ret_info['ret_desc'] = "没有数据"
        ret_info['ret_code'] = 201
        print("没有数据")
        return False

    for row in sel_data:
        print(row)
        lot_list = get_print_lot(row)

        for pce_id in lot_list:
            label_content = f'''"ITEM","{row['part_no']}";"INVENTORY_ID","{pce_id}"'''
            print_label(label_content, sel_data['entry_no'])

    ret_info['ret_desc'] = "标签打印成功"
    ret_info['ret_code'] = 200
    return True


def print_label(label_content, entry_no):
    sql = f''' insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Content,Flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY) 
               values('HT_ST','MATERIAL.btw','{label_content}','0',GetDate(),'STORE','MATERIAL','{entry_no}','1')"
           '''
    print(sql)
    # conn.MssConn.exec(sql)


def get_print_lot(row):
    inventory_lot = row['part_no'] + '_' + row['lot_id']
    print_list = []
    for i in range(row['lbl_printing_qty']):
        sql = f"select nvl(max(max_id) + 1, 1) from TBL_MATERIAL_SEQ_ID  WHERE inventory_lot = '{inventory_lot}'"
        ret = conn.OracleConn.query(sql)
        if ret == '1':
            conn.OracleConn.exec(
                f"insert into TBL_MATERIAL_SEQ_ID(INVENTORY_LOT,MAX_ID) values('{inventory_lot}',1)")
        else:
            conn.OracleConn.exec(
                f"update TBL_MATERIAL_SEQ_ID set MAX_ID = {ret} where INVENTORY_LOT = '{inventory_lot}''")

        print_list.append(inventory_lot + ('0000' + str(ret))[-4:])

    return print_list
