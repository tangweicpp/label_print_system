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

    sql = f'''SELECT distinct t1.到货单编号 FROM erpbase..tblToRec t1 inner join erpbase..tblToRecEntry t2
    on t1.到货单编号 = t2.到货单编号 where t1.收货日期 < '{po_query['end_date']}' and t1.收货日期 > '{po_query['start_date']}'
    and t2.物料编号 not like '01.01.01%'  order by t1.到货单编号 '''

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
        select t2.F_101 as 料号, t2.FName as 物料名称,t1.到货批号, sum(t1.到货数量) as 总数量, t3.单位,sum(t1.到货数量) / t3.单位 as 标签数量,t1.有效期至
        from erpbase..tblToRecEntry t1
        inner join AIS20141114094336.dbo.t_ICItem  t2 on t2.FNumber = t1.物料编号
        left join erpbase.dbo.unitlist t3 on t3.料号 = t2.F_101
        where t1.到货单编号 = '{po_query['entry_number']}' and substring(t2.F_101,1,2) <> '60'
        group by t2.F_101,t2.FName,t1.到货批号,t3.单位,t1.有效期至
    '''
    # print(sql)
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
        result['lbl_term'] = xstr(row[6])
        result['lbl_print_again_qty'] = '0'
        result['lbl_printing_qty'] = '0'
        result['print_reason'] = ''

        if not result['unit_qty']:
            ret_info['ret_desc'] = f"物料：{result['part_name']}  料号：{result['part_no']} 没有维护单位数量，请先维护好，否则无法打印标签"
            ret_info['ret_code'] = 201
            ret_info['ret_part_name'] = result['part_name']
            ret_info['ret_part_no'] = result['part_no']
            return False

        sql = f"SELECT count(*) FROM TBL_MATERIAL_PRINT_HISTORY WHERE remark = '{po_query['entry_number']}' AND part_id = '{result['part_no']}' AND lot_id = '{result['lot_id']}' AND printed_flag = '1'"
        result['lbl_printed_qty'] = conn.OracleConn.query(sql)[0][0]
        result['lbl_non_printed_qty'] = float(result['lbl_qty']) - \
            result['lbl_printed_qty']

        json_data.append(result)

    ret_info['ret_desc'] = "success"
    ret_info['ret_code'] = 200

    return json_data


# Print label
def print_handle(sel_data, ret_info, flag):
    if not sel_data:
        ret_info['ret_desc'] = "没有数据"
        ret_info['ret_code'] = 201
        print("没有数据")
        return False

    for row in sel_data:
        print(row)
        lot_list = get_print_lot(row)
        # print(lot_list)

        for pce_id in lot_list:
            label_content = f'''"ITEM","{row['part_no']}";"INVENTORY_ID","{pce_id}"'''

            print_label(label_content, row, pce_id, flag)

        time.sleep(2)

    ret_info['ret_desc'] = "标签打印成功"
    ret_info['ret_code'] = 200
    return True


def print_label(label_content, row, pce_id, flag):
    # insert to print table
    sql = f''' insert into erpdata.dbo.tblME_PrintInfo(PrinterNameID,BartenderName,Content,Flag,Createdate,EVENT_SOURCE,EVENT_ID,LABEL_ID,PRINT_QTY)
               values('HT_ST','MATERIAL.btw','{label_content}','0',GetDate(),'STORE','MATERIAL','{row['entry_no']}','1')
           '''
    # print(sql)
    conn.MssConn.exec(sql)

    # insert to mes
    sql = f'''insert into ERPBASE..TblERPFLToME(STOCK_TYPE,STOCK_ID,PRD_ID,PRD_VER,QTY,PRD_DATE,EFF_DATE,flag,CreateDate,FStauts)
            values('M','{pce_id}','{row['part_no']}','A','{row['unit_qty']}',getdate(),'{row['lbl_term']}',0,getdate(),0)
        '''

    # print(sql)
    conn.MssConn.exec(sql)

    # insert to print history
    sql = f'''insert into TBL_MATERIAL_PRINT_HISTORY(PART_ID,PART_NAME,LOT_ID,PCE_ID,PRINTED_DATE,PRINTED_BY,PRINTED_FLAG,REMARK,REASON)
    values('{row['part_no']}','{row['part_name']}','{row['lot_id']}','{pce_id}',sysdate,'{row['user_name']}','{flag}','{row['entry_no']}','{row['print_reason']}')
    '''
    # print(sql)
    conn.OracleConn.exec(sql)


def get_print_lot(row):
    inventory_lot = row['part_no'] + '_' + row['lot_id']
    print_list = []
    for i in range(int(row['lbl_printing_qty'])):
        sql = f"select nvl(max(max_id) + 1, 1) from TBL_MATERIAL_SEQ_ID  WHERE inventory_lot = '{inventory_lot}'"
        ret = conn.OracleConn.query(sql)[0][0]
        if ret == 1:
            conn.OracleConn.exec(
                f"insert into TBL_MATERIAL_SEQ_ID(INVENTORY_LOT,MAX_ID) values('{inventory_lot}',1)")
        else:
            conn.OracleConn.exec(
                f"update TBL_MATERIAL_SEQ_ID set MAX_ID = {ret} where INVENTORY_LOT = '{inventory_lot}' ")

        print_list.append(inventory_lot + ('00000' + str(ret))[-5:])

    return print_list


def set_unit_qty(sel_data, ret_info):
    sql = f"insert into erpbase.dbo.unitlist(料号,单位) values('{sel_data['partID']}','{sel_data['unitQty']}')"
    conn.MssConn.exec(sql)

    ret_info['ret_desc'] = "单位用量维护成功"
    ret_info['ret_code'] = 200
