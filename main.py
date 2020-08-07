'''
@File    :   main.py
@Time    :   2020/07/09 15:30:24
@Author  :   Tang wei
@Version :   1.0
@Contact :   wei.tang_ks@ht-tech.com
@License :   (C)Copyright 2020-2021
@Desc    :   None
'''
from flask import Flask
from flask import jsonify
from flask import request
from flask import make_response
from flask_cors import CORS
import handle as h
import handle_po_mgr as hpm
import handle_print as hpt
import json

# CORS
app = Flask(__name__)
CORS(app)


# Login
@app.route('/login', methods=['GET', 'POST'])
def r_login():
    if request.method == 'POST':
        username = request.values.get('username')
        password = request.values.get('password')
        if h.check_account(username, password):
            return make_response("success", 200)
        else:
            return make_response("用户名或密码不存在", 201)


# Get customer list
@app.route('/cust_code_list', methods=['GET', 'POST'])
def r_get_cust_code_list():
    if request.method == 'GET':
        json_data = h.get_cust_code_list()
        return make_response(jsonify(json_data), 200)


# Get customer po template
@app.route('/po_template', methods=['GET', 'POST'])
def r_get_po_template():
    if request.method == 'POST':
        cust_code = request.values.get('custCode')
        json_data = h.get_po_template(cust_code)
        return make_response(jsonify(json_data), 200)


# Upload po file
@app.route('/upload_po_file', methods=['GET', 'POST'])
def r_upload_po_file():
    if request.method == 'POST':
        f = request.files.get('poFile')
        po_header = {}
        po_header['user_name'] = request.values.get('userName')
        po_header['cust_code'] = request.values.get('custCode')
        po_header['po_type'] = request.values.get('poType')
        po_header['bonded_type'] = request.values.get('bondedType')
        po_header['offer_sheet'] = request.values.get('offerSheet')
        po_header['need_delay'] = request.values.get('needDelay')
        po_header['delay_days'] = request.values.get('delayDays')
        po_header['need_mail_tip'] = request.values.get('needMailTip')
        po_header['mail_tip'] = request.values.get('mailTip')
        po_header['file_id'] = request.values.get('fileID')
        po_header['err_desc'] = '错误测试'

        json_data = h.upload_po_file(f, po_header)
        if not json_data:
            return make_response(jsonify({'err_desc': po_header['err_desc'], 'status': 201}))
        else:
            return make_response(jsonify({'data': json_data, 'status': 200}))


# Update progress
@ app.route('/update_progress', methods=['GET', 'POST'])
def r_update_progress():
    if request.method == 'GET':
        user_key = request.args.get('userKey')
        num = h.get_progress(user_key)
        return make_response(jsonify({"progress": num}), 200)


# Query po data
@app.route('/query_po_data', methods=['GET', 'POST'])
def r_query_po_data():
    if request.method == 'GET':
        po_query = {}
        po_query['cust_code'] = request.args.get('custCode')
        po_query['cust_lot_id'] = request.args.get('custLotID')
        ret = hpm.get_po_data(po_query)
        return make_response(jsonify({"info": ret}), 200)


# Query entry no
@app.route('/query_entry_no', methods=['GET', 'POST'])
def r_query_entry_no():
    if request.method == 'GET':
        po_query = {}
        ret_info = {}
        po_query['start_date'] = request.args.get('startDate')
        po_query['end_date'] = request.args.get('endDate')

        ret_data = hpt.get_entry_no(po_query, ret_info)
        if not ret_data:
            return make_response(jsonify(ret_info))

        else:
            ret_info['ret_data'] = ret_data
            return make_response(jsonify(ret_info))


# Query entry no
@app.route('/query_entry_data', methods=['GET', 'POST'])
def r_query_entry_data():
    if request.method == 'GET':
        po_query = {}
        ret_info = {'ret_part_name': '', 'ret_part_no': ''}

        po_query['entry_number'] = request.args.get('entryNumber')

        ret_data = hpt.get_entry_data(po_query, ret_info)
        if not ret_data:
            return make_response(jsonify(ret_info))

        else:
            ret_info['ret_data'] = ret_data
            return make_response(jsonify(ret_info))


# Print lables
@app.route('/print_label', methods=['GET', 'POST'])
def r_print_label():
    if request.method == 'POST':
        sel_data = json.loads(request.get_data(as_text=True))
        # print(sel_data)
        ret_info = {}
        hpt.print_handle(sel_data, ret_info, flag='1')
        return make_response(jsonify(ret_info))


# Print lables agagin
@app.route('/print_label_again', methods=['GET', 'POST'])
def r_print_label_again():
    if request.method == 'POST':
        sel_data = json.loads(request.get_data(as_text=True))
        # print(sel_data)
        ret_info = {}
        for row in sel_data:
            row['lbl_printing_qty'] = row['lbl_print_again_qty']

        hpt.print_handle(sel_data, ret_info, flag='2')
        return make_response(jsonify(ret_info))


# Set unit qty
@app.route('/set_unit_qty', methods=['GET', 'POST'])
def r_set_unit_qty():
    if request.method == 'POST':
        sel_data = json.loads(request.get_data(as_text=True))
        print(sel_data)
        ret_info = {}
        hpt.set_unit_qty(sel_data, ret_info)
        return make_response(jsonify(ret_info))


# Run server
if __name__ == "__main__":
    app.run(host='0.0.0.0', debug=True, threaded=True)
