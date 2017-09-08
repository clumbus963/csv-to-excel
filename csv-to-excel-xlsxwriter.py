#!/usr/bin/python
# -*- coding: utf-8 -*-
# ======================================================================
# Project Name    : csv-to-excel-xlswriter
# File Name       : csv-to-excel-xlswriter.py
# Encoding        : utf-8
# Creation Date   : Sep. 8, 2017
#
# Copyright @ 2017 Masaru Kawabata. All rights reserved.
#
# This source code or any portion thereof must not be
# reproduced or used in any manner whatsoever.
# ======================================================================

# Import Module
import sys
import os
import glob
import csv
from xlsxwriter.workbook import Workbook

reload(sys)
sys.setdefaultencoding('utf-8')

for csvfile in glob.glob(os.path.join('.', '<CSV_FILE>')):
    workbook = Workbook('<EXCEL_FILE>')
    worksheet = workbook.add_worksheet()
#    # 書式設定を変更する場合はworksheet.write()の時に使用するためここに定義
#    format_num = workbook.add_format()
#    format_num.set_num_format('#,##0')
    with open(csvfile, 'rb') as f:
        reader = csv.reader(f)
        for r, row in enumerate(reader):
            for c, col in enumerate(row):
#                 worksheet.write(r, c, col)
                # 列ごとに書式設定するときはここで分岐
                if r == 0:
                    worksheet.write(r, c, col, workbook.add_format({'bg_color': '#00CC00', 'align': 'center'}))
                else:
                    worksheet.write(r, c, col)

    # 列ごとの幅を設定
    worksheet.set_column(0, 0, 11)
    worksheet.set_column(1, 1, 12)
    worksheet.set_column(2, 2, 9)
    worksheet.set_column(3, 3, 115)
    worksheet.set_column(4, 4, 13)
    worksheet.set_column(5, 5, 15)
#    # 右揃えをしてみたけど１回入力モードにしないといけないのでコメント
#    format = workbook.add_format()
#    format.set_align('right')
#    worksheet.set_column(4, 4, 14, format)
#    worksheet.set_column(5, 5, 16, format)
    worksheet.set_column(6, 6, 24)
    worksheet.set_column(7, 7, 11)

    # 行ごとの書式を設定
    worksheet.set_row(0, 18)
    format = workbook.add_format({'bold': True, 'font_color': 'red'})
    worksheet.set_row(1, 18, format)
    format = workbook.add_format({'font_color': 'red'})
    worksheet.set_row(2, 18, format)
    worksheet.set_row(3, 18, format)
    worksheet.set_row(4, 18)
    worksheet.set_row(5, 18)
    worksheet.set_row(6, 18)
    worksheet.set_row(7, 18)
    worksheet.set_row(8, 18)
    worksheet.set_row(9, 18)
    worksheet.set_row(10, 18)

#    worksheet.write('A1:H1', None, workbook.add_format({'border': 1, 'border_color': 'green', 'bg_color': '#00CC00'}))

    workbook.close()
