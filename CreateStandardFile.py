import os

import xlrd
import xlsxwriter

from ParseData import ParseData
from SrcFormat import ObjItem


class createStandardFile(object):

    dict_cap_col_no = {}
    dict_res_col_no = {}

    cap_no = 0
    cap_other = 0
    res_no = 0
    res_other = 0

    dict_standard_cap = {}
    dict_standard_res = {}
    dict_standard_other = {}

    def __init__(self):
        pass

    @classmethod
    def col_width(cls, sheet, row_count):
        sheet.set_row(0, 44)
        for i in range(1, row_count + 4):
            sheet.set_row(i, 27)
        sheet.set_column('A:A', 8)  # 项次
        sheet.set_column('B:B', 12)  # 宝时得料号
        sheet.set_column('C:C', 14)  # 组件料号
        sheet.set_column('D:D', 48)  # 规格描述
        sheet.set_column('E:E', 25)  # 制造商型号
        sheet.set_column('F:F', 22)  # 制造商
        sheet.set_column('G:G', 20)  # 封装
        sheet.set_column('H:H', 11)  # 装配方式
        sheet.set_column('I:I', 25)  # 位置
        sheet.set_column('J:J', 10)  # 用量
        sheet.set_column('K:K', 5)  # 单位
        sheet.set_column('L:L', 15)  # 备注

    @classmethod
    def write_head(cls, book):
        sheet = book.get_worksheet_by_name('电子BOM')
        # heading
        png = os.getcwd() + '/res/positec.png'
        sheet.insert_image('A1', png, {'x_offset': 10, 'y_offset': 20})
        style = book.add_format({'font_name': '微软雅黑',
                                 'font_size': 10,
                                 'bold': True,
                                 'border': 1,
                                 'text_wrap': True,
                                 'align': 'right',  # 对齐方式
                                 'valign': 'vcenter',  # 字体对齐方式
                                 })  # 初始化样式
        sheet.merge_range(0, 0, 0, 11, '电子BOM\nBill of Material（Electronic）', style)
        style = book.add_format({'font_name': '微软雅黑',
                                 'font_size': 10,
                                 'bold': True,
                                 'border': 1,
                                 'text_wrap': True,
                                 'align': 'left',  # 对齐方式
                                 'valign': 'vcenter',  # 字体对齐方式
                                 })  # 初始化样式
        sheet.write(1, 0, 'Version\n内控版本', style)
        sheet.write(1, 1, 'V2.3', style)
        sheet.write(1, 2, 'PCBA PN\n组件料号', style)
        sheet.write(1, 3, 'A409008500', style)
        sheet.merge_range(1, 4, 1, 11, '', style)

        sheet.write(2, 0, 'ITEM\n项次', style)
        sheet.write(2, 1, 'Positec PN\n宝时得料号', style)
        sheet.write(2, 2, 'Name\n名称', style)
        sheet.write(2, 3, 'Description\n规格描述', style)
        sheet.write(2, 4, 'MPN\n制造商型号', style)
        sheet.write(2, 5, 'MFR\n制造商', style)
        sheet.write(2, 6, 'Package\n封装', style)
        sheet.write(2, 7, 'Assembly\n装配方式', style)
        sheet.write(2, 8, 'Location\n位置', style)
        sheet.write(2, 9, 'Quantity\n用量', style)
        sheet.write(2, 10, 'Unit\n单位', style)
        sheet.write(2, 11, 'Remark\n备注', style)

    @classmethod
    def write_data(cls, book, dict_standard, offset):
        sheet = book.get_worksheet_by_name('电子BOM')
        style = book.add_format({'font_name': '微软雅黑',
                                 'font_size': 10,
                                 'bold': False,
                                 'border': 1,
                                 'text_wrap': True,
                                 'align': 'left',  # 对齐方式
                                 'valign': 'vcenter',  # 字体对齐方式
                                 })  # 初始化样式
        num = len(dict_standard)
        for i in range(0, num):
            item = dict_standard[i]
            sheet.write(i + 3 + offset, 0, item.ITEM, style)
            sheet.write(i + 3 + offset, 1, item.PositecPN, style)
            sheet.write(i + 3 + offset, 2, item.Name, style)
            sheet.write(i + 3 + offset, 3, item.Description, style)
            sheet.write(i + 3 + offset, 4, item.MPN, style)
            sheet.write(i + 3 + offset, 5, item.MFR, style)
            sheet.write(i + 3 + offset, 6, item.Package, style)
            sheet.write(i + 3 + offset, 7, item.Assembly, style)
            sheet.write(i + 3 + offset, 8, item.Location, style)
            sheet.write(i + 3 + offset, 9, item.Quantity, style)
            sheet.write(i + 3 + offset, 10, item.Unit, style)
            sheet.write(i + 3 + offset, 11, item.Remark, style)
            str_col = 'H' + str(i + 3 + offset + 1)
            sheet.data_validation(str_col, {'validate': 'list',
                                            'source': ['贴片',
                                                       '插件',
                                                       '其它']})
            str_col = 'K' + str(i + 3 + offset + 1)
            sheet.data_validation(str_col, {'validate': 'list',
                                            'source': ['pcs',
                                                       'g',
                                                       'kg',
                                                       'mm',
                                                       'm',
                                                       'mL',
                                                       'L']})

    @classmethod
    def write_content(cls, book):
        cls.write_data(book, cls.dict_standard_cap, 0)
        cls.write_data(book, cls.dict_standard_res, len(cls.dict_standard_cap))
        cls.write_data(book, cls.dict_standard_other,
                       len(cls.dict_standard_cap) + len(cls.dict_standard_res))

    @classmethod
    def isChange(cls, string):
        if string.find('.') != -1:  # is decimal
            if int(string.lstrip()[0]) == 0:
                return 1
            else:
                return 0
        else:
            return 0

    @classmethod
    def isChange_res(cls, string):
        if string.find('.') != -1:  # is decimal
            if str(string.strip()[-1]).upper() == 'R' and str(string.strip()[-2]) != 'm':
                if int(string.lstrip()[0]) == 0:
                    return 1
                else:
                    return 0
            else:
                return 0
        else:
            return 0

    @classmethod
    def changeF(cls, string):
        str_change = ''
        i = len(string) - 2
        if string.rstrip()[-2:].upper() == 'UF':  # uF
            str_change = str(int(float(string.strip()[:i]) * 1000)) + 'nF'
        elif string.rstrip()[-2:].upper() == 'NF':  # nF
            str_change = str(int(float(string.strip[:i]) * 1000)) + 'pF'
        return str_change

    @classmethod
    def changeM(cls, string):
        i = len(string) - 1
        return str(int(float(string.strip()[:i]) * 1000)) + 'M'

    @classmethod
    def get_res_col_vector(cls):
        list_col = ["名称",
                    "规格描述",
                    "制造商型号",
                    "制造商",
                    "阻值",
                    "精度",
                    "功率",
                    "封装"]
        return list_col

    @classmethod
    def get_res_col_no(cls, sheet, list_res_col):
        col_max = sheet.ncols
        for i in range(0, col_max):
            for j in list_res_col:
                if sheet.cell(1, i).value.strip() == j:
                    cls.dict_res_col_no[j] = i

    @classmethod
    def get_cap_col_vector(cls):
        list_col = ["名称",
                    "规格描述",
                    "制造商型号",
                    "制造商",
                    "电压",
                    "材质",
                    "封装",
                    "容值",
                    "精度"]
        return list_col

    @classmethod
    def get_cap_col_no(cls, sheet, list_cap_col):
        col_max = sheet.ncols
        for i in range(0, col_max):
            for j in list_cap_col:
                if sheet.cell(1, i).value.rstrip() == j:
                    cls.dict_cap_col_no[j] = i

    @classmethod
    def update_dict(cls, sheet, srcData, i, no, dict_col_no):
        obj_item = ObjItem()
        obj_item.ITEM = no + 1
        obj_item.PositecPN = ''
        obj_item.Name = sheet.cell(i, dict_col_no["名称"]).value
        obj_item.Description = sheet.cell(i, dict_col_no["规格描述"]).value
        obj_item.MPN = sheet.cell(i, dict_col_no["制造商型号"]).value
        obj_item.MFR = sheet.cell(i, dict_col_no["制造商"]).value
        obj_item.Package = sheet.cell(i, dict_col_no["封装"]).value
        obj_item.Assembly = '贴片'
        obj_item.Location = srcData.Designator
        obj_item.Quantity = srcData.Quantity
        obj_item.Unit = 'pcs'
        if srcData.Footprint[0].upper() == "C":
            cls.dict_standard_cap[no] = obj_item
        else:
            cls.dict_standard_res[no] = obj_item
        pass

    @classmethod
    def update_dict_other(cls, srcData, no):
        obj_item = ObjItem()
        obj_item.ITEM = str(no)
        obj_item.PositecPN = ''
        obj_item.Name = ''
        obj_item.Description = str(srcData.TC) + ' ' \
                               + str(srcData.Value) + ' ' \
                               + str(srcData.Tolerance) + ' ' \
                               + str(srcData.Voltage) + ' ' \
                               + str(srcData.Footprint)
        obj_item.MPN = ''
        obj_item.MFR = ''
        obj_item.Package = srcData.Footprint
        obj_item.Assembly = ''
        obj_item.Location = srcData.Designator
        obj_item.Quantity = srcData.Quantity
        obj_item.Unit = 'pcs'
        cls.dict_standard_other[no] = obj_item

    @classmethod
    def add_other_to_dict(cls):
        for i in

    @classmethod
    def find_first(cls, str_value, sheet, col):
        total_row = sheet.nrows
        list_data = []
        for i in range(2, total_row):
            if str_value.upper() == sheet.cell(i, col).value.upper():
                list_data.append(i)
        return list_data

    @classmethod
    def find_out_first(cls, list_value, str_value, sheet, col):
        list_data = []
        for i in list_value:
            if str_value == str(sheet.cell(i, col).value).upper():
                list_data.append(i)
        return list_data

    @classmethod
    def find_cap_data_item(cls, srcData, sheet):
        list_vol = cls.find_first(str(srcData.Voltage).upper(), sheet, cls.dict_cap_col_no["电压"])
        if len(list_vol) != 0:
            list_tc = cls.find_out_first(list_vol,
                                         str(srcData.TC).upper(),
                                         sheet, cls.dict_cap_col_no["材质"])
            if len(list_tc) != 0:
                footprint = str(srcData.Footprint).upper()[1:]
                list_footprint = cls.find_out_first(list_tc, footprint,
                                                    sheet, cls.dict_cap_col_no["封装"])
                if len(list_footprint) != 0:
                    if cls.isChange(srcData.Value):
                        str_value = cls.changeF(srcData.Value)
                    else:
                        str_value = str(srcData.Value)
                    list_value = cls.find_out_first(list_footprint,
                                                    str_value.upper(),
                                                    sheet, cls.dict_cap_col_no["容值"])
                    if len(list_value) != 0:
                        list_tol = cls.find_out_first(list_value,
                                                      str(srcData.Tolerance).upper(),
                                                      sheet, cls.dict_cap_col_no["精度"])
                        if len(list_tol) != 0:
                            cls.update_dict(sheet, srcData,
                                            int(list_tol[0]),
                                            cls.cap_no, cls.dict_cap_col_no)
                            cls.cap_no += 1
                        else:
                            cls.update_dict_other(srcData, cls.cap_other)
                            cls.cap_other += 1
                    else:
                        cls.update_dict_other(srcData, cls.cap_other)
                        cls.cap_other += 1
                else:
                    cls.update_dict_other(srcData, cls.cap_other)
                    cls.cap_other += 1
            else:
                cls.update_dict_other(srcData, cls.cap_other)
                cls.cap_other += 1
        else:
            cls.update_dict_other(srcData, cls.cap_other)
            cls.cap_other += 1

    @classmethod
    def find_res_data_item(cls, srcData, sheet):
        if cls.isChange_res(str(srcData.Value).upper()):
            str_res = cls.changeM(str(srcData.Value).upper())
        else:
            str_res = str(srcData.Value).upper()
        list_res = cls.find_first(str_res, sheet, cls.dict_res_col_no["阻值"])
        if len(list_res) != 0:
            footprint = str(srcData.Footprint).upper()[1:]
            list_footprint = cls.find_out_first(list_res, footprint,
                                                sheet, cls.dict_res_col_no["封装"])
            if len(list_footprint) != 0:
                list_tol = cls.find_out_first(list_footprint,
                                              str(srcData.Tolerance)[1:],
                                              sheet,
                                              cls.dict_res_col_no["精度"])
                if len(list_tol) != 0:
                    cls.update_dict(sheet, srcData,
                                    int(list_tol[0]),
                                    cls.res_no,
                                    cls.dict_res_col_no)
                    cls.res_no += 1
                else:
                    cls.update_dict_other(srcData, cls.cap_other)
                    cls.cap_other += 1
            else:
                cls.update_dict_other(srcData, cls.cap_other)
                cls.cap_other += 1
        else:
            cls.update_dict_other(srcData, cls.cap_other)
            cls.cap_other += 1

    @classmethod
    def write_excel(cls, lib_cas, lib_res, dict_para_cap, dict_para_res):
        file = '123.xlsx'
        book_obj = xlsxwriter.Workbook(file)
        sheet = book_obj.add_worksheet('电子BOM')
        cls.col_width(sheet, 20)
        cls.write_head(book_obj)  # write sheet head
        # write data
        # find cap and update dict_cap
        book_cas = xlrd.open_workbook(lib_cas)
        sheet_cas = book_cas.sheet_by_index(0)
        cls.get_cap_col_no(sheet_cas, cls.get_cap_col_vector())
        num = len(dict_para_cap)
        for i in range(1, num + 1):
            src_data = dict_para_cap[i]
            cls.find_cap_data_item(src_data, sheet_cas)
        # find cap and update dict_cap
        book_res = xlrd.open_workbook(lib_res)
        sheet_res = book_res.sheet_by_index(0)
        cls.get_res_col_no(sheet_res, cls.get_res_col_vector())
        num = len(dict_para_res)
        for i in range(1, num + 1):
            src_data = dict_para_res[i]
            cls.find_res_data_item(src_data, sheet_res)

        cls.write_content(book_obj)

        sheet.freeze_panes(3, 0)
        book_obj.close()
        print('write excel successfully!!!')


if __name__ == "__main__":
    parseData = ParseData()
    parseData.readSrcFile('/Users/user-b016/Desktop/111/test/LX850_MB_V2.3(AD18 Export).xls')
    dict_cap = parseData.get_dict_cap()
    dict_res = parseData.get_dict_res()
    dict_other = parseData.get_dict_other()
    file_cas = '/Users/user-b016/Desktop/111/test/普通贴片电容_Yageo.xls'
    file_res = '/Users/user-b016/Desktop/111/test/普通贴片电阻_Yageo.xls'
    create_file = createStandardFile()
    create_file.write_excel(file_cas, file_res, dict_cap, dict_res)
