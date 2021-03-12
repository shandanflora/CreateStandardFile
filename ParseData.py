import SrcFormat
import xlrd
from Common import Common
from Common import Component


class ParseData(object):
    col_no_tolerance = 0
    col_no_tc = 0
    col_no_value = 0
    col_no_voltage = 0
    col_no_designator = 0
    col_no_footprint = 0
    col_no_quantity = 0
    dict_cap = {}
    dict_res = {}
    dict_other = {}

    def __init__(self):
        pass

    def __del__(self):
        pass

    @classmethod
    def get_dict_cap(cls):
        return cls.dict_cap

    @classmethod
    def get_dict_res(cls):
        return cls.dict_res

    @classmethod
    def get_dict_other(cls):
        return cls.dict_other

    @classmethod
    def get_col_no(cls, book):
        sheet = book.sheet_by_index(0)
        col_max = sheet.ncols
        for i in range(1, col_max):
            if sheet.cell(0, i).value == "Tolerance":
                cls.col_no_tolerance = i
            elif sheet.cell(0, i).value == "TC":
                cls.col_no_tc = i
            elif sheet.cell(0, i).value == "Value":
                cls.col_no_value = i
            elif sheet.cell(0, i).value == "Voltage":
                cls.col_no_voltage = i
            elif sheet.cell(0, i).value == "Designator":
                cls.col_no_designator = i
            elif sheet.cell(0, i).value == "Footprint":
                cls.col_no_footprint = i
            elif sheet.cell(0, i).value == "Quantity":
                cls.col_no_quantity = i

    @classmethod
    def set_format(cls, sheet, row_no):
        src_format = SrcFormat.SrcData()
        src_format.Tolerance = str(sheet.cell(row_no, cls.col_no_tolerance).value)
        src_format.TC = str(sheet.cell(row_no, cls.col_no_tc).value)
        src_format.Value = str(sheet.cell(row_no, cls.col_no_value).value)
        src_format.Voltage = str(sheet.cell(row_no, cls.col_no_voltage).value)
        src_format.Designator = str(sheet.cell(row_no, cls.col_no_designator).value)
        src_format.Footprint = str(sheet.cell(row_no, cls.col_no_footprint).value)
        src_format.Quantity = sheet.cell(row_no, cls.col_no_quantity).value
        return src_format

    @classmethod
    def readSrcFile(cls, file):
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(0)
        row_max = sheet.nrows
        cls.get_col_no(book)
        n_cap = 1
        n_res = 1
        n_other = 1

        for i in range(1, row_max):
            component = Common.isComponent(sheet.cell(i, cls.col_no_value).value,
                                           sheet.cell(i, cls.col_no_voltage).value,
                                           sheet.cell(i, cls.col_no_footprint).value)
            if component == Component.CAPACITANCE:
                src_format = cls.set_format(sheet, i)
                cls.dict_cap[n_cap] = src_format
                n_cap += 1
            elif component == Component.RESISTANCE:
                src_format = cls.set_format(sheet, i)
                cls.dict_res[n_res] = src_format
                n_res += 1
            else:
                src_format = cls.set_format(sheet, i)
                cls.dict_other[n_other] = src_format
                n_other += 1


if __name__ == "__main__":
    ParseData().readSrcFile('/Users/user-b016/Desktop/111/test/LX850_MB_V2.3(AD18 Export).xls')
