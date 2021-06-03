"""
Advanced Python Project
"""

import openpyxl
from openpyxl import Workbook


class CheckPs:
    def __init__(self, n_1):
        self.n = n_1

    def check_ps(self):
        path = "data.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj["Sheet1"]
        max_r = sheet_obj.max_row
        c = 0
        for i in range(1, max_r + 1):
            cell_obj = sheet_obj.cell(row=i, column=1)
            if cell_obj.value == self.n:
                c = 1
        return c


class PrintPs:
    @staticmethod
    def print_ps():
        path = "data.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj["Sheet1"]
        max_r = sheet_obj.max_row
        for i in range(2, max_r + 1):
            cell_obj = sheet_obj.cell(row=i, column=1)
            print(cell_obj.value)


class ExactPs:
    def __init__(self, n_1):
        self.n = n_1

    def exactps(self):
        path = "data.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj["Sheet2"]
        max_r = sheet_obj.max_row
        for i in range(1, max_r + 1):
            cell_obj = sheet_obj.cell(row=i, column=1)
            if cell_obj.value == self.n:
                return i


class Detailsps:
    def __init__(self, n):
        self.n = n

    def detailsPs(self):
        path = "data.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj_1 = wb_obj["Sheet1"]
        max_c = sheet_obj_1.max_column
        # out put file creation
        wb = Workbook()
        filepath = "output.xlsx"
        wb.save(filepath)
        wb = openpyxl.load_workbook(filepath)
        sheet = wb.active
        for j in range(2, max_c + 1):
            cell_obj_1 = sheet_obj_1.cell(row=1, column=j)
            sheet.cell(row=j, column=1).value = cell_obj_1.value
            cell_obj_2 = sheet_obj_1.cell(row=self.n, column=j)
            sheet.cell(row=j, column=2).value = cell_obj_2.value
            str = "The {c1} Grade is {c2}".format(c1=cell_obj_1.value, c2=cell_obj_2.value)
            print(str)
        sheet_obj_2 = wb_obj["Sheet2"]
        max_c = sheet_obj_2.max_column
        for j in range(2, max_c + 1):
            cell_obj_1 = sheet_obj_2.cell(row=1, column=j)
            sheet.cell(row=j, column=3).value = cell_obj_1.value
            cell_obj_2 = sheet_obj_2.cell(row=self.n, column=j)
            sheet.cell(row=j, column=4).value = cell_obj_2.value
            str = "The {c1}  is {c2}".format(c1=cell_obj_1.value, c2=cell_obj_2.value)
            print(str)
        wb.save(filepath)

def main():
    print("The List of Ps ID :")
    PrintPs.print_ps()
    print("Enter The Ps ID")
    n_1 = int(input())
    c_1 = CheckPs(n_1)
    val = c_1.check_ps()
    if val == 1:
        print("Present")
        E_1 = ExactPs(n_1)
        A = E_1.exactps()
        D_1 = Detailsps(A)
        D_1.detailsPs()
    else:
        print("Not Present")


if __name__ == "__main__":
    main()
