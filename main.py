'''
Advanced Python Project
'''
import openpyxl

class CheckPs:
    def __init__(self,n_1):
        self.n=n_1
    def check_ps(self):
        path = "data.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj["Sheet1"]
        max_r = sheet_obj.max_row
        c = 0
        for i in range(1, max_r + 1):
            cell_obj = sheet_obj.cell(row=i, column=1)
            if(cell_obj.value == self.n):
                c = 1
        return c

class PrintPs:

    def print_ps(self):
        path = "data.xlsx"
        wb_obj = openpyxl.load_workbook(path)
        sheet_obj = wb_obj["Sheet1"]
        max_r = sheet_obj.max_row
        for i in range(2, max_r + 1):
            cell_obj = sheet_obj.cell(row=i, column=1)
            print(cell_obj.value)


def main():
    print("The List of Ps ID :")
    p_1=PrintPs()
    p_1.print_ps()
    print("Enter The Ps ID")
    n_1 = int(input())
    c_1 = CheckPs(n_1)
    val = c_1.check_ps()
    if(val == 1):
        print("Present")
    else:
        print("Not Present")

if __name__ == "__main__":
    main()

