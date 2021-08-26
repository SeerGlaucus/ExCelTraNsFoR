import xlrd
import xlutils.copy


class Xls:
    path = ''
    reader = None
    writer = None
    changed = False

    def __init__(self, path):
        self.path = path
        self.reader = xlrd.open_workbook(path, formatting_info=True)
        self.writer = xlutils.copy.copy(self.reader)

    def write(self, sheet, row, col, val):
        self.writer.get_sheet(sheet - 1).write(row - 1, col - 1, val)
        self.changed = True

    def save(self):
        self.writer.save(self.path)

    def read(self, sheet, row, col):
        return self.reader.sheet_by_index(sheet - 1).cell(row - 1, col - 1).value

    def read_row_values(self, sheet, row, start_col=1, end_col=None):
        if end_col is not None:
            end_col = end_col - 1
        return self.reader.sheet_by_index(sheet - 1).row_values(row - 1, start_col - 1, end_col)

    def read_col_values(self, sheet, col, start_row=1, end_row=None):
        if end_row is not None:
            end_row = end_row - 1
        return self.reader.sheet_by_index(sheet - 1).col_values(col - 1, start_row - 1, end_row)


# def main():
#     xls = Xls("D:\\PycharmProjects\\pythonProject\\env\\excel\\xread2.xls")
#     print(xls.read(1, 1, 1))
#     print(xls.read(2, 2, 2))
#     xls.write(1, 3, 3, "hello")
#     xls.save()
#     print(xls.read(1, 3, 3))
#
#
# main()
