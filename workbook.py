import xlrd
import os

from xls import Xls


def load(path):
    abs_path = []
    for root, dirs, files in os.walk(path, topdown=False):
        for name in files:
            abs_path.append(os.path.join(root, name))

    abs_path.sort()
    wb = []
    for name in abs_path:
        wb.append(Xls(name))
    return wb

