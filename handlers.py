import re

from handler.col_summary import col_summary
from handler.col_to_col import col_to_col
from handler.map_col_operator import map_col_operator
from handler.point_to_point import point_to_point
from handler.row_summary import row_summary
from handler.row_to_row import row_to_row
from handler.search_row_based_col_operator import search_row_based_col_operator

handler_list = [
    point_to_point,
    row_to_row,
    col_to_col,
    row_summary,
    col_summary,
    map_col_operator,
    search_row_based_col_operator
]


def handler(wbs, name, script):
    eval(name)(wbs, script)
