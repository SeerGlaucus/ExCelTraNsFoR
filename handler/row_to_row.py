import re


def row_to_row(wbs, script):
    groups = re.search(r'(\d+)\[(\d+)\]\[(\d+),(\d+)\.\.\.\]->(\d+)\[(\d+)\]\[(\d+),(\d+)\.\.\.\]', script)
    from_wb = int(groups.group(1))
    from_sheet = int(groups.group(2))
    from_row = int(groups.group(3))
    from_start_col = int(groups.group(4))
    values = wbs[from_wb - 1].read_row_values(from_sheet, from_row, from_start_col)
    to_wb = int(groups.group(5))
    to_sheet = int(groups.group(6))
    to_row = int(groups.group(7))
    to_start_col = int(groups.group(8))
    for i in range(len(values)):
        wbs[to_wb - 1].write(to_sheet, to_row, to_start_col + i, values[i])
