import re


def col_summary(wbs, script):
    groups = re.search(r'(\d+)\[(\d+)\]\[(\d+)\.\.\.,(\d+)\]->(\d+)\[(\d+)\]\[(\d+),(\d+)\]', script)
    from_wb = int(groups.group(1))
    from_sheet = int(groups.group(2))
    from_start_row = int(groups.group(3))
    from_col = int(groups.group(4))
    values = wbs[from_wb - 1].read_col_values(from_sheet, from_col, from_start_row)
    val = 0.0
    for v in values:
        val = val + float(v)
    to_wb = int(groups.group(5))
    to_sheet = int(groups.group(6))
    to_row = int(groups.group(7))
    to_col = int(groups.group(8))
    wbs[to_wb - 1].write(to_sheet, to_row, to_col, val)
