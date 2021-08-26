import re


def point_to_point(wbs, script):
    group_list = re.findall(r'(\d+)\[(\d+)\]\[(\d+),(\d+)\]]', script)
    operator_list = re.findall(r'[\\+-]', script)
    if len(group_list) == 2:
        tp1 = group_list[0]
        tp2 = group_list[1]
        val = wbs[int(tp1[0]) - 1].read(int(tp1[1]), int(tp1[2]), int(tp1[3]))
        wbs[int(tp2[0]) - 1].write(int(tp2[1]), int(tp2[2]), int(tp2[3]), val)
        return
    val = 0.0
    for i in range(0, len(group_list) - 1):
        tp = group_list[i]
        tmp = wbs[int(tp[0]) - 1].read(int(tp[1]), int(tp[2]), int(tp[3]))
        opr = '-' if i > 0 and operator_list[i - 1] == '-' else '+'
        val = val + float(tmp) if opr == '+' else val - float(tmp)
    tgt = group_list[len(group_list) - 1]
    wbs[int(tgt[0]) - 1].write(int(tgt[1]), int(tgt[2]), int(tgt[3]), val)


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


def col_to_col(wbs, script):
    groups = re.search(r'(\d+)\[(\d+)\]\[(\d+)\.\.\.,(\d+)\]->(\d+)\[(\d+)\]\[(\d+)\.\.\.,(\d+)\]', script)
    from_wb = int(groups.group(1))
    from_sheet = int(groups.group(2))
    from_start_row = int(groups.group(3))
    from_col = int(groups.group(4))
    values = wbs[from_wb - 1].read_col_values(from_sheet, from_col, from_start_row)
    to_wb = int(groups.group(5))
    to_sheet = int(groups.group(6))
    to_start_row = int(groups.group(7))
    to_col = int(groups.group(8))
    for i in range(len(values)):
        wbs[to_wb - 1].write(to_sheet, to_start_row + i, to_col, values[i])


def row_summary(wbs, script):
    groups = re.search(r'(\d+)\[(\d+)\]\[(\d+),(\d+)\.\.\.\]->(\d+)\[(\d+)\]\[(\d+),(\d+)\]', script)
    from_wb = int(groups.group(1))
    from_sheet = int(groups.group(2))
    from_row = int(groups.group(3))
    from_start_col = int(groups.group(4))
    values = wbs[from_wb - 1].read_row_values(from_sheet, from_row, from_start_col)
    val = 0.0
    for v in values:
        val = val + float(v)
    to_wb = int(groups.group(5))
    to_sheet = int(groups.group(6))
    to_row = int(groups.group(7))
    to_col = int(groups.group(8))
    wbs[to_wb - 1].write(to_sheet, to_row, to_col, val)


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


def map_col_operator(wbs, script):
    group_list = re.findall(r'(\d+)\[(\d+)\]\[(\d+)\.\.\.,(\d+)\]\[(\d+)\]', script)
    operator_list = re.findall(r'[\\+-]', script)
    map_data = {}
    for i in range(0, len(group_list) - 1):
        tp = group_list[i]
        keys = wbs[int(tp[0]) - 1].read_col_values(int(tp[1]), int(tp[4]), int(tp[2]))
        values = wbs[int(tp[0]) - 1].read_col_values(int(tp[1]), int(tp[3]), int(tp[2]))
        opr = '-' if i > 0 and operator_list[i - 1] == '-' else '+'
        for j in range(0, len(keys)):
            val = float(values[j]) if opr == '+' else -1 * float(values[j])
            if keys[j] in map_data:
                map_data[keys[j]] = map_data[keys[j]] + val
            else:
                map_data[keys[j]] = val
    target = group_list[len(group_list) - 1]
    start_row = int(target[2])
    keys = wbs[int(target[0]) - 1].read_col_values(int(target[1]), int(target[4]), start_row)
    for key in keys:
        wbs[int(target[0]) - 1].write(int(target[1]), start_row, int(target[3]), map_data[key])
        start_row = start_row + 1


def search_row_based_col_operator(wbs, script):
    group_list = re.findall(r'(\d+)\[(\d+)\]\[(\d+)\[([^\]]+)\],(\d+)\]', script)
    operator_list = re.findall(r'[\\+-]', script)
    if len(group_list) == 2:
        tp1 = group_list[0]
        tp2 = group_list[1]
        from_row = search_row_by_col_and_content(wbs[int(tp1[0]) - 1], int(tp1[1]), int(tp1[2]), tp1[3])
        to_row = search_row_by_col_and_content(wbs[int(tp2[0]) - 1], int(tp2[1]), int(tp2[2]), tp2[3])
        val = wbs[int(tp1[0]) - 1].read(int(tp1[1]), from_row, int(tp1[4]))
        wbs[int(tp2[0]) - 1].write(int(tp2[1]), to_row, int(tp1[4]), val)
        return
    val = 0.0
    for i in range(0, len(group_list) - 1):
        tp = group_list[i]
        from_row = search_row_by_col_and_content(wbs[int(tp[0]) - 1], int(tp[1]), int(tp[2]), tp[3])
        tmp = wbs[int(tp[0]) - 1].read(int(tp[1]), from_row, int(tp[4]))
        opr = '-' if i > 0 and operator_list[i - 1] == '-' else '+'
        val = val + float(tmp) if opr == '+' else val - float(tmp)
    tgt = group_list[len(group_list) - 1]
    to_row = search_row_by_col_and_content(wbs[int(tgt[0]) - 1], int(tgt[1]), int(tgt[2]), tgt[3])
    wbs[int(tgt[0]) - 1].write(int(tgt[1]), to_row, int(tgt[4]), val)


def search_row_by_col_and_content(wb, sheet, col, content):
    values = wb.read_col_values(sheet, col, 1)
    for i in range(0, len(values)):
        if str(values[i]) == str(content):
            return i + 1
    return -1


def handler(wbs, name, script):
    if name == "point_to_point":
        point_to_point(wbs, script)
    elif name == "row_to_row":
        row_to_row(wbs, script)
    elif name == "col_to_col":
        col_to_col(wbs, script)
    elif name == "row_summary":
        row_summary(wbs, script)
    elif name == "col_summary":
        col_summary(wbs, script)
    elif name == "map_col_operator":
        map_col_operator(wbs, script)
    elif name == "search_row_based_col_operator":
        search_row_based_col_operator(wbs, script)
    else:
        return
