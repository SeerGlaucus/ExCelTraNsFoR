import re


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
