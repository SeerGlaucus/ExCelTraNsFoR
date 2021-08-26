import re


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
