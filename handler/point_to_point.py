import re


def point_to_point(wbs, script):
    group_list = re.findall(r'(\d+)\[(\d+)\]\[(\d+),(\d+)\]', script)
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
