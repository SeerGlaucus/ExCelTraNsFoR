import re

import handlers
import match
import workbook


def call(excel_path, script_path):
    wbs = workbook.load(excel_path)
    file = open(script_path, 'r', encoding='UTF-8')
    line = file.readline()
    while line:
        line = re.sub(re.compile(r'\s+'), '', line)
        if not line.startswith("#"):
            handlers.handler(wbs, match.match(line), line)
        line = file.readline()
    for wb in wbs:
        if wb.changed is True:
            wb.save()
