import xlrd
import sys
import json
from parser import get_text_and_colors, advance

book = xlrd.open_workbook(sys.argv[1], formatting_info=True)
start_row = 7
start_col = 4
num_time_slots = 16
num_days = 30
output = []
skipped_cx = []
i = 1

for row in get_text_and_colors(book, start_row, start_col, num_days ,num_time_slots):
    hour = 6
    minute = 0
    d = '2017-12-%.2d' % (i)
    schedule = []
    j = 0
    for time_slot in row:
        text, color, column = time_slot
        if column in skipped_cx:
            continue
        if color == '#FF0000':
            text = '预约'
        else:
            text = '可用'
        for k in range(0, 3 if j == 15 else 2):
            schedule.append({
                'time': '%.2d:%.2d' % (hour, minute),
                'status': text,
            }) 
            hour, minute = advance(hour, minute)
        j += 1
    output.append({'date': d, 'schedule': schedule, 'id': 63})
    i += 1

with open(sys.argv[2], 'w') as f:
    f.write(json.dumps(output, indent=4))
