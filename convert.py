import xlrd
import sys
import json
from parser import get_text_and_colors, advance

book = xlrd.open_workbook(sys.argv[1], formatting_info=True)
start_row = 11
start_col = 3
num_time_slots = 16
num_days = 30
num_merged_columns = 5
output = []
skipped_cx = [12, 20, 22, 18, 5]
i = 1

for row in get_text_and_colors(book, start_row, start_col, num_days ,num_time_slots + num_merged_columns):
    hour = 6
    minute = 30
    d = '2017-12-%.2d' % (i)
    schedule = []
    j = 0
    for time_slot in row:
        text, color, column = time_slot
        if column in skipped_cx:
            continue
        if color == '#FFCC9A':
            text = '可用'
        for k in range(0, 2 if j > 0 else 1):
            schedule.append({
                'time': '%.2d:%.2d' % (hour, minute),
                'status': text,
            }) 
            hour, minute = advance(hour, minute)
        j += 1
    output.append({'date': d, 'schedule': schedule, 'id': 488})
    i += 1

with open(sys.argv[2], 'w') as f:
    f.write(json.dumps(output, indent=4))
