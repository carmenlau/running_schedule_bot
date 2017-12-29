import xlrd
import sys
import json

book = xlrd.open_workbook(sys.argv[1], formatting_info=True)
sh = book.sheet_by_index(0)


def get_text_and_colors(book, start_row, start_column, num_rows, num_columns):
    output = []
    for rx in range(start_row, start_row + num_rows):
        row = []
        for cx in range(start_column, start_column + num_columns):
            xfx = sh.cell_xf_index(rowx=rx, colx=cx)
            xf = book.xf_list[xfx]
            rgb = book.colour_map[xf.background.pattern_colour_index]
            rgb_str = "#%.2X%.2X%.2X" % rgb
            text = sh.cell_value(rowx=rx, colx=cx).strip()
            row += [(text, rgb_str, cx)]
        output.append(row)
    return output

start_row = 11
start_col = 3
num_time_slots = 16
num_days = 30
num_merged_columns = 5
output = []
skipped_cx = [12, 20, 22, 18, 5]
i = 1

def advance(hour, minute):
    minute += 30
    if minute == 60:
        hour += 1
        minute = 0
    return hour, minute

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
                'cx': column
            }) 
            hour, minute = advance(hour, minute)
        j += 1
    output.append({'date': d, 'schedule': schedule})
    i += 1

with open(sys.argv[2], 'w') as f:
    f.write(json.dumps(output, indent=4))
