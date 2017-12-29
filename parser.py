def get_text_and_colors(book, start_row, start_column, num_rows, num_columns):
    sh = book.sheet_by_index(0)
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

def advance(hour, minute):
    minute += 30
    if minute == 60:
        hour += 1
        minute = 0
    return hour, minute


