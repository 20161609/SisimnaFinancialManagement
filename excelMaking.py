from openpyxl.styles import Side, Border

Alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

def line_bordors(ws, begin_spot, end_spot):
    side1 = Side(style="medium", color="000000")
    side2 = Side(style="dotted", color="000000")
    side3 = Side(style="double", color="000000")

    for rows in ws[begin_spot:end_spot]:
        for cell in rows:
            cell.border = Border(left=side1, right=side1, top=side1, bottom=side1)