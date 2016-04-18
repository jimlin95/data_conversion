#!/usr/bin/env python
# coding=UTF-8
#
from openpyxl import Workbook
from openpyxl.styles import (
    Side,
    Font,
    PatternFill,
    Border,
    Alignment,
    fills,
    colors
    )
from openpyxl.chart import Reference, BarChart, LineChart
# import openpyxl
import glob

FOLDER_PREFIX_NAME = "EY"
SHEET_SFR = "SFR"
SHEET_SFR_2 = "SFR-2"
SHEET_CTF = "CTF"
AREA_TAG = ("UL-0.5", "UR-0.5", "UL-0.3", "UR-0.3", "Center",
            "LL-0.3", "LR-0.3", "LL-0.5", "LR-0.5")
SFR2_TAG = ("0.5f", "0.3f", "0", "0.3f", "0.5f")
EY_FOLDERS = []


def find_between(s, first, last):
    try:
        start = s.index(first) + len(first)
        end = s.index(last, start)
        return s[start:end]
    except ValueError:
        return ""


def find_between_r(s, first, last):
    try:
        start = s.rindex(first) + len(first)
        end = s.rindex(last, start)
        return s[start:end]
    except ValueError:
        return ""


def excel_create():
    wb = Workbook()
    ws = wb.active
    # remove the sheet named "Sheet"
    wb.remove_sheet(ws)
    return wb


def excel_save(wb, filename):
    # Save the file
    wb.save(filename)


def set_allborder(ws, cell_range):
    rows = list(ws.iter_rows(cell_range))
    side = Side(border_style='thin', color="FF000000")
    for pos_y, cells in enumerate(rows):
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            border.left = side
            border.right = side
            border.top = side
            border.bottom = side
            cell.border = border


def set_border(ws, cell_range):
    rows = list(ws.iter_rows(cell_range))
    side = Side(border_style='thin', color="FF000000")
    max_y = len(rows) - 1  # index of the last row
    for pos_y, cells in enumerate(rows):
        max_x = len(cells) - 1  # index of the last cell
        for pos_x, cell in enumerate(cells):
            border = Border(
                left=cell.border.left,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )
            if pos_x == 0:
                border.left = side
            if pos_x == max_x:
                border.right = side
            if pos_y == 0:
                border.top = side
            if pos_y == max_y:
                border.bottom = side
            # set new border only if it's one of the edge cells
            if pos_x == 0 or pos_x == max_x or pos_y == 0 or pos_y == max_y:
                cell.border = border


def set_font(ws, cell_range):
    rows = list(ws.iter_rows(cell_range))
    font = Font(name='Calibri', size=11, bold=False, italic=False,
                vertAlign=None, underline='none', strike=False,
                color='FFFFFFFF')

    for pos_y, cells in enumerate(rows):
        for pos_x, cell in enumerate(cells):
            cell.font = font


def set_alignment(ws, cell_range, position='center'):
    rows = list(ws.iter_rows(cell_range))
    align_center = Alignment(horizontal=position)
    for pos_y, cells in enumerate(rows):
        for pos_x, cell in enumerate(cells):
            cell.alignment = align_center


def set_background_color(ws, cell_range, color_string):
    rows = list(ws.iter_rows(cell_range))
    backgroundcolor = PatternFill(
        fill_type="solid",
        start_color='FF' + color_string,
        end_color='FF' + color_string)
    for pos_y, cells in enumerate(rows):
        for pos_x, cell in enumerate(cells):
            cell.fill = backgroundcolor


def roi_mtp_dealwith(ws):
    ws['A1'] = "MTF"
    ey_folders = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*/")
#    print(ey_folders)
    for folder in ey_folders:
        ws['B1'] = folder[:-1]
        filefullpath = folder + "mtf/" + folder[0:6] + "-H-MTFOUT.txt"
        with open(filefullpath, "r") as f:
            for line in f:
                roi, mtf = line.split('=')
                index = find_between(roi, "ROI_", "_MTF")
                roi_location = 'A' + str(int(index)+2)
                ws[roi_location] = roi
                mtf_location = 'B' + str(int(index)+2)
                ws[mtf_location] = (float(mtf))
        filefullpath = folder + "mtf/" + folder[0:6] + "-V-MTFOUT.txt"
        with open(filefullpath, "r") as f:
            for line in f:
                roi, mtf = line.split('=')
                index = find_between(roi, "ROI_", "_MTF")
                roi_location = 'A' + str(int(index)+2)
                ws[roi_location] = roi
                mtf_location = 'B' + str(int(index)+2)
                ws[mtf_location] = (float(mtf))
        for i in range(5, 37, 4):
            ul = 'C' + str(i)
            ws[ul] = "UL-0.5"


def sfr_dealwith(ws, eyfile_index):
    adjust_index = 3 * (eyfile_index) + 2
    ws.cell(column=adjust_index, row=1, value=EY_FOLDERS[eyfile_index])
    ws.cell(column=adjust_index+2, row=1, value=EY_FOLDERS[eyfile_index])
    filefullpath = EY_FOLDERS[eyfile_index] + "/sfr/" + "SFROUT_shopfloor.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, sfr = line.split('=')
            index = find_between(roi, "ROI", "_SFR_RESULT")
            ws.cell(column=1, row=int(index)+2, value=roi)
            ws.cell(column=adjust_index, row=int(index)+2, value=float(sfr))
            for ul_index, i in enumerate(range(5, 38, 4)):
                ws.cell(column=3*(eyfile_index+1), row=i,
                        value=AREA_TAG[ul_index])
                # write SUM(4 values)/4 in UL/UR/LL/LR cell
                pos = chr(ord('B') + eyfile_index*3)
                ws.cell(column=(3*(eyfile_index+1))+1, row=i,
                        value="=SUM({0}{1}:{2}{3})/4".format(pos, i-3, pos, i))
                # paint in blue and green
                color_string = 'C6D9F1'  # blue color hex string
                datarow = chr(ord('C') + 3 * (eyfile_index))
                set_background_color(ws, "{0}2:{0}9".format(datarow),
                                     color_string)
                set_background_color(ws, "{0}30:{0}37".format(datarow),
                                     color_string)
                color_string = 'C3D69B'  # green color
                set_background_color(ws, "{0}10:{0}17".format(datarow),
                                     color_string)
                set_background_color(ws, "{0}22:{0}29".format(datarow),
                                     color_string)


def sfr_linechart(ws, data_index):
    chart1 = LineChart()
    chart1.title = "Line Chart"
    chart1.style = 2
    chart1.y_axis.title = ''
    chart1.x_axis.title = ''

    for idx, i in enumerate(range(2, 3*data_index, 3)):
        # Select all data include title
        data = Reference(ws, min_col=i, min_row=1, max_row=37, max_col=i)
        chart1.add_data(data, titles_from_data=True)
        s1 = chart1.series[idx]
        s1.marker.symbol = "triangle"
        s1.marker.graphicalProperties.solidFill = "FF0000"  # Marker filling
        # Marker outline
        s1.marker.graphicalProperties.line.solidFill = "FF0000"
        s1.graphicalProperties.line.noFill = False
    ws.add_chart(chart1, "H5")


def sfr_handle(ws):
    "Deal with all SFR files"
    global EY_FOLDERS
    ws['{0}'.format('A')+'1'] = ""
    set_alignment(ws, "A1:AZ1")
    EY_FOLDERS = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*")
    for index, folder in enumerate(EY_FOLDERS):
        sfr_dealwith(ws, index)
    sfr_linechart(ws, len(EY_FOLDERS))


def excel_creatsheet(wb, ws_title):
    ws = wb.create_sheet(title=ws_title)
    return ws


def excel_mtf_barchart(ws):
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "MTF Chart"
    chart1.y_axis.title = 'MTF'
    chart1.x_axis.title = 'ROI'
# Select all data include title
    data = Reference(ws, min_col=2, min_row=1, max_row=19, max_col=2)
# Select data only
    cats = Reference(ws, min_col=1, min_row=2, max_row=18)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.x_axis.scaling.min = 0
    chart1.x_axis.scaling.max = 18
    chart1.y_axis.scaling.min = 0
    chart1.y_axis.scaling.max = 1
    ws.add_chart(chart1, "G1")


def excel_sfr_barchart(ws):
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 12
    chart1.title = "SFR Chart"
    chart1.y_axis.title = 'SFR'
    chart1.x_axis.title = 'ROI'
# Select all data include title
    data = Reference(ws, min_col=5, min_row=1, max_row=37, max_col=5)
# Select data only
    cats = Reference(ws, min_col=4, min_row=2, max_row=37)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.x_axis.scaling.min = 0
    chart1.x_axis.scaling.max = 37
    chart1.y_axis.scaling.min = 0
    chart1.y_axis.scaling.max = 1
    ws.add_chart(chart1, "G21")


def mtp_linechart(ws):
    chart1 = LineChart()
    chart1.title = "Line Chart"
    chart1.style = 9
    chart1.y_axis.title = 'Size'
    chart1.x_axis.title = 'Test Number'
# Select all data include title
    data = Reference(ws, min_col=2, min_row=1, max_row=19, max_col=2)
# Select data only
    cats = Reference(ws, min_col=1, min_row=2, max_row=18)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    # Style the lines
    s1 = chart1.series[0]
    s1.marker.symbol = "triangle"
    s1.marker.graphicalProperties.solidFill = "FF0000"  # Marker filling
    s1.marker.graphicalProperties.line.solidFill = "FF0000"  # Marker outline
    s1.graphicalProperties.line.noFill = False
    ws.add_chart(chart1, "A10")


def create_working_sheets(wb):
    SHEETS_LIST = (SHEET_SFR, SHEET_SFR_2, SHEET_CTF)
    worksheets = []
    for sheet in SHEETS_LIST:
        worksheets.append(excel_creatsheet(wb, sheet))
    # print(worksheets)


def find_directoies_with_substring(ey):
    return glob.glob(ey)


def sfr_sheet_initiate(ws):
    "A function to initiate SFR sheet"
    # paste "ROI0_SFR_RESULT" to A2:A37
    for index, row_index in enumerate(range(2, 38)):
        ws.cell(column=1, row=row_index,
                value="ROI{0}_SFR_RESULT".format(index))
        ws.column_dimensions['A'].width = 18


def sfr2_sheet_initiate(ws):
    "A function to initiate SFR-2 sheet"
    for index, i in enumerate(range(2, 11)):
        # format column B2:B10
        ws.cell(column=2, row=i).style.fill.fill_type = fills.FILL_SOLID
        ws.cell(column=2, row=i).style.fill.start_color = colors.DARKRED
        ws.cell(column=2, row=i).value = AREA_TAG[index]
    for index, i in enumerate(range(13, 22)):
        ul = 'C' + str(i)
        ws[ul] = AREA_TAG[index]
    for f_index, f in enumerate(range(6, 11)):
        ws.cell(column=f, row=13, value=SFR2_TAG[f_index])
    for f_index, f in enumerate(range(14, 19)):
        ws.cell(column=5, row=f, value=SFR2_TAG[f_index])
    ws['F14'] = "=D13"
    ws['J14'] = "=D14"
    ws['G15'] = "=D15"
    ws['I15'] = "=D16"
    ws['H16'] = "=D17"
    ws['G17'] = "=D18"
    ws['I17'] = "=D19"
    ws['F18'] = "=D20"
    ws['J18'] = "=D21"
    set_alignment(ws, "E13:J18")
    set_allborder(ws, "E13:J18")
    color_string = 'C6D9F1'  # blue color hex string
    set_background_color(ws, "B2:B3", color_string)
    set_background_color(ws, "B9:B10", color_string)
    set_background_color(ws, "C13:C14", color_string)
    set_background_color(ws, "C20:C21", color_string)
    color_string = 'C3D69B'  # green color
    set_background_color(ws, "B4:B5", color_string)
    set_background_color(ws, "B7:B8", color_string)
    set_background_color(ws, "C15:C16", color_string)
    set_background_color(ws, "C18:C19", color_string)

    color_string = 'DBEEF4'  # light blue color hex string
    set_background_color(ws, "F14:F18", color_string)
    set_background_color(ws, "G14:I14", color_string)
    set_background_color(ws, "G18:I18", color_string)
    set_background_color(ws, "J14:J18", color_string)

    color_string = 'C3D69B'  # light green color hex string
    set_background_color(ws, "G15:I17", color_string)
    color_string = 'D99694'  # light red color hex string
    set_background_color(ws, "H16:H16", color_string)
    set_background_color(ws, "C17:C17", color_string)


if __name__ == '__main__':
    wb = excel_create()
    create_working_sheets(wb)
    sfr_sheet_initiate(wb[SHEET_SFR])
    sfr2_sheet_initiate(wb[SHEET_SFR_2])
    sfr_handle(wb[SHEET_SFR])
    excel_save(wb, "ey3.xlsx")
