#!/usr/bin/env python
# coding=UTF-8
#
from openpyxl import Workbook
from openpyxl.styles import Side, Font, PatternFill, Border, Alignment
from openpyxl.chart import Reference, BarChart, LineChart
# import openpyxl
import glob

FOLDER_PREFIX_NAME = "EY"
SHEET_SFR = "SFR"
SHEET_SFR_2 = "SFR-2"
SHEET_CTF = "CTF"


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


def set_alignment(ws, cell_range):
    rows = list(ws.iter_rows(cell_range))
    align_center = Alignment(horizontal='center')
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


def mtp_cell_format(ws):
    set_allborder(ws, "A1:B19")
    set_font(ws, "A1:B1")
    set_alignment(ws, "A1:B19")
    color_string = '8b8989'  # color hex string
    set_background_color(ws, "A1:B1", color_string)
    color_string = 'cdc9c9'  # light gray
    set_background_color(ws, "A2:B19", color_string)


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


def sfr_cell_format(ws):
    color_string = 'cdb79e'  # light gray
    set_font(ws, "D1:E1")
    set_allborder(ws, "D1:E37")
    set_alignment(ws, "D1:E37")
    set_background_color(ws, "D1:E1", color_string)
    color_string = 'ffdab9'  # light gray
    set_background_color(ws, "D2:E37", color_string)


def sfr_dealwith(ws):

    AREA_TAG = ("UL-0.5", "UR-0.5", "UL-0.3", "UR-0.3", "Center",
                "LL-0.3", "LR-0.3", "LL-0.5", "LR-0.5")
    ws['{0}'.format('A')+'1'] = ""
    ey_folders = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*/")
    for folder in ey_folders:
        ws['B1'] = folder[:-1]
        filefullpath = folder + "/sfr/" + "SFROUT_shopfloor.txt"
        with open(filefullpath, "r") as f:
            for line in f:
                roi, sfr = line.split('=')
                index = find_between(roi, "ROI", "_SFR_RESULT")
                roi_location = 'A' + str(int(index)+2)
                ws[roi_location] = roi
                mtf_location = 'B' + str(int(index)+2)
                ws[mtf_location] = float(sfr)

                for index, i in enumerate(range(5, 38, 4)):
                    ul = 'C' + str(i)
                    ws[ul] = AREA_TAG[index]
                    uavg = 'D' + str(i)
                    ws[uavg] = "=SUM(B{0}:B{1})/4".format(i-3, i)


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


if __name__ == '__main__':
    ey_folders = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*/")
    wb = excel_create()
    create_working_sheets(wb)
    active_sheet = wb[SHEET_SFR]
    sfr_dealwith(active_sheet)
    #  roi_mtp_dealwith(active_sheet)
    """
    for ey in ey_folders:
        mtp_cell_format(ws_sn)
        roi_mtp_dealwith(ws_sn, ey)
        mtp_linechart(ws_sn)
        sfr_cell_format(ws_sn)
        sfr_dealwith(ws_sn, ey)
        excel_mtf_barchart(ws_sn)
        excel_sfr_barchart(ws_sn)
    """
    excel_save(wb, "ey3.xlsx")
