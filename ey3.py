#
import openpyxl
import glob

FOLDER_PREFIX_NAME = "EY"


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
    wb = openpyxl.Workbook()
    ws = wb.active
    # remove the sheet named "Sheet"
    wb.remove_sheet(ws)
    return wb


def excel_save(wb, filename):
    # Save the file
    wb.save(filename)


def roi_mtp_dealwith(ws, folder):
    ws['A1'] = "MTF"
    ws['B1'] = "Value"
    filefullpath = folder + "mtf/" + folder[0:6] + "-H-MTFOUT.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, mtf = line.split('=')
            index = find_between(roi, "ROI_", "_MTF")
            roi_location = 'A' + str(int(index)+2)
            ws[roi_location] = index
            mtf_location = 'B' + str(int(index)+2)
            ws[mtf_location] = (float(mtf))
    filefullpath = folder + "mtf/" + folder[0:6] + "-V-MTFOUT.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, mtf = line.split('=')
            index = find_between(roi, "ROI_", "_MTF")
            roi_location = 'A' + str(int(index)+2)
            ws[roi_location] = index
            mtf_location = 'B' + str(int(index)+2)
            ws[mtf_location] = (float(mtf))


def sfr_dealwith(ws, folder):

    ws['D1'] = "SFR"
    ws['E1'] = "Value"
    filefullpath = folder + "/sfr/" + "SFROUT_shopfloor.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, sfr = line.split('=')
            index = find_between(roi, "ROI", "_SFR_RESULT")
            roi_location = 'D' + str(int(index)+2)
            ws[roi_location] = index
            mtf_location = 'E' + str(int(index)+2)
            ws[mtf_location] = float(sfr)


def excel_creatsheet(wb, ws_title):
    ws = wb.create_sheet(title=ws_title)
    return ws


def excel_mtf_barchart(ws):
    chart1 = openpyxl.chart.BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "MTF Chart"
    chart1.y_axis.title = 'MTF'
    chart1.x_axis.title = 'ROI'
# Select all data include title
    data = openpyxl.chart.Reference(ws, min_col=2, min_row=1,
                                    max_row=19, max_col=2)
# Select data only
    cats = openpyxl.chart.Reference(ws, min_col=1, min_row=2, max_row=18)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.x_axis.scaling.min = 0
    chart1.x_axis.scaling.max = 18
    chart1.y_axis.scaling.min = 0
    chart1.y_axis.scaling.max = 1
    ws.add_chart(chart1, "G1")


def excel_sfr_barchart(ws):
    chart1 = openpyxl.chart.BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "SFR Chart"
    chart1.y_axis.title = 'SFR'
    chart1.x_axis.title = 'ROI'
# Select all data include title
    data = openpyxl.chart.Reference(ws, min_col=5, min_row=1,
                                    max_row=37, max_col=5)
# Select data only
    cats = openpyxl.chart.Reference(ws, min_col=4, min_row=2, max_row=37)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.x_axis.scaling.min = 0
    chart1.x_axis.scaling.max = 37
    chart1.y_axis.scaling.min = 0
    chart1.y_axis.scaling.max = 1
    ws.add_chart(chart1, "G21")


def find_directoies_with_substring(ey):
    return glob.glob(ey)


if __name__ == '__main__':

    ey_folders = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*/")
    wb = excel_create()
    for ey in ey_folders:
        ws_sn = excel_creatsheet(wb, ey[:-1])
        roi_mtp_dealwith(ws_sn, ey)
        sfr_dealwith(ws_sn, ey)
        excel_mtf_barchart(ws_sn)
        excel_sfr_barchart(ws_sn)
    excel_save(wb, "ey3.xlsx")
