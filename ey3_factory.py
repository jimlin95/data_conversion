#!/usr/bin/env python
# coding=UTF-8
#
from openpyxl import Workbook
import glob

FOLDER_PREFIX_NAME = "EY"

FILENAME = "ey3_factory.xlsx"
SHEET_SFR = "SFR"
SHEET_SFR_2 = "SFR-2"
SHEET_CTF = "CTF"
AREA_TAG = ("UL-0.5", "UR-0.5", "UL-0.3", "UR-0.3", "Center",
            "LL-0.3", "LR-0.3", "LL-0.5", "LR-0.5")
SFR2_TAG = ("0.5f", "0.3f", "0", "0.3f", "0.5f")
EY_FOLDERS = []
EY_FOLDERS_NUM = 0

ROW_OFFSET = 1
MTF_START_INDEX1 = 2
MTF_1_LEN = 9
MTF_START_INDEX2 = (MTF_START_INDEX1 + MTF_1_LEN)
MTF_2_LEN = 9
SFR_START_INDEX = (MTF_START_INDEX2 + MTF_2_LEN)
CTF_INITIATE_TABLE = (
        ("cy/mm", "ROI"),
        (77.265, "ROI_0_MTF"),
        (57.433, "ROI_1_MTF"),
        (56.759, "ROI_2_MTF"),
        (74.585, "ROI_3_MTF"),
        (52.066, "ROI_9_MTF"),
        (56.428, "ROI_13_MTF"),
        (56.759, "ROI_14_MTF"),
        (75.679, "ROI_16_MTF"),
        (73.523, "ROI_17_MTF"),
        (71.486, "ROI_4_MTF"),
        (70.703, "ROI_5_MTF"),
        (57.263, "ROI_6_MTF"),
        (56.759, "ROI_7_MTF"),
        (51.523, "ROI_8_MTF"),
        (56.926, "ROI_10_MTF"),
        (56.593, "ROI_11_MTF"),
        (71.486, "ROI_12_MTF"),
        (70.317, "ROI_15_MTF"),
    )

TGT_FREQ = (
        (77.265),
        (57.433),
        (56.759),
        (74.585),
        (52.066),
        (56.428),
        (56.759),
        (75.679),
        (73.523),
        (71.486),
        (70.703),
        (57.263),
        (56.759),
        (51.523),
        (56.926),
        (56.593),
        (71.486),
        (70.317)
)

ROI_table = (
    "ROI_0_MTF",
    "ROI_1_MTF",
    "ROI_2_MTF",
    "ROI_3_MTF",
    "ROI_9_MTF",
    "ROI_13_MTF",
    "ROI_14_MTF",
    "ROI_16_MTF",
    "ROI_17_MTF",
    "ROI_4_MTF",
    "ROI_5_MTF",
    "ROI_6_MTF",
    "ROI_7_MTF",
    "ROI_8_MTF",
    "ROI_10_MTF",
    "ROI_11_MTF",
    "ROI_12_MTF",
    "ROI_15_MTF"
)


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
    return wb


def excel_save(wb, filename):
    # Save the file
    wb.save(filename)


def excel_creatsheet(wb, ws_title):
    ws = wb.create_sheet(title=ws_title)
    return ws


def create_working_sheets(wb):
    SHEETS_LIST = (SHEET_SFR, SHEET_SFR_2, SHEET_CTF)
    worksheets = []
    for sheet in SHEETS_LIST:
        worksheets.append(excel_creatsheet(wb, sheet))
    # print(worksheets)


def find_directoies_with_substring(ey):
    return glob.glob(ey)


def mtf_dealwith(ws, ey_folder, ey_folder_idx):
    ws.cell(column=1, row=ey_folder_idx + ROW_OFFSET, value=ey_folder)
    filefullpath = ey_folder + "/mtf/" + ey_folder + "-H-MTFOUT.txt"
    with open(filefullpath, "r") as f:
        for mtf_idx, line in enumerate(f, MTF_START_INDEX1):
            roi, mtf = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=mtf_idx, row=1, value=roi)
            ws.cell(column=mtf_idx , row=ey_folder_idx + ROW_OFFSET,
                    value=float(mtf))
    filefullpath = ey_folder + "/mtf/" + ey_folder + "-V-MTFOUT.txt"
    with open(filefullpath, "r") as f:
        for mtf_idx, line in enumerate(f, MTF_START_INDEX2):
            roi, mtf = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=mtf_idx, row=1, value=roi)
            ws.cell(column=mtf_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(mtf))


def sfr_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + "/sfr/" + "SFROUT_shopfloor.txt"
    with open(filefullpath, "r") as f:
        for sfr_idx, line in enumerate(f, SFR_START_INDEX):
            roi, sfr = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=sfr_idx, row=1, value=roi)
            ws.cell(column=sfr_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(sfr))


def tp1_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + "/tp1/" + "tp1-black-out.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, sfr = line.split('=')


def tp2_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + "/tp2/" + "chart-out.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, sfr = line.split('=')


def color_fidelity_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + "/" + "color_fidelity.txt"
    with open(filefullpath, "r") as f:
        for line in f:
            roi, sfr = line.split('=')


if __name__ == '__main__':
    wb = excel_create()
    ws = wb.active
    EY_FOLDERS = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*")
    for ey_folder_idx, ey_folder in enumerate(EY_FOLDERS, 1):
        mtf_dealwith(ws, ey_folder, ey_folder_idx)
        sfr_dealwith(ws, ey_folder, ey_folder_idx)
    #    tp1_dealwith(ws, ey_folder, ey_folder_idx)
    #    tp2_dealwith(ws, ey_folder, ey_folder_idx)
    #    color_fidelity_dealwith(ws, ey_folder, ey_folder_idx)
    excel_save(wb, FILENAME)
