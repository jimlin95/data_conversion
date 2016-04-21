#!/usr/bin/env python
# coding=UTF-8
from openpyxl import Workbook
import glob
#
# Copyright (C) 2016 Quanta Computer Inc.
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software
# distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and
# limitations under the License.

# increment this whenever we make important changes to this script
VERSION = (0, 2)


# Maintainer <jim_lin@quantatw.com>


FOLDER_PREFIX_NAME = "*EY3_"

FILENAME = "ey3_factory.xlsx"
MTF_PATH = ""
SFR_PATH = ""
TP1_PATH = ""
TP2_PATH = ""
ROW_OFFSET = 1
MTF_START_POS1 = 2
MTF_1_LEN = 9
MTF_START_POS2 = (MTF_START_POS1 + MTF_1_LEN)
MTF_2_LEN = 9
SFR_START_POS = (MTF_START_POS2 + MTF_2_LEN)
SFR_LEN = 36
TP1_START_POS1 = (SFR_START_POS + SFR_LEN)
TP1_1_LEN = 3
TP1_START_POS2 = (TP1_START_POS1 + TP1_1_LEN)
TP1_2_LEN = 42
TP2_START_POS = (TP1_START_POS2 + TP1_2_LEN)
TP2_LEN = 13
COLOR_FIDELITY_START_POS = (TP2_START_POS + TP2_LEN)
COLOR_FIDELITY_LEN = 12


def excel_create():
    wb = Workbook()
    return wb


def excel_save(wb, filename):
    # Save the file
    wb.save(filename)


def excel_creatsheet(wb, ws_title):
    ws = wb.create_sheet(title=ws_title)
    return ws


def find_directoies_with_substring(ey):
    return glob.glob(ey)


def mtf_dealwith(ws, ey_folder, ey_folder_idx):
    ws.cell(column=1, row=ey_folder_idx + ROW_OFFSET, value=ey_folder[:-1])
    filefullpath = ey_folder + MTF_PATH + ey_folder[:-1] + "-H-MTFOUT.txt"
    with open(filefullpath, "r") as f:
        for mtf_idx, line in enumerate(f, MTF_START_POS1):
            roi, mtf = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=mtf_idx, row=1, value=roi)
            ws.cell(column=mtf_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(mtf))
    filefullpath = ey_folder + MTF_PATH + ey_folder[:-1] + "-V-MTFOUT.txt"
    with open(filefullpath, "r") as f:
        for mtf_idx, line in enumerate(f, MTF_START_POS2):
            roi, mtf = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=mtf_idx, row=1, value=roi)
            ws.cell(column=mtf_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(mtf))


def sfr_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + SFR_PATH + "SFROUT_shopfloor.txt"
    with open(filefullpath, "r") as f:
        for sfr_idx, line in enumerate(f, SFR_START_POS):
            roi, sfr = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=sfr_idx, row=1, value=roi)
            ws.cell(column=sfr_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(sfr))


def tp1_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + TP1_PATH + "tp1-black-out.txt"
    with open(filefullpath, "r") as f:
        for tp1_idx, line in enumerate(f, TP1_START_POS1):
            roi, black = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=tp1_idx, row=1, value=roi)
            ws.cell(column=tp1_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(black))
    filefullpath = ey_folder + TP1_PATH + "tp1-white-out.txt"
    with open(filefullpath, "r") as f:
        for tp1_idx, line in enumerate(f, TP1_START_POS2):
            roi, white = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=tp1_idx, row=1, value=roi)
            ws.cell(column=tp1_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(white))


def tp2_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + TP2_PATH + "chart-out.txt"
    with open(filefullpath, "r") as f:
        for tp2_idx, line in enumerate(f, TP2_START_POS):
            roi, tp2 = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=tp2_idx, row=1, value=roi)
            ws.cell(column=tp2_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(tp2))


def color_fidelity_dealwith(ws, ey_folder, ey_folder_idx):
    filefullpath = ey_folder + "color_fidelity.txt"
    with open(filefullpath, "r") as f:
        for cfd_idx, line in enumerate(f, COLOR_FIDELITY_START_POS):
            roi, cfd = line.split('=')
            if ey_folder_idx == 1:
                ws.cell(column=cfd_idx, row=1, value=roi)
            ws.cell(column=cfd_idx, row=ey_folder_idx + ROW_OFFSET,
                    value=float(cfd))


if __name__ == '__main__':
    wb = excel_create()
    ws = wb.active
    EY_FOLDERS = find_directoies_with_substring(FOLDER_PREFIX_NAME + "*/")
    for ey_folder_idx, ey_folder in enumerate(EY_FOLDERS, 1):
        mtf_dealwith(ws, ey_folder, ey_folder_idx)
        sfr_dealwith(ws, ey_folder, ey_folder_idx)
        tp1_dealwith(ws, ey_folder, ey_folder_idx)
        tp2_dealwith(ws, ey_folder, ey_folder_idx)
        color_fidelity_dealwith(ws, ey_folder, ey_folder_idx)
    excel_save(wb, FILENAME)
