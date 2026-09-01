#!/usr/bin/env python3
import os, sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from openpyxl import Workbook
from build_cover_overview_types import build_cover, build_overview, build_types
from build_step1_2cc import build_2cc
from build_step2_5cc import build_3cc, build_4cc, build_5cc
from build_spectrum_coord import build_spectrum_coord

OUT = "/workspace/Carrier_Aggregation_Deployment.xlsx"


def main():
    wb = Workbook()
    # remove default sheet after we create ours
    default = wb.active
    default.title = "_tmp"

    print("cover...")
    build_cover(wb)
    print("overview...")
    build_overview(wb)
    print("types...")
    build_types(wb)
    print("2cc...")
    build_2cc(wb)
    print("3cc...")
    build_3cc(wb)
    print("4cc...")
    build_4cc(wb)
    print("5cc...")
    build_5cc(wb)
    print("spectrum coordination...")
    build_spectrum_coord(wb)

    if "_tmp" in wb.sheetnames:
        del wb["_tmp"]

    # consistent tab colors
    colors = ["1F4E79", "2E75B6", "0D7377", "C00000", "C65911", "548235", "7030A0", "0D7377"]
    for i, ws in enumerate(wb.worksheets):
        ws.sheet_properties.tabColor = colors[i % len(colors)]
        ws.sheet_view.showGridLines = False

    print("saving", OUT)
    wb.save(OUT)
    print("ok", os.path.getsize(OUT))


if __name__ == "__main__":
    main()
