#!/usr/bin/env python3

import csv
from enum import Enum

import openpyxl


class Version(Enum):
    V1 = 1
    V2 = 2


def set_count_to_one(row, version):
    if version == Version.V1:
        row[26].value = 1
    else:
        row[4].value = 1


def expand(input_file, output_file, version):
    workbook = openpyxl.load_workbook(input_file)
    ws = workbook.active

    with open(output_file, 'w') as csvfile:
        csv_writer = csv.writer(csvfile)

        first = True
        for row in ws.rows:
            if first:
                csv_writer.writerow([cell.value for cell in row])
                first = False
            else:
                count = get_count(row, version)
                set_count_to_one(row, version)
                for j in range(count):
                    csv_writer.writerow([cell.value for cell in row])


def get_count(row, version):
    if version == Version.V1:
        count = int(row[26].value)
    else:
        count = int(row[4].value)
    return count


if __name__ == "__main__":
    expand("Bruinvis alle waarnemingen 1991_2013.xlsx", "bruinvis_waarnemingen_1991_2013_expanded.csv", Version.V1)
    expand("export_BV_waarnemingen_14270_20171025.xlsx", "bruinvis_waarnemingen_2013_heden_expanded.csv", Version.V2)
