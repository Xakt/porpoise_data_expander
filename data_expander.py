#!/usr/bin/env python3

import csv
from enum import Enum

import openpyxl


# Enum to differentiate between the two file formats
class Version(Enum):
    V1 = 1
    V2 = 2


# Reads an xlsx file with harbour porpoise observations and copies it to a csv. If multiple harbour porpoises are
# spotted the row is duplicated so that each row represents one observation.
def expand(input_file, output_file, version):
    workbook = openpyxl.load_workbook(input_file)
    ws = workbook.active

    with open(output_file, 'w') as csv_file:
        csv_writer = csv.writer(csv_file)

        first = True
        for row in ws.rows:
            if first:
                # Copy the header row to the csv
                csv_writer.writerow([cell.value for cell in row])
                first = False
            else:
                # Copy the row to csv. The row is copied as many times as the number of harbour porpoises found.
                count = get_count(row, version)
                set_count_to_one(row, version)
                for j in range(count):
                    csv_writer.writerow([cell.value for cell in row])


# Sets the harbour porpoise count to 1 in the row
def set_count_to_one(row, version):
    if version == Version.V1:
        row[26].value = 1
    else:
        row[4].value = 1


# Retrieves the harbour porpoise count from the row
def get_count(row, version):
    if version == Version.V1:
        count = int(row[26].value)
    else:
        count = int(row[4].value)
    return count


if __name__ == "__main__":
    expand("Bruinvis alle waarnemingen 1991_2013.xlsx", "bruinvis_waarnemingen_1991_2013_expanded.csv", Version.V1)
    expand("export_BV_waarnemingen_14270_20171025.xlsx", "bruinvis_waarnemingen_2013_heden_expanded.csv", Version.V2)
