#!/usr/bin/env python3

import csv
from enum import Enum

import openpyxl
import utm


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
        row_count = 1
        for row in ws.rows:
            if first:
                if version == Version.V2:
                    convert_lat_lon(row, True)
                # Prepend the ID header and copy the header row to the csv
                id_header = ["ID"]
                header_row = [cell.value for cell in row]
                id_header.extend(header_row)
                csv_writer.writerow(id_header)
                first = False
            else:
                # Copy the row to csv. The row is copied as many times as the number of harbour porpoises found.
                count = get_count(row, version)
                set_count_to_one(row, version)
                if version == Version.V2:
                    convert_lat_lon(row)
                for j in range(count):
                    # Prepend the row count
                    row_cell = [str(row_count)]
                    row_list = [cell.value for cell in row]
                    row_cell.extend(row_list)
                    csv_writer.writerow(row_cell)
                    row_count += 1


# Convert latitude and longitude columns to UTM format.
def convert_lat_lon(row, first=False):
    if first:
        row[5].value = "XUTM"
        row[6].value = "YUTM"
    else:
        lat = row[5].value
        lon = row[6].value
        utm_value = utm.from_latlon(lat, lon)
        row[5].value = int(utm_value[0])
        row[6].value = int(utm_value[1])


# Sets the harbour porpoise count to 1 in the row
def set_count_to_one(row, version):
    if version == Version.V1:
        row[25].value = 1
    else:
        row[4].value = 1


# Retrieves the harbour porpoise count from the row
def get_count(row, version):
    if version == Version.V1:
        count = int(row[25].value)
    else:
        count = int(row[4].value)
    return count


if __name__ == "__main__":
    expand("data/in/Bruinvis alle waarnemingen 1991_2013.xlsx", "data/out/bruinvis_waarnemingen_1991_2013_expanded.csv",
           Version.V1)
    expand("data/in/export_BV_waarnemingen_14270_20171025.xlsx",
           "data/out/bruinvis_waarnemingen_2013_heden_expanded.csv",
           Version.V2)
