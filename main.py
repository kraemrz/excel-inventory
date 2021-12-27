#!/usr/bin/env python3
import openpyxl
import argparse
import os


def main():
    parser = argparse.ArgumentParser(
        description="Add filename of input and output files")
    parser.add_argument("-i", "--infile", type=str, metavar="",
                        required=True, help="Name of input file")
    parser.add_argument("-o", "--outfile", type=str, metavar="",
                        required=True, help="Name of output file")
    args = parser.parse_args()

    cwd = os.getcwd() + "/"
    input_file = (cwd + args.infile)
    parts = parts_in_car(cwd + args.outfile)

    input_workbook = openpyxl.load_workbook(input_file)
    sheet = input_workbook.active
    sheet2 = input_workbook.create_sheet("FINNS EJ")

    for part in parts:
        print(part)

    for i in range(7, sheet.max_row + 1):
        if sheet.cell(row=i, column=1).value in parts.keys():
            sheet['E' + str(i)].value = parts[sheet.cell(row=i,
                                                         column=1).value]
            del parts[sheet.cell(row=i, column=1).value]
        else:
            continue

    sheet2['A1'] = "artnr"
    sheet2['A2'] = "antal"
    for row, (artnr, qty) in enumerate(parts.items(), start=2):
        sheet2[f'A{row}'] = artnr
        sheet2[f'A{row}'] = qty

    input_workbook.save('updated.xlsx')


def parts_in_car(inv_file):
    inv_workbook = openpyxl.load_workbook(inv_file)
    inv_sheet = inv_workbook.active

    inventoried_parts = {}
    for i in range(1, inv_sheet.max_row + 1):
        inventoried_parts.update({
            inv_sheet.cell(row=i, column=1).value:
            inv_sheet.cell(row=i, column=2).value})

    return inventoried_parts


def test():
    pass


if __name__ == '__main__':
    main()
