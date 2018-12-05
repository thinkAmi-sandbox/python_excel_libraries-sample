from datetime import timedelta, datetime

import openpyxl

BEFORE = "zero_time.xlsx"
AFTER = "zero_time_after.xlsx"
FIXED = "zero_time_fix.xlsx"
ISO = "zero_time_iso.xlsx"

EXCEL_ZERO_DATE = datetime(1899, 12, 30, 0, 0, 0)


def load_and_save():
    wb = openpyxl.load_workbook(BEFORE)
    wb.save(AFTER)


def print_before_and_after(before, after):
    print('--- before ---')
    print_a1_a2(before)

    print('--- after ---')
    print_a1_a2(after)


def print_a1_a2(path):
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    a1_value = sheet['A1'].value
    a2_value = sheet['A2'].value
    print(f'A1: {a1_value}({type(a1_value)}), A2: {a2_value}({type(a2_value)})')
    wb.close()


def load_and_fix():
    wb = openpyxl.load_workbook(BEFORE)
    wb = fix_default_date(wb, "A1")
    wb = fix_default_date(wb, "A2")
    wb.save(FIXED)


def fix_default_date(workbook, cell):
    sheet = workbook.active
    value = sheet[cell].value
    if value == EXCEL_ZERO_DATE:
        sheet[cell].value = EXCEL_ZERO_DATE + timedelta(days=1)
    return workbook


def load_and_fix_iso_date_but_ng():
    wb = openpyxl.load_workbook(BEFORE)
    wb.iso_dates = True
    wb.save(ISO)


if __name__ == '__main__':
    print('===== save only ====')
    load_and_save()
    print_before_and_after(BEFORE, AFTER)
    # =>
    # --- before ---
    # A1: 1899-12-30 00:00:00(<class 'datetime.datetime'>), A2: 00:01:00(<class 'datetime.time'>)
    # --- after ---
    # A1: 1899-12-29 00:00:00(<class 'datetime.datetime'>), A2: 00:01:00(<class 'datetime.time'>)

    print('===== save and fix =====')
    load_and_fix()
    print_before_and_after(BEFORE, FIXED)
    # =>
    # --- before ---
    # A1: 1899-12-30 00:00:00(<class 'datetime.datetime'>), A2: 00:01:00(<class 'datetime.time'>)
    # --- after ---
    # A1: 1899-12-30 00:00:00(<class 'datetime.datetime'>), A2: 00:01:00(<class 'datetime.time'>)

    print('===== save and iso_date=True =====')
    load_and_fix_iso_date_but_ng()
    print_before_and_after(BEFORE, ISO)
    # =>
    # --- before ---
    # A1: 1899-12-30 00:00:00(<class 'datetime.datetime'>), A2: 00:01:00(<class 'datetime.time'>)
    # --- after ---
    # A1: 1899-12-30 00:00:00(<class 'datetime.datetime'>), A2: 00:01:00(<class 'datetime.time'>)
