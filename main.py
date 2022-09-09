import os
import xlrd
import datetime

def marg_excel_to_py_list(excel_path: str):
    result = list()
    assert os.path.exists(excel_path)
    sheet: xlrd.sheet.Sheet = xlrd.open_workbook(excel_path).sheet_by_index(0)
    rows = sheet.get_rows()
    row = next(rows)

    def get_start_and_end_dates(row, rows):
        while "RECEIPT BOOK" not in row[0].value:
            row = next(rows)
        row = next(rows)
        start_date_str, _, end_date_str, *_ = row[0].value.split()
        # print(start_date_str, end_date_str)
        date_format = "%d-%m-%Y"
        start_date = datetime.datetime.strptime(start_date_str, date_format)
        end_date = datetime.datetime.strptime(end_date_str, date_format)
        # print(start_date, end_date)
        return start_date, end_date
    start_date, end_date = get_start_and_end_dates(row, rows)

    def get_transaction_date(row, start_date: datetime.datetime, end_date: datetime.datetime):
        date_format = "%b %d"
        try:
            transaction_date = datetime.datetime.strptime(row[0].value, date_format)
        except Exception as e:
            return False
        else:
            transaction_date = transaction_date.replace(year=start_date.year)
            if start_date <= transaction_date <= end_date:
                return transaction_date
            transaction_date = transaction_date.replace(year=end_date.year)
            if start_date <= transaction_date <= end_date:
                return transaction_date
            assert False, f"Invalid start or end date: start_date={start_date}, end_date={end_date}, transaction_date={transaction_date}."
            
    for row in rows:
        transaction_date = get_transaction_date(row, start_date, end_date)
        if transaction_date:
            # print(transaction_date)
            d = dict()
            d["date"] = transaction_date.isoformat()
            d["party"] = row[1].value
            row = next(rows)
            d["bank"] = row[1].value
            d["amount"] = row[2].value
            row = next(rows)
            d["desc"] = "".join([str(row[1].value), str(row[2].value), str(row[3].value)])
            # print(d)
            result.append(d)
        # print(transaction_date)

    return result