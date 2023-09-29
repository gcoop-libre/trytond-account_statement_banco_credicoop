# The COPYRIGHT file at the top level of this repository contains
# the full copyright notices and license terms.
import xlrd
from io import BytesIO
from datetime import datetime
from decimal import Decimal


def _date(value):
    v = value.strip()
    return datetime.strptime(v, '%d/%m/%Y').date()


def _string(value):
    return value.strip()


def _amount(value):
    return Decimal(str(value).replace(',', '.'))


def _code_to_string(value):
    try:
        return str(int(value))
    except ValueError:
        return str(value)


COL = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6}

MOVE = {
    'date': ('A', _date),
    'concept': ('B', _string),
    'debit': ('D', _amount),
    'credit': ('E', _amount),
    'code': ('G', _code_to_string),
    }


class Credicoop(object):

    def __init__(self, name, encoding='windows-1252'):
        self.statements = []
        self._parse(name)

    def _parse(self, infile):
        statement = Statement()
        self.statements.append(statement)

        filedata = BytesIO(infile)
        workbook = xlrd.open_workbook(file_contents=filedata.getvalue())
        worksheets = workbook.sheet_names()
        for worksheet_name in worksheets:
            worksheet = workbook.sheet_by_name(worksheet_name)

            num_rows = worksheet.nrows - 1
            curr_row = 0
            while curr_row < num_rows:
                # Start reading at the second row
                curr_row += 1
                row = worksheet.row(curr_row)
                move = Move()
                self._parse_move(row, move, MOVE)
                # 'Date from', in first row
                if curr_row == 1:
                    statement.date_from = move.date
                statement.date_to = move.date
                move.op_number = curr_row
                statement.moves.append(move)

    def _parse_move(self, row, move, desc):
        for name, (col, parser) in desc.items():
            value = parser(row[COL[col]].value)
            setattr(move, name, value)


class Statement(object):
    __slots__ = ['date_from', 'date_to', 'moves']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.moves = []


class Move(object):
    __slots__ = list(MOVE.keys()) + ['op_number']

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
