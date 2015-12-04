# -*- coding: utf-8 -*-

from ofxstatement.plugin import Plugin
from ofxstatement.parser import StatementParser
from ofxstatement.statement import StatementLine, Statement

from xlrd import open_workbook
import csv, re, datetime, sys


class ZunoPlugin(Plugin):
    """
    Zuno plugin
    """

    def get_parser(self, filename):
        return ZunoParser(filename)


class ZunoParser(StatementParser):
    """
    0.Dátum transakcie:
    1.Typ transakcie
    2.Názov účtu v zozname
    3.Číslo účtu
    4.Kód banky
    5.Popis
    6.Suma
    7.Zostatok po transakcii
    29.11.2015	Platba kartou				BILLA  SPOL  S R O       BRATISLAVA   	-6,86	EUR	16 160,90	EUR
    """
    mappings = {"date":0,
                "payee":[3, 4], #"id": 2,
                "memo": 5,
                "amount": 6}
    encoding = 'windows-1250' # 'utf-8'
    date_format = "%d.%m.%Y"
    filename = ''
    def __init__(self, filename):
        self.filename = filename
        self.statement = Statement()
        self.statement.currency = 'EUR'
        return super(ZunoParser, self).__init__()

    def parse_float(self, value):
        return float(value.replace(",","."))

    def parse(self):
        """
        Main entry point for parsers

        super() implementation will call to split_records and parse_record to
        process the file.
        """
        self.book = open_workbook(self.filename, on_demand=True)
        return super(ZunoParser, self).parse()

    def split_records(self):
        """
        Return iterable object consisting of a line per transaction
        """
        for name in self.book.sheet_names():
            sheet = self.book.sheet_by_name(name)

            for i, rx in enumerate(range(sheet.nrows)):
                if i == 0:
                    continue
                yield sheet.row(rx)                #self.statement.lines.append(self.parse_record(sheet.row(rx)))
                #self.statement.lines.append(self.parse_record(sheet.row(rx)))

            self.book.unload_sheet(name)
        #return self.statement

    def parse_record(self, line):
        """
        Parse given transaction line and return StatementLine object
        """
        stmt_line = StatementLine()
        for field, col in self.mappings.items():
            if type(col) is list:
                rawvalue = []
                for ii in col:
                    rawvalue.append(line[ii].value)
                rawvalue = "|".join(rawvalue)
            else:
                rawvalue = line[col].value
            value = self.parse_value(rawvalue, field)
            setattr(stmt_line, field, value)
        return stmt_line
