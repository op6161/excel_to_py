import openpyxl
import os
import sys

class ExcelData:
    def __init__(self, file_path = None):
        self.workbook = None
        self.sheetnames = []
        self.sheets = {}
        if file_path:
            self.load(file_path)

    def __getitem__(self, sheet_name):
        """
        シート名でシートを取得
        """
        if self.workbook is None:
            raise ValueError("エクセルファイルが読み込んでいません。load()を使用してファイルを読み込んでください。")
        if sheet_name not in self.sheetnames:
            raise ValueError(f"シート名'{sheet_name}'が存在していません。")
        return self.sheets[sheet_name]

    def load(self, file_path):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"指定されたファイル'{file_path}'が存在していません。")
        
        # Excelファイル
        self.workbook = openpyxl.load_workbook(file_path)
        # シート名
        self.sheetnames = self.workbook.sheetnames
        # シートの辞書
        self.sheets = {name: self.workbook[name] for name in self.sheetnames}


class DataHandler:
    def __init__(self, excel_data, sheet_name = None):
        self.data = excel_data
        if not isinstance(self.data, ExcelData):
            raise ValueError("ExcelFileのインスタンスではありません。ExcelFileのインスタンスを渡してください。")
        self._current_sheet_name = None

        if sheet_name is not None:
            if sheet_name not in self.data.sheetnames:
                raise ValueError(f"シート名 '{sheet_name}' が存在していません。")
            self.set_sheet(sheet_name)

    def _resolve_sheet_name(self, sheet_name):
        if sheet_name is None:
            if self._current_sheet_name is None:
                raise ValueError(f"set_sheet('sheet_name')の実行、または {sys._getframe(1).f_code.co_name}(sheet_name = 'sheet_name')を指定してください")
            return self._current_sheet_name
        if sheet_name not in self.data.sheetnames:
            raise ValueError(f"シート名 '{sheet_name}' が存在していません。")
        return sheet_name

    def set_sheet(self, sheet_name) -> None:
        """
        シート名を設定
        """
        if sheet_name not in self.data.sheetnames:
            raise ValueError(f"シート名 '{sheet_name}' が存在しません。")
        self._current_sheet_name = sheet_name

    def get_sheet(self, sheet_name = None):
        """
        シート名でシートを取得
        """
        sheet_name = self._resolve_sheet_name(sheet_name)
        sheet_data = []
        for i in self.data[sheet_name].iter_rows(values_only=True):
            sheet_data.append(i)
        return sheet_data

    def get_cell(self, cell, sheet_name=None):
        """
        """
        sheet_name = self._resolve_sheet_name(sheet_name)
        return self.data[sheet_name][cell].value
    
    def get_row(self, row, sheet_name=None):
        sheet_name = self._resolve_sheet_name(sheet_name)
        if row < 1:
            raise ValueError("行番号は1以上でなければなりません。")
        return [cell.value for cell in self.data[sheet_name][row]]
    
    def get_column(self, column, sheet_name=None):
        sheet_name = self._resolve_sheet_name(sheet_name)
        if column < 1:
            raise ValueError("列番号は1以上でなければなりません。")
        column_data = []
        for cell_tuple in self.data[sheet_name].iter_cols(min_col=column, max_col=column, values_only=True):
            column_data.extend(cell_tuple) # iter_cols는 튜플의 튜플을 반환할 수 있으므로 extend 사용
        return column_data