import src.excel_to_py.excel_module as ex

file_path = "files/sample.xlsx"
output_path = "files/output.py"

if __name__ == "__main__":
    data = ex.ExcelData(file_path)
    handler = ex.DataHandler(data)
    cel1 = handler.get_cell(cell = "A1")
    sheet1 = handler.get_sheet("Sheet1")
    cel1 = handler.get_cell(sheet_name = "Sheet1", cell = "A1")
    print("sheet1",sheet1)
    print("cel1",cel1)