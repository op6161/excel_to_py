import src.excel_to_py.excel_module as ex
import src.excel_to_py.excel_extend as ext

file_path = "files/test_data.xlsx"
output_path = "files/output.py"

if __name__ == "__main__":
    data = ex.ExcelData(file_path)
    handler = ex.DataHandler(data)
    handler.set_sheet("Sheet1")
    cel1 = handler.get_cell(cell = "A1")
    sheet1 = handler.get_sheet()
    cel2 = handler.get_cell(cell = "A5")

    handler2 = ext.ExtendHandler(data)
    handler2.set_sheet("Sheet1")

    blocks = handler2.get_block()
    
    for block in blocks:
        print("block", block)