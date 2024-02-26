from RPA.Excel.Files import Files
from RPA.Excel.Application import Application
from RPA.PDF import PDF
import pandas as pd
import shutil

def open_workbook(file_path)->Files:
    excel = Files()
    excel.open_workbook(file_path)
    
    return excel
def write_in_cell(excel: Files, sheet_name: str, row: str, column:str, value:str):
    print(excel.list_worksheets())
    excel.set_cell_value(row=row, column=column, value=value,name=sheet_name)

def get_clients(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    return df

def save_workbook(excel: Files):
    excel.save_workbook()
    excel.close_workbook()
    

def create_copy(file_path: str, new_name:str):
    shutil.copy(file_path, new_name)


def create_dict_from_csv(filepath: str) -> dict:
    df = pd.read_csv(filepath)
    dictionary = {}
    for _, row in df.iterrows():
        cell_list = str(row['cells']).split(',')
        dictionary[row['Info']] = cell_list
    return dictionary


def get_row_column(cell_id: str) -> list:
    column = ''.join(filter(str.isalpha, cell_id))  
    row = ''.join(filter(str.isdigit, cell_id))
    return column,row

def fill_data(excel: Files, df: pd.DataFrame, cells: dict,sheet_name:str): 
        for key, value in cells.items():
            data = df[key]
            if len(value)>1:
                length=0
                for cell_id in value:
                    column,row = get_row_column(cell_id)
                    write_in_cell(excel=excel, sheet_name=sheet_name, row=row, column=column, value=str(data)[length])
                    length=length+1
                
            else:
                
                cell_id = value[0]
                column,row = get_row_column(cell_id)
                exists_data = excel.get_cell_value(row,column,name=sheet_name)
                
                if exists_data is not None:
                    write_in_cell(excel=excel, sheet_name=sheet_name, row=row, column=column, value=exists_data+" "+str(data))
                else:
                    write_in_cell(excel=excel, sheet_name=sheet_name, row=row, column=column, value=str(data))
                    
                    
                    
def export_pdf(excel_name: str, pdf_name: str):
    pdf = Application()
    pdf.export_as_pdf(excel_filename=excel_name, pdf_filename=pdf_name)