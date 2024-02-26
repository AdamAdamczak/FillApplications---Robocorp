from robocorp.tasks import task
from utils.excel_hande import export_pdf,open_workbook, save_workbook, get_clients, create_dict_from_csv,create_copy,fill_data
@task
def fill_excel():
    df = get_clients('template/clients.xlsx')
    cells = create_dict_from_csv('template/cells.csv')
    
    for _,each in df.iterrows():
        file_name ='output_files/'+each['Nazwisko']+'_'+each['Imię']+'.xlsx'
    
        create_copy('template/template.xlsx',file_name)
    
        excel = open_workbook(file_name)    
        fill_data(excel=excel, df=each, cells=cells,sheet_name='str. 1')
        save_workbook(excel=excel)
        export_pdf(file_name, 'output_files/'+each['Nazwisko']+'_'+each['Imię']+'.pdf')

