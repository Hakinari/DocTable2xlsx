import win32com.client as win32   
import os
import sys
import argparse

def inicio(my_file,mult_tabs, sheet_name):
    try:
        xl = win32.Dispatch('Excel.Application')
        xl.Visible = 0
        xlsx_name = my_file.rstrip('.doc')
        _workbook = xl.Workbooks.Add()
        # abre word invisivel
        print("abriu excel")
        try:
            word = win32.Dispatch('Word.Application')
            word.Visible = 0 
            word.Documents.Open(os.path.join(os.getcwd(),my_file))
            doc = word.ActiveDocument
            print("abriu word")
        except:
            print("Problema para abrir o doc")
            exit(doc,word,xl,_workbook)
            return
			
        if mult_tabs:
            print("mult")
            multiple_tabs(doc,_workbook,xlsx_name, sheet_name)
        else:
            one_tab(doc,_workbook,xlsx_name, sheet_name)

        exit(doc,word,xl,_workbook)      
    except:
        print(sys.exc_info()[0])
        exit(doc,word,xl,_workbook)
        pass


def one_tab(doc,_workbook,xlsx_name, sheet_name):
    print("Total de tabelas: "+str(doc.tables.Count))
    row_record = 0
    _worksheet = _workbook.Sheets(1)
    _worksheet.Name = sheet_name
    _workbook.SaveAs(os.path.join(os.getcwd(), xlsx_name +'.xlsx'))
    for i in range(1,doc.tables.Count+1):
        table = doc.Tables(i)
        for row in range(1,table.Rows.Count+1):
            row_record += 1
            for column in range(1,table.Columns.Count+1):
                _worksheet.Cells(row_record,column).Value = table.Cell(Row=row, Column=column).Range.Text


def multiple_tabs(doc,_workbook, xlsx_name, sheet_name):
    print("Total de tabelas/abas: "+str(doc.tables.Count))
    i = 1 
    for table in doc.tables:
        if i == 1:
            _worksheet = _workbook.Sheets(i)
            _worksheet.Name = sheet_name+"1"
            _workbook.SaveAs(os.path.join(os.getcwd(), xlsx_name +'.xlsx'))
        else:
            _worksheet =_workbook.Worksheets.Add(Before=_worksheet)
            _worksheet.Name = sheet_name+str(i)
     
        table = doc.Tables(i)
        for row in range(1,table.Rows.Count+1):
            for column in range(1,table.Columns.Count+1):
                _worksheet.Cells(row,column).Value = table.Cell(Row=row, Column=column).Range.Text
        i+=1        
def exit(doc,word,xl,_workbook):
    doc.Close()
    word.Quit()
    del word
    _workbook.Close(True)
    xl.Quit()
    del xl
            
if __name__=="__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument('file_name', help="File name",)
    parser.add_argument('-m', help="Cada tabela em 1 aba", action='store_true', default=False)
    parser.add_argument('-sn', help="Nome da aba", default="Table")
	
    options = parser.parse_args()
    print("\n###############  Funciona apenas no windows com office ###############\n###############  Feche todas planilhas do excel, ou pode dar erro  ###############\n")
    print("Iniciando....\nArquivo Selecionado: "+options.file_name)
    inicio(options.file_name,options.m,str(options.sn))