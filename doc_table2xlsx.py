import win32com.client as win32   
import os
import sys
import argparse

def inicio(my_file,mult_tabs):
    # abre excel invisivel
    try:
        xl = win32.Dispatch('Excel.Application')
        xl.Visible = 1
        xlsx_name = my_file.rstrip('.doc')
        _workbook = xl.Workbooks.Add()

        # abre word invisivel
        word = win32.Dispatch('Word.Application')
        word.Visible = 0 
        try:
            word.Documents.Open(os.path.join(os.getcwd(),my_file))
            doc = word.ActiveDocument
        except:
            print("Problema para abrir o doc")
            exit(doc,word,xl,_workbook)
            return

        if mult_tabs:
            multiple_tabs(doc,_workbook,xlsx_name)
        else:
            one_tab(doc,_workbook,xlsx_name)

        exit(doc,word,xl,_workbook)

        
    except:
        exit(doc,word,xl,_workbook)
        pass


def one_tab(doc,_workbook,xlsx_name):
    row_record = 0
    _worksheet = _workbook.Sheets(1)
    _worksheet.Name = "Tabela"
    _workbook.SaveAs(os.path.join(os.getcwd(), xlsx_name +'.xlsx'))
    for i in range(1,doc.tables.Count+1):
        table = doc.Tables(i)
        for row in range(1,table.Rows.Count+1):
            row_record += 1
            for column in range(1,table.Columns.Count+1):
                _worksheet.Cells(row_record,column).Value = table.Cell(Row=row, Column=column).Range.Text


def multiple_tabs(doc,_workbook, xlsx_name):
    print(doc.tables.Count)
    #for i in range(1,doc.tables.Count+1):
    
    for i in range(1,doc.tables.Count+1):
        if i == 1:
            _worksheet = _workbook.Sheets(i)
            _worksheet.Name = "Tabela1"
            _workbook.SaveAs(os.path.join(os.getcwd(), xlsx_name +'.xlsx'))
        else:
            _worksheet =_workbook.Worksheets.Add(After=_worksheet)
            _worksheet.Name = "Tabela"+str(i)
     
        table = doc.Tables(i)
        for row in range(1,table.Rows.Count+1):
            for column in range(1,table.Columns.Count+1):
                _worksheet.Cells(row,column).Value = table.Cell(Row=row, Column=column).Range.Text
        
def exit(doc,word,xl,_workbook):
    doc.Close()
    word.Quit()
    del word
    _workbook.Close(True)
    xl.Quit()
    del xl
            
if __name__=="__main__":

    parser = argparse.ArgumentParser()
    parser.add_argument('string', help="File name")
    parser.add_argument('-m', help="Each table per sheets", action='store_true', default=False)
    
    options = parser.parse_args()
    try:
        mult_tabs = options.m
    except:
        mult_tabs = False
    if mult_tabs: ## se for separar por multiplas abas
        my_file = sys.argv[2]    
    else:
        my_file = sys.argv[1]

    print("Feche todas planilhas do excel, ou pode dar erro")
    print("")
    print my_file
    inicio(my_file,mult_tabs)
    
