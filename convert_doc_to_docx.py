from glob import glob
import os
import win32com.client as win32
from win32com.client import constants

def save_as_docx(path):
   
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(path)
    doc.Activate ()
    
    new_file_abs = os.path.splitext(os.path.abspath(path))[0] + "_convert.docx"
    print( new_file_abs)
    
    word.ActiveDocument.SaveAs(
        new_file_abs, FileFormat=constants.wdFormatXMLDocument
    )
    doc.Close(False)

paths = glob(r'E:\Отдел корпоративной защиты\Персонал\Персонал\Я\**\*.doc', recursive=True)

for path in paths:
    save_as_docx(path)