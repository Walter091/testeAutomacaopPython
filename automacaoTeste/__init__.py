import AutoBase
import AutoExplorer
import openpyxl

def getAndSetPlanilha():
    book = openpyxl.load_workbook('Planilha de compras.xlsx')
    pagina = book['Frutas']
    for rows in pagina.iter_rows(min_row=2, max_row=5):
        for cell in rows:
            print(cell.value)

    book.save("Planilha de compras.xlsx")

obj = AutoExplorer.TesteAutomacao()
obj.Index()

# AutoBase.localizationPositionMouse()


# Verificar este erro:
# AttributeError: 'Workbook' object has no attribute 'write'
# Exception ignored in: <function ZipFile.__del__ at 0x000001BBBDA39160>
# Traceback (most recent call last):
#   File "C:\Users\hp\AppData\Local\Programs\Python\Python39\lib\zipfile.py", line 1807, in __del__
#     self.close()
#   File "C:\Users\hp\AppData\Local\Programs\Python\Python39\lib\zipfile.py", line 1825, in close
#     self._write_end_record()
#   File "C:\Users\hp\AppData\Local\Programs\Python\Python39\lib\zipfile.py", line 1919, in _write_end_record
#     self.fp.write(endrec)
#   File "C:\Users\hp\AppData\Local\Programs\Python\Python39\lib\zipfile.py", line 759, in write
#     n = self.fp.write(data)
# AttributeError: 'Workbook' object has no attribute 'write'
#
# Process finished with exit code 1
