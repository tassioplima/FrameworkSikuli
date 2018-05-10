import xlwt, xlrd, os
from itertools import islice
#Utils for Sikuli  - Tassio Lima


#Method printscreen         
def getPrint(arq):
    newpath = r'.\\evidence\\print'    
    finalpath = path.join(newpath)
    img = capture (SCREEN)
    pathing = path.join(finalpath, arq+'.png')
    shutil.move(img, pathing)

def readXls(linha, coluna):
    dir = '.\\read'    
    if not os.path.exists(dir):
        os.makedirs(dir)
    excel = '.\\read\\massa.xls'
    book = xlrd.open_workbook(excel)
    sh = book.sheet_by_index(0)
    escreva = Env.getClipboard().strip()
    escreva = sh.cell_value(rowx=linha, colx=coluna)
    return escreva

#Method write XLS
def writXls(linha, coluna, valor):
    dir = '.\\write'
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('Conteudo', cell_overwrite_ok=True)
    worksheet.write(linha, coluna, valor)
    if not os.path.exists(dir):
        os.makedirs(dir)
    workbook.save('C:\\esc\\write\\arquivo.xls')

#Method read text img
def capText(image):
    valor = find(image).text()
    return valor
