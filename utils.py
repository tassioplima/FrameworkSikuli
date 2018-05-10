#Utils for Sikuli  - Tassio Lima


#Method printscreen         
def getPrint(arq):
    newpath = r'C:\\evidencias\\testes'    
    finalpath = path.join(newpath)
    img = capture (SCREEN)
    pathing = path.join(finalpath, arq+'.png')
    shutil.move(img, pathing)

#Method read XLS 
def readXls(linha, coluna):
    excel = 'C:\\Users\\BR3TPL\\Documents\\Projetos\\testesikuli.sikuli\\test.xls'
    book = xlrd.open_workbook(excel)
    sh = book.sheet_by_index(0)
    escreva = Env.getClipboard().strip()
    escreva = sh.cell_value(rowx=linha, colx=coluna)
    return escreva

#Method write XLS
def writeXls(linha, coluna, valor):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('Plan1', cell_overwrite_ok=True)
    worksheet.write(linha, coluna, valor )
    workbook.save('./evidencias/')

#Method read text img
def captText(image):
    valor = find(image).text()
    return valor


    
