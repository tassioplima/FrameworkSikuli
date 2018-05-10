import unittest, HTMLTestRunner, shutil, xlwt, xlrd
from itertools import islice
from os import path
from sikuli import *
import utils
#Variáveis de print

#Variáveis de report
outfile = open("C:\\evidencias\\testes\\report.html", "w+") # path to report folder
screenshotfile = "C:\\evidencias\\testes"
#Google Chrome
chrome = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'

#Suite test
class testeChrome(unittest.TestCase):
    
    #Constructor 
    def setUp(self):
        App.open(chrome)
        wait(1)
        
    #Test Case    
    def test_AcessarGoogle(self):  
        paste('1525954718691.png', readXls(1,0))
        type(Key.ENTER)
        wait(3)
        valor = capText('1525970454192.png')
        print('Passo 1 retorno de resultado: '+ valor )
        getPrint('Passo2')

    #Destructor            
    def tearDown(self):
        #pass
        type(Key.F4, KeyModifier.ALT)



#Report    
suite = unittest.TestLoader().loadTestsFromTestCase(testeChrome)
runner = HTMLTestRunner.HTMLTestRunner(stream=outfile, title=' Reporte de Execucao de Teste', description='Descrição do teste',dirTestScreenshots=screenshotfile)
runner.run(suite)
