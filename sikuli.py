from sikuli import *
import unittest, HTMLTestRunner, shutil, utils


#Google Chrome
chrome = 'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'

dir = '.\\evidence\\test'
if not os.path.exists(dir):
        os.makedirs(dir)
outfile = open(".\\evidence\\test\\report.html", "w+") # path to report folder
screenshotfile = ".\\evidence\\test"

#Suite test
class testeChrome(unittest.TestCase):
    
    #Constructor 
    def setUp(self):
        App.open(chrome)
        wait(1)
        
    #Test Case    
    def test_AcessarGoogle(self):
      #  readXls(1,0)
        paste('1525954718691.png', readXls(1,0))
        type(Key.ENTER)
        wait(3)
        valor = capText('1525970454192.png')
        print('Passo 1 retorno de resultado: '+ valor )
        writXls(0,0, valor)
        getPrint('Passo2')     

    #Destructor            
    def tearDown(self):
        #pass
        type(Key.F4, KeyModifier.ALT)

#Report    
suite = unittest.TestLoader().loadTestsFromTestCase(testeChrome)
runner = HTMLTestRunner.HTMLTestRunner(stream=outfile, title=' Reporte de Execucao de Teste', description='Descrição do teste',dirTestScreenshots=screenshotfile)
runner.run(suite)
