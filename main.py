import urllib.request as urllib2
import urllib.parse
from bs4 import BeautifulSoup
from urllib.request import Request, urlopen
from datetime import date
import sys
import xlrd 
import pandas as pd
import openpyxl
import time

# busca primera coincidencia
def find_between( s, first, last ):
    try:
        start = s.index( first ) + len( first )
        end = s.index( last, start )
        return s[start:end]
    except ValueError:
        return ""
# busca ultima coincidencia
def find_between_r( s, first, last ):
    try:
        start = s.rindex( first ) + len( first )
        end = s.rindex( last, start )
        return s[start:end]
    except ValueError:
        return ""

origen = "Angel-Scrapped7days.xlsx"
destino = "Angel-Scrapped7days_fix.xlsx"
#https://analisisydecision.es/leer-archivos-excel-con-python/
#wb = xlrd.open_workbook(origen) 

#hoja = wb.sheet_by_index(0) 
#print(hoja.nrows) 
#print(hoja.ncols) 
#print(hoja.cell_value(0, 0))
#hoja = wb.sheet_by_index(0) 
#nombres = hoja.row(0)  
#print(nombres[0])

wb = xlrd.open_workbook(origen) 

hoja = wb.sheet_by_index(0) 

# Creamos listas
filas = []
for fila in range(1,hoja.nrows):
    columnas = []
    for columna in range(0,3):
        columnas.append(hoja.cell_value(fila,columna))
    filas.append(columnas)

import pandas as pd
df = pd.DataFrame(filas)
df.head()
#print(df)
for index, row in df.iterrows():
    time.sleep(5)
    opener = urllib.request.build_opener()
    #opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36'), ('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'), ('Accept-Encoding','gzip, deflate, br'),\
    #    ('Accept-Language','en-US,en;q=0.5' ), ("Connection", "keep-alive"), ("Upgrade-Insecure-Requests",'1')]
    opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Linux; Android 4.3; Nexus 7 Build/JSS15Q) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'), ('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'), \
        ('Accept-Language','en-US,en;q=0.5' ), ("Connection", "keep-alive"), ("Upgrade-Insecure-Requests",'1')]
    #opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Linux; U; Android 2.3.6; en-us; Nexus S Build/GRK39F) AppleWebKit/533.1 (KHTML, like Gecko) Version/4.0 Mobile Safari/533.1'), ('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'), \
    #    ('Accept-Language','en-US,en;q=0.5' ), ("Connection", "keep-alive"), ("Upgrade-Insecure-Requests",'1')]
    urllib.request.install_opener(opener)
    error = 0
    try:
        page = urllib2.urlopen(row[0])
    except urllib2.URLError as err:
        print (err.code)
        #print (err.read())
        error = 1
    if (error == 0 ):
        soup = BeautifulSoup(page)
        x = soup.__str__()
        website = find_between( x, 'styles_links__VvYv7"><ul><li class="styles_websiteLink___Rnfc"><a href="', '" rel="nofollow ugc" target="_blank">' )
        print('primera')
        print (website)
        row[2] = website
    else:
        break
        #row[2] = 'Error'
#print(df)
df.to_excel(destino, "Sheet1")




