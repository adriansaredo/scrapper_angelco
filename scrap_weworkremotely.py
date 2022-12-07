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

origen = "weworkremotelytotal.xlsx"
destino = "weworkremotelytotalfix.xlsx"

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
    if str(row[2]).strip(): 
        print('Linea ya procesada: ' + str(row[0]) + ' : '+ str(row[1]) + ' : '+ str(row[2]))
    else:
        time.sleep(20)
        opener = urllib.request.build_opener()
        opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Linux; Android 4.3; Nexus 7 Build/JSS15Q) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'), ('Accept','text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'), \
            ('Accept-Language','en-US,en;q=0.5' ), ("Connection", "keep-alive"), ("Upgrade-Insecure-Requests",'1')]
        urllib.request.install_opener(opener)
        error = 0
        try:
            #print(row[0])
            page = urllib2.urlopen(row[0])
        except urllib2.URLError as err:
            print (err.code)
            #print (err.read())
            error = 1
        #print(error)
        if (error == 0 ):
            soup = BeautifulSoup(page)
            #print(soup.encode('utf-8'))
            x = soup.__str__()
            #print(x.encode('utf-8'))
            website = find_between( str(x.encode('utf-8')), '</div><div style="margin-top: -38px;"><h3><a href="', '" target="_blank">Website</a></h3></div><div>' )
            if (website == "" ):
                website = find_between( str(x.encode('utf-8')), '</div><div style="margin-top: -38px;"><h3><a target=_blank href="', '">Website</a></h3></div><div>' )
            #print('primera')
            print (website)
            row[2] = website
            #break
            df.to_excel(destino, "Sheet1", index=False)
        else:
            break
            #row[2] = 'Error'
        

print('Avance Grabado en el Excel' + origen)
#df = df.drop(df.columns[[0]], axis='columns')
df.to_excel(origen, sheet_name="Sheet1", index=False)
