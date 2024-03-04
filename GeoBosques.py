## 1.  PAQUETES ###############################################################
# Clasicos
import pandas as pd
import numpy as np

# Paths
import re
import os
from pathlib import Path

# For simulate human behavior.
import time
from time import sleep
import random

# Clear data
import unidecode

# DIVERSOS
from IPython.core.display import display, HTML
#display(HTML("<style>.container { width:80% !important; }</style>"))
from datetime import date
import pytest
import json

# SELENIUM
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.desired_capabilities import DesiredCapabilities

# Options driver
from webdriver_manager.chrome import ChromeDriverManager
from win32com.client import Dispatch


## 2. Área de trabajo #########################################################

# DEFINIR EL PATH PRINCIPAL
os.chdir(r"D:\CIUP\TRABAJO\SEMANA 18\geobosques_scrapping")

# GUARDAD EL PATH COMO OBJETO Y ABREVIR PARA RUTAS FUTURAS
inicio = os.getcwd()
new_dir = 'Prueba1_' + str(date.today())

# CREAR UNA CARPETA Y DEFINIR SU PATH
Path(new_dir).mkdir(exist_ok=True)
descargas = os.path.join(inicio, new_dir); descargas

## 3. Extracción de datos #####################################################
# Iniciar selenium
options = webdriver.ChromeOptions()
options.add_argument("--start-maximized")
driver = webdriver.Chrome(r"C:\Users\JOEL\AppData\Local\Programs\Python\Python38\Lib\site-packages\selenium\chromedriver.exe", chrome_options=options)

# Abrir página de interés
driver.get("http://geobosques.minam.gob.pe/geobosque/view/perdida.php")

#######################################################################################

# OBTENCIÓN DE DATOS: BRUTO
all_tables=[]
list_dptos = driver.find_element_by_id( 'dr_departamento_chosen')
list_dptos.click()
n_dpto = len(list_dptos.find_elements_by_class_name( "active-result" ))
nombres_dptos_pre = list((driver.find_element_by_id( 'dr_departamento_chosen').text).split("\n"))
nombres_dptos = nombres_dptos_pre[1:]
list_dptos.click()
p = 0
for dpto_index in range( 1, n_dpto ):
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]//input[contains(@class, 'chosen-search-input')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]//input[contains(@class, 'chosen-search-input')]").send_keys(nombres_dptos[dpto_index])
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]//li[contains(@class, 'active-result')]").click()
    time.sleep( 2 )
 
    # Total provinces
    list_prov = driver.find_element_by_id( 'dr_provincia_chosen')
    list_prov.click()
    n_prov = len(list_prov.find_elements_by_class_name( "active-result"))
    nombres_prov_pre = list((driver.find_element_by_id( 'dr_provincia_chosen').text).split("\n"))
    nombres_prov = nombres_prov_pre[1:]
    list_prov.click()
    
    for prov_index in range( 1, n_prov ):
        driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]").click()
        driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]//input[contains(@class, 'chosen-search-input')]").click()
        driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]//input[contains(@class, 'chosen-search-input')]").send_keys(nombres_prov[prov_index])
        driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]//li[contains(@class, 'active-result')]").click()
        time.sleep( 2 )
        
        list_dist = driver.find_element_by_id( 'dr_distrito_chosen')
        list_dist.click()
        n_dist = len(list_dist.find_elements_by_class_name( "active-result"))
        nombres_dist_pre = list((driver.find_element_by_id( 'dr_distrito_chosen').text).split("\n"))
        nombres_dist = nombres_dist_pre[1:]
        list_dist.click()
        
        for dist_index in range( 1, n_dist ):
            driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]").click()
            driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]//input[contains(@class, 'chosen-search-input')]").click()
            driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]//input[contains(@class, 'chosen-search-input')]").send_keys(nombres_dist[dist_index])
            driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]//li[contains(@class, 'active-result')]").click()
            ubigeo_path  = "//select[contains(@id,'dr-distrito')]//option[contains(text(),"+ "'"+nombres_dist[dist_index]+"'"+")]"
            ubigeo_val = driver.find_element(By.XPATH, ubigeo_path).get_attribute('value')
            time.sleep( 2 )
            
            tabla_html = driver.find_element_by_id("pannel-perdida-t-ha")
            tabla_final = pd.read_html( tabla_html.get_attribute('outerHTML') )[0]
            tabla_final.drop(tabla_final[tabla_final['Rango']=="Total"].index, inplace=True)
            tabla_final['dpto']=nombres_dptos[dpto_index]
            tabla_final['prov']=nombres_prov[prov_index]
            tabla_final['dist']=nombres_dist[dist_index]
            tabla_final['ubigeo']=ubigeo_val
            
            all_tables.append(tabla_final)
            
            print(str(p)+" | "+nombres_dptos[dpto_index]+" |  "+nombres_prov[prov_index]+" | "+nombres_dist[dist_index]+" | "+ubigeo_val)
            p = p + 1
            
parte_1 = pd.concat(all_tables, axis=0)

# Algunos distritos han  tomado un UBIGEO incorrecto y solo uno se ha repetido.
parte_1 = parte_1.drop_duplicates(subset=['Rango','ubigeo'])

# Los distritos eliminados anteriormente, tomarán el valor corecto
faltantes  =  {'dpto_f':['SAN MARTIN','CUSCO','JUNiN','JUNiN','JUNiN','LORETO','PIURA','AYACUCHO'],
               'prov_f':['HUALLAGA','ESPINAR','JAUJA','JAUJA','TARMA','REQUENA','HUANCABAMBA','LUCANAS'],
               'dist_f':['SAPOSOA','PICHIGUA','MASMA','MUQUI','PALCA','TAPICHE','SONDOR','SAN PEDRO'],
               'ubigeo_f':['220401','080806','120416','120420','120706','160509','200307','050617']}

faltantes_df = pd.DataFrame.from_dict(faltantes)

all_tables_f = []
for i in range(0,len(faltantes_df)):
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]//input[contains(@class, 'chosen-search-input')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]//input[contains(@class, 'chosen-search-input')]").send_keys(faltantes_df['dpto_f'][i])
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_departamento_chosen')]//li[contains(@class, 'active-result')]").click()
    time.sleep( 2 )
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]//input[contains(@class, 'chosen-search-input')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]//input[contains(@class, 'chosen-search-input')]").send_keys(faltantes_df['prov_f'][i])
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_provincia_chosen')]//li[contains(@class, 'active-result')]").click()
    time.sleep( 2 )
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]//input[contains(@class, 'chosen-search-input')]").click()
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]//input[contains(@class, 'chosen-search-input')]").send_keys(faltantes_df['dist_f'][i])
    time.sleep( 2 )
    driver.find_element(By.XPATH, "//div[contains(@id, 'dr_distrito_chosen')]//li[contains(@class, 'active-result')][2]").click()
    time.sleep( 2 )
    
    tabla_html_f = driver.find_element_by_id("pannel-perdida-t-ha")
    tabla_final_f = pd.read_html( tabla_html_f.get_attribute('outerHTML') )[0]
    tabla_final_f.drop(tabla_final_f[tabla_final_f['Rango']=="Total"].index, inplace=True)
    tabla_final_f['dpto']=faltantes_df['dpto_f'][i]
    tabla_final_f['prov']=faltantes_df['prov_f'][i]
    tabla_final_f['dist']=faltantes_df['dist_f'][i]
    tabla_final_f['ubigeo']=faltantes_df['ubigeo_f'][i]
    
    all_tables_f.append(tabla_final_f)
    
parte_2 = pd.concat(all_tables_f, axis=0)   
    
# JUNTAMOS LOS  VALORES FINALES
data = pd.concat([parte_1, parte_2], ignore_index=True)

# ORDENAMOS Y CAMBIAMOS NOMBRES
old_cols = list(data.columns)
old_cols = old_cols[1:22]
for i in range(0, len(old_cols)):
    data = data.rename({old_cols[i]: 'y'+old_cols[i]}, axis=1)
cols = list(data.columns)
cols = cols[22:27] + cols[0:22]
data_final = data[cols]

# EXPORTAMOS LOS RESULTADOS
data_final.to_excel('GeoBosques_final.xlsx', index=False)

