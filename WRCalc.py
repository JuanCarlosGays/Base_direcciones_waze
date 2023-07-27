import WazeRouteCalculator
import subprocess
import logging
import pandas as pd
import os


logger = logging.getLogger('WazeRouteCalculator.WazeRouteCalculator')
logger.setLevel(logging.DEBUG)
handler = logging.StreamHandler()
logger.addHandler(handler)

os.system('cls')
       
print("ESTIMADOR DE DISTANCIAS CON WAZE, VERIFICAR QUE EL ARCHIVO base_direcciones.xlsx está en el mismo directorio de ejecución")
print("por Juan C Gays")

dat = pd.read_excel('base_direcciones.xlsx', sheet_name="inicio",usecols= 'A, E,F,G,H,I')

p_intermedios=input("registrar puntos intermedios?(si/no):")
b_referencia=input("registrar contra base de referencia?(si/no):")
if(b_referencia=="si"):
    base_referencia=input("base referencia (ej. Talleres Varela, Aysa, Buenos Aires, Argentina ):")



region = 'AR'

for linea in dat.index:
    if(p_intermedios=="si"):
        try:
            route = WazeRouteCalculator.WazeRouteCalculator(dat.iloc[linea,0],dat.iloc[linea,1], region)
            print(route.calc_route_info()[0],"min,", route.calc_route_info()[1], "km")
            dat.iloc[linea,2]=route.calc_route_info()[0]
            dat.iloc[linea,3]=route.calc_route_info()[1]
        except:
            dat.iloc[linea,2]="ver error"
            dat.iloc[linea,3]="ver error"
    if(b_referencia=="si"):   
        try:
            route = WazeRouteCalculator.WazeRouteCalculator(base_referencia,dat.iloc[linea,1], region)
            print(route.calc_route_info()[0],"min,", route.calc_route_info()[1], "km")
            dat.iloc[linea,4]=route.calc_route_info()[0]
            dat.iloc[linea,5]=route.calc_route_info()[1]
        except:
            dat.iloc[linea,4]="ver error"
            dat.iloc[linea,5]="ver error"
        
with pd.ExcelWriter("base_direcciones.xlsx", mode="a", engine="openpyxl",if_sheet_exists='replace') as writer:
    dat.to_excel(writer, sheet_name="final")
            