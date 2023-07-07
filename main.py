# %% 
# Imports
import os

import geopandas as gpd
import pandas as pd
from openpyxl import load_workbook

# %% roots
# roots
rootFolder = os.getcwd()
rootUser = rootFolder.split('\\')

ABEIUser = '\\'.join(rootUser[0:3]) + "\\ABEIOneDrive"
try:
    os.path.exists(ABEIUser)   
except Exception as e:
    ABEIUser = '\\'.join(rootUser[0:3])
# C:\Users\usuario\ABEI Energy\Prom. Estudios Técnicos - Documentos\01- ESPAÑA
if not os.path.exists(ABEIUser):
    ABEIUser = '\\'.join(rootUser[0:3])

# Excel Lineas
rootLineas = "Lineas Overpass.xlsx"


# %% Excel file
# Excel file
# Read the xlsx
lineas = load_workbook(rootLineas, data_only = True)

# %%
dfLineas = pd.DataFrame(lineas['Lineas Overpass'].values) 
dfLineas.columns = dfLineas.iloc[0]
dfLineas.drop([0], inplace = True, axis = 0)

dfLineas = dfLineas[['Name', 'voltage']]
dfLineas = dfLineas.dropna(subset=['Name'])
dfLineas = dfLineas.dropna(subset=['voltage'])
dfLineas = dfLineas[dfLineas['Name'].str.contains('-', na=False)]
dfLineas = dfLineas[~dfLineas['Name'].str.contains(';')]
dfLineas = dfLineas[~dfLineas['Name'].duplicated()]
dfOrig = pd.DataFrame(lineas['Lineas Overpass'].values) # Hacemos una copia del bruto.

########################################################################################
########################################################################################


# %% 
# Criba por contener /. Se separan (es igual que cuando hay ;)
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace('/ 1', '/1'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace(' / 1', '/1'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace(' /1', '/1'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace('/ 2', '/2'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace(' / 2', '/2'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace(' /2', '/2'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace('/ 3', '/3'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace(' / 3', '/3'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace(' /3', '/3'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace('Melilli-Isab', 'Melilli_Isab'))
dfLineas['Name'] = dfLineas['Name'].apply(lambda x: x.replace('Melilli-Sio', 'Melilli_Sio'))
dfLineas = dfLineas[~dfLineas['Name'].str.contains('/1-2')]

# dfLineas = dfLineas.loc[dfLineas['Name'].str.contains(' / ').split(" / ")[0]]

# %%
#
dfLineas['Name 1'] = dfLineas['Name'].apply(lambda x: x.split(" / ")[0])
dfLineas['Name 2'] = dfLineas['Name'].apply(lambda x: x.split(" / ")[-1])

dfLineas1 = dfLineas[['Name 1', 'voltage']]
dfLineas2 = dfLineas[['Name 2', 'voltage']]
dfLineas1.columns = ['Name', 'voltage']
dfLineas2.columns = ['Name', 'voltage']
dfLineas = pd.concat([dfLineas1, dfLineas2])
dfLineas = dfLineas.drop_duplicates(subset='Name', keep='first')

# %%
# Separamos en función del número de líneas que hay en la string
dfLineasTres = dfLineas[dfLineas['Name'].str.count('-') >= 2]
dfLineasCuatro = dfLineasTres[dfLineasTres['Name'].str.count('-') >= 3]
dfLineas = dfLineas[dfLineas['Name'].str.count('-') <= 1]


####################################################################################################
####################################################################################################
# %%
# Aquí se incluirá el cálculo de las íneas. Hay que dejar todo metido en líneas de A-B. 




####################################################################################################
####################################################################################################
# Aquí sería para calcular el número de líneas que hay entrando en cada subestación
# %%
# Separamos para tener todas las SETs
dfLineas['A'] = dfLineas['Name'].apply(lambda x: x.split("-")[0])
dfLineas['B'] = dfLineas['Name'].apply(lambda x: x.split("-")[1])

dfLineasTres['A'] = dfLineasTres['Name'].apply(lambda x: x.split("-")[0])
dfLineasTres['B'] = dfLineasTres['Name'].apply(lambda x: x.split("-")[1])
dfLineasTres['C'] = dfLineasTres['Name'].apply(lambda x: x.split("-")[2])

dfLineasCuatro['A'] = dfLineasCuatro['Name'].apply(lambda x: x.split("-")[0])
dfLineasCuatro['B'] = dfLineasCuatro['Name'].apply(lambda x: x.split("-")[1])
dfLineasCuatro['C'] = dfLineasCuatro['Name'].apply(lambda x: x.split("-")[2])
dfLineasCuatro['D'] = dfLineasCuatro['Name'].apply(lambda x: x.split("-")[3])


# %%
# Listados

dfLineas['SET A'] = dfLineas['A'] + ';' + dfLineas['voltage']
dfLineas['SET B'] = dfLineas['B'] + ';' + dfLineas['voltage']
dfSETS = pd.DataFrame()
dfSETS['SET'] = pd.concat([dfLineas['SET A'], dfLineas['SET B']])
# # dfSETS= dfSETS['SET'].unique()
dfSETS = pd.DataFrame(dfSETS, columns = ['SET'])
dfSETS = dfSETS.astype(str)
dfSETS['Voltage'] = dfSETS['SET'].apply(lambda x: x.split(";")[-1])
dfSETS['SET'] = dfSETS['SET'].apply(lambda x: x.split(";")[0])


dfLineasTres['SET A'] = dfLineasTres['A'] + ';' + dfLineasTres['voltage']
dfLineasTres['SET B'] = dfLineasTres['B'] + ';' + dfLineasTres['voltage']
dfLineasTres['SET C'] = dfLineasTres['C'] + ';' + dfLineasTres['voltage']
dfSETS3 = pd.DataFrame()
dfSETS3['SET'] = pd.concat([dfLineasTres['SET A'], dfLineasTres['SET B'], dfLineasTres['SET C']])
# dfSETS3= dfSETS3['SET'].unique()
dfSETS3 = pd.DataFrame(dfSETS3, columns = ['SET'])
dfSETS3 = dfSETS3.astype(str)
dfSETS3['Voltage'] = dfSETS3['SET'].apply(lambda x: x.split(";")[-1])
dfSETS3['SET'] = dfSETS3['SET'].apply(lambda x: x.split(";")[0])

dfLineasCuatro['SET A'] = dfLineasCuatro['A'] + ';' + dfLineasCuatro['voltage']
dfLineasCuatro['SET B'] = dfLineasCuatro['B'] + ';' + dfLineasCuatro['voltage']
dfLineasCuatro['SET C'] = dfLineasCuatro['C'] + ';' + dfLineasCuatro['voltage']
dfLineasCuatro['SET D'] = dfLineasCuatro['D'] + ';' + dfLineasCuatro['voltage']
dfSETS4 = pd.DataFrame()
dfSETS4['SET'] = pd.concat([dfLineasCuatro['SET A'], dfLineasCuatro['SET B'], dfLineasCuatro['SET C'], dfLineasCuatro['SET D']])
# dfSETS4= dfSETS4['SET'].unique()
dfSETS4 = pd.DataFrame(dfSETS4, columns = ['SET'])
dfSETS4 = dfSETS4.astype(str)
dfSETS4['Voltage'] = dfSETS4['SET'].apply(lambda x: x.split(";")[-1])
dfSETS4['SET'] = dfSETS4['SET'].apply(lambda x: x.split(";")[0])
# %%
# We get the list
df = pd.concat([dfSETS, dfSETS3, dfSETS4])
df['SET'] = df['SET'].apply(lambda x: x.strip())

# %%
# We clean the dataframe
name_counts = df['SET'].value_counts()

# Add a new column 'Count' to the dataframe
df['Count'] = df['SET'].map(name_counts)
df = df.drop_duplicates()

# %%
