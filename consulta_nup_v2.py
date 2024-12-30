# %%
import openpyxl
import pandas as pd
import os
df=pd.read_excel("C:\\Users\\luis.lizana\\OneDrive - Coordinador Eléctrico Nacional\\Archivos actualizables para automatización\\Revisiones DAOP a MR, MNR y otras modificaciones.xlsx", sheet_name="Revisiones MR y MNR")
df2=df

columns=df.loc[0]
df=df.rename(columns=dict(zip(df.columns, columns)))

df=df.drop(0)
df=df.reset_index(drop=True)
df=df.loc[~(df["Estado"]=="Respondido")&(df["Revisor"]=="LLQ")]
df2=pd.DataFrame()
df2["Asunto"]=df["NUP"].astype(str)+" "+df["Nombre del Proyecto - Descripción"]+" - "+df["Tipo de Documento"]+" - Iter.N°"+df["N°de revisión"].astype(str)
df2["Fecha de inicio"]=df["Fecha de recepción Coordinador"]
df2["Fecha de vencimiento"]=df["Fecha máxima  DAOP (KPI)"]
df2["Prioridad"]=0
df2["% Completado"]=0
df2["Estado"]="No comenzada"
df2["Categoría"]="Categoría azul"
df2["Mensaje"]=""
df2.to_excel("salida_nup_v2.xlsx", index=False)
#os.remove("C:\\Users\\luis.lizana\\Downloads\\NUP.xlsx")