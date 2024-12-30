# %%
import openpyxl
import pandas as pd
import os
import streamlit as st


st.title("📄 Índice de NUP V1.0")
file = st.file_uploader("Sube el planilla de Revisiones NUP (.xlsx)", type=("xlsx"))


if file is not None and file.name.startswith("Revisiones DAOP"):

  df=pd.read_excel(file, sheet_name="Revisiones MR y MNR")
  df2=df

    option = st.selectbox(
    "Selecciona al ingeniero/a DAOP",
    ("LLQ", "JGM", "PPV", "RSP", "CGL", "NGM", "JMB" , "RGD"), )

  columns=df.loc[0]
  df=df.rename(columns=dict(zip(df.columns, columns)))

  df=df.drop(0)
  df=df.reset_index(drop=True)
  df=df.loc[~(df["Estado"]=="Respondido")&(df["Revisor"]==option)]
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
  st.dataframe(df2)

