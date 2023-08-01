# CIEE
# Tratamento de Dados para Planilhas e turmas existentes
# criacao de CSV
# 
# NOME ABA, DIA, TURNO, DATA, TOTAL, CICLO, INSTRUTOR, TURMA, QUANTIDADE
# tabela, dia, turno, data, total, ciclo, instrutor, df[colunas[colunaInicial]][j], str(df[colunas[colunaInicial+3]][j])
#

import pandas as pd
import csv 
import streamlit as st
import io

def obtemDia ( df, colunas, linhas ):
  for i in range(linhas):
    if df[colunas[0]][i].upper() == "DIA:" :
      dia = df[colunas[1]][i]
      return dia, i
  return "Nan",0

def obtemTurno ( df, colunas, linhas ):
  for i in range(linhas):
    if df[colunas[0]][i].upper() == "TURNO:" :
      turno = df[colunas[1]][i]
      return turno, i
  return "Nan",0

def obtemData ( df, colunas, linhas ):
  return df[colunas[2]][2], 2

def obtemTotalAprendizes( df, colunas, linhas ):
  return df[colunas[5]][2], 2

def obtemCiclo( df, colunas, linhas, linhaInicial, colunaInicial , tabela, dia, turno, data, total):
  linha = linhaInicial
  for i in range( linha , linhas ):
    if df[colunas[colunaInicial]][i] == "CICLO" :
      break
  if i < linhas :
    ciclo = df[colunas[colunaInicial+1]][i]
    turmas = []
    i +=1
    if i < linhas :
      if ( pd.isna( df[colunas[colunaInicial+1]][i] )) :
        return "","","",[], i
    else :
      return "","","",[], i

    instrutor = df[colunas[colunaInicial+1]][i]

    i +=1
    quantidade = df[colunas[colunaInicial]][i]

    i +=2
    j=i
    if ( i < linhas ):
      for j in range( i, linhas):
        if ( not pd.isna(df[colunas[colunaInicial]][j]) ) :
          dados =[ tabela, dia, turno, data, total, ciclo, instrutor, df[colunas[colunaInicial]][j], str(df[colunas[colunaInicial+3]][j]) ]
          turmas.append(dados)
        else :
          break
    return  ciclo, instrutor, quantidade, turmas, j

  else :
    return "","","",[], i
  

def principal( arquivo, abas ):

    # file_buffer = io.StringIO()

    #with file_buffer :  
    dados = []
    header = ["NOME ABA", "DIA", "TURNO", "DATA", "TOTAL", "CICLO", "INSTRUTOR", "TURMA", "QUANTIDADE"]
    dados.append(header)

    #    writer = csv.writer( file_buffer ) 
    #    writer.writerow( header ) 

    # looping para abrir abas
    # abas = xls.sheet_names
    # df = pd.read_excel( arquivo , sheet_name=abas[0])
    for tabela in abas :
        df = pd.read_excel ( arquivo, sheet_name = tabela )

        # Iteração direta na série "Nome" para percorrer linha por linha
        colunas =  df.columns
        linhas = len(df)
        dia, linha = obtemDia( df, colunas, linhas)
        turno, linha = obtemTurno( df, colunas, linhas)
        data, linha = obtemData ( df, colunas, linhas )
        total, linha = obtemTotalAprendizes( df, colunas, linhas)

        for coluna in range( 0, len(colunas), 5) :
            linha = 0
            while linha < len(df) :
                ciclo, instrutor, quantidade, turmas, ultimalinha = obtemCiclo( df, colunas, linhas, linha, coluna , tabela, dia, turno, data, total)
                # print( linha, turmas)
                # file = open(arquivo_destino, 'a+', newline ='') 
                #with file:     
                # writer = csv.writer(file) 
                #for turma in turmas :
                #    writer.writerow(turma) 
                dados.extend(turmas)
                linha = ultimalinha

    csv_buffer = io.StringIO()
    writer = csv.writer(csv_buffer)
    writer.writerows(dados)

    # Exibindo o botão de download no Streamlit
    st.download_button(
        label="Baixar CSV",
        data=csv_buffer.getvalue().encode(),
        file_name='turmas_em_linhas.csv',
        mime='text/csv',
    )
    return
    
    
# ------------------------------------
st.write(" Conversão Tabela")
arquivo = st.file_uploader('Escolha a tabela origem (.csv)', type = 'xlsx')
if arquivo is not None:
    xls = pd.ExcelFile(arquivo)
    # Obtenha os nomes das abas (sheets) da planilha
    abas = xls.sheet_names
    principal ( arquivo, abas )

