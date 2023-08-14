#!/usr/bin/env python
# -*- coding: utf-8 -*-
#Brunodgk 2023
####################################################
#Importar bibliotecas

import MetaTrader5 as mt5;
import pandas as pd;
import numpy as np;
#import time;
from tzlocal import get_localzone;
from datetime import datetime, timezone;
import xlwings as xw;
import platform;
import subprocess;
import xlrd;
import openpyxl;

####################################################
#Diferenciar pequenos comandos para cada sistema operacional

if (platform.system().lower() == "windows"):
    command = "CLS";
else:
    command = "clear";

####################################################
#Declarando variáveis globais

planilha = "CotacaoAtivosB3MT5.xlsx";
aba = "PythonMetaTrader5";
aba2 = pd.read_excel(planilha, sheet_name="AtivosB3", engine='openpyxl');
dias = 1;

####################################################
#Pegando o código dos 85 ativos e do vencimento atual Índice Cheio
#Estes valores devem ser alterados pelo usuário na tabela

tabela_AtivosB3 = aba2.iloc[1:86, 0];
tabela_AtivosB3_cod_ind = aba2.iloc[1, 6];

####################################################
#Inicializando o MetaTrader 5

if not (mt5.initialize()):
    print("Problemas na conexão com o MetaTrader 5.\n");

####################################################
#Funções de loop

while (True):

    #Pegando valores do Índice com vencimento e IBOV e printando na tabela
    preco_IBOV = mt5.symbol_info_tick("IBOV");
    preco_IND = mt5.symbol_info_tick(tabela_AtivosB3_cod_ind);

    preco_IBOV_atual = preco_IBOV[3];
    preco_IND_atual = preco_IND[3];

    xw.Book(planilha).sheets[aba].range("N2").options(index=False).value = preco_IBOV_atual;
    xw.Book(planilha).sheets[aba].range("O2").options(index=False).value = preco_IND_atual;

    #Loop para preencher as tabelas conforme o código do ativo 
    for i in range(1, len(tabela_AtivosB3)+1):
        cod_ativo = tabela_AtivosB3[i];
        cont_while = 1;

        while (cont_while == 1):

            #Conferindo o horário e o TimeZone
            #hoje = time.time();
            #hoje = pd.to_datetime(hoje, unit='s');
            tz = get_localzone();
            now = datetime.utcnow();
            tzoffset = tz.utcoffset(now);
            hoje = (now + tzoffset);

            #Pegando valores de 1 candle anterior e candles atuais no TIMEFRAME Diário
            candle_anterior = mt5.copy_rates_from_pos(cod_ativo, mt5.TIMEFRAME_D1, 1, dias);
            fechamento_anterior = candle_anterior['close'];


            candle_atual = mt5.copy_rates_from(cod_ativo, mt5.TIMEFRAME_D1, hoje, dias);
            abertura_atual = candle_atual['open'];
            fechamento_atual = candle_atual['close'];


            preco_ativo_info = mt5.symbol_info_tick(cod_ativo);
            preco_ativo_BID = preco_ativo_info[1];
            preco_ativo_ASK = preco_ativo_info[2];


            #Printando informações dos ativos na tabela
            xw.Book(planilha).sheets[aba].range(f"A{i+1}").options(index=False).value = cod_ativo; 
            xw.Book(planilha).sheets[aba].range(f"B{i+1}").options(index=False).value = hoje; 
            xw.Book(planilha).sheets[aba].range(f"C{i+1}").options(index=False).value = fechamento_anterior;
            xw.Book(planilha).sheets[aba].range(f"D{i+1}").options(index=False).value = abertura_atual; 
            xw.Book(planilha).sheets[aba].range(f"E{i+1}").options(index=False).value = fechamento_atual;
            xw.Book(planilha).sheets[aba].range(f"F{i+1}").options(index=False).value = preco_ativo_BID; 
            xw.Book(planilha).sheets[aba].range(f"G{i+1}").options(index=False).value = preco_ativo_ASK;



            #Printando algumas informações extras no terminal
            subprocess.call(command,shell=True);
            print(cod_ativo);            
            print("Done.");
            print(hoje);
            print("TimeZone: ", tz, "\n");

            cont_while = 0;

        #fim While

    #fim FOR

#fim While
