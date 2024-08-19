import pyautogui
import cv2
import time
from datetime import datetime, timedelta
import pygetwindow as gw
import os
import psycopg2
import pandas as pd
import numpy as np
from io import StringIO
import pyperclip

start_time= datetime.now()
print("Inicio da execução:", start_time)

data_ontem = datetime.now() - timedelta(days=1)
data_ontem_formatada = data_ontem.strftime("%d%m%Y")
data_contrato = data_ontem.strftime("%d-%m-%Y")
# data_ontem_formatada = '02072024'
# data_contrato = '02-07-2024'

pyautogui.PAUSE = 0.7

######## ABRIR GOGLOBAL #########
pyautogui.press('win')
pyautogui.typewrite('appcontroller')
pyautogui.press('enter')

goglobal = pyautogui.locateOnScreen('1 - goglobal.PNG', confidence=0.8)
contador = 0
print("Estou aguardando o GoGlobal aparecer")
while goglobal is None:
    goglobal = pyautogui.locateOnScreen('1 - goglobal.PNG', confidence=0.8)
    time.sleep(0.5)
    contador += 1
    
    if contador >= 12000:
        informError(i)

pyautogui.press('enter')
  
######## PRIMEIRO LOGIN #########
login = pyautogui.locateOnScreen('2 - login.PNG', confidence=0.8)
contador = 0
print("Estou aguardando a primeira tela de login aparecer")
while login is None:
    login = pyautogui.locateOnScreen('2 - login.PNG', confidence=0.8)
    time.sleep(0.5)
    contador += 1
    if contador >= 12000:
        informError(i)

pyautogui.typewrite('gabriela.lopes')
pyautogui.press('tab')
pyautogui.typewrite('********')
pyautogui.press('enter')

######## ESPERAR ICONE DE PRODUÇÃO APARECER E CLICAR NELE ########
entrada = pyautogui.locateOnScreen('3 - entradaproducao.PNG', confidence=0.5)
contador = 0
print("Estou aguardando o ícone de produção aparecer")
while entrada is None:
    entrada = pyautogui.locateOnScreen('3 - entradaproducao.PNG', confidence=0.5)
    time.sleep(0.1)
    contador += 1
    if contador >= 12000:
        informError(i)

pyautogui.press('right')
pyautogui.press('right')
pyautogui.press('enter')

######## SEGUNDO LOGIN #########
login2 = pyautogui.locateOnScreen('5 - logindatasul.PNG', confidence=0.5)
contador = 0
print("Estou aguardando a tela de login do DataSul")
while login2 is None:
    login2 = pyautogui.locateOnScreen('5 - logindatasul.PNG', confidence=0.5)
    time.sleep(0.1)
    contador += 1
    if contador >= 12000:
        informError(i)

pyautogui.typewrite('01-cardosog')
pyautogui.press('tab')
pyautogui.typewrite('*********')
pyautogui.press('enter')
time.sleep(10)

######## ATIVAR DI #########
pyautogui.locateCenterOnScreen('8 - abrirDI.PNG', confidence=0.5)
pyautogui.press('right')
pyautogui.press('enter')
time.sleep(10)

######## ALTERNAR PARA DATASUL INTERACTIVE E PROCURAR PA3040 #########
atTELAS = gw.getWindowsWithTitle('DATASUL Interactive')[0]
time.sleep(1)
atTELAS.restore()
atTELAS.activate()
time.sleep(1)
pyautogui.hotkey('ctrl', 'alt', 'x')
time.sleep(1)
pyautogui.typewrite('PA3040')
pyautogui.press('enter')

######## ESPERAR TELA DO PA3040 APARECER #########
telaPA3040 = pyautogui.locateOnScreen('9 - telaPA3040.PNG', confidence=0.5)
contador = 0
print("Estou aguardando a tela do PA3040")
while telaPA3040 is None:   
    telaPA3040 = pyautogui.locateOnScreen('9 - telaPA3040.PNG', confidence=0.5)
    time.sleep(0.1)
    contador += 1
    if contador >= 12000:
        informError(i)

######## SELECIONAR AS REGIONAIS #########
pyautogui.press('enter')
for _ in range(4):
    pyautogui.press('tab')

pyautogui.press('enter')

for _ in range(2):
    pyautogui.press('tab')
pyautogui.press('enter')

######## SELECIONAR AS UR's #########
pyautogui.press('tab')
pyautogui.press('enter')
for _ in range(4):
    pyautogui.press('tab')
pyautogui.press('enter')

time.sleep(2)

for _ in range(2):
    pyautogui.press('tab')
pyautogui.press('enter')

######## INCLUIR PERIODO #########
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1) 
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
    

######## INCLUIR FAMILIA SOJA #########
pyautogui.press('tab')
pyautogui.typewrite('2040000')
pyautogui.press('tab')
pyautogui.typewrite('2049999')
time.sleep(5)
for _ in range(5):
    pyautogui.press('tab')

######## INCLUIR DATA DE PAGAMENTO #########

pyautogui.typewrite('01011000', interval=0.1)

######## EXECUTAR A JANELA #########
executar = pyautogui.locateCenterOnScreen('12 - clickEXECUTAR.PNG', confidence=0.7)
pyautogui.moveTo(executar.x, executar.y)
pyautogui.click(executar.x, executar.y)

######## ESPERAR TXT ABRIR #########
time.sleep(5)
notepadWindows = gw.getWindowsWithTitle('pa3040.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('pa3040.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(2)
pyperclip.copy('')
# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

# Verifique se o conteúdo copiado está vazio (arquivo vazio)
if not conteudo_copiado:
    print("O arquivo de texto está vazio. Não há dados para processar no PA3040 Soja.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\PA3040\SOJA'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-PA3040SOJA.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)

    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )

    ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_pa3040soja = [(0, 13),(14, 26), (27, 37), (38, 82), (83, 97), (98, 110), (111, 117), (112, 130), (131, 146), (147, 162), (163, 173), (174, 182), (183, 196), (197, 205), (206, 217), (218, 223), (224, 230)]

    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    df3040soja = pd.read_fwf(caminho_arquivo, colspecs=colunas_pa3040soja, skiprows=10, header=None)

    # Renomeie as colunas
    df3040soja.columns = ['Regional', 'Local Recebimento', 'ITEM', 'Descrição', 'Safra Composta', 'Dt Pagto', 'Refer.', 'Renda Liquida', 'Valor Bruto', 'Valor líquido', 'Preço unitário','Valor Tabela', 'Valor Cond. Especial', 'Valor GMOFree', 'Valor Ajuda Frete', 'Outros Valores', 'Tp. Contrato']

    # Remova ou substitua o caractere inválido (Form Feed) dos dados
    df3040soja = df3040soja.apply(lambda x: x.str.replace('\x0c', '', regex=True))

    # Encontre os índices das linhas onde 'Regional' contém 'Total Fixação' E 'Local Recebimento' contém 'Normal'
    indices = df3040soja[(df3040soja['Regional'].str.contains('Total Fixação', case=False, regex=True, na=False)) & (df3040soja['Local Recebimento'] == 'Normal')].index

    # Criar uma nova coluna 'Tipo Fixação' com valores 'Normal' e 'Contrato'
    df3040soja['Tipo Fixação'] = 'Normal'  # Inicialmente, definimos todos os valores como 'Normal'

    # Verifique se o índice está dentro dos índices onde 'Regional' contém 'Total Fixação'
    df3040soja['Tipo Fixação'] = np.where(df3040soja.index <= indices.max(), 'FN', 'Contrato')

    df3040soja['Data do contrato'] = data_contrato

    # Valores que você deseja filtrar na coluna 'ITEM' (com 'AF' e 'af')
    valores_filtrados = ['2041001AF', '2041002AF', '2041009AF', '2041010AF', '2041011AF', '2041012AF', '2041013AF', '2041014AF', '2041015AF', '2041016AF',
                        '2041001af', '2041002af', '2041009af', '2041010af', '2041011af', '2041012af', '2041013af', '2041014af', '2041015af', '2041016af']

    df_filtradosoja = df3040soja[df3040soja['ITEM'].isin(valores_filtrados)]
    df_filtradosoja.to_excel(f'{caminho_salvar}\{data_contrato}-PA3040SOJA.xlsx',
    index=False, 
    sheet_name='PA3040SOJA', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradosoja.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.pa3040soja FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()


# Fechar a janela de notepad
notepadWindow.close()

# Fechar o programa
gw.getWindowsWithTitle('Relatório Sintético de Fixação - pa3040')[0].close()




####### COMEÇAR MILHO AGORA #########


######## ALTERNAR PARA DATASUL INTERACTIVE E PROCURAR PA3040 #########
atTELAS = gw.getWindowsWithTitle('DATASUL Interactive')[0]
time.sleep(1)
atTELAS.restore()
atTELAS.activate()
time.sleep(1)
pyautogui.hotkey('ctrl', 'alt', 'x')
time.sleep(1)
pyautogui.typewrite('PA3040')
pyautogui.press('enter')
######## ESPERAR TELA DO PA3040 APARECER #########
telaPA3040 = pyautogui.locateOnScreen('9 - telaPA3040.PNG', confidence=0.5)
contador = 0
print("Estou aguardando a tela do PA3040")
while telaPA3040 is None:   
    telaPA3040 = pyautogui.locateOnScreen('9 - telaPA3040.PNG', confidence=0.5)
    time.sleep(0.1)
    contador += 1
    if contador >= 12000:
        informError(i)

######## SELECIONAR AS REGIONAIS #########
pyautogui.press('enter')
for _ in range(4):
    pyautogui.press('tab')

pyautogui.press('enter')

for _ in range(2):
    pyautogui.press('tab')
pyautogui.press('enter')

######## SELECIONAR AS UR's #########
pyautogui.press('tab')
pyautogui.press('enter')
for _ in range(4):
    pyautogui.press('tab')
pyautogui.press('enter')

time.sleep(2)

for _ in range(2):
    pyautogui.press('tab')
pyautogui.press('enter')

######## INCLUIR PERIODO #########
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)


######## INCLUIR FAMILIA MILHO #########
pyautogui.press('tab')
pyautogui.typewrite('2010000')
pyautogui.press('tab')
pyautogui.typewrite('2019999')
time.sleep(5)
for _ in range(5):
    pyautogui.press('tab')

######## INCLUIR DATA DE PAGAMENTO #########
pyautogui.typewrite('01011000', interval=0.1)


######## EXECUTAR A JANELA #########
executar = pyautogui.locateCenterOnScreen('12 - clickEXECUTAR.PNG', confidence=0.7)
pyautogui.moveTo(executar.x, executar.y)
pyautogui.click(executar.x, executar.y)

######## ESPERAR TXT ABRIR #########
time.sleep(5)
notepadWindows = gw.getWindowsWithTitle('pa3040.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('pa3040.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(2)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar (você pode ajustar o tempo conforme necessário)
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:
    print("O arquivo de texto está vazio. Não há dados para processar no PA3040 Milho.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\PA3040\MILHO'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-PA3040MILHO.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)
    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )

    ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_pa3040milho = [(0, 13), (14, 26), (27, 37), (38, 83), (84, 99), (100, 110), (111, 117), (112, 130), (131, 146), (147, 162), (163, 171), (172, 180), (181, 194), (195, 203), (204, 215), (216, 223), (224, 229)]
    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfmilho = pd.read_fwf(caminho_arquivo, colspecs=colunas_pa3040milho, skiprows=10, header=None)

    # Renomeie as colunas
    dfmilho.columns = ['Regional', 'Local Recebimento', 'ITEM', 'Descrição', 'Safra Composta', 'Dt Pagto', 'Refer.', 'Renda Liquida', 'Valor Bruto', 'Valor líquido', 'Preço unitário','Valor Tabela', 'Valor Cond. Especial', 'Valor GMOFree', 'Valor Ajuda Frete', 'Outros Valores', 'Tp. Contrato']

    # Remova ou substitua o caractere inválido (Form Feed) dos dados
    dfmilho = dfmilho.apply(lambda x: x.str.replace('\x0c', '', regex=True))

    # Encontre os índices das linhas onde 'Regional' contém 'Total Fixação' E 'Local Recebimento' contém 'Normal'
    indices = dfmilho[(dfmilho['Regional'].str.contains('Total Fixação', case=False, regex=True, na=False)) & (dfmilho['Local Recebimento'] == 'Normal')].index

    # Criar uma nova coluna 'Tipo Fixação' com valores 'Normal' e 'Contrato'
    dfmilho['Tipo Fixação'] = 'Normal'  # Inicialmente, definimos todos os valores como 'Normal'

    # Verifique se o índice está dentro dos índices onde 'Regional' contém 'Total Fixação'
    dfmilho['Tipo Fixação'] = np.where(dfmilho.index <= indices.max(), 'FN', 'Contrato')

    dfmilho['Data do contrato'] = data_contrato

    # Valores que você deseja filtrar na coluna 'ITEM' (com 'AF' e 'af')  

    valores_filtrados = ['2010001AF', '2010002AF', '2010002Af', '2010002aF', '2010002af','2010001Af','2010001aF','2010001af','2010007AF', '2010007Af', '2010007aF',
                        '2010007af','2010011AF', '2010011Af', '2010011aF', '2010011af','2010012AF', '2010012Af', '2010012aF', '2010012af']

    df_filtradomilho = dfmilho[dfmilho['ITEM'].isin(valores_filtrados)]
    df_filtradomilho.to_excel(f'{caminho_salvar}\{data_contrato}-PA3040MILHO.xlsx',
    index=False, 
    sheet_name='PA3040MILHO', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradomilho.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.pa3040milho FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()



notepadWindow.close()

# Fechar o programa
gw.getWindowsWithTitle('Relatório Sintético de Fixação - pa3040')[0].close()



####### PA3044 MILHO PERMUTA #########

######## ALTERNAR PARA DATASUL INTERACTIVE E PROCURAR PA3044 #########
atTELAS = gw.getWindowsWithTitle('DATASUL Interactive')[0]
time.sleep(1)
atTELAS.restore()
atTELAS.activate()
time.sleep(5)
pyautogui.hotkey('ctrl', 'alt', 'x')
time.sleep(1)
pyautogui.typewrite('PA3044')
pyautogui.press('enter')
time.sleep(2)

######## ESPERAR TELA DO PA3044 APARECER #########
telaPA3044 = pyautogui.locateOnScreen('14 - telaPA3044.PNG', confidence=0.5)
contador = 0
print("Estou aguardando a tela do PA3044")
while telaPA3044 is None:   
    telaPA3044 = pyautogui.locateOnScreen('14 - telapa3044.PNG', confidence=0.5)
    time.sleep(0.1)
    contador += 1
    if contador >= 12000:
        informError(i)

#Apertar tab 6x
for _ in range(6):
    pyautogui.press('tab')

#Incluir safras
pyautogui.typewrite('2000/2000')
pyautogui.press('tab')
pyautogui.typewrite('2050/2050')

#Incluir familia
pyautogui.press('tab')
pyautogui.typewrite('2010000')
pyautogui.press('tab')
pyautogui.typewrite('2019999')

#Incluir data do contrato
pyautogui.press('tab')

for dt in data_ontem_formatada:
    pyautogui.typewrite(data_ontem_formatada, interval=0.1)
    time.sleep(0.1)

pyautogui.press('tab')

for dt in data_ontem_formatada:
    pyautogui.typewrite(data_ontem_formatada, interval=0.1)
    time.sleep(0.1)


par3044 = pyautogui.locateCenterOnScreen('15 - parametros3044.PNG', confidence=0.9)
pyautogui.moveTo(par3044.x, par3044.y)
pyautogui.click(par3044.x, par3044.y)

time.sleep(2)

pyautogui.press('down')

canc = pyautogui.locateCenterOnScreen('16 - cancelados.PNG', confidence=0.7)
pyautogui.moveTo(canc.x, canc.y)
pyautogui.click(canc.x, canc.y)

######## EXECUTAR A JANELA #########
executar = pyautogui.locateCenterOnScreen('12 - clickEXECUTAR.PNG', confidence=0.7)
pyautogui.moveTo(executar.x, executar.y)
pyautogui.click(executar.x, executar.y)

######## ESPERAR TXT ABRIR #########
time.sleep(5)
notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(2)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:
    print("O arquivo de texto está vazio. Não há dados para processar no PA3044 Milho.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\PA3044\MILHO'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-PA3044MILHOper.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)

    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )

    ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_pa3044milhoPER = [(0, 8), (9, 49), (50, 53), (54, 64), (65, 73), (74, 83), (84, 96), (97, 108), (109, 119),
                             (120, 130), (131, 146), (147, 157), (158, 171), (172, 185), (186, 202), (203, 216), (217, 222),
                             (223, 233)]
    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfmilho3044PER = pd.read_fwf(caminho_arquivo, colspecs=colunas_pa3044milhoPER, header=None)

    # Renomeie as colunas
    dfmilho3044PER.columns = ['Produtor', 'Nome', 'Est.Ent', 'Item', 'Contrato', 'Operação', 'Quant.(Kg) Contrato',
                             'Dt.Contrato', 'Dt.Vencimento', 'Prazo Entr', 'Valor Negociad.Liq','Vl Sc.Brut',
                             'Quant.(Kg) Liquidada', 'Quant.(Kg) Saldo', 'Vl.Liquidado', 'Quant.(Kg) Inadimplente',
                             'Verif','Situação']

    dfmilho3044PER['Item'] = dfmilho3044PER['Item'].str.strip()

    valores_filtrados = [f'{i}AF' for i in range(2010000, 2050000)] + [f'{i}Af' for i in range(2010000, 2050000)] + [f'{i}aF' for i in range(2010000, 2050000)] + [f'{i}af' for i in range(2010000, 2050000)]


    df_filtradomilho3044PER = dfmilho3044PER[dfmilho3044PER['Item'].isin(valores_filtrados)]
    df_filtradomilho3044PER.to_excel(f'{caminho_salvar}\{data_contrato}-PA3044MILHOper.xlsx',
    index=False, 
    sheet_name='PA3044MILHOper', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradomilho3044PER.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.pa3044milhoper FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()
    
notepadWindow.close()

####### PA3044 MILHO PREÇO FIXO REAL #######

# Voltar para tela PA3044
atTELAS2 = gw.getWindowsWithTitle('Relatório de Contratos - PA3044\xa0')[0]
time.sleep(1)
atTELAS2.restore()
time.sleep(1)
atTELAS2.activate()
time.sleep(1)

botao = pyautogui.locateCenterOnScreen('18 - botaoseta.PNG', confidence=0.6)
pyautogui.moveTo(botao.x, botao.y)
pyautogui.click()
#Selecionar preço fixo real
for i in range(2):
    pyautogui.press('down')

time.sleep(1)
pyautogui.press('enter')

######## EXECUTAR A JANELA #########
executar2 = pyautogui.locateCenterOnScreen('19 - executar2.PNG', confidence=0.7)
pyautogui.moveTo(executar2.x, executar2.y)
pyautogui.click(executar2.x, executar2.y)

executar = pyautogui.locateCenterOnScreen('12 - clickEXECUTAR.PNG', confidence=0.7)
pyautogui.moveTo(executar.x, executar.y)
pyautogui.click(executar.x, executar.y)

######## ESPERAR TXT ABRIR #########
time.sleep(5)
notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(5)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:

    print("O arquivo de texto está vazio. Não há dados para processar no PA3044 Milho.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\PA3044\MILHO'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-PA3044MILHOpf.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)
    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )
        ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_pa3044milhoPF= [(0, 8), (9, 49), (50, 53), (54, 64), (65, 73), (74, 83), (84, 94), (95, 106), (107, 118),
                             (120, 133), (134, 146), (147, 157), (158, 171), (172, 185), (186, 202), (203, 214), (215, 220),
                             (221, 230), (231, 240)]

    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfmilho3044PF = pd.read_fwf(caminho_arquivo, colspecs=colunas_pa3044milhoPF, header=None)

    # Renomeie as colunas
    dfmilho3044PF.columns = ['Produtor', 'Nome', 'Est.Ent', 'Item', 'Refer', 'Contrato', 'Quantd. (Kg) Contrato', 'Dt. Contrato', 'Dt. Pagto',
                             'Vl. Contrato', 'Prazo Entrega', 'Preço Fixo', 'Quant. (Kg) Liquidada', 'Quant. (Kg) Saldo', 'Vl. Liquidado',
                             'Quant. (Kg) Inadimplente', 'Verif', 'Situação', '']

    dfmilho3044PF['Item'] = dfmilho3044PF['Item'].str.strip()

    valores_filtrados = [f'{i}AF' for i in range(2010000, 2050000)] + [f'{i}Af' for i in range(2010000, 2050000)] + [f'{i}aF' for i in range(2010000, 2050000)] + [f'{i}af' for i in range(2010000, 2050000)]


    df_filtradomilho3044PF = dfmilho3044PF[dfmilho3044PF['Item'].isin(valores_filtrados)]
    df_filtradomilho3044PF.to_excel(f'{caminho_salvar}\{data_contrato}-PA3044MILHOpf.xlsx',
    index=False, 
    sheet_name='PA3044MILHOpf', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradomilho3044PF.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.pa3044milhopf FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()

notepadWindow.close()

######## PREÇO FIXO REAL SOJA #########
time.sleep(2)
atTELAS2.restore()
time.sleep(1)
atTELAS2.activate()
time.sleep(1)
selecao = pyautogui.locateCenterOnScreen('17 - selecao3044.PNG', confidence=0.5)
pyautogui.moveTo(selecao.x, selecao.y)
pyautogui.click(selecao.x, selecao.y)

#Apertar tab 8x
for _ in range(8):
    pyautogui.press('tab')

pyautogui.typewrite('2040000')
pyautogui.press('tab')
pyautogui.typewrite('2049999')

######## EXECUTAR A JANELA #########
executar = pyautogui.locateCenterOnScreen('12 - clickEXECUTAR.PNG', confidence=0.7)
pyautogui.moveTo(executar.x, executar.y)
pyautogui.click(executar.x, executar.y)

executar2 = pyautogui.locateCenterOnScreen('19 - executar2.PNG', confidence=0.7)
pyautogui.moveTo(executar2.x, executar2.y)
pyautogui.click(executar2.x, executar2.y)

######## ESPERAR TXT ABRIR #########
time.sleep(5)
notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(5)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:

    print("O arquivo de texto está vazio. Não há dados para processar no PA3044 Soja.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\PA3044\SOJA'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-PA3044SOJApf.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)
    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )
        ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_pa3044sojaPF = [(0, 8), (9, 49), (50, 53), (54, 64), (65, 73), (74, 83), (84, 94), (95, 106), (107, 118),
                             (120, 133), (134, 146), (147, 157), (158, 171), (172, 185), (186, 202), (203, 214), (215, 220),
                             (221, 230), (231, 240)]
    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfsoja3044PF = pd.read_fwf(caminho_arquivo, colspecs=colunas_pa3044sojaPF, header=None)

    # Renomeie as colunas
    dfsoja3044PF.columns = ['Produtor', 'Nome', 'Est.Ent', 'Item', 'Refer', 'Contrato', 'Quantd. (Kg) Contrato', 'Dt. Contrato', 'Dt. Pagto',
                             'Vl. Contrato', 'Prazo Entrega', 'Preço Fixo', 'Quant. (Kg) Liquidada', 'Quant. (Kg) Saldo', 'Vl. Liquidado',
                             'Quant. (Kg) Inadimplente', 'Verif', 'Situação', '']

    dfsoja3044PF['Item'] = dfsoja3044PF['Item'].str.strip()

    valores_filtrados = [f'{i}AF' for i in range(2010000, 2050000)] + [f'{i}Af' for i in range(2010000, 2050000)] + [f'{i}aF' for i in range(2010000, 2050000)] + [f'{i}af' for i in range(2010000, 2050000)]


    df_filtradosoja3044PF = dfsoja3044PF[dfsoja3044PF['Item'].isin(valores_filtrados)]
    df_filtradosoja3044PF.to_excel(f'{caminho_salvar}\{data_contrato}-PA3044SOJApf.xlsx',
    index=False, 
    sheet_name='PA3044SOJApf', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradosoja3044PF.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.pa3044sojapf FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()

notepadWindow.close()

######## PERMUTA SOJA #########
time.sleep(2)
atTELAS2.restore()
time.sleep(1)
atTELAS2.activate()
time.sleep(1)

pyautogui.moveTo(par3044.x, par3044.y)
pyautogui.click(par3044.x, par3044.y)

pyautogui.press('up')
pyautogui.press('up')

pyautogui.moveTo(executar.x, executar.y)
pyautogui.click(executar.x, executar.y)

time.sleep(5)
notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('pa3044.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(5)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:

    print("O arquivo de texto está vazio. Não há dados para processar no PA3044 Soja.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\PA3044\SOJA'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-PA3044SOJAper.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)
    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )
        ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_pa3044sojaPER = [(0, 8), (9, 49), (50, 53), (54, 64), (65, 73), (74, 83), (84, 94), (95, 106), (107, 118),
                             (120, 133), (134, 146), (147, 157), (158, 171), (172, 185), (186, 202), (203, 214), (215, 220),
                             (221, 230), (231, 240)]
    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfsoja3044PER = pd.read_fwf(caminho_arquivo, colspecs=colunas_pa3044sojaPER, header=None)

    # Renomeie as colunas
    dfsoja3044PER.columns = ['Produtor', 'Nome', 'Est.Ent', 'Item', 'Refer', 'Contrato', 'Quantd. (Kg) Contrato', 'Dt. Contrato', 'Dt. Pagto',
                             'Vl. Contrato', 'Prazo Entrega', 'Preço Fixo', 'Quant. (Kg) Liquidada', 'Quant. (Kg) Saldo', 'Vl. Liquidado',
                             'Quant. (Kg) Inadimplente', 'Verif', 'Situação', '']

    dfsoja3044PER['Item'] = dfsoja3044PER['Item'].str.strip()

    valores_filtrados = [f'{i}AF' for i in range(2010000, 2050000)] + [f'{i}Af' for i in range(2010000, 2050000)] + [f'{i}aF' for i in range(2010000, 2050000)] + [f'{i}af' for i in range(2010000, 2050000)]


    df_filtradosoja3044PER = dfsoja3044PER[dfsoja3044PER['Item'].isin(valores_filtrados)]
    df_filtradosoja3044PER.to_excel(f'{caminho_salvar}\{data_contrato}-PA3044SOJAper.xlsx',
    index=False, 
    sheet_name='PA3044SOJAper', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradosoja3044PER.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.pa3044sojaper FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()

notepadWindow.close()
time.sleep(2)
atTELAS2.close()

######## COMEÇAR VG3018 #########
atTELAS = gw.getWindowsWithTitle('DATASUL Interactive')[0]
time.sleep(1)
atTELAS.restore()
atTELAS.activate()
time.sleep(1)
pyautogui.hotkey('ctrl', 'alt', 'x')
time.sleep(1)
pyautogui.typewrite('VG3018')
pyautogui.press('enter')


######## Família VG3018 #########
time.sleep(2)
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.typewrite('20101,20110,20120,20401,20410')

#Entrar na sessão parametros
time.sleep(2)
par3018 = pyautogui.locateCenterOnScreen('20 - parametros3018.PNG', confidence=0.7)
pyautogui.moveTo(par3018.x, par3018.y)
pyautogui.click(par3018.x, par3018.y)  
pyautogui.press('down')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.press('up')
pyautogui.press('up')
pyautogui.press('tab')
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.typewrite('01012000', interval=0.1)

codCliente = pyautogui.locateCenterOnScreen('21 - codCliente.PNG', confidence=0.7)
pyautogui.moveTo(codCliente.x, codCliente.y)
pyautogui.click(codCliente.x, codCliente.y)
pyautogui.press('enter')

notepadWindows = gw.getWindowsWithTitle('VG3018.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('VG3018.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(5)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:

    print("O arquivo de texto está vazio. Não há dados para processar no VG3018.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\VG3018'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-VG3018.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)
        
    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )
        ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_VG3018 = [(0, 9), (10,35), (36,44), (45,67), (68,83), (84,99), (100,115), (116,131),
                    (132,147), (148,163), (164,179),  (180,183), (184,189), (190,205)]
    
    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo

    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfVG3018 = pd.read_fwf(caminho_arquivo, colspecs=colunas_VG3018, header=None)
    dfVG3018.columns = ['codigo', 'cliente', 'contrato', 'produto', 'Qtcontrato', 'TotalContrato',
               'TotalFaturado', 'icms', 'QtEmbarcada', 'VlrEmbarcado',
               'ICMEmbarcado', 'Tip', 'SEmb', 'modelo']
    
    filtrosVG3018 = ['MILHO TIPO 2','SOJA GMO']
    df_filtradoVG3018 = dfVG3018[dfVG3018['produto'].isin(filtrosVG3018)]
    df_filtradoVG3018
    df_filtradoVG3018.to_excel(f'{caminho_salvar}\{data_contrato}-VG3018.xlsx',
    index=False, 
    sheet_name='VG3018', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradoVG3018.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.vg3018 FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()

notepadWindow.close()

vg3018 = gw.getWindowsWithTitle('Relatório de Vendas -  VG3018\xa0')[0]
vg3018.close()

####### COMEÇAR VG3011 #########
#Alternar para Interactive DI e procurar por VG3011
atTELAS = gw.getWindowsWithTitle('DATASUL Interactive')[0]
time.sleep(1)
atTELAS.restore()
time.sleep(1)
atTELAS.activate()
time.sleep(1)
pyautogui.hotkey('ctrl', 'alt', 'x')
time.sleep(1)
pyautogui.typewrite('VG3011')
pyautogui.press('enter')

#Informar parametros do VG3011
time.sleep(1)
for _ in range(4):
    pyautogui.press('tab')

pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite(data_ontem_formatada, interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite('2000', interval=0.1)
pyautogui.press('divide')
pyautogui.typewrite('2000', interval=0.1)
pyautogui.press('tab')
pyautogui.typewrite('2050', interval=0.1)
pyautogui.press('divide')
pyautogui.typewrite('2050', interval=0.1)
pyautogui.press('tab')
pyautogui.press('enter')
time.sleep(5)
item201 = pyautogui.locateOnScreen('22 - 201.PNG', confidence=0.8)
pyautogui.click(item201)
pyautogui.keyDown('shift')
item204 = pyautogui.locateOnScreen('23 - 204.PNG', confidence=0.9)
pyautogui.click(item204)
pyautogui.keyUp('shift')
pyautogui.press('enter')
time.sleep(5)
clickOk = pyautogui.locateOnScreen('11 - clickOK.PNG', confidence=0.7)
pyautogui.click(clickOk)
time.sleep(5)
executar = pyautogui.locateOnScreen('12 - clickEXECUTAR.PNG', confidence=0.7)
pyautogui.click(executar)

notepadWindows = gw.getWindowsWithTitle('.tmp - Bloco de notas')

while True:
    notepadWindows = gw.getWindowsWithTitle('.tmp - Bloco de notas')
    if not notepadWindows:
        time.sleep(10)
        print('Aguardando notepad abrir')
    else:
        notepadWindow = notepadWindows[0]
        print('Notepad aberto')
        break

time.sleep(5)
pyperclip.copy('')

# Selecione todo o conteúdo (Ctrl+A)
pyautogui.hotkey('ctrl', 'a')

# Aguarde um momento antes de copiar
time.sleep(2)

# Copie o conteúdo selecionado (Ctrl+C)
pyautogui.hotkey('ctrl', 'c')

# Aqui você pode colar o conteúdo da área de transferência em uma variável
conteudo_copiado = pyperclip.paste()

if not conteudo_copiado:

    print("O arquivo de texto está vazio. Não há dados para processar no VG3018.")
else:
    # Caminho onde você deseja salvar o arquivo com o mesmo conteúdo
    caminho_salvar = r'P:\caixinha caixinha\VG3011'
    nome_arquivo = os.path.join(caminho_salvar, data_contrato + '-VG3011.txt')
    # Crie um novo arquivo no destino com o conteúdo copiado
    with open(nome_arquivo, 'w', encoding='utf-8') as arquivo_destino:
        arquivo_destino.write(conteudo_copiado)
        
    ###### CONEXÃO COM O BANCO #########
    # Configurar a conexão com o banco de dados PostgreSQL
    conexao = psycopg2.connect(
        host="192.168.1.30",
        database="postgres",
        user="postgres",
        password="654876"
    )
        ####### LER E TRATAR O ARQUIVO #########
    # Defina as larguras das colunas com base no padrão da linha
    colunas_VG3011 = [(0, 10), (11, 33), (34, 49), (50, 60), (61, 69), (70, 79), (80, 99),
                        (100, 108), (109, 119), (120, 134), (135, 153), (154, 172)]
    
    # Caminho para o arquivo de texto
    caminho_arquivo = nome_arquivo
    # Carregue o arquivo de texto em um DataFrame usando as larguras definidas
    dfVG3011 = pd.read_fwf(caminho_arquivo, colspecs=colunas_VG3011, header=None, encoding='latin-1')
    dfVG3011.columns = [
        "Nr.Cont.",
        "Nome Abreviado Cliente",
        "Nr.Cont.Rep",
        "Nr.Fixa",
        "Data",
        "Data Rep.",
        "Qt.Fixada (kg)",
        "Preco kg",
        "Preco Saca",
        "Preço Tonelada",
        "Vl.Total Fixado",
        "PTAX"
    ]

    dfVG3011['Nr.Cont.']= pd.to_numeric(dfVG3011['Nr.Cont.'], errors='coerce')
    df_filtradosVG3011 = dfVG3011[dfVG3011['Nr.Cont.'].notna()]
    df_filtradosVG3011.to_excel(f'{caminho_salvar}\{data_contrato}-VG3011.xlsx',
    index=False, 
    sheet_name='VG3011', 
    engine='openpyxl')

    # Converter o DataFrame tratado para um arquivo CSV no formato adequado para o PostgreSQL
    csv_buffer = StringIO()
    df_filtradosVG3011.to_csv(csv_buffer, sep=';', index=False, header=False, quotechar='"')
    csv_buffer.seek(0)

    # Abrir um cursor e executar o comando COPY para inserir os dados no PostgreSQL
    cursor = conexao.cursor()
    cursor.copy_expert("COPY caixinha.vg3011 FROM stdin WITH CSV DELIMITER as ';'", csv_buffer)
    conexao.commit()

    # Fechar a conexão e o cursor
    cursor.close()
    conexao.close()

notepadWindow.close()

end_time = datetime.now()
print("Fim da execução:", end_time)

duration = end_time - start_time
print("Duração da execução:", duration)
