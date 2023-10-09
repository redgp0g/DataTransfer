import logging
from contextlib import nullcontext
from datetime import datetime
from decimal import Decimal
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import xlwings as xw

log_directory = r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Registros PDF'

log_filename = 'log_relatorio_dimensional.log'

log_path = os.path.join(log_directory, log_filename)

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

pastas_planilhas = [r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Spacer',
                    r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Taurus',
                    r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Bosch',
                    r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Schaeffler']

pastas_monitoradas = [r'\\Sch-fns03a\ds1\Qualidade1\2021\05.Metrologia\2021\23 Cep Zeiss',
                        r'\\Sch-fns03a\ds1\Qualidade1\2021\05.Metrologia\2021\23 Cep Dea',]

def buscar_primeira_linha_cota(linhas):
    for i, linha in enumerate(linhas):
        if "Cota" in linha or "COTA" in linha:
            return linhas[i:]
    return None
                        
def encontrar_caminho_planilha(codigo):
    caminho_planilha = None
    for pasta in pastas_planilhas:
        for arquivo in os.listdir(pasta):
            if arquivo.endswith('.xlsm'):
                if codigo in arquivo and not '~' in arquivo:
                    caminho_planilha = os.path.join(pasta, arquivo)
    return caminho_planilha

def buscar_dados_zeiss(todas_linhas,palavras_chave):
    padrao = "M" + str(datetime.now().year)[2:]
    dados_medicao = None
    codigo = None
    for i, linha in enumerate(todas_linhas):
        for palavra in palavras_chave:
            if palavra in linha:
                dados_medicao = todas_linhas[i + 1].strip()
                codigo = dados_medicao[dados_medicao.index(padrao):][:10].strip()
                return codigo,dados_medicao

def ler_arquivo_txt_zeiss(caminho):
    try:
        with open(caminho, 'r', encoding='latin-1') as arquivo:
            todas_linhas = arquivo.readlines()
            
            codigo,dados_medicao = buscar_dados_zeiss(todas_linhas,["Plano Medição","ID Teste","Data","Comentario"])

            if(codigo != None):
                caminho_planilha = encontrar_caminho_planilha(codigo)
                if(caminho_planilha == None):
                    logging.warning("Planilha não encontrada!")
                else:
                    nomes_cotas = []
                    valores_encontrados = []
                    nominais = []
                    tolerancias_superiores = []
                    tolerancias_inferiores = []
                    desvios = []

                    cotas_linhas = buscar_primeira_linha_cota(todas_linhas)

                    for cota in cotas_linhas:
                        if cota and not cota[:25].isspace(): 
                            valor_encontrado = cota[35:46].strip()
                            nominal = cota[49:58].strip()
                            tolerancia_superior = cota[58:67].strip()
                            tolerancia_inferior = cota[67:76].strip()
                            desvio = cota[76:85].strip()
                            peca = dados_medicao[66:].strip()

                            if not tolerancia_superior:
                                tolerancia_superior = "0.000"

                            if not tolerancia_inferior:
                                tolerancia_inferior = "0.000"
                            
                            if not valor_encontrado:
                                nominal  = "0.000"
                                valor_encontrado = desvio

                            nomes_cotas.append(cota[:25].strip() + "_" + cota[25:35].strip())
                            valores_encontrados.append(valor_encontrado)
                            nominais.append(nominal)
                            tolerancias_superiores.append(tolerancia_superior)
                            tolerancias_inferiores.append(tolerancia_inferior)
                            desvios.append(desvio)

                    app = xw.App(visible=False)
                    workbook = app.books.open(caminho_planilha)
                    planilha = workbook.sheets.active
                    planilha.api.Unprotect()
                    rng = planilha.range("A10:A698")
                    rng.api.Rows.Hidden = False

                    primeira_celula_cota = None
                    for linha_celula in range(10, 388 + 1):                    
                        if planilha.range("A" + str(linha_celula)).color == (0, 255, 0):
                            primeira_celula_cota = planilha.range("A" + str(linha_celula)).address
                            break
                    if primeira_celula_cota == None:
                        primeira_celula_cota = planilha.range("A10").end("down").offset(row_offset=1).address
                    linha_celula = int(primeira_celula_cota.split('$')[2])

                    celula_numero_peca = planilha.range("B699")

                    while celula_numero_peca.column < 26:  
                        if celula_numero_peca.value == float(peca):
                            break
                        celula_numero_peca = celula_numero_peca.offset(column_offset=1)

                    coluna_celula_numero_peca = xw.utils.col_name(celula_numero_peca.column + 1)

                    for i, valor_encontrado in enumerate(valores_encontrados):
                        valor_encontrado = valores_encontrados[i]
                        nominal = nominais[i]
                        tolsup = tolerancias_superiores[i]
                        tolinf = tolerancias_inferiores[i]
                        desvio = desvios[i]
                        nome = nomes_cotas[i]
                            
                        value_cell = planilha.range(coluna_celula_numero_peca + str(linha_celula))

                        planilha.range("A" + str(linha_celula)).value = nome
                        planilha.range("A" + str(linha_celula)).color = (0, 255, 0)

                        if nominal == "0.000":
                            planilha.range("Z" + str(linha_celula)).value = Decimal(nominal)
                            planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                            planilha.range("AD" + str(linha_celula)).value = Decimal(nominal)
                            planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                        else:
                            planilha.range("Z" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolinf)
                            planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)

                            planilha.range("AD" + str(linha_celula)).value = Decimal(nominal) + (Decimal(tolinf) * Decimal(0.7))
                            planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + (Decimal(tolsup) * Decimal(0.7))

                        value_cell.value = valor_encontrado
                            
                        linha_celula = linha_celula + 1

                    primeira_celula_cota = planilha.range("A10").end("down").offset(row_offset=1)
                    next_last_cell = primeira_celula_cota.end("down").offset(row_offset=-1)
                    next_last_cell_position = next_last_cell.address
                    primeira_celula_cota = primeira_celula_cota.address
                    range_address = primeira_celula_cota.replace("$", "") + ":" + next_last_cell_position.replace("$", "") 
                    planilha.range(range_address).api.Rows.Hidden = True
                    planilha.api.Protect()
                    workbook.save()
                    workbook.close()
                    app.quit()
                    logging.info("Dados transferidos com sucesso!")
            else:
                logging.warning("Código não encontrado!")
    except Exception as e:
        logging.error(f"Ocorreu um erro em ler_arquivo_txt: {e}")
    finally:
        try:
            workbook.close()
            app.quit()
        except:
            pass

def buscar_dados_mea(todas_linhas):
    padrao = "M" + str(datetime.now().year)[2:]
    codigo = None
    peca = None
    for i, linha in enumerate(todas_linhas):
        if '%52' in linha:
            if padrao in linha:
                index_padrao = linha.index(padrao)
                codigo = linha[index_padrao:index_padrao + 9].strip()
                peca = linha[index_padrao + 9:].strip()
                break 
    return codigo, peca

def ler_arquivo_mea(caminho):
    try:
        with open(caminho, 'r', encoding='latin-1') as arquivo:
            todas_linhas = arquivo.readlines()
            
            codigo, peca = buscar_dados_mea(todas_linhas)       
                    
            if codigo != None:
                caminho_planilha = encontrar_caminho_planilha(codigo)
                if not caminho_planilha:
                    logging.warning("Planilha não encontrada!")
                else:         
                    cotas_linhas = buscar_primeira_linha_cota(todas_linhas)           

                    nomes_cotas = []
                    valores_encontrados = []
                    nominais = []
                    tolerancias_superiores = []
                    tolerancias_inferiores = []

                    for cota in cotas_linhas:
                            if cota and not cota[:25].isspace():
                                if not(Decimal(cota[56:68].strip()) == 0.000000 and Decimal(cota[56:68].strip()) == Decimal(cota[56:68].strip())):
                                    nomes_cotas.append(cota[:14].strip() + "_" + cota[14:16].strip())
                                    valores_encontrados.append(Decimal(cota[16:31].strip()))
                                    nominais.append(Decimal(cota[30:44].strip()))
                                    tolerancias_superiores.append(Decimal(cota[56:68].strip()))
                                    tolerancias_inferiores.append(Decimal(cota[43:56].strip()))

                    app = xw.App(visible=False)
                    workbook = app.books.open(caminho_planilha)
                    planilha = workbook.sheets.active
                    planilha.api.Unprotect()
                    rng = planilha.range("A10:A698")
                    rng.api.Rows.Hidden = False

                    primeira_celula_cota = nullcontext
                    for linha_celula in range(10, 388 + 1):
                        
                        cor_celula = planilha.range("A" + str(linha_celula)).color
                        
                        if cor_celula == (0, 255, 0):
                            primeira_celula_cota = planilha.range("A" + str(linha_celula))
                            break
                    if primeira_celula_cota == nullcontext:
                        primeira_celula_cota = planilha.range("A10").end("down").offset(row_offset=1)

                    primeira_celula_cota = primeira_celula_cota.address
                    linha_celula = int(primeira_celula_cota.split('$')[2])
                    
                    celula_numero_peca = planilha.range("B699")
                    while celula_numero_peca.column < 26:  # Coluna Z é a coluna de número 26
                        if celula_numero_peca.value == float(peca):
                            break
                        celula_numero_peca = celula_numero_peca.offset(column_offset=1)

                    coluna_celula_numero_peca = xw.utils.col_name(celula_numero_peca.column + 1)

                    for i, valor_encontrado in enumerate(valores_encontrados):
                            valor_encontrado = valores_encontrados[i]
                            nominal = nominais[i]
                            tolsup = tolerancias_superiores[i]
                            tolinf = tolerancias_inferiores[i]
                            nome = nomes_cotas[i]
                                
                            value_cell = planilha.range(coluna_celula_numero_peca + str(linha_celula))

                            planilha.range("A" + str(linha_celula)).value = nome
                            planilha.range("A" + str(linha_celula)).color = (0, 255, 0)

                            if nominal == "0.000":
                                planilha.range("Z" + str(linha_celula)).value = nominal
                                planilha.range("AA" + str(linha_celula)).value = nominal + tolsup
                                planilha.range("AD" + str(linha_celula)).value = nominal
                                planilha.range("AE" + str(linha_celula)).value = nominal + tolsup
                            else:
                                planilha.range("Z" + str(linha_celula)).value = nominal - tolinf
                                planilha.range("AA" + str(linha_celula)).value = nominal + tolsup

                                planilha.range("AD" + str(linha_celula)).value = nominal - (tolinf * Decimal(0.7))
                                planilha.range("AE" + str(linha_celula)).value = nominal + (tolsup * Decimal(0.7))

                            value_cell.value = str(valor_encontrado)
                                
                            linha_celula = linha_celula + 1

                    primeira_celula_cota = planilha.range("A10").end("down").offset(row_offset=1)
                    next_last_cell = primeira_celula_cota.end("down").offset(row_offset=-1)
                    next_last_cell_position = next_last_cell.address
                    primeira_celula_cota = primeira_celula_cota.address
                    range_address = primeira_celula_cota.replace("$", "") + ":" + next_last_cell_position.replace("$", "") 
                    planilha.range(range_address).api.Rows.Hidden = True
                    planilha.api.Protect()
                    workbook.save()
                    workbook.close()
                    app.quit()
                    logging.info("Dados transferidos com sucesso!")
            else:
                logging.warning("Código não encontrado!")
    except Exception as e:
        logging.error(f"Ocorreu um erro em ler_arquivo_mea: {e}")
    finally:
        try:
            workbook.close()
            app.quit()
        except:
            pass

class ArquivoHandler(FileSystemEventHandler):
    def process_file(self, file_path):
        time.sleep(1)
        if file_path.endswith('.txt'):
            ler_arquivo_txt_zeiss(file_path)
            arquivo_nome = os.path.basename(file_path)
            logging.info(f'Arquivo processado: {arquivo_nome}')
        elif file_path.endswith('.MEA'):
            ler_arquivo_mea(file_path)
            arquivo_nome = os.path.basename(file_path)
            logging.info(f'Arquivo processado: {arquivo_nome}')

    def on_created(self, event):
        if event.is_directory:
            return
        self.process_file(event.src_path)

if __name__ == "__main__":
    observer = Observer()
    event_handler = ArquivoHandler()

    for pasta in pastas_monitoradas:
        observer.schedule(event_handler, path=pasta, recursive=False)

    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()

    observer.join() 