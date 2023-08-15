from contextlib import nullcontext
from decimal import Decimal
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import xlwings as xw

pastas_monitoradas = [r'\\Sch-fns03a\ds1\Qualidade1\2021\05.Metrologia\2021\23 Cep Zeiss',
                      r'\\Sch-fns03a\ds1\Qualidade1\2021\05.Metrologia\2021\23 Cep Dea']

def ler_arquivo_txt_zeiss(caminho):
    try:
        with open(caminho, 'r', encoding='latin-1') as arquivo:
            todas_linhas = arquivo.readlines()
            
            for i, linha in enumerate(todas_linhas):
                if "Plano Medição" in linha or "ID Teste" in linha or "Data" in linha or "Comentario" in linha:
                    if i + 1 < len(todas_linhas):
                        dados_medicao = todas_linhas[i + 1].strip()
                        break

            codigo = dados_medicao[55:66].strip()
            if(codigo != ""):
                planilha = encontrar_caminho_planilha(codigo)
                if(planilha == None):
                    print("Planilha não encontrada")
                else:
                    nomes_cotas = []
                    valores_encontrados = []
                    nominais = []
                    tolerancias_superiores = []
                    tolerancias_inferiores = []
                    desvios = []

                    for i, linha in enumerate(todas_linhas):
                        if "Cota" in linha or "COTA" in linha:
                            cotas_linhas = todas_linhas[i:]
                            break

                    for info_cotas in cotas_linhas:
                        if info_cotas and not info_cotas[:25].isspace(): 
                            valor_encontrado = info_cotas[35:46].strip()
                            nominal = info_cotas[49:58].strip()
                            tolerancia_superior = info_cotas[58:67].strip()
                            tolerancia_inferior = info_cotas[67:76].strip()
                            desvio = info_cotas[76:85].strip()
                            peca = dados_medicao[66:].strip()

                            if not tolerancia_superior:
                                tolerancia_superior = "0.000"

                            if not tolerancia_inferior:
                                tolerancia_inferior = "0.000"
                            
                            if not valor_encontrado:
                                nominal  = "0.000"
                                valor_encontrado = desvio

                            nomes_cotas.append(info_cotas[:25].strip() + "_" + info_cotas[25:35].strip())
                            valores_encontrados.append(valor_encontrado)
                            nominais.append(nominal)
                            tolerancias_superiores.append(tolerancia_superior)
                            tolerancias_inferiores.append(tolerancia_inferior)
                            desvios.append(desvio)

                    #Abrir planilha e configurar para inserir dados
                    app = xw.App(visible=False)
                    workbook = app.books.open(planilha)
                    planilha = workbook.sheets.active
                    planilha.api.Unprotect()
                    rng = planilha.range("A10:A698")
                    rng.api.Rows.Hidden = False

                    primeira_celula_cota = nullcontext
                    for linha_celula in range(10, 388 + 1):                    
                        if planilha.range("A" + str(linha_celula)).color == (0, 255, 0):
                            primeira_celula_cota = planilha.range("A" + str(linha_celula)).address
                            break
                    if primeira_celula_cota == nullcontext:
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
            else:
                print("Código não encontrado!")
    except Exception as e:
        print(f"Ocorreu um erro em ler_arquivo_txt: {e}")

def ler_arquivo_mea_mistral(caminho):
    try:
        with open(caminho, 'r', encoding='latin-1') as arquivo:
            todas_linhas = arquivo.readlines()
            
            for i, linha in enumerate(todas_linhas):
                if '02' in linha:
                    codigo = linha[7:].strip()
                    peca = todas_linhas[i + 1][5:].strip()
                    break               
                    
            if codigo != "":
                caminho_planilha = encontrar_caminho_planilha(codigo)
                if not caminho_planilha:
                    print("Planilha não encontrada")
                else:
                    nomes_cotas = []
                    valores_encontrados = []
                    nominais = []
                    tolerancias_superiores = []
                    tolerancias_inferiores = []

                    for i, linha in enumerate(todas_linhas):
                        if "Cota" in linha or "COTA" in linha:
                            cotas_linhas = todas_linhas[i:]
                            break

                    for info_cotas in cotas_linhas:
                            if info_cotas and not info_cotas[:25].isspace(): 
                                nomes_cotas.append(info_cotas[:14].strip() + "_" + info_cotas[14:16].strip())
                                valores_encontrados.append(info_cotas[16:31].strip())
                                nominais.append(info_cotas[30:44].strip())
                                tolerancias_superiores.append(info_cotas[56:68].strip())
                                tolerancias_inferiores.append(info_cotas[43:56].strip())

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
                                planilha.range("Z" + str(linha_celula)).value = Decimal(nominal)
                                planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                                planilha.range("AD" + str(linha_celula)).value = Decimal(nominal)
                                planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                            else:
                                planilha.range("Z" + str(linha_celula)).value = Decimal(nominal) - Decimal(tolinf)
                                planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)

                                planilha.range("AD" + str(linha_celula)).value = Decimal(nominal) - (Decimal(tolinf) * Decimal(0.7))
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
            else:
                print("Código não encontrado!")
    except Exception as e:
        print(f"Ocorreu um erro em ler_arquivo_txt: {e}")


def encontrar_caminho_planilha(codigo):
    pastas_planilhas = [r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Bosch',r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Spacer']
    for pasta in pastas_planilhas:
        for arquivo in os.listdir(pasta):
            if arquivo.endswith('.xlsm'):
                if arquivo.find(codigo) != -1:
                    return os.path.join(pasta, arquivo)
    return None    

class ArquivoHandler(FileSystemEventHandler):
    def process_file(self, file_path):
        time.sleep(1)
        if file_path.endswith('.txt'):
            ler_arquivo_txt_zeiss(file_path)
        elif file_path.endswith('.MEA'):
            ler_arquivo_mea_mistral(file_path)
        print(f'Arquivo processado: {file_path}')

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