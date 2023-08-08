from contextlib import nullcontext
from decimal import Decimal
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import xlwings as xw

pastas_monitoradas = [r'\\Sch-fns03a\ds1\Inovacao1\Guilherme\Projetos\Novo IVR\Teste']

def encontrar_valor_linha_abaixo(linhas, palavra_buscadas, get_valor_proxima_linha=False):
    for i, linha in enumerate(linhas):
        if any(palavra_buscada in linha for palavra_buscada in palavra_buscadas):
            if get_valor_proxima_linha and i + 1 < len(linhas):
                return linhas[i + 1].strip()
            else:
                return linha.strip()
    return None

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
                caminho_planilha = encontrar_caminho_planilha(codigo)
                if(caminho_planilha == None):
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
                            cotas_linhas = todas_linhas[i].strip()
                            break
                    for info_cotas in cotas_linhas[14:]:
                        if info_cotas and not info_cotas[:25].isspace(): 
                            nomes_cotas.append(info_cotas[:25].strip() + "_" + info_cotas[25:35].strip())
                            valores_encontrados.append(info_cotas[35:46].strip())
                            nominais.append(info_cotas[49:58].strip())
                            tolerancias_superiores(info_cotas[58:67].strip())
                            tolerancias_inferiores(info_cotas[67:76].strip())
                            desvios.append(info_cotas[76:85].strip())

                            if not tolsup:
                                tolsup = "0.000"

                            if not tolinf:
                                tolinf = "0.000"
                            
                            
                            planilha.range("A" + str(linha_celula)).value = nome
                            planilha.range("A" + str(linha_celula)).color = (0, 255, 0)


                            if not atual:
                                nominal  = "0.000"
                                atual = desvio
                                planilha.range("Z" + str(linha_celula)).value = Decimal(nominal)
                                planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                                planilha.range("AD" + str(linha_celula)).value = Decimal(nominal)
                                planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                            else:
                                planilha.range("Z" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolinf)
                                planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)

                                planilha.range("AD" + str(linha_celula)).value = Decimal(nominal) + (Decimal(tolinf) * Decimal(0.7))
                                planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + (Decimal(tolsup) * Decimal(0.7))

                            value_cell.value = atual
                            

                    peca = dados_medicao[66:].strip()

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

                    for linha in todas_linhas[14:]:
                    
                        if linha and not linha[:25].isspace(): 
                            cota = linha[:25].strip()
                            descricao = linha[25:35].strip()
                            nome = cota + "_" + descricao
                            atual = linha[35:46].strip()
                            nominal = linha[49:58].strip()
                            tolsup = linha[58:67].strip()
                            tolinf = linha[67:76].strip()
                            desvio = linha[76:85].strip()


                            if not tolsup:
                                tolsup = "0.000"

                            if not tolinf:
                                tolinf = "0.000"
                            
                            value_cell = planilha.range(coluna_celula_numero_peca + str(linha_celula))
                            
                            planilha.range("A" + str(linha_celula)).value = nome
                            planilha.range("A" + str(linha_celula)).color = (0, 255, 0)


                            if not atual:
                                nominal  = "0.000"
                                atual = desvio
                                planilha.range("Z" + str(linha_celula)).value = Decimal(nominal)
                                planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                                planilha.range("AD" + str(linha_celula)).value = Decimal(nominal)
                                planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)
                            else:
                                planilha.range("Z" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolinf)
                                planilha.range("AA" + str(linha_celula)).value = Decimal(nominal) + Decimal(tolsup)

                                planilha.range("AD" + str(linha_celula)).value = Decimal(nominal) + (Decimal(tolinf) * Decimal(0.7))
                                planilha.range("AE" + str(linha_celula)).value = Decimal(nominal) + (Decimal(tolsup) * Decimal(0.7))

                            value_cell.value = atual
                            
                            linha_celula = linha_celula + 1
                            print(nome + " " + atual + " " + nominal + " " + tolsup + " " + tolinf + " " + desvio)

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
        with open(caminho, 'r', encoding='latin-1') as arquivo:
            todas_linhas = arquivo.readlines()

            linha4 = todas_linhas[4]
            linha5 = todas_linhas[5]
            peca = linha5[5:].strip()
                
            codigo = linha4[7:].strip()
            if(codigo != ""):
                caminho_planilha = encontrar_caminho_planilha(codigo)
            if(planilha == None):
                print("Planilha não encontrada")
            else:
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

                for linha in todas_linhas[6:]:
                
                    if linha and not linha[:14].isspace():
                        cota = linha[:14].strip()
                        descricao = linha[14:16].strip()
                        nome = cota + "_" + descricao
                        atual = linha[16:31].strip()
                        nominal = linha[30:44].strip()
                        tolsup = linha[56:68].strip()
                        tolinf = linha[43:56].strip()


                        if not tolsup:
                            tolsup = "0.000"

                        if not tolinf:
                            tolinf = "0.000"
                        
                        print(nome + " " + atual + " " + nominal + " " + tolsup + " " + tolinf)
                        value_cell = planilha.range(coluna_celula_numero_peca + str(linha_celula))
                        
                        planilha.range("A" + str(linha_celula)).value = nome
                        planilha.range("A" + str(linha_celula)).color = (0, 255, 0)



def encontrar_caminho_planilha(codigo):
    pastas_planilhas = [r'\\Sch-fns03a\ds1\Inovacao1\Guilherme\Projetos\Novo IVR\Teste']
    for pasta in pastas_planilhas:
        for arquivo in os.listdir(pasta):
            if arquivo.endswith('.xlsx') or arquivo.endswith('.xlsm'):
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