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
            linhas = arquivo.readlines()
            
            linha = linhas[5]

            codigo = linha[55:66].strip()
            caminho_planilha = encontrar_caminho_planilha(codigo)
            if(caminho_planilha == None):
                print("Planilha não encontrada")
            else:
                produto = linha[:9].strip()
                situacao = linha[28:42].strip()
                data = linha[42:52].strip()
                peca = linha[66:].strip()
                
                print("Produto:", produto)
                print("Situação:", situacao)
                print("Data:", data)
                print("Código:", codigo)
                print("Peça:", peca)

                app = xw.App(visible=False)
                workbook = app.books.open(planilha)
                planilha = workbook.sheets.active

                planilha.api.Unprotect()

                rng = planilha.range("A10:A698")

                rng.api.Rows.Hidden = False

                next_cell = nullcontext
                for row_cell in range(10, 388 + 1):
                    cor_celula = planilha.range("A" + str(row_cell)).color
                    
                    if cor_celula == (0, 255, 0):
                        next_cell = planilha.range("A" + str(row_cell))
                        break
                if next_cell == nullcontext:
                    next_cell = planilha.range("A10").end("down").offset(row_offset=1)
                next_cell_position = next_cell.address
                row_cell = int(next_cell_position.split('$')[2])
                
                current_cell = planilha.range("B699")
                while current_cell.column < 26:  
                    if current_cell.value == float(peca):
                        break
                    current_cell = current_cell.offset(column_offset=1)

                column_letter = xw.utils.col_name(current_cell.column + 1)

                for linha in linhas[14:]:
                
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
                        
                        value_cell = planilha.range(column_letter + str(row_cell))
                        
                        planilha.range("A" + str(row_cell)).value = nome
                        planilha.range("A" + str(row_cell)).color = (0, 255, 0)


                        if not atual:
                            nominal  = "0.000"
                            atual = desvio
                            planilha.range("Z" + str(row_cell)).value = Decimal(nominal)
                            planilha.range("AA" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)
                            planilha.range("AD" + str(row_cell)).value = Decimal(nominal)
                            planilha.range("AE" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)
                        else:
                            planilha.range("Z" + str(row_cell)).value = Decimal(nominal) + Decimal(tolinf)
                            planilha.range("AA" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)

                            planilha.range("AD" + str(row_cell)).value = Decimal(nominal) + (Decimal(tolinf) * Decimal(0.7))
                            planilha.range("AE" + str(row_cell)).value = Decimal(nominal) + (Decimal(tolsup) * Decimal(0.7))

                        value_cell.value = atual
                        
                        row_cell = row_cell + 1
                        print(nome + " " + atual + " " + nominal + " " + tolsup + " " + tolinf + " " + desvio)

                next_cell = planilha.range("A10").end("down").offset(row_offset=1)
                next_last_cell = next_cell.end("down").offset(row_offset=-1)
                next_last_cell_position = next_last_cell.address
                next_cell_position = next_cell.address
                range_address = next_cell_position.replace("$", "") + ":" + next_last_cell_position.replace("$", "") 
                planilha.range(range_address).api.Rows.Hidden = True
                planilha.api.Protect()
                workbook.save()
                workbook.close()
                app.quit()
    except Exception as e:
        print(f"Ocorreu um erro em ler_arquivo_txt: {e}")

def ler_arquivo_mea_mistral(caminho):
        with open(caminho, 'r', encoding='latin-1') as arquivo:
            linhas = arquivo.readlines()

            linha4 = linhas[4]
            linha5 = linhas[5]
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

                next_cell = nullcontext
                for row_cell in range(10, 388 + 1):
                    
                    cor_celula = planilha.range("A" + str(row_cell)).color
                    
                    if cor_celula == (0, 255, 0):
                        next_cell = planilha.range("A" + str(row_cell))
                        break
                if next_cell == nullcontext:
                    next_cell = planilha.range("A10").end("down").offset(row_offset=1)

                next_cell_position = next_cell.address
                row_cell = int(next_cell_position.split('$')[2])
                
                current_cell = planilha.range("B699")
                while current_cell.column < 26:  # Coluna Z é a coluna de número 26
                    if current_cell.value == float(peca):
                        break
                    current_cell = current_cell.offset(column_offset=1)

                column_letter = xw.utils.col_name(current_cell.column + 1)

                for linha in linhas[6:]:
                
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
                        value_cell = planilha.range(column_letter + str(row_cell))
                        
                        planilha.range("A" + str(row_cell)).value = nome
                        planilha.range("A" + str(row_cell)).color = (0, 255, 0)



def encontrar_caminho_planilha(codigo):
    pastas_planilhas = [r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Bosch',r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Spacer']
    for pasta in pastas_planilhas:
        for arquivo in os.listdir(pasta):
            if arquivo.endswith('.xlsx') or arquivo.endswith('.xlsm'):
                if arquivo.find(codigo) != -1:
                    return os.path.join(pasta, arquivo)
        return None    

class ArquivoHandler(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return
        elif event.src_path.endswith('.txt'):
            time.sleep(1)
            ler_arquivo_txt_zeiss(event.src_path)
            print(f'Arquivo processado: {event.src_path}')
        elif event.src_path.endswith('.MEA'):
            time.sleep(1)
            ler_arquivo_mea_mistral(event.src_path)
            print(f'Arquivo processado: {event.src_path}')

observers = [Observer() for _ in range(len(pastas_monitoradas))]
for i, pasta in enumerate(pastas_monitoradas):
    event_handler = ArquivoHandler()
    observers[i].schedule(event_handler, path=pasta, recursive=False)
    observers[i].start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    for observer in observers:
        observer.stop()

for observer in observers:
    observer.join()