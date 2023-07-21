from contextlib import nullcontext
from decimal import Decimal
import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import xlwings as xw

# Pasta para monitorar
pasta_monitorada = r'\\Sch-fns03a\ds1\Qualidade1\2021\05.Metrologia\2021\23 Cep Zeiss'

def ler_arquivo_txt_zeiss(caminho):
    with open(caminho, 'r', encoding='latin-1') as arquivo:
        # Ler a quinta linha do arquivo
        linhas = arquivo.readlines()
        
        linha = linhas[5]  # Remove espaços em branco no início e no fim da linha

        codigo = linha[55:66].strip()
        planilha = encontrar_planilha(codigo)
        if(planilha != None):
            # Separar os valores da linha
            produto = linha[:9].strip()
            situacao = linha[28:42].strip()
            data = linha[42:52].strip()
            peca = linha[66:].strip()
            
            # Exibir os valores separados
            print("Produto:", produto)
            print("Situação:", situacao)
            print("Data:", data)
            print("Código:", codigo)
            print("Peça:", peca)

        # Abrir a planilha com o xlwings
            app = xw.App(visible=False)
            workbook = app.books.open(planilha)
            sheet = workbook.sheets.active

            # Desbloquear a planilha
            sheet.api.Unprotect()

            # Selecionar o range "A10:A255"
            rng = sheet.range("A10:A698")

            # Verificar e exibir as células ocultas
            rng.api.Rows.Hidden = False

            next_cell = nullcontext
            for row_cell in range(10, 388 + 1):
                # Verificar a cor da célula
                cor_celula = sheet.range("A" + str(row_cell)).color
                # Verificar se a cor é verde
                if cor_celula == (0, 255, 0):
                    next_cell = sheet.range("A" + str(row_cell))
                    break
            if next_cell == nullcontext:
                next_cell = sheet.range("A10").end("down").offset(row_offset=1)
            # Armazenar a posição da célula abaixo da célula atual
            next_cell_position = next_cell.address
            row_cell = int(next_cell_position.split('$')[2])
            
            # Percorrer as células da linha 699 na direção da direita
            current_cell = sheet.range("B699")
            while current_cell.column < 26:  # Coluna Z é a coluna de número 26
                if current_cell.value == float(peca):
                    break
                #G257
                current_cell = current_cell.offset(column_offset=1)

            # Obter a coluna da célula atual
            column_letter = xw.utils.col_name(current_cell.column + 1)

            for linha in linhas[12:]:  #Começar a partir da linha 11
            
                if linha and not linha[:25].isspace(): # Verificar se a linha não está vazia
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
                    
                    # Selecionar a célula correspondente à coluna armazenada e à linha armazenada
                    value_cell = sheet.range(column_letter + str(row_cell))
                    
                    # Inserir o valor do "nome" na coluna A da linha armazenada
                    sheet.range("A" + str(row_cell)).value = nome
                    sheet.range("A" + str(row_cell)).color = (0, 255, 0)

                    # Copiar o valor do desvio para atual se atual for nulo
                    if not atual:
                        nominal  = "0.000"
                        atual = desvio
                        # Inserir o valor do "nome" na coluna A da linha armazenada
                        sheet.range("Z" + str(row_cell)).value = Decimal(nominal)
                        sheet.range("AA" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)
                        # Inserir o valor do "nome" na coluna A da linha armazenada
                        sheet.range("AD" + str(row_cell)).value = Decimal(nominal)
                        sheet.range("AE" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)
                    else:
                        # Inserir o valor do "nome" na coluna A da linha armazenada
                        sheet.range("Z" + str(row_cell)).value = Decimal(nominal) + Decimal(tolinf)
                        sheet.range("AA" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)

                        sheet.range("AD" + str(row_cell)).value = Decimal(nominal) + (Decimal(tolinf) * Decimal(0.7))
                        sheet.range("AE" + str(row_cell)).value = Decimal(nominal) + (Decimal(tolsup) * Decimal(0.7))

                    # Inserir o valor do "atual" na coluna correspondente à coluna armazenada e à linha armazenada
                    value_cell.value = atual
                    
                    row_cell = row_cell + 1
                    print(nome + " " + atual + " " + nominal + " " + tolsup + " " + tolinf + " " + desvio)

            # Deslocar até a última célula com valor na coluna A
            next_cell = sheet.range("A10").end("down").offset(row_offset=1)
            next_last_cell = next_cell.end("down").offset(row_offset=-1)
            next_last_cell_position = next_last_cell.address
            next_cell_position = next_cell.address
            range_address = next_cell_position.replace("$", "") + ":" + next_last_cell_position.replace("$", "") 
            # Armazenar a posição da célula abaixo da célula atual
            sheet.range(range_address).api.Rows.Hidden = True
            sheet.api.Protect()
            workbook.save()
            workbook.close()
            app.quit()
        else:
            print("Planilha não encontrada")

def ler_arquivo_mea_mistral(caminho):
    with open(caminho, 'r', encoding='latin-1') as arquivo:
        linhas = arquivo.readlines()
        linha4 = linhas[4]
        linha5 = linhas[5]
        peca = linha5[5:].strip()
            
        codigo = linha4[7:].strip()

        
        planilha = encontrar_planilha(codigo)
        if(planilha != None):

        # Abrir a planilha com o xlwings
            app = xw.App(visible=False)
            workbook = app.books.open(planilha)
            sheet = workbook.sheets.active

            # Desbloquear a planilha
            sheet.api.Unprotect()

            # Selecionar o range "A10:A255"
            rng = sheet.range("A10:A698")

            # Verificar e exibir as células ocultas
            rng.api.Rows.Hidden = False

            next_cell = nullcontext
            for row_cell in range(10, 388 + 1):
                # Verificar a cor da célula
                cor_celula = sheet.range("A" + str(row_cell)).color
                # Verificar se a cor é verde
                if cor_celula == (0, 255, 0):
                    next_cell = sheet.range("A" + str(row_cell))
                    break
            if next_cell == nullcontext:
                next_cell = sheet.range("A10").end("down").offset(row_offset=1)
            # Armazenar a posição da célula abaixo da célula atual
            next_cell_position = next_cell.address
            row_cell = int(next_cell_position.split('$')[2])
            
            # Percorrer as células da linha 699 na direção da direita
            current_cell = sheet.range("B699")
            while current_cell.column < 26:  # Coluna Z é a coluna de número 26
                if current_cell.value == float(peca):
                    break
                #G257
                current_cell = current_cell.offset(column_offset=1)

            # Obter a coluna da célula atual
            column_letter = xw.utils.col_name(current_cell.column + 1)

            for linha in linhas[6:]:  #Começar a partir da linha 7
            
                if linha and not linha[:14].isspace(): # Verificar se a linha não está vazia
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
                    # Selecionar a célula correspondente à coluna armazenada e à linha armazenada
                    value_cell = sheet.range(column_letter + str(row_cell))
                    
                    # Inserir o valor do "nome" na coluna A da linha armazenada
                    sheet.range("A" + str(row_cell)).value = nome
                    sheet.range("A" + str(row_cell)).color = (0, 255, 0)

                    # Copiar o valor do desvio para atual se atual for nulo
                    if not atual:
                        nominal  = "0.000"
                        sheet.range("Z" + str(row_cell)).value = Decimal(nominal)
                        sheet.range("AA" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)

                        sheet.range("AD" + str(row_cell)).value = Decimal(nominal)
                        sheet.range("AE" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)
                    else:
                        sheet.range("Z" + str(row_cell)).value = Decimal(nominal) - Decimal(tolinf)
                        sheet.range("AA" + str(row_cell)).value = Decimal(nominal) + Decimal(tolsup)

                        sheet.range("AD" + str(row_cell)).value = Decimal(nominal) - (Decimal(tolinf) * Decimal(0.7))
                        sheet.range("AE" + str(row_cell)).value = Decimal(nominal) + (Decimal(tolsup) * Decimal(0.7))

                    # Inserir o valor do "atual" na coluna correspondente à coluna armazenada e à linha armazenada
                    value_cell.value = atual
                    
                    row_cell = row_cell + 1
                    print(nome + " " + atual + " " + nominal + " " + tolsup + " " + tolinf)

            # Deslocar até a última célula com valor na coluna A
            next_cell = sheet.range("A10").end("down").offset(row_offset=1)
            next_last_cell = next_cell.end("down").offset(row_offset=-1)
            next_last_cell_position = next_last_cell.address
            next_cell_position = next_cell.address
            range_address = next_cell_position.replace("$", "") + ":" + next_last_cell_position.replace("$", "") 
            # Armazenar a posição da célula abaixo da célula atual
            sheet.range(range_address).api.Rows.Hidden = True
            sheet.api.Protect()
            workbook.save()
            workbook.close()
            app.quit()
        else:
            print("Planilha não encontrada")


def encontrar_planilha(codigo):
    for arquivo in os.listdir(r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Bosch'):
        if arquivo.endswith('.xlsx') or arquivo.endswith('.xlsm'):
            if arquivo.find(codigo) != -1:
                return os.path.join(r'\\Sch-fns03a\ds1\Producao2\Registro de Inspeção\Bosch', arquivo)
    return None    

# Classe para monitorar eventos de adição de arquivo na pasta
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

# Configurar o observador
event_handler = ArquivoHandler()
observer = Observer()
observer.schedule(event_handler, path=pasta_monitorada, recursive=False)

# Iniciar o observador
observer.start()

try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer

