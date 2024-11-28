import time
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, ElementNotInteractableException, NoSuchElementException
from datetime import datetime
import pandas as pd
import os


# 1. Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta que o Chrome está utilizando para depuração
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


# 2. Função auxiliar para verificar se um elemento existe com tempo de espera reduzido
def elemento_existe(driver, xpath, tempo_espera=1):
    try:
        WebDriverWait(driver, tempo_espera).until(EC.presence_of_element_located((By.XPATH, xpath)))
        return True
    except TimeoutException:
        print(f"Elemento não encontrado: {xpath}")
        return False


# 3. Função para garantir que o elemento esteja visível e clicável antes de interagir
def clicar_elemento(driver, xpath):
    try:
        # Esperar o elemento estar visível e clicável
        elemento = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView();", elemento)  # Scroll até o elemento
        elemento.click()
        print(f"Elemento clicado com sucesso: {xpath}")
    except (TimeoutException, ElementNotInteractableException) as e:
        print(f"Erro ao clicar no elemento: {xpath}, Erro: {e}")
        raise e


# 4. Função para reiniciar a navegação a partir da proposta seguinte
def reiniciar_navegacao(driver):
    try:
        WebDriverWait(driver, 50).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]'))).click()
        WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[2]/div[1]/div[1]/ul/li[3]/a'))).click()
    except TimeoutException:
        print("Elemento não encontrado ou levou muito tempo para carregar durante a navegação. Tentando novamente.")
        driver.refresh()
        WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[3]'))).click()


# 5. Função para buscar a data mais recente em uma tabela (Certidões, Declarações, etc.)
def buscar_data_mais_recente(driver, xpath_coluna_datas, secao):
    try:
        if elemento_existe(driver, xpath_coluna_datas, tempo_espera=0):  # Sem tempo de espera
            datas_upload = driver.find_elements(By.XPATH, xpath_coluna_datas)
            if datas_upload:
                lista_datas = []
                for data_element in datas_upload:
                    data_text = data_element.text
                    try:
                        data_convertida = datetime.strptime(data_text, "%d/%m/%Y %H:%M:%S")
                        lista_datas.append(data_convertida)
                    except Exception as e:
                        print(f"Erro ao converter data na seção '{secao}': {data_text}, Erro: {e}")
                if lista_datas:
                    data_mais_recente = max(lista_datas)
                    print(f"Data mais recente na seção '{secao}': {data_mais_recente.strftime('%d/%m/%Y %H:%M:%S')}")
                    return data_mais_recente.strftime('%d/%m/%Y %H:%M:%S')
            print(f"Nenhuma data encontrada na seção '{secao}'.")
        else:
            print(f"Seção '{secao}' não localizada.")
        return None
    except Exception as e:
        print(f"Erro ao buscar a data mais recente na seção '{secao}': {e}")
        return None


# 6. Função para buscar o valor do status, sem tentar converter para data
def buscar_status(driver, xpath_coluna_status, secao):
    try:
        if elemento_existe(driver, xpath_coluna_status, tempo_espera=0):  # Sem tempo de espera
            status_valor = driver.find_element(By.XPATH, xpath_coluna_status).text
            print(f"Status na seção '{secao}': {status_valor}")
            return status_valor
        else:
            print(f"Seção '{secao}' não localizada.")
        return None
    except Exception as e:
        print(f"Erro ao buscar o status na seção '{secao}': {e}")
        return None


# 7. Função para processar uma proposta
def processar_proposta(driver, proposta_numero):
    print(f"Iniciando o processamento da proposta: {proposta_numero}")

    # Digitar o número da proposta e consultar
    driver.find_element(By.XPATH, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input').clear()
    driver.find_element(By.XPATH,
                        '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/input').send_keys(
        proposta_numero)

    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input',
                       tempo_espera=2):
        clicar_elemento(driver, '/html/body/div[3]/div[15]/div[3]/div/div/form/table/tbody/tr[1]/td[2]/span/input')

    # Clicar no número da proposta
    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a', tempo_espera=2):
        clicar_elemento(driver, '/html/body/div[3]/div[15]/div[3]/div[3]/table/tbody/tr/td[1]/div/a')

    # Clicar na aba "Requisitos"
    if elemento_existe(driver, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span', tempo_espera=2):
        clicar_elemento(driver, '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span')
        print("Aba 'Requisitos' clicada com sucesso.")
    else:
        print("Aba 'Requisitos' não localizada.")
        return False

    # Tentar clicar em "Requisitos para celebração" com cinco XPaths
    xpaths = [
        '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[12]/div/span/span',  # Primeiro XPath
        '/html/body/div[3]/div[15]/div[1]/div/div[1]/a[3]/div/span/span',  # Segundo XPath
        '/html/body/div[3]/div[16]/div[1]/div/div[2]/a[29]/div/span/span',  # Terceiro XPath
        '/html/body/div[3]/div[16]/div[1]/div/div[2]/a[13]/div/span/span',  # Quarto XPath
        '/html/body/div[3]/div[15]/div[1]/div/div[2]/a[13]/div/span/span'  # Quinto XPath
    ]

    clicado = False
    for xpath in xpaths:
        if elemento_existe(driver, xpath, tempo_espera=1):
            try:
                clicar_elemento(driver, xpath)
                clicado = True
                print(f"Clicado em 'Requisitos para celebração' usando XPath: {xpath}")
                break
            except Exception as e:
                print(
                    f"Erro ao tentar clicar em 'Requisitos para celebração' usando XPath: {xpath}. Tentando o próximo...")

    if not clicado:
        print("Requisitos para celebração não localizado.")
        return False

    # Consultar valores nas seções, agora com pulos rápidos em caso de não localização
    dados_proposta = {
        "Proposta": proposta_numero,
        "Certidões": buscar_data_mais_recente(driver,
                                              "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[1]/table/tbody/tr[1]/td[2]",
                                              "Certidões"),
        "Declarações": buscar_data_mais_recente(driver,
                                                "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[2]/table/tbody/tr[1]/td[2]",
                                                "Declarações"),
        "Comprovantes de Execução": buscar_data_mais_recente(driver,
                                                             "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[3]/table/tbody/tr[1]/td[2]",
                                                             "Comprovantes de Execução"),
        "Outros": buscar_data_mais_recente(driver,
                                           "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[4]/table/tbody/tr[1]/td[2]",
                                           "Outros"),
        "Históricos - Data": buscar_data_mais_recente(driver,
                                                      "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[3]",
                                                      "Históricos"),
        "Históricos - Status": buscar_status(driver,
                                             "/html/body/div[3]/div[16]/div[2]/div[2]/form/div[1]/div[5]/table/tbody/tr[1]/td[1]",
                                             "Históricos Status")
    }

    return dados_proposta


# 8. Função para salvar o progresso no Excel, agora limpando a planilha no início
def salvar_progresso(resultado):
    caminho_arquivo = r"C:\Users\d-deb\OneDrive\Documents\dev\robov1\Consulta Transferegov Requisitos.xlsx"

    try:
        if not os.path.exists(caminho_arquivo):
            # Se o arquivo não existir, criar a primeira vez
            pd.DataFrame([resultado]).to_excel(caminho_arquivo, index=False)
        else:
            # Concatenar novos dados ao arquivo existente
            df_existente = pd.read_excel(caminho_arquivo)
            df_atualizado = pd.concat([df_existente, pd.DataFrame([resultado])], ignore_index=True)
            df_atualizado.to_excel(caminho_arquivo, index=False)
        print(f"Progresso salvo com sucesso no arquivo '{caminho_arquivo}'!")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")


# 9. Função para clicar no botão "Nova Pesquisa" com fallback para XPaths diferentes
def clicar_nova_pesquisa(driver):
    xpaths_nova_pesquisa = [
        '/html/body/div[3]/div[3]/div[6]/a[2]',  # Primeiro XPath
        '/html/body/div[3]/div[2]/div[6]/a[2]',  # Segundo XPath
        '/html/body[1]/div[3]/div[3]/div[6]/a[2]'  # Terceiro XPath
    ]

    botao_encontrado = False

    for xpath in xpaths_nova_pesquisa:
        if elemento_existe(driver, xpath, tempo_espera=2):
            try:
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpath))).click()
                print(f"Botão 'Nova Pesquisa' clicado com sucesso usando XPath: {xpath}")
                botao_encontrado = True
                break
            except (NoSuchElementException, TimeoutException):
                print(f"Falha ao clicar no botão 'Nova Pesquisa' com o XPath: {xpath}. Tentando o próximo...")
        else:
            print(f"Botão 'Nova Pesquisa' não encontrado com o XPath: {xpath}. Tentando o próximo...")

    # Se nenhum dos XPaths funcionar, recarregar a página e reiniciar a navegação
    if not botao_encontrado:
        print(f"Botão 'Nova Pesquisa' não encontrado após tentar todos os XPaths. Recarregando a página.")
        reiniciar_navegacao(driver)


# 10. Função principal para processar todas as propostas
def processar_todas_propostas():
    resultados = []
    propostas_consultadas = 0
    tempo_acumulado = 0  # Para somar o tempo total

    caminho_planilha = r"C:\Users\d-deb\OneDrive\Documents\dev\robov1\propostas_iniciais.xlsx"
    df_propostas = pd.read_excel(caminho_planilha)

    # Limpar espaços desnecessários
    df_propostas.columns = df_propostas.columns.str.strip()
    df_propostas['NºProposta'] = df_propostas['NºProposta'].str.strip()

    driver = conectar_navegador_existente()

    # Limpar a planilha antes de começar
    caminho_arquivo_resultados = r"C:\Users\d-deb\OneDrive\Documents\dev\robov1\Consulta Transferegov Requisitos.xlsx"
    if os.path.exists(caminho_arquivo_resultados):
        os.remove(caminho_arquivo_resultados)  # Apagar o arquivo existente para começar do zero
        print(f"A planilha '{caminho_arquivo_resultados}' foi limpa.")

    try:
        reiniciar_navegacao(driver)

        for index, row in df_propostas.iterrows():
            proposta_numero = row['NºProposta']
            propostas_consultadas += 1

            print(f"\nConsultando proposta {proposta_numero} (Proposta {propostas_consultadas})...")

            try:
                # Medir o tempo de início da consulta
                tempo_inicio = time.time()

                # Processar a proposta
                dados_proposta = processar_proposta(driver, proposta_numero)
                resultados.append(dados_proposta)

                # Medir o tempo de fim da consulta
                tempo_fim = time.time()
                tempo_consulta = tempo_fim - tempo_inicio  # Tempo em segundos para essa proposta
                tempo_acumulado += tempo_consulta  # Acumular o tempo total

                # Convertendo o tempo total acumulado para minutos e segundos
                minutos_acumulados = int(tempo_acumulado // 60)
                segundos_acumulados = int(tempo_acumulado % 60)

                # Exibir informações de monitoramento
                print(f"Proposta {proposta_numero} consultada em {tempo_consulta:.2f} segundos.")
                print(f"Tempo total acumulado: {minutos_acumulados}m:{segundos_acumulados:02d}s.")
                print(f"Propostas consultadas até agora: {propostas_consultadas}")

                # Salvar progresso
                salvar_progresso(dados_proposta)

                # Clicar no botão "Nova Pesquisa"
                clicar_nova_pesquisa(driver)

            except Exception as e:
                print(f"Erro ao processar a proposta {proposta_numero}: {e}")
                salvar_progresso({"Proposta": proposta_numero, "Erro": str(e)})
                reiniciar_navegacao(driver)

    finally:
        driver.quit()


# 11. Executar o processamento de todas as propostas
processar_todas_propostas()

print("Execução concluída e dados salvos com sucesso!")
