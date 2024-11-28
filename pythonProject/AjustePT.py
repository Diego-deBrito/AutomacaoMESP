import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill


def conectar_navegador_existente():
    """
    Conecta ao navegador Chrome já aberto utilizando a porta de depuração 9222.
    """
    try:
        print("Tentando conectar ao navegador na porta 9222...")

        opcoes_navegador = webdriver.ChromeOptions()
        opcoes_navegador.debugger_address = "localhost:9222"

        navegador = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opcoes_navegador)
        print("Conectado ao navegador existente com sucesso.")
        return navegador

    except WebDriverException as erro:
        print(f"Erro ao conectar ao navegador existente: {erro}")
        return None


def clicar_elemento(navegador, xpath, tempo_espera=10):
    """
    Clica em um elemento identificado pelo XPath.
    """
    try:
        elemento = WebDriverWait(navegador, tempo_espera).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        elemento.click()
        print(f"Elemento clicado: {xpath}")
    except (TimeoutException, NoSuchElementException) as erro:
        print(f"Erro ao clicar no elemento {xpath}: {erro}")


def clicar_e_colar_texto(navegador, xpath):
    """
    Clica em um campo e cola o conteúdo da área de transferência.
    """
    try:
        elemento = WebDriverWait(navegador, 10).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        elemento.click()
        elemento.send_keys(Keys.CONTROL, 'v')  # Cola o texto copiado
        print(f"Texto colado no campo: {xpath}")
    except (TimeoutException, NoSuchElementException) as erro:
        print(f"Erro ao clicar e colar texto no elemento {xpath}: {erro}")


def criar_nova_planilha_excel(caminho_arquivo, dados, colunas):
    """
    Cria um novo arquivo Excel com os dados fornecidos.
    """
    try:
        nova_planilha = Workbook()
        aba_resultados = nova_planilha.active
        aba_resultados.title = "Resultados"

        # Adiciona cabeçalhos na planilha
        aba_resultados.append(colunas)

        # Adiciona os dados processados
        for linha in dados:
            aba_resultados.append(linha)

        # Salva o novo arquivo Excel
        nova_planilha.save(caminho_arquivo)
        print(f"Novo arquivo criado em: {caminho_arquivo}")
    except Exception as erro:
        print(f"Erro ao criar o novo arquivo Excel: {erro}")


def executar_processo_principal():
    """
    Fluxo principal para carregar dados do Excel, processar informações e gerar uma nova planilha.
    """
    print("Iniciando o processo principal...")

    # Conectar ao navegador existente
    navegador = conectar_navegador_existente()
    if not navegador:
        print("Não foi possível conectar ao navegador. Encerrando o processo.")
        return

    # Definir caminhos de entrada e saída
    caminho_arquivo_entrada = r'C:/Users/diego.brito/Downloads/robov1/CONTROLE DE PARCERIAS CGAP.xlsx'
    caminho_arquivo_saida = r'C:/Users/diego.brito/Downloads/Resultados_Processados.xlsx'

    # Carregar dados do arquivo de entrada
    try:
        dataframe = pd.read_excel(caminho_arquivo_entrada, sheet_name='PARCERIAS CGAP', engine='openpyxl')

        # Exibir nomes de colunas para verificação
        print("Colunas disponíveis no arquivo:", dataframe.columns.tolist())

        # Verificar e converter a coluna "Instrumento nº" para string
        if "Instrumento nº" in dataframe.columns:
            dataframe["Instrumento nº"] = dataframe["Instrumento nº"].astype(str).str.replace(r'\.0$', '', regex=True)
        else:
            print("A coluna 'Instrumento nº' não foi encontrada no arquivo. Encerrando o processo.")
            return

        colunas_interesse = ["Instrumento nº", "Técnico", "e-mail do Técnico"]
        if not all(coluna in dataframe.columns for coluna in colunas_interesse):
            print(f"Colunas necessárias não encontradas: {colunas_interesse}. Encerrando o processo.")
            return

        dados_saida = []

        # Loop através de cada linha do DataFrame
        for _, linha in dataframe.iterrows():
            instrumento_numero = linha["Instrumento nº"]
            tecnico = linha["Técnico"]
            email_tecnico = linha["e-mail do Técnico"]

            # Copiar número do instrumento para a área de transferência
            pyperclip.copy(instrumento_numero)

            try:
                # Realizar interações no navegador
                clicar_elemento(navegador, '//*[@id="menuPrincipal"]/div[1]/div[4]')
                clicar_elemento(navegador, '//*[@id="contentMenu"]/div[1]/ul/li[5]/a')
                clicar_e_colar_texto(navegador, '//*[@id="consultarNumeroConvenio"]')
                clicar_elemento(navegador, '//*[@id="form_submit"]')
                clicar_elemento(navegador, '//*[@id="instrumentoId"]/a')
                clicar_elemento(navegador, '//*[@id="div_-173460853"]/span/span')
                clicar_elemento(navegador, '//*[@id="menu_link_-173460853_-1293190284"]/div/span/span')



                # Verificar situação "Em Análise" ou "Em Análise (aguardando parecer)"
                try:
                    elemento_situacao = WebDriverWait(navegador, 5).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="row"]//td[contains(text(),"Em Análise")]'))
                    )
                    situacao = elemento_situacao.text
                    clicar_elemento(navegador, '//*[@id="tbodyrow"]/tr[5]/td[4]/nobr/a')
                    data_solicitacao = navegador.find_element(By.XPATH, '//*[@id="tr-editarDataSolicitacao"]/td[2]').text
                except TimeoutException:
                    situacao = "Sem ajuste"
                    data_solicitacao = ""
                clicar_elemento(navegador, '//*[@id="logo"]/a/span')
                # Adicionar linha processada à lista de dados para saída
                dados_saida.append([instrumento_numero, tecnico, email_tecnico, situacao, data_solicitacao])
                print(f"Instrumento {instrumento_numero}: {situacao}")
            except Exception as erro:
                print(f"Erro ao processar o instrumento {instrumento_numero}: {erro}")
                continue

        # Criar novo arquivo Excel com os resultados processados
        colunas_saida = ["Instrumento nº", "Técnico", "e-mail do Técnico", "AjustesPT", "Data da Solicitação"]
        criar_nova_planilha_excel(caminho_arquivo_saida, dados_saida, colunas_saida)

    except Exception as erro:
        print(f"Erro ao carregar ou processar o arquivo de entrada: {erro}")

    finally:
        # Fechar o navegador
        navegador.quit()
        print("Processo concluído.")


if __name__ == "__main__":
    executar_processo_principal()
