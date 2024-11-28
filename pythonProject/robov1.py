from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import time


# Função para conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"  # Porta que o Chrome está utilizando para depuração
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


# Função para buscar os cargos "Presidente" ou "Prefeito" na tabela, encontrar o botão e clicar
def identificar_cargo_e_clicar_botao(driver):
    try:
        # Esperar até que a tabela esteja presente
        tabela_membros = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="tblMembros"]'))  # Substitua pelo XPath da tabela
        )

        # Buscar todas as linhas da tabela
        linhas_tabela = tabela_membros.find_elements(By.XPATH, './/tr')

        # Percorrer todas as linhas e verificar o campo de cargo
        for linha in linhas_tabela:
            try:
                # Verificar se existe um campo de cargo dentro da linha com atributo title
                cargo_element = linha.find_element(By.XPATH, './/td[@title]')  # Procura a célula que contém o 'title'
                title_texto = cargo_element.get_attribute("title")
                print(f"Cargo encontrado: {title_texto}")

                # Verificar se o cargo é "Presidente" ou "Prefeito"
                if title_texto in ["Presidente", "Prefeito"]:
                    print(f"Cargo '{title_texto}' encontrado. Identificando o botão correspondente...")

                    # Localizar o botão dentro da mesma linha pelo ID e clicar
                    botao_element = linha.find_element(By.XPATH, './/button[contains(@id, "tblMembros_acoes")]')
                    botao_id = botao_element.get_attribute("id")
                    print(f"Botão encontrado com ID: {botao_id}")

                    # Clicar no botão correspondente
                    botao_element.click()
                    print("Botão clicado com sucesso!")
                    time.sleep(1)  # Pausa para garantir que a ação ocorra corretamente

                    break  # Para de procurar após encontrar e clicar no cargo correto

            except NoSuchElementException:
                # Se não encontrar um campo com o 'title' ou botão na linha, continua para a próxima linha
                continue

        print("Verificação concluída.")

    except (TimeoutException, NoSuchElementException) as e:
        print(f"Erro ao verificar o cargo e clicar no botão: {e}")

# Função principal para executar o fluxo de automação
def executar_automacao():
    driver = conectar_navegador_existente()

    # Acessar a URL onde a tabela está localizada
    driver.get("URL_DO_SEU_SITE")  # Substitua pela URL do seu site

    # Esperar 5 segundos para o site carregar completamente
    time.sleep(5)

    # Buscar o cargo e clicar no botão correspondente
    identificar_cargo_e_clicar_botao(driver)

    # Fechar o navegador após o teste
    time.sleep(5)  # Opcional: Mantenha o navegador aberto para verificar
    driver.quit()

# Executar o fluxo de automação
executar_automacao()