import pygsheets #  biblioteca para mexer com o google sheets
from playwright.sync_api import sync_playwright  #  biblioteca para mexer no navegador de forma dinâmica
import logging  # gerador de logs

#**************** Inicio ****************

#  local do token de autorização da API do Google Sheets
path = (credenciais.json)
gc = pygsheets.authorize(service_account_file=path)

# abrindo a planilha, e escolhendo a aba
planilha = gc.open_by_key('id da planilha no Google Sheets')
aba = planilha.worksheet_by_title('Links')  # sh[0] seleciona a primeira sheet

#pegando a ultima linha disponivel para atualizar
linha = aba.get_value('H1')

# selecionando a aba onde ficarão os dados
abadados = planilha.worksheet_by_title('Dados')

# log config básico. W é de Write, A de append. W ele apaga o ultimo log, A, ele vai somando
logging.basicConfig(filename='Holmes_bot.log', filemode='a', format='%(asctime)s - %(levelname)s - %(message)s')

# Definindo os links de cada tipo de documento, de acordo com a céluda do Drive
nf = aba.get_value('C8')
ap = aba.get_value('C9')
federal = aba.get_value('C10')
icms = aba.get_value('C11')
iss = aba.get_value('C12')
difal = aba.get_value('C13')
comp = aba.get_value('C14')
adm = aba.get_value('C15')
peri = aba.get_value('C16')
ponto = aba.get_value('C17')
demissao = aba.get_value('C18')
contratos = aba.get_value('C19')
energia = aba.get_value('C20')
agua = aba.get_value('C21')
taxa = aba.get_value('C22')
iptu = aba.get_value('C23')
aluguel = aba.get_value('C24')
charge = aba.get_value('C25')
total = aba.get_value('C26')

print('Planilha logada')

#informando no log que essa etapa foi completada
logging.warning('Links da planilha resgatados com sucesso')

#iniciando módulo do navegador
with sync_playwright() as p:
    navegador = p.chromium.launch(headless=True)  #Headless = False. Significa: não fazer no modo invisvel.

    pagina = navegador.new_page()

    #inicio da operação - Fazer o Login no Holmes
    pagina.goto("https://empresa.holmesdoc.com/#home")
    pagina.locator('//*[@id="tiUser"]').fill(cred.email)
    pagina.locator('//*[@id="tiPass"]').fill(cred.senha)
    pagina.locator('//*[@id="login"]/div[4]/button').click()
    print('login efetuado no Holmes')

    try:
        #indo para a página de pesquisa. O link fica na planilha
        pagina.goto(nf)

        #locaizando a quantidade de documentos na página, e removendo o que é texto (deixando só numeros)
        nf1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        nf2 = nf1.replace(" ARQUIVO(S) ENCONTRADO(S).","")

        #primeiro vem a linha, depois coluna. E com isso, já Armazemos os dados na planilha
        abadados.update_value((linha, 2), nf2)

        # informando no log que essa etapa foi completada
        logging.warning('Quantidade de NFs obtido com sucesso')
    except:
        # se der erro ou não tiver valores, preencher com zero na relação, e avisar no log
        abadados.update_value((linha, 2), "0")
        # informando no log sobre o erro
        logging.warning('Erro: Quantidade de NF é zero, ou link não está disponivel')


    #Demais itens / documentos. É o mesmo código e log do passo anterior

    try:
        pagina.goto(ap)
        ap1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        ap2 = ap1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 3), ap2)

        logging.warning('Quantidade de APs obtido com sucesso')
    except:
        abadados.update_value((linha, 3), "0")
        logging.warning('Erro: Quantidade de AP é zero, ou link não está disponivel')


    try:
        pagina.goto(federal)
        federal1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        federal2 = federal1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 5), federal2)
        logging.warning('Quantidade de Federal obtido com sucesso')
    except:
        abadados.update_value((linha, 5), "0")
        logging.warning('Erro: Quantidade de Federal é zero, ou link não está disponivel')


    try:
        pagina.goto(icms)
        icms1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        icms2 = icms1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 6), icms2)
        logging.warning('Quantidade de ICMS obtido com sucesso')
    except:
        abadados.update_value((linha, 6), "0")
        logging.warning('Erro: Quantidade de ICMS é zero, ou link não está disponivel')


    try:
        pagina.goto(iss)

        iss1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        iss2 = iss1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 7), iss2)
        logging.warning('Quantidade de ISS obtido com sucesso')
    except:
        abadados.update_value((linha, 7), "0")
        logging.warning('Erro: Quantidade de ISS é zero, ou link não está disponivel')

    try:
        pagina.goto(difal)

        difal1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        difal2 = difal1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 8), difal2)
        logging.warning('Quantidade de Difal obtido com sucesso')
    except:
        abadados.update_value((linha, 8), "0")
        logging.warning('Erro: Quantidade de Difal é zero, ou link não está disponivel')


    try:
        pagina.goto(comp)

        comp1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        comp2 = comp1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 9), comp2)
        logging.warning('Quantidade de Comprovantes obtido com sucesso')
    except:
        abadados.update_value((linha, 9), "0")
        logging.warning('Erro: Quantidade de Comprovantes é zero, ou link não está disponivel')


    try:
        pagina.goto(adm)

        adm1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        adm2 = adm1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 11), adm2)
        logging.warning('Quantidade de Admissão obtido com sucesso')
    except:
        abadados.update_value((linha, 11), "0")
        logging.warning('Erro: Quantidade de Admissão é zero, ou link não está disponivel')

    try:
        pagina.goto(peri)

        peri1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        peri2 = peri1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 12), peri2)
        logging.warning('Quantidade de Periódico obtido com sucesso')
    except:
        abadados.update_value((linha, 12), "0")
        logging.warning('Erro: Quantidade de Periódico é zero, ou link não está disponivel')


    try:
        pagina.goto(ponto)

        ponto1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        ponto2 = ponto1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 13), ponto2)
        logging.warning('Quantidade de Ponto obtido com sucesso')
    except:
        abadados.update_value((linha, 13), "0")
        logging.warning('Erro: Quantidade de Ponto é zero, ou link não está disponivel')


    try:
        pagina.goto(demissao)

        demissao1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        demissao2 = demissao1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 14), demissao2)
        logging.warning('Quantidade de Demissões obtido com sucesso')
    except:
        abadados.update_value((linha, 14), "0")
        logging.warning('Erro: Quantidade de Demissões é zero, ou link não está disponivel')


    try:
        pagina.goto(contratos)

        contratos1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        contratos2 = contratos1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 15), contratos2)
        logging.warning('Quantidade de Contratos obtido com sucesso')
    except:
        abadados.update_value((linha, 15), "0")
        logging.warning('Erro: Quantidade de Contratos é zero, ou link não está disponivel')

    try:
        pagina.goto(energia)

        energia1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        energia2 = energia1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 17), energia2)
        logging.warning('Quantidade de Energia obtido com sucesso')
    except:
        abadados.update_value((linha, 17), "0")
        logging.warning('Erro: Quantidade de Energia é zero, ou link não está disponivel')

    try:
        pagina.goto(agua)

        agua1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        agua2 = agua1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 18), agua2)
        logging.warning('Quantidade de Contas de agua obtido com sucesso')
    except:
        abadados.update_value((linha, 18), "0")
        logging.warning('Erro: Quantidade de Contas é zero, ou link não está disponivel')

    try:
        pagina.goto(taxa)

        taxa1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        taxa2 = taxa1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 19), taxa2)
        logging.warning('Quantidade de Taxas obtido com sucesso')
    except:
        abadados.update_value((linha, 19), "0")
        logging.warning('Erro: Quantidade de Taxas é zero, ou link não está disponivel')


    try:
        pagina.goto(iptu)

        iptu1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        iptu2 = iptu1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 20), iptu2)
        logging.warning('Quantidade de IPTU obtido com sucesso')
    except:
        abadados.update_value((linha, 20), "0")
        logging.warning('Erro: Quantidade de IPTU é zero, ou link não está disponivel')

    try:
        pagina.goto(aluguel)

        aluguel1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        aluguel2 = aluguel1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 21), aluguel2)
        logging.warning('Quantidade de Aluguel obtido com sucesso')
    except:
        abadados.update_value((linha, 21), "0")
        logging.warning('Erro: Quantidade de Aluguel é zero, ou link não está disponivel')

    try:
        pagina.goto(charge)

        charge1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        charge2 = charge1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 22), charge2)
        logging.warning('Quantidade de Chargeback obtido com sucesso')
    except:
        abadados.update_value((linha, 22), "0")
        logging.warning('Erro: Quantidade de Chargeback é zero, ou link não está disponivel')

    try:
        pagina.goto(total)

        total1 = pagina.locator('//*[@id="results"]/div/div[5]/div[6]').inner_text()
        total2 = total1.replace(" ARQUIVO(S) ENCONTRADO(S).", "")
        abadados.update_value((linha, 24), total2)
        logging.warning('Quantidade Total obtida com sucesso')

    except:
        abadados.update_value((linha, 24), "0")
        logging.warning('Erro: Quantidade Total é zero, ou link não está disponivel')

logging.warning('Bot finalizado!')
print('Finalizado!')
