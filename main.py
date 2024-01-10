import requests, json, os, limit_timer, sys, keyboard, re, time
from configparser import ConfigParser


# Função para formatar à uma variável um objeto JSON
def jsonfy(directory, variavel):
    with open(directory, "w+") as file:
        responseParsed = json.loads(variavel)
        file.write(json.dumps(responseParsed, indent=2))
        file.close()
        return responseParsed

# Função para pegar o padrão correto do input
def isValidDateFormat(dateStr):
    pattern = re.compile(r'^\d{2}/\d{2}/\d{4}$')
    return bool(pattern.match(dateStr))

while True:

    # URL's para as requests
    urlFetchNF = "https://api.tiny.com.br/api2/notas.fiscais.pesquisa.php"
    urlGetXML = "https://api.tiny.com.br/api2/nota.fiscal.obter.xml.php"

    # Dados da request
    headers = {"Content-Type": "application/x-www-form-urlencoded"}
    formato = "JSON"

    # Importação do token
    dirToken = "C:\\TinyAPI\\token.cfg"
    try:
        configObject = ConfigParser()
        configObject.read(dirToken)
        key = configObject["KEY"]
        token = key["token"]
    except Exception as error:
        print("\n", error)
        time.sleep(60)
        sys.exit()

    # Input do usuario para saber o intervalo de data para pegar as notas fiscais
    while True:
        firstDate = input("Digite a data inicial no formato (dd/mm/yyyy): ")   
        if isValidDateFormat(firstDate):
            break
        else:
            print("Formato inválido. Por favor, digite no formato correto (dd/mm/yyyy).")

    while True:
        lastDate = input("Digite a data final no formato (dd/mm/yyyy): ")       
        if isValidDateFormat(lastDate):
            break
        else:
            print("Formato inválido. Por favor, digite no formato correto (dd/mm/yyyy).")
    print("\n")

    # Ajusta a data para salvar
    firstDateReplaced = firstDate.replace("/","-")
    lastDateReplaced = lastDate.replace("/","-")

    # Cria o diretório se não existe
    dirNF = "C:\\Users\\User\\Desktop\\extract_NFe_Tiny\\Notas Fiscais"
    dirProject = "C:\\Users\\User\\Desktop\\extract_NFe_Tiny"
    os.makedirs(f"{dirNF}\\{firstDateReplaced}_{lastDateReplaced}", exist_ok=True)

    # Request para pegar as notas fiscais geradas no intervalo selecionado
    data = f"token={token}&formato={formato}&dataInicial={firstDate}&dataFinal={lastDate}"
    try:
        response = requests.post(urlFetchNF, headers=headers, data=data)
        resposta = response.text
    except Exception as error:
        print("\n", error)
        time.sleep(60)
        sys.exit()

    # Salva em um JSON as informações das notas e extrai os id's e números de nota
    try:
        notasJSON = jsonfy(f"{dirProject}\\tempFile.json", resposta)
        nfe = notasJSON["retorno"]["notas_fiscais"]
        idNota = {}
    except KeyError:
        print("Não foram encontrados XMLs de notas fiscais no periodo selecionado.")
        time.sleep(60)
        sys.exit()

    for nota in nfe:
        try:
            idNota[nota["nota_fiscal"]["id"]] = nota["nota_fiscal"]["numero"]
        except:
            idNota[nota["nota_fiscal"]["id"]] = nota["nota_fiscal"]["numero_ecommerce"]

    # Variaveis para o controle do limite de requests
    requestCounter = 0
    requestLimit = 29
    requestSize = len(idNota)

    print("Baixando XMLs das Notas Fiscais...")

    # Loop para baixar e salvar os XMLs
    for id, value in idNota.items():
        # Variavel para controle da quantide de requests feitas
        requestCounter += 1

        # Manda a request para obter o XML da nota
        data = f"token={token}&id={id}"
        try:
            response = requests.post(urlGetXML, headers=headers, data=data)
            resposta = response.text
        except Exception as error:
            print("\n", error)
            time.sleep(60)
            sys.exit()

        try: # Salva os XMLs das notas
            with open(f"{dirNF}\\{firstDateReplaced}_{lastDateReplaced}\\{value}.xml", "w+") as file:
                file.write(resposta)
        except Exception as error:
            print("\n", error)
            time.sleep(60)
            sys.exit()

        print(f"Nota {value} baixada.")

        if requestCounter == requestLimit: # Ativa o timer se o limite de requests for alcançado
            requestCounter = 0
            print("Limite de importações atingido, esperando o timer finalizar...")
            limit_timer.create_timer_window()
            print("Timer finalizado, continuando as importações...")
            time.sleep(2)

    # Remove o arquivo temporario
    try:
        os.remove(f"{dirProject}\\tempFile.json")
    except Exception as error:
        print("\n", error)
        print("Não foi possivel apagar o arquivo temporario.")
        time.sleep(60)

    # Resumo de informações para o usuario
    print("\n")
    print(f"Foram importados {requestSize} arquivos XML de notas fiscais do dia {firstDate} ao dia {lastDate}. \n")

    # Dá a opção para o usuario de fechar o programa ou fazer outra operação
    print("Pressione [1] para sair ou [2] para realizar outra importação: ", end='', flush=True)

    keyb = keyboard.read_event(suppress=True).name

    if keyb == "1":
        # Se o usuário pressionar 1 o programa será encerrado
        break
    elif keyb == "2":
        # Se o usuário pressionar 2 o programa será reiniciado
        print("\n")
        continue
    else:
        # Se o usuário pressionar qualquer outra tecla o programa será encerrado
        break

sys.exit()
