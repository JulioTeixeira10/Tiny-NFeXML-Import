import requests, json, os, limit_timer, sys, keyboard, re, time, zipfile
from modulo_email import sendEmail
from configparser import ConfigParser
from lxml import etree


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

# Função para validar o email digitado
def is_valid_email(email):
    email_regex = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
    match = re.match(email_regex, email)
    return bool(match)

# Função para zipar os arquivos XML
def zip_files(files_to_zip, zip_name):
    with zipfile.ZipFile(zip_name, 'w') as zipf:
        for file in files_to_zip:
            zipf.write(file)


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

    while True:
        firstDate = input("Digite a data inicial no formato (dd/mm/yyyy): ")   
        if isValidDateFormat(firstDate):
            break
        else:
            print("Formato inválido. Por favor, digite no formato correto (dd/mm/yyyy).")
            print("\n")

    while True:
        lastDate = input("Digite a data final no formato (dd/mm/yyyy): ")       
        if isValidDateFormat(lastDate):
            break
        else:
            print("Formato inválido. Por favor, digite no formato correto (dd/mm/yyyy).")
            print("\n")
    print("\n")

    # Ajusta a data para salvar os arquivos
    firstDateReplaced = firstDate.replace("/","-")
    lastDateReplaced = lastDate.replace("/","-")

    # Cria o diretório se não existe
    dirNF = "C:\\Notas Fiscais"
    dirMain = f"{dirNF}\\{firstDateReplaced}_{lastDateReplaced}"
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

    # Salva em um JSON as informações das notas
    try:
        notasJSON = jsonfy(f"{dirNF}\\tempFile.json", resposta)
        nfe = notasJSON["retorno"]["notas_fiscais"]
        idNota = {}
    except KeyError:
        print("Não foram encontrados XMLs de notas fiscais no periodo selecionado.")
        time.sleep(60)
        sys.exit()

    # extrai e armazena os id's e números das notas
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
                
            # Parse o arquivo xml
            tree = etree.parse(f"{dirNF}\\{firstDateReplaced}_{lastDateReplaced}\\{value}.xml")

            # Pega o elemento raiz
            root = tree.getroot()

            # Encontra a tag CFe
            cfe = root.find(".//CFe")

            # Cria uma nova tree com a tag CFe como raiz tornando a tag a nova tag raiz
            new_tree = etree.ElementTree(cfe)

            # Guarda o arquivo alterado
            new_tree.write(f"{dirNF}\\{firstDateReplaced}_{lastDateReplaced}\\{value}.xml", pretty_print=True)
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
        os.remove(f"{dirNF}\\tempFile.json")
    except Exception as error:
        print("\n", error)
        print("Não foi possivel apagar o arquivo temporario.")
        time.sleep(60)

    # Loop para determinar se o usuário deseja mandar o arquivo por email ou não
    """
    while True:
        print("\n")
        askEmail = int(input("Deseja enviar os arquivos XML por email ?: [1]SIM [2]NÃO "))

        if askEmail == 1:
            print("\n")
            nomeRemetente = input("Insira seu nome ou o nome do seu negócio: ")

            while True:
                emailDestinatario = input("Insira o email do destinatario: ")
                if is_valid_email(emailDestinatario):
                    print("\n")
                    print(f"Enviando o arquivo compactado para o endereço de email: {emailDestinatario}")
                    break
                else:
                    print("Formato inválido. Por favor, certifique-se que seu email esteja correto e tente novamente.")
                    print("\n")
            
            # Comprime em um arquivo .zip os arquivos XML
            folder_path = dirMain
            files_to_zip = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]
            zip_name = f'{dirMain}.zip'
            zip_files(files_to_zip, zip_name)

            # Manda o email com as informações necessárias e o arquivo .zip em anexo
            sendEmail(nomeRemetente, emailDestinatario, firstDateReplaced, lastDateReplaced, zip_name, f"{firstDateReplaced}_{lastDateReplaced}")
            break

        elif askEmail == 2:
            break

        else:
            print("Entrada inválida, tente novamente.")
            print("\n")
    """
    # Resumo de informações para o usuario
    print("\n")
    print(f"Foram importados {requestSize} arquivos XML de notas fiscais do dia {firstDate} ao dia {lastDate}. \n")

    # Dá a opção para o usuario de fechar o programa ou fazer outra operação
    print("Pressione [1] para sair ou [2] para realizar outra importação: ", end='', flush=True)

    keyb = keyboard.read_event(suppress=True).name
    if keyb == "1":
        break
    elif keyb == "2":
        print("\n")
        continue
    else:
        break

sys.exit()
