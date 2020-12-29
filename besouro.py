import requests
import time
import json
import sys
import re
import csv
import os
import datetime
from bs4 import BeautifulSoup

# Disables SSL truststore warnings, as we use sessions' verify=False
import urllib3
urllib3.disable_warnings()

IS_DEBUG=False

OPERACAO_COMPRA='1'
OPERACAO_VENDA='3'

def do_login(session, user, password):
    response_ini = session.get('https://portalinvestidor.tesourodireto.com.br', verify=False)
    soup = BeautifulSoup(response_ini.content, 'html.parser')
    verif_token = soup.find_all('input')[0].get("value")

    print ("Logging in Portal do Investidor")
    session.post('https://portalinvestidor.tesourodireto.com.br/Login/ValidateLogin',
            data={'UserCpf': user,
                'UserPassword': password,
                'g-recaptcha-response': '',
                '__RequestVerificationToken': verif_token},
            headers={
                'Origin': 'https://portalinvestidor.tesourodireto.com.br',
                'Referer': 'https://portalinvestidor.tesourodireto.com.br/'
            },
            verify=False)

def consulta_operacoes_json(tipo_operacao, session):
    response_ini = session.get('https://portalinvestidor.tesourodireto.com.br/Consulta', verify=False)
    soup = BeautifulSoup(response_ini.content, 'html.parser')
    verif_token = soup.find_all('form', action='/Consulta')[0].find_all("input", type="hidden")[0].get("value")

    print ("Listing operacoes={}".format(tipo_operacao))
    response_consulta = session.post("https://portalinvestidor.tesourodireto.com.br/Consulta/ConsultarOperacoes",
            headers={
                '__RequestVerificationToken': verif_token,
                'Accept': 'application/json, text/javascript, */*; q=0.01',
                'Origin': 'https://portalinvestidor.tesourodireto.com.br',
                'Referer': 'https://portalinvestidor.tesourodireto.com.br/Consulta',
                'X-Requested-With': 'XMLHttpRequest'
            },
            data={"Operacao": tipo_operacao,"InstituicaoFinanceira":"386","DataInicial":"01/01/2016","DataFinal":"19/07/2020"})
    ops_json = json.loads(response_consulta.content)['Operacoes']

    print("Found {} operations with tipo={}".format(len(ops_json), tipo_operacao))
    return ops_json


def get_info_titulo(soup, index):
    content_str = soup.find_all("p", class_="td-protocolo-info")[index].span.text
    regex_result = re.findall('([0-9.,]+)', content_str)[0]
    cleaned = regex_result.replace('.','').replace(',', '.')
    return cleaned

def processa_titulos (session, titulos, operacao):
    csv=[]

    for tit in titulos:
        # Skipa os nao liquidados
        if tit['Situacao'] != 'Realizado':
            continue

        url = "https://portalinvestidor.tesourodireto.com.br/Protocolo/{}/{}".format(tit['CodigoProtocolo'], operacao)
        print ("Fetching protocolo={}, url={}".format(tit['CodigoProtocolo'], url))
        response = session.get(url,
            headers={
            'Referer': 'https://portalinvestidor.tesourodireto.com.br/Consulta',
            })
        soup = BeautifulSoup(response.content, 'html.parser')

        titulo         = soup.find(class_='td-protocolo-info-titulo').text
        quantidade     = float(get_info_titulo(soup, 0))
        valor_unitario = str(float(get_info_titulo(soup, 1))).replace('.', ',')
        rentabilidade  = str(float(get_info_titulo(soup, 2))/100).replace('.', ',')
        liquido        = float(get_info_titulo(soup, 3))
        taxa_corretora = str(-float(get_info_titulo(soup, 4))).replace('.', ',')
        taxa_b3        = str(-float(get_info_titulo(soup, 5))).replace('.', ',')
        valor_bruto    = float(get_info_titulo(soup, 6))

        if operacao == OPERACAO_VENDA:
            quantidade = -quantidade
            liquido = -liquido
            valor_bruto = -valor_bruto

        quantidade = str(quantidade).replace('.', ',')
        liquido = str(liquido).replace('.', ',')
        valor_bruto = str(valor_bruto).replace('.', ',')
        
        data_operacao = str(datetime.datetime.strptime(tit["DataOperacao"], "%d/%m/%Y"))[0:10]
        print(data_operacao)
        line = [tit["TipoOperacao"], data_operacao, tit["CodigoProtocolo"],
                titulo, rentabilidade, 
                quantidade, valor_unitario, valor_bruto,
                taxa_corretora, taxa_b3, liquido]

        if IS_DEBUG:
            file = open ('protocolos/{}.html'.format(tit['CodigoProtocolo']), "w")
            file.write(str(response.content))
            file.close()
            print(url)
            print(line)

        csv.append(line)

    return csv

def add_header_formulas(csv):
    csv_with_header=[
        ['Tipo', 'Data', 'Protocolo', 'Titulo',
        'Rent %', 'Qtd', '$ Unit', '$ Compra Bruto',
        'Tx Corret', 'Taxa B3', '$ Compra Liq',
        'Compras Anteriores', 'Vendas Anteriores', 'Qtd Vendas Futuro',
        'Saldo', 'Saldo Compra Liq',
        'PU Atual', 'Saldo Atual', 'Lucro/Prej']]

    line_num = 2 # Formula reference, starts at line 2 in spreadsheet, as 1 is header
    csv_index = 0
    for line in csv:
        line_str = str(line_num)
        csv_with_header.append(csv[csv_index] + [
            # Col L - Compras Anteriores
            '=if(AND(F@<>"";F@>0);sumifs(F$1:F#;F$1:F#;">0";D$1:D#;D@);"")'.replace('@', line_str).replace('#', str(line_num - 1)),
            # Col M - Vendas Anteriores
            '=if(AND(F@<>"";F@>0);SUMIFS(F$1:F#;F$1:F#;"<0";D$1:D#;D@);"")'.replace('@', line_str).replace('#', str(line_num - 1)),
            # Col N - Vendas Futuras
            '=if(and(F@<>"";F@>0);sumifs(F@:F$9999;F@:F$9999;"<0";D@:D$9999;D@);"")'.replace('@', line_str).replace('#', str(line_num - 1)),
            # Col O - Saldo
            '=if(and(F@<>"";F@>0);floor(min(max(L@+M@+F@+N@;0);F@);0,01);"")'.replace('@', line_str), 
            # Col P - Saldo Compra Liq
            '=if(O@>0;O@*G@+J@+I@;"")'.replace('@', line_str),                                        
            # Col Q - PU Atual
            '=IF(ISERROR(VLOOKUP(D@;\'Preços Tesouro Direto\'!$A$36:$F$100;6;FALSE));"";VLOOKUP(D@;\'Preços Tesouro Direto\'!$A$36:$F$100;6;FALSE))'.replace('@', line_str),
            '=if(and(O@>0;O@<>"");O@*Q@;"")'.replace('@', line_str), # Saldo Atual
            '=if(O@>0;R@-K@;"")'.replace('@', line_str)])            # Lucro/Prej
        line_num += 1
        csv_index += 1

    return csv_with_header

with requests.Session() as session:
    user = os.environ.get('BESOURO_CPF')
    password = os.environ.get('BESOURO_PWD')
    do_login(session, user, password)

    json_compras = consulta_operacoes_json(OPERACAO_COMPRA, session)
    compras = processa_titulos(session, json_compras, OPERACAO_COMPRA)

    json_vendas = consulta_operacoes_json(OPERACAO_VENDA, session)
    vendas = processa_titulos(session, json_vendas, OPERACAO_VENDA)

    print ("Generating CSV")
    sorted_csv = add_header_formulas(sorted(compras + vendas, key=lambda row: row[1]))
    with open('operacoes.csv', 'w') as csvfile:
        csvwriter = csv.writer(csvfile, delimiter=';', quoting=csv.QUOTE_MINIMAL)
        csvwriter.writerows(sorted_csv)
