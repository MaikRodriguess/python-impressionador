import xmltodict

# Abrir e ler arquivo
with open('NFs Finais/DANFENespresso.xml', 'rb') as arquivo:
    documento = xmltodict.parse(arquivo)

dic_notafiscal = documento['nfeProc']['NFe']['infNFe']

valor_total = dic_notafiscal['total']['ICMSTot']['vNF']
cnpj_vendeu = dic_notafiscal['emit']['CNPJ']
nome_vendeu = dic_notafiscal['emit']['xNome']
cpf_comprou = dic_notafiscal['dest']['CPF']
nome_comprou = dic_notafiscal['dest']['xNome']
produtos = dic_notafiscal['det']
lista_produtos = []
for produto in produtos:
    nome_produto = produto['prod']['xProd']
    valor_produto = produto['prod']['vProd']
    lista_produtos.append((nome_produto, valor_produto))

resposta = {'valor_total': [valor_total],
            'cnpj_vendeu': [cnpj_vendeu],
            'nome_vendeu': [nome_vendeu],
            'cpf_comprou': [cpf_comprou],
            'nome_comprou': [nome_comprou],
            'Lista_produtos': [lista_produtos]
            }
print(resposta)
# Convertendo o dicionario em Tupla
resposta = resposta.items()

for item in resposta:
    dicionario, valor = item
    print(dicionario, ' = ', valor)

import pandas as pd

tabela = pd.DataFrame.from_dict(resposta)
tabela.to_excel('NFs.xlsx')
























