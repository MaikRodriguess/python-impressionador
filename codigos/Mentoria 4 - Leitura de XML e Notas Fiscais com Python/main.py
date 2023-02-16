import xmltodict

with open('NFs Finais/DANFEBrota.xml', 'rb') as arquivo:
    documento = xmltodict.parse(arquivo)
print(documento)
dic_notafiscal = documento['nfeProc']['NFe']['infNFe']

valor_total = dic_notafiscal['total']['ICMSTot']['vNF']
cnpj_vendeu = dic_notafiscal['emit']['CNPJ']
nome_vendeu = dic_notafiscal['emit']['xNome']
cpf_comprou = dic_notafiscal['dest']['CPF']
nome_comprou = dic_notafiscal['dest']['xNome']
resposta = {'valor_total': valor_total,
            'cnpj_vendeu': cnpj_vendeu,
            'nome_vendeu': nome_vendeu,
            'cpf_comprou': cpf_comprou,
            'nome_comprou': nome_comprou
            }
# # Convertendo o dicionario em Tupla
# resposta = resposta.items()
#
# for item in resposta:
#     dicionario, valor = item
#     print(dicionario + ' = ' + valor)

print(resposta)



























