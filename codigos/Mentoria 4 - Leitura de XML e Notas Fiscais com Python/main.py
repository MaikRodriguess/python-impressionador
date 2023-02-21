import xmltodict

# Abrir e ler arquivo
def ler_xml_danfe(nota):
    with open(nota, 'rb') as arquivo:
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
    return resposta

def ler_xml_servico(nota):
    with open(nota, 'rb') as arquivo:
        documento = xmltodict.parse(arquivo)

    dic_notafiscal = documento['ConsultarNfseResposta']['ListaNfse']['CompNfse']['Nfse']['InfNfse']

    valor_total = dic_notafiscal['Servico']['Valores']['ValorServicos']
    cnpj_vendeu = dic_notafiscal['PrestadorServico']['IdentificacaoPrestador']['Cnpj']
    nome_vendeu = dic_notafiscal['PrestadorServico']['RazaoSocial']
    cpf_comprou = dic_notafiscal['TomadorServico']['IdentificacaoTomador']['CpfCnpj']['Cnpj']
    nome_comprou = dic_notafiscal['TomadorServico']['RazaoSocial']
    produtos = dic_notafiscal['Servico']['Discriminacao']

    resposta = {'valor_total': [valor_total],
                'cnpj_vendeu': [cnpj_vendeu],
                'nome_vendeu': [nome_vendeu],
                'cpf_comprou': [cpf_comprou],
                'nome_comprou': [nome_comprou],
                'Lista_produtos': [produtos]
                }
    return resposta

import os

lista_arquivos = os.listdir('NFs Finais') # Lista os nomes de arquivos de uma pasta

for arquivo in lista_arquivos: # Para cada arquivo
    if 'xml' in arquivo: # Se tem xml no nome do arquivo
        if 'DANFE' in arquivo: # Se tem DANFE no nome do arquivo
            print(ler_xml_danfe(f'NFs Finais/{arquivo}')) # rodar o leitor de XML de DANFE para esse arquivo
        else:
            print(ler_xml_servico(f'NFs Finais/{arquivo}'))

# import pandas as pd
# # tabela = pd.DataFrame.from_dict(resposta)
# # tabela.to_excel('NFs.xlsx')
























