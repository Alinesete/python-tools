import os
import docx
import datetime
import calendar
import shutil

data_atual = datetime.datetime.now()
nome_mes = calendar.month_name[data_atual.month].upper()

diretorio = "./Modelos-pol"
diretorio_novo = "./Politicas-Procedimentos"
if not os.path.exists(diretorio_novo):
    os.makedirs(diretorio_novo)

default_num_cart = "NUM_DOC"
numero_cart = input("Qual o número do cartório: ")
default_raz_soc= "CONTROLADOR-CLIENTE"
raz_soc = input("Digite a razão social do cliente: ")
default_fantasia = "CONTROLADOR-FANTASIA"
fantasia = input("Digite o nome-fantasia do cliente: ")
default_cidade = "COD_MUNICIPIO"
cidade = input("Digite a cidade: ")
default_uf = "COD_UF"
uf = "GO"
data_default = "dd/mm/aaaa"
data_hoje = data_atual.strftime("%d/%m/%Y")
mes_default = "MES/ANO"
mes_hoje = f"{nome_mes}/{str(data_atual.year)[-2:]}"

def main():
    for arquivo in os.listdir(diretorio):

        if arquivo.endswith(".docx"):
            caminho_arquivo = os.path.join(diretorio, arquivo)

            caminho_arquivo_novo = os.path.join(diretorio_novo, numero_cart + " - " + arquivo)
            shutil.copy(caminho_arquivo, caminho_arquivo_novo)

            doc = docx.Document(caminho_arquivo_novo)

            head = doc.sections[0].header

            Dictionary = {default_num_cart : numero_cart, default_raz_soc : raz_soc, default_fantasia : fantasia, default_cidade : cidade, default_uf : uf, data_default : data_hoje, mes_default : mes_hoje}
            for i in Dictionary:
                for p in doc.paragraphs:
                    if p.text.find(i)>=0:
                        p.text=p.text.replace(i,Dictionary[i])
                        
                for tabela in doc.tables:
                    for linha in tabela.rows:
                        for celula in linha.cells:
                            for i in Dictionary.keys():
                                if i in celula.text:
                                    celula.text = celula.text.replace(i, Dictionary[i])

                for tabela in head.tables:
                    for linha in tabela.rows:
                        for celula in linha.cells:
                            for i in Dictionary.keys():
                                if i in celula.text:
                                    celula.text = celula.text.replace(i, Dictionary[i])

            doc.save(caminho_arquivo_novo)

if __name__ == '__main__':
    main()