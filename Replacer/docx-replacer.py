import os
import docx
import datetime
import calendar
import shutil
import locale
from docx.enum.text import WD_COLOR_INDEX
from docx.shared import RGBColor

data_atual = datetime.datetime.now()

locale.setlocale(locale.LC_ALL, 'pt_PT.utf8')
nome_mes = calendar.month_name[data_atual.month].upper()

diretorio = "./Modelos-pol"

default_num_cart = "NUM_DOC"
numero_cart = input("Qual o número do cartório: ")
default_raz_soc= "CONTROLADOR-CLIENTE"
raz_soc = input("Digite a razão social do cliente: ")
default_fantasia = "CONTROLADOR-FANTASIA"
fantasia = input("Digite o nome-fantasia do cliente (Caso não haja, deixar em branco): ")
if not fantasia:
    fantasia = default_fantasia

default_cidade = "CONTROLADOR-MUNICIPIO"
cidade = input("Digite a cidade: ").upper()
default_uf = "CONTROLADOR-ESTADO"
uf = "GO"
default_email = "CONTROLADOR-EMAIL"
email = input("Digite o email (Caso não haja, deixar em branco): ").lower()
if not email:
    email = default_email

default_site = "CONTROLADOR-SITE"
site = input("Digite o site (Caso não haja, deixar em branco): ").lower()
if not site:
    site = "cartorio2oficio.not.br "

data_default = "dd/mm/aaaa"
data_hoje = data_atual.strftime("%d/%m/%Y")
mes_default = "MES/ANO"
mes_hoje = f"{nome_mes}/{str(data_atual.year)[-2:]}"

diretorio_novo = "./Politicas-Procedimentos_" + numero_cart
if not os.path.exists(diretorio_novo):
    os.makedirs(diretorio_novo)

def main():
    for arquivo in os.listdir(diretorio):

        if arquivo.endswith(".docx"):
            caminho_arquivo = os.path.join(diretorio, arquivo)

            caminho_arquivo_novo = os.path.join(diretorio_novo, numero_cart + " - " + arquivo)
            shutil.copy(caminho_arquivo, caminho_arquivo_novo)

            doc = docx.Document(caminho_arquivo_novo)

            head = doc.sections[0].header

            Dictionary = {default_num_cart : numero_cart, default_raz_soc : raz_soc, default_fantasia : fantasia, default_cidade : cidade, default_uf : uf, default_site : site, default_email : email, data_default : data_hoje, mes_default : mes_hoje}

            for i in Dictionary:
                for p in doc.paragraphs:
                    for run in p.runs:
                        if i in run.text:
                            text = run.text.replace(i, Dictionary[i])
                            run.text = ""
                            for idx, part in enumerate(text.split(i)):
                                run.add_text(part)
                                if idx < len(text.split(i)) - 1:
                                    run.add_text(i)
                                    font = run.font
                                    font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
                                    font.highlight_color = WD_COLOR_INDEX.YELLOW
                                    font.bold = True
                        
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