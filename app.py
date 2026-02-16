from flask import Flask, request, render_template, send_from_directory
import pdfplumber
import pandas as pd
import re
import os


app = Flask(__name__)
PLANILHA_DIR = "static/planilhas"
os.makedirs(PLANILHA_DIR, exist_ok=True)

import re
import pdfplumber

def extrair_dados_fgts(arquivo_pdf):
    dados_fgts = []

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue

            linhas = texto.split('\n')

            cnpj = None
            nome = None
            vencimento = None
            valor = None

            for i, linha in enumerate(linhas):

                if "CPF/CNPJ DO EMPREGADOR" in linha.upper() and i + 1 < len(linhas):
                    linha_empregador = linhas[i + 1].strip()

   
                    cnpj_match = re.search(r'\d{2}\.?\d{3}\.?\d{3}', linha_empregador)
                    if cnpj_match:
                        cnpj = cnpj_match.group()


                    linha_limpa = re.sub(r'às\s+\d{2}:\d{2}:\d{2}.*', '', linha_empregador, flags=re.IGNORECASE)

                    nome = linha_limpa.replace(cnpj, '').strip()

            for i, linha in enumerate(linhas):
                if "PAGAR ESTE DOCUMENTO" in linha.upper():

                    data_match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                    if data_match:
                        vencimento = data_match.group()
                        break

                    for j in range(i + 1, min(i + 6, len(linhas))):
                        data_match = re.search(r'\d{2}/\d{2}/\d{4}', linhas[j])
                        if data_match:
                            vencimento = data_match.group()
                            break

                    break

            for i, linha in enumerate(linhas):
                if "VALOR A RECOLHER" in linha.upper():
                    for j in range(i, min(i + 5, len(linhas))):
                        valor_match = re.search(r'\d+\.\d{3},\d{1,2}|\d+,\d{1,2}', linhas[j])
                        if valor_match:
                            valor = valor_match.group()
                            break

            if cnpj and nome and vencimento and valor:
                dados_fgts.append({
                    "CPF/CNPJ": cnpj,
                    "Empregador": nome,
                    "Vencimento": vencimento,
                    "Valor": valor
                })

    return dados_fgts

def extrair_dados_inss(arquivo_pdf):
    dados_inss = []

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if not texto:
                continue

            linhas = texto.split('\n')

            razao_social = ""
            valor_total = ""
            codigo_barras = ""
            vencimento = ""

            for i, linha in enumerate(linhas):

                if "Razão Social" in linha:
                    if i + 1 < len(linhas):
                        razao_social = linhas[i + 1].strip()

                if "Valor Total do Documento" in linha:
                    if i + 1 < len(linhas):
                        valor_total = linhas[i + 1].strip()

                if "DATA DE VENCIMENTO" in linha.upper() or "PAGAR ESTE DOCUMENTO" in linha.upper():

                    data_match = re.search(r'\d{2}/\d{2}/\d{4}', linha)
                    if data_match:
                        vencimento = data_match.group()

                    else:
                        for j in range(i + 1, min(i + 6, len(linhas))):
                            data_match = re.search(r'\d{2}/\d{2}/\d{4}', linhas[j])
                            if data_match:
                                vencimento = data_match.group()
                                break

                if "Documento de Arrecadação de Receitas Federais" in linha:
                    if i + 1 < len(linhas):
                        codigo_barras = linhas[i + 1].strip()[:55]

            if razao_social or valor_total:
                dados_inss.append({
                    "Razão Social": razao_social,
                    "Valor Total do Documento": valor_total,
                    "Código de Barras": codigo_barras,
                    "Vencimento": vencimento
                })

    return dados_inss

def extrair_dados_vinculo(arquivo_pdf):
    dados_vinculo = []
    palavras_ignorar = ["Situação:", "Trabalhando", "CPF:", "Adm:", "Doença"]

    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto = pagina.extract_text()
            if texto:
                linhas = texto.split('\n')
                funcionario = ""
                tipo_vinculo = ""
                tipo_cargo = ""
                valor_liquido = ""

                for i in range(len(linhas)):
                    if "Empr.:" in linhas[i]:
                        funcionario = linhas[i].split("Empr.:")[1].strip()
                        for palavra in palavras_ignorar:
                            funcionario = funcionario.replace(palavra, "")
                        funcionario = re.sub(r'[^a-zA-Z\s]', '', funcionario)
                        funcionario = re.sub(r'\s+', ' ', funcionario).strip()

                    if "Vínculo:" in linhas[i]:
                        tipo_vinculo = linhas[i].split("Vínculo:")[1].strip()
                        if "celetista" in tipo_vinculo.lower():
                            tipo_vinculo = "Celetista"

                    if "CARGO" in linhas[i]:
                        tipo_cargo = linhas[i].split("CARGO:")[1].strip()
                        if "sindico" in tipo_cargo.lower():
                            tipo_cargo = "Síndico"

                    if "Líquido:" in linhas[i]:
                        valor_liquido = linhas[i].split("Líquido:")[1].strip()

                        if (tipo_vinculo.lower() == "celetista" or tipo_cargo.lower() == "síndico") and funcionario and valor_liquido:
                            dados_vinculo.append({
                                "Funcionário": funcionario,
                                "Vínculo": tipo_vinculo,
                                "Líquido": valor_liquido
                            })
                            funcionario = ""
                            tipo_vinculo = ""
                            tipo_cargo = ""
                            valor_liquido = ""

    return dados_vinculo



def exportar_para_planilhas(dados_fgts, dados_inss, dados_vinculo):
    pd.DataFrame(dados_fgts).to_excel(os.path.join(PLANILHA_DIR, "dados_fgts.xlsx"), index=False)
    pd.DataFrame(dados_inss).to_excel(os.path.join(PLANILHA_DIR, "dados_inss.xlsx"), index=False)
    pd.DataFrame(dados_vinculo).to_excel(os.path.join(PLANILHA_DIR, "dados_vinculo.xlsx"), index=False)

@app.route("/", methods=["GET", "POST"])
def index():
    arquivos_gerados = False
    if request.method == "POST":
        arquivos_pdf1 = request.files.getlist("pdf1")
        arquivos_pdf2 = request.files.getlist("pdf2")
        arquivos_pdf3 = request.files.getlist("pdf3")

        todos_dados_fgts = []
        todos_dados_inss = []
        todos_dados_vinculo = []

        for arquivo_pdf in arquivos_pdf1:
            if arquivo_pdf and arquivo_pdf.filename != '':
                todos_dados_fgts.extend(extrair_dados_fgts(arquivo_pdf))

        for arquivo_pdf in arquivos_pdf2:
            if arquivo_pdf and arquivo_pdf.filename != '':
                todos_dados_inss.extend(extrair_dados_inss(arquivo_pdf))

        for arquivo_pdf in arquivos_pdf3:
            if arquivo_pdf and arquivo_pdf.filename != '':
                todos_dados_vinculo.extend(extrair_dados_vinculo(arquivo_pdf))

        exportar_para_planilhas(todos_dados_fgts, todos_dados_inss, todos_dados_vinculo)
        arquivos_gerados = True

    return render_template("index.html", arquivos_gerados=arquivos_gerados)

@app.route("/download/<filename>")
def download(filename):
    return send_from_directory(PLANILHA_DIR, filename, as_attachment=True)

    

if __name__ == "__main__":
    app.run(debug=True, host='0.0.0.0', port=5000)
