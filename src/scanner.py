"""
scanner.py
-----------

Responsável por localizar normas técnicas em documentos Word (.docx)

Autor: NormFinder
"""

import re
from pathlib import Path
from docx import Document


class Scanner:

    def __init__(self):
        pass

    # ---------------------------------------------------------

    def pesquisar(self, pasta, norma):

        """
        Pesquisa uma norma em todos os documentos Word.

        Retorna uma lista de dicionários.
        """

        resultados = []

        arquivos = Path(pasta).rglob("*.docx")

        regex = self.criar_regex(norma)

        for arquivo in arquivos:

            try:

                texto = self.ler_docx(arquivo)

                ocorrencias = list(regex.finditer(texto))

                if not ocorrencias:
                    continue

                resultado = {

                    "arquivo": arquivo.name,

                    "caminho": str(arquivo),

                    "norma": norma,

                    "quantidade": len(ocorrencias),

                    "trechos": []

                }

                for item in ocorrencias:

                    inicio = max(0, item.start() - 70)
                    fim = min(len(texto), item.end() + 100)

                    trecho = texto[inicio:fim]

                    trecho = trecho.replace("\n", " ")

                    resultado["trechos"].append(trecho)

                resultados.append(resultado)

            except Exception as erro:

                print(f"Erro em {arquivo}")

                print(erro)

        return resultados

    # ---------------------------------------------------------

    def ler_docx(self, caminho):

        documento = Document(caminho)

        texto = []

        for paragrafo in documento.paragraphs:

            texto.append(paragrafo.text)

        return "\n".join(texto)

    # ---------------------------------------------------------

    def criar_regex(self, norma):

        """
        Cria uma expressão regular flexível.

        Exemplo:

        IEC 60335-1

        também encontra

        IEC60335-1

        IEC-60335-1

        IEC 60335 1
        """

        norma = norma.strip()

        partes = re.split(r"[\s\-]+", norma)

        regex = r"[\s\-]*".join(map(re.escape, partes))

        return re.compile(regex, re.IGNORECASE)

    # ---------------------------------------------------------

    def pesquisar_todas_normas(self, pasta):

        """
        Localiza automaticamente todas as normas IEC,
        ISO, EN, UL, ABNT e NBR existentes.

        Muito útil para criar um índice de normas.
        """

        padrao = re.compile(

            r"(IEC|ISO|EN|UL|ABNT|NBR)"
            r"[\s\-]*"
            r"\d{3,6}"
            r"(?:[\-\.]\d+)*",

            re.IGNORECASE

        )

        resultados = []

        for arfor arquivo in Path(pasta).rglob("*.docx"):

    # Ignora arquivos temporários do Word
    if arquivo.name.startswith("~$"):
        continue

    # Garante que é realmente um arquivo
    if not arquivo.is_file():
        continue

            try:

                texto = self.ler_docx(arquivo)

                encontrados = sorted(

                    set(padrao.findall(texto))

                )

                if encontrados:

                    resultados.append({

                        "arquivo": arquivo.name,

                        "normas": encontrados

                    })

            except:

                pass

        return resultados


# ---------------------------------------------------------
# Teste
# ---------------------------------------------------------

if __name__ == "__main__":

    scanner = Scanner()

    pasta = r"C:\Documentos"

    resultados = scanner.pesquisar(
        pasta,
        "IEC 60335-2-15"
    )

    for r in resultados:

        print("=" * 60)

        print(r["arquivo"])

        print(r["quantidade"])

        for trecho in r["trechos"]:

            print()

            print(trecho)