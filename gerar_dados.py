"""
Gerador de dados FICTÍCIOS do extrator de PER/DCOMP.
Semente fixa: rodar de novo gera exatamente os mesmos dados.

Todos os nomes, CNPJs, processos e números de PER/DCOMP são inventados.
Semente fixa: rodar de novo gera exatamente os mesmos dados.
"""
import csv
import json
import random
import shutil
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

random.seed(2026)


# ------------------------------------------------------------
# Utilidades
# ------------------------------------------------------------

def dv_cnpj(base12):
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6] + pesos1
    d1 = sum(int(n) * p for n, p in zip(base12, pesos1)) % 11
    d1 = 0 if d1 < 2 else 11 - d1
    d2 = sum(int(n) * p for n, p in zip(base12 + str(d1), pesos2)) % 11
    d2 = 0 if d2 < 2 else 11 - d2
    return f"{d1}{d2}"


def gerar_cnpj():
    base8 = "".join(random.choice("0123456789") for _ in range(8))
    base12 = base8 + "0001"
    return base12 + dv_cnpj(base12)


def cnpj_valido(cnpj):
    n = "".join(c for c in cnpj if c.isdigit())
    return len(n) == 14 and n[12:] == dv_cnpj(n[:12])


def fmt_cnpj(n):
    return f"{n[:2]}.{n[2:5]}.{n[5:8]}/{n[8:12]}-{n[12:]}"


def gerar_perdcomp():
    # Formato ilustrativo: 5.5.6.1.1.2-4 (ajuste ao formato que você usa)
    b1 = f"{random.randint(0, 99999):05d}"
    b2 = f"{random.randint(0, 99999):05d}"
    d = date(2023, 1, 1) + timedelta(days=random.randint(0, 1000))
    b3 = d.strftime("%d%m%y")
    t1 = random.choice("123")
    t2 = random.choice("1234")
    cc = f"{random.randint(1, 99):02d}"
    dvs = f"{random.randint(0, 9999):04d}"
    return f"{b1}.{b2}.{b3}.{t1}.{t2}.{cc}-{dvs}", d


def gerar_processo(usados):
    while True:
        p = (f"{random.randint(10000, 19999)}.{random.randint(100000, 999999)}"
             f"/{random.randint(2022, 2025)}-{random.randint(10, 99)}")
        if p not in usados:
            usados.add(p)
            return p


def nome_pasta(processo):
    return processo.replace("/", "_")


# ------------------------------------------------------------
# PROJETO 1: extrator-perdcomp
# ------------------------------------------------------------

NOMES = [
    "Comercial Aurora Exemplo Ltda", "Metalúrgica Horizonte Fictícia S.A.",
    "Distribuidora Pedra Azul Ltda", "Tecnologia Vale Verde Ltda",
    "Transportes Rio Claro Exemplo Ltda", "Indústria de Plásticos Sol Nascente Ltda",
    "Alimentos Campo Dourado Fictícios Ltda", "Construtora Serra Alta Exemplo S.A.",
    "Farmácia Bom Remédio Fictícia Ltda", "Confecções Fio de Ouro Ltda",
    "Gráfica Tinta Viva Exemplo Ltda", "Móveis Cedro Bravo Fictícios Ltda",
    "Editora Folha Branca Exemplo Ltda",
]
UFS = ["SP", "RJ", "MG", "PR", "RS", "SC", "BA", "GO", "PE", "CE", "ES", "DF", "MT"]
CIDADES = ["Cidade Exemplo", "Vila Fictícia", "Porto Imaginário", "Monte Teste",
           "Campos de Demonstração", "Lagoa Simulada", "Serra do Exemplo",
           "Boa Vista de Mentira", "Ribeirão Fictício", "Santa Amostra",
           "São Dado Falso", "Nova Simulação", "Jardim das Provas"]

TEMPLATES = {
    "A": (
        "RECIBO DE TRANSMISSÃO - PER/DCOMP (DOCUMENTO FICTÍCIO)\n"
        "------------------------------------------------------\n"
        "Contribuinte: {razao}\n"
        "CNPJ: {cnpj}\n"
        "Processo administrativo: {proc}\n"
        "Número do PER/DCOMP: {perd}\n"
        "Número do recibo: {recibo}\n"
        "Data da transmissão: {data}\n"
        "Tipo de crédito: Pagamento indevido ou a maior\n"
        "Valor do crédito: R$ {valor}\n"
    ),
    "B": (
        "DECLARAÇÃO DE COMPENSAÇÃO (FICTÍCIA)\n\n"
        "Empresa: {razao} - CNPJ {cnpj}\n"
        "Protocolo interno: {recibo}\n"
        "Processo: {proc}\n\n"
        "PER/DCOMP nº {perd}\n"
        "Transmitido em {data}\n"
        "Valor original do crédito: R$ {valor}\n"
        "Observação: documento gerado para fins de demonstração.\n"
    ),
    "C": (
        "COMPROVANTE - PEDIDO ELETRÔNICO (FICTÍCIO)\n"
        "Identificação do declarante\n"
        "  Razão social .......: {razao}\n"
        "  CNPJ ...............: {cnpj}\n"
        "Dados do pedido\n"
        "  N.º do processo ....: {proc}\n"
        "  Protocolo ..........: {recibo}\n"
        "  Data ...............: {data}\n"
        "  Valor (R$) .........: {valor}\n"
        "Identificação do documento\n"
        "  PERDCOMP: {perd}\n"
    ),
}


def valor_br():
    v = random.randint(1500, 480000) + random.choice([0, 0.25, 0.5, 0.75])
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def recibo():
    return "".join(random.choice("0123456789") for _ in range(13))


def corromper_ocr(txt):
    trocas = {"0": "O", "1": "l"}
    saida, qtd = [], 0
    for ch in txt:
        if ch in trocas and qtd < 3 and random.random() < 0.3:
            saida.append(trocas[ch])
            qtd += 1
        else:
            saida.append(ch)
    out = "".join(saida)
    if out == txt:
        out = txt.replace("0", "O", 1)
    return out[:3] + " " + out[3:]


def quebrar_numero(txt):
    pos = txt.index(".", txt.index(".") + 1) + 1
    return txt[:pos] + "\n" + txt[pos:]


def projeto_extrator(raiz):
    pasta = raiz
    if (pasta / "processos").exists():
        shutil.rmtree(pasta / "processos")
    (pasta / "processos").mkdir(parents=True)

    # Cadastro de CNPJs (a "outra planilha")
    empresas = []
    for i, nome in enumerate(NOMES):
        empresas.append({"cnpj": gerar_cnpj(), "razao_social": nome,
                         "uf": UFS[i], "municipio": CIDADES[i]})
    cadastro = pd.DataFrame(empresas)
    cadastro["cnpj"] = cadastro["cnpj"].apply(fmt_cnpj)

    # Empresa que NÃO está no cadastro
    cnpj_fora = gerar_cnpj()
    razao_fora = "Importadora Estrela do Norte Fictícia Ltda"

    # Cenários dos 30 processos
    cenarios = (["limpo"] * 17 + ["quebrado"] * 3 + ["ocr"] * 3 + ["outro_cnpj"] * 2 +
                ["sem_documento", "dois_perdcomp", "sem_numero",
                 "cnpj_invalido", "cnpj_fora_cadastro"])
    random.shuffle(cenarios)

    usados, linhas, gabarito = set(), [], []
    templates = list(TEMPLATES)

    for idx, cen in enumerate(cenarios):
        emp = empresas[idx % len(empresas)]
        cnpj_num = emp["cnpj"]
        razao = emp["razao_social"]
        proc = gerar_processo(usados)
        perd, dt = gerar_perdcomp()
        tpl = TEMPLATES[templates[idx % 3]]

        cnpj_planilha_num = cnpj_num
        status, motivo, esperado = "ok", "", perd

        if cen == "cnpj_invalido":
            cnpj_planilha_num = cnpj_num[:13] + str((int(cnpj_num[13]) + 1) % 10)
            status, motivo, esperado = "revisar", "CNPJ da planilha com dígito verificador inválido", ""
        elif cen == "cnpj_fora_cadastro":
            cnpj_num, razao = cnpj_fora, razao_fora
            cnpj_planilha_num = cnpj_fora
            status, motivo, esperado = "revisar", "CNPJ não consta no cadastro", ""

        # Formato do CNPJ na planilha: metade formatado, metade só dígitos
        cnpj_planilha = fmt_cnpj(cnpj_planilha_num) if idx % 2 == 0 else cnpj_planilha_num

        pasta_proc = pasta / "processos" / nome_pasta(proc)
        pasta_proc.mkdir()

        dados = dict(razao=razao, cnpj=fmt_cnpj(cnpj_num), proc=proc, perd=perd,
                     recibo=recibo(), data=dt.strftime("%d/%m/%Y"), valor=valor_br())

        if cen == "sem_documento":
            status, motivo, esperado = "revisar", "Pasta do processo sem documento", ""
        else:
            if cen == "quebrado":
                dados["perd"] = quebrar_numero(perd)
            elif cen == "ocr":
                dados["perd"] = corromper_ocr(perd)
            elif cen == "outro_cnpj":
                outro = gerar_cnpj()
                dados["cnpj"] = fmt_cnpj(outro)
                status, motivo, esperado = "revisar", "CNPJ do documento diverge da planilha", ""
            elif cen == "dois_perdcomp":
                perd2, _ = gerar_perdcomp()
                dados["perd"] = f"{perd} (original)\n  Retificadora: {perd2}"
                status, motivo, esperado = "revisar", "Mais de um número de PER/DCOMP no documento", ""
            elif cen == "sem_numero":
                dados["perd"] = "(não localizado)"
                status, motivo, esperado = "revisar", "Documento sem número de PER/DCOMP", ""
            texto = tpl.format(**dados)
            (pasta_proc / "recibo.txt").write_text(texto, encoding="utf-8")

        linhas.append({"processo": proc, "cnpj": cnpj_planilha,
                       "perdcomp": "", "status": "", "observacao": ""})
        gabarito.append({"processo": proc, "cnpj_planilha": cnpj_planilha,
                         "perdcomp_esperado": esperado, "status_esperado": status,
                         "motivo": motivo, "dificuldade": cen})

    df = pd.DataFrame(linhas)
    xlsx = pasta / "planilha_processos.xlsx"
    with pd.ExcelWriter(xlsx, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="processos", index=False)
        cadastro.to_excel(w, sheet_name="cadastro_cnpj", index=False)
        pd.DataFrame({"Legenda": [
            "Aba 'processos': o programa deve preencher as colunas perdcomp, status e observacao.",
            "Aba 'cadastro_cnpj': base usada para conferir o CNPJ de cada processo.",
            "Metade dos CNPJs da planilha está formatada e metade só com dígitos (de propósito).",
            "Todos os dados são fictícios.",
        ]}).to_excel(w, sheet_name="legenda", index=False)

    with open(pasta / "gabarito.csv", "w", newline="", encoding="utf-8") as f:
        wr = csv.DictWriter(f, fieldnames=list(gabarito[0].keys()))
        wr.writeheader()
        wr.writerows(gabarito)

    # sanidade: CNPJs do cadastro devem ser válidos
    assert all(cnpj_valido(c) for c in cadastro["cnpj"])
    return len(linhas)


# ------------------------------------------------------------
# Formatação das planilhas
# ------------------------------------------------------------

def formatar(caminho):
    wb = load_workbook(caminho)
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                c.font = Font(name="Arial", size=10)
        for c in ws[1]:
            c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="305496")
            c.alignment = Alignment(horizontal="left")
        for col in ws.columns:
            largura = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max(largura + 2, 12), 70)
        ws.freeze_panes = "A2"
    wb.save(caminho)


if __name__ == "__main__":
    n = projeto_extrator(Path("."))
    formatar(Path("planilha_processos.xlsx"))
    print(f"{n} processos gerados")
