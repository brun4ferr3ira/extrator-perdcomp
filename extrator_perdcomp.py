"""
Extrator de números de PER/DCOMP (dados fictícios).

Para cada processo da planilha, o programa abre a pasta do processo, confere o CNPJ
e extrai o número do PER/DCOMP do documento. O fluxo é:

    1. regra (regex) primeiro;
    2. IA só quando o texto está sujo demais para a regra;
    3. conferência por código sempre: a resposta da IA só vale se estiver no documento;
    4. dúvida vai para revisão humana, com o motivo.

A planilha original nunca é alterada: o resultado é salvo na pasta 'saida'.
"""
import argparse
import json
import os
import re
from collections import Counter
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import PatternFill

PLANILHA_PADRAO = Path("planilha_processos.xlsx")
PASTA_PROCESSOS = Path("processos")
PASTA_SAIDA = Path("saida")

# Formato ILUSTRATIVO: 5.5.6.1.1.2-4. Ajuste ao formato real, se quiser.
PADRAO_PERDCOMP = r"(?<!\d)\d{5}\.\d{5}\.\d{6}\.\d\.\d\.\d{2}-\d{4}(?!\d)"
PADRAO_CNPJ = r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}"

CONFUSOES = str.maketrans({"O": "0", "o": "0", "l": "1", "I": "1"})

MODELO_PADRAO = "claude-haiku-5-5"
INSTRUCAO_IA = (
    "Você extrai o número de um PER/DCOMP de um documento fiscal. "
    "O número tem o formato NNNNN.NNNNN.DDMMAA.N.N.NN-NNNN (dígitos separados por pontos e hífen). "
    "O texto pode ter quebras de linha, espaços no meio e erros de leitura, como a letra O no lugar "
    "do zero ou a letra l no lugar do um. Ignore outros números, como protocolo, recibo e processo. "
    'Responda SOMENTE com JSON, no formato {"perdcomp": "<número no formato correto>"}. '
    'Se não houver número, responda {"perdcomp": null}. Nunca invente um número.'
)


# ------------------------------------------------------------
# CNPJ
# ------------------------------------------------------------

def so_digitos(texto):
    return "".join(c for c in str(texto) if c.isdigit())


def dv_cnpj(base12):
    pesos1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
    pesos2 = [6] + pesos1
    d1 = sum(int(n) * p for n, p in zip(base12, pesos1)) % 11
    d1 = 0 if d1 < 2 else 11 - d1
    d2 = sum(int(n) * p for n, p in zip(base12 + str(d1), pesos2)) % 11
    d2 = 0 if d2 < 2 else 11 - d2
    return f"{d1}{d2}"


def cnpj_valido(cnpj):
    n = so_digitos(cnpj)
    return len(n) == 14 and n[12:] == dv_cnpj(n[:12])


def cnpjs_no_texto(texto):
    return {so_digitos(c) for c in re.findall(PADRAO_CNPJ, texto)}


def mascarar_cnpj(texto):
    return re.sub(PADRAO_CNPJ, "[CNPJ]", texto)


# ------------------------------------------------------------
# Extração por regra
# ------------------------------------------------------------

def juntar_quebras(texto):
    """Remove espaços e quebras de linha ao redor de pontos e hífens."""
    return re.sub(r"\s*([.\-])\s*", r"\1", texto)


def achar_numeros_por_regra(texto):
    return sorted(set(re.findall(PADRAO_PERDCOMP, juntar_quebras(texto))))


def texto_tolerante(texto):
    """Corrige confusões comuns de leitura (O/0, l/1) e tira todos os espaços."""
    return re.sub(r"\s+", "", texto.translate(CONFUSOES))


def fundamentado(numero, texto):
    """A resposta da IA só vale se o número realmente estiver no documento."""
    return numero in texto_tolerante(texto)


# ------------------------------------------------------------
# IA
# ------------------------------------------------------------

class ClienteSimulado:
    """NÃO é IA: serve para testar o fluxo sem chave de API."""
    nome = "simulada"

    def responder(self, texto):
        achados = re.findall(PADRAO_PERDCOMP, texto_tolerante(texto))
        return json.dumps({"perdcomp": achados[0] if len(achados) == 1 else None})


class ClienteClaude:
    nome = "claude"

    def __init__(self):
        try:
            import anthropic
        except ImportError:
            raise SystemExit("Instale o pacote da IA: pip install anthropic")
        if not os.environ.get("ANTHROPIC_API_KEY"):
            raise SystemExit("Defina a variável de ambiente ANTHROPIC_API_KEY.")
        self.cliente = anthropic.Anthropic()
        self.modelo = os.environ.get("CLAUDE_MODEL", MODELO_PADRAO)

    def responder(self, texto):
        resposta = self.cliente.messages.create(
            model=self.modelo, max_tokens=200, system=INSTRUCAO_IA,
            messages=[{"role": "user", "content": texto}])
        return resposta.content[0].text


def interpretar_resposta(bruta):
    limpa = re.sub(r"^```(?:json)?|```$", "", bruta.strip(), flags=re.MULTILINE).strip()
    try:
        dados = json.loads(limpa)
    except json.JSONDecodeError:
        return None, "resposta da IA não é um JSON válido"
    valor = dados.get("perdcomp") if isinstance(dados, dict) else None
    if not valor:
        return None, "a IA não encontrou número no documento"
    return str(valor).strip(), ""


def extrair_com_ia(texto, cliente):
    """Devolve (número, motivo). O número só é devolvido se passar na conferência."""
    try:
        bruta = cliente.responder(mascarar_cnpj(texto))
    except Exception as erro:
        return None, f"falha ao chamar a IA ({type(erro).__name__})"
    numero, motivo = interpretar_resposta(bruta)
    if not numero:
        return None, motivo
    if not re.fullmatch(PADRAO_PERDCOMP, numero):
        return None, "número da IA fora do formato esperado"
    if not fundamentado(numero, texto):
        return None, "número da IA não consta no documento"
    return numero, ""


# ------------------------------------------------------------
# Processamento
# ------------------------------------------------------------

def carregar_planilha(caminho):
    wb = load_workbook(caminho)
    cadastro = {so_digitos(c[0]) for c in wb["cadastro_cnpj"].iter_rows(min_row=2, values_only=True) if c[0]}
    return wb, cadastro


def ler_documentos(processo):
    pasta = PASTA_PROCESSOS / processo.replace("/", "_")
    if not pasta.is_dir():
        return None
    arquivos = sorted(p for p in pasta.iterdir() if p.suffix.lower() == ".txt")
    if not arquivos:
        return None
    return "\n".join(a.read_text(encoding="utf-8", errors="replace") for a in arquivos)


def revisar(motivo):
    return {"perdcomp": "", "status": "revisar", "observacao": motivo, "origem": ""}


def processar_processo(processo, cnpj, cadastro, cliente):
    cnpj_limpo = so_digitos(cnpj)
    if not cnpj_valido(cnpj_limpo):
        return revisar("CNPJ da planilha com dígito verificador inválido")
    if cnpj_limpo not in cadastro:
        return revisar("CNPJ não consta no cadastro")

    texto = ler_documentos(processo)
    if texto is None:
        return revisar("Pasta do processo sem documento")
    if cnpj_limpo not in cnpjs_no_texto(texto):
        return revisar("CNPJ do documento diverge da planilha")

    numeros = achar_numeros_por_regra(texto)
    if len(numeros) > 1:
        return revisar(f"Mais de um número de PER/DCOMP no documento ({len(numeros)})")
    if len(numeros) == 1:
        return {"perdcomp": numeros[0], "status": "ok",
                "observacao": "Extraído por regra", "origem": "regra"}

    if cliente is None:
        return revisar("Número não encontrado pela regra (IA desligada)")
    numero, motivo = extrair_com_ia(texto, cliente)
    if numero:
        return {"perdcomp": numero, "status": "ok",
                "observacao": "Extraído pela IA e conferido no documento", "origem": "ia"}
    return revisar(f"Número não localizado: {motivo}")


def processar_planilha(planilha, cliente):
    wb, cadastro = carregar_planilha(planilha)
    ws = wb["processos"]
    cab = {c.value: c.column for c in ws[1]}
    verde = PatternFill("solid", fgColor="C6E0B4")
    vermelho = PatternFill("solid", fgColor="FFC7CE")

    resultados = []
    for linha in range(2, ws.max_row + 1):
        processo = ws.cell(row=linha, column=cab["processo"]).value
        if not processo:
            continue
        cnpj = ws.cell(row=linha, column=cab["cnpj"]).value
        r = processar_processo(str(processo), cnpj, cadastro, cliente)
        ws.cell(row=linha, column=cab["perdcomp"], value=r["perdcomp"])
        ws.cell(row=linha, column=cab["status"], value=r["status"]).fill = (
            verde if r["status"] == "ok" else vermelho)
        ws.cell(row=linha, column=cab["observacao"], value=r["observacao"])
        resultados.append({"processo": str(processo), **r})
    ws.column_dimensions[ws.cell(row=1, column=cab["perdcomp"]).column_letter].width = 34
    ws.column_dimensions[ws.cell(row=1, column=cab["observacao"]).column_letter].width = 60
    return wb, resultados


def resumo(resultados):
    c = Counter((r["status"], r["origem"]) for r in resultados)
    linhas = [f"{len(resultados)} processos analisados",
              f"  ok por regra: {c[('ok', 'regra')]}",
              f"  ok pela IA:   {c[('ok', 'ia')]}",
              f"  para revisão: {c[('revisar', '')]}"]
    for r in resultados:
        if r["status"] == "revisar":
            linhas.append(f"    - {r['processo']}: {r['observacao']}")
    return "\n".join(linhas)


def main():
    ap = argparse.ArgumentParser(description="Extrator de números de PER/DCOMP")
    ap.add_argument("--planilha", type=Path, default=PLANILHA_PADRAO)
    ap.add_argument("--ia", choices=["simulada", "claude", "nenhuma"], default="simulada",
                    help="simulada (padrão, sem chave), claude (IA real) ou nenhuma (só regra)")
    args = ap.parse_args()

    cliente = {"simulada": ClienteSimulado, "claude": ClienteClaude, "nenhuma": lambda: None}[args.ia]()
    wb, resultados = processar_planilha(args.planilha, cliente)

    PASTA_SAIDA.mkdir(exist_ok=True)
    destino = PASTA_SAIDA / "planilha_processos_preenchida.xlsx"
    wb.save(destino)
    print(f"Modo de IA: {args.ia}")
    print(resumo(resultados))
    print(f"\nPlanilha preenchida salva em {destino}")


if __name__ == "__main__":
    main()
