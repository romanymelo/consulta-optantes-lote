import os
import re

_MAX_NOME = 120
_RESERVADOS = {"CON", "PRN", "AUX", "NUL"} | {f"COM{i}" for i in range(1, 10)} | {f"LPT{i}" for i in range(1, 10)}


def _limpar(nome: str) -> str:
    nome = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "-", nome)
    nome = re.sub(r"\s+", " ", nome).strip()
    return nome[:_MAX_NOME].rstrip(". ")


def caminho_pdf(pasta: str, nome_empresa: str | None, cnpj: str) -> str:
    """Caminho livre para o PDF da empresa, criando a pasta; nomes repetidos ganham ' (2)', ' (3)'..."""
    os.makedirs(pasta, exist_ok=True)
    base = _limpar(nome_empresa or "") or cnpj
    if base.upper() in _RESERVADOS:
        base += "-"
    caminho = os.path.join(pasta, f"{base}.pdf")
    n = 2
    while os.path.exists(caminho):
        caminho = os.path.join(pasta, f"{base} ({n}).pdf")
        n += 1
    return caminho
