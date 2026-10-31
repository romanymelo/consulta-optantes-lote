import math
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# (cabeçalho na planilha, campo do portal de onde vem; None = calculado pelo programa)
_PRINCIPAIS = [
    ("CNPJ", None),
    ("Nome Empresarial", "Nome Empresarial"),
    ("Status", None),
    ("Simples Nacional", "Situação no Simples Nacional"),
    ("SIMEI", "Situação no SIMEI"),
    ("IBS", "Situação no Regime Regular do IBS"),
    ("CBS", "Situação no Regime Regular da CBS"),
    ("IBS - evento futuro", "IBS - evento futuro"),
    ("CBS - evento futuro", "CBS - evento futuro"),
    ("Evento futuro Simples", "Eventos Futuros (Simples Nacional)"),
    ("Data da consulta", "Data da consulta"),
    ("PDF", "PDF"),
    ("Erro", None),
]
_DETALHES = [
    "Períodos Anteriores",
    "Eventos Futuros (Regime Regular IBS/CBS)",
    "Eventos Futuros (SIMEI)",
    "Informações de Períodos como MEI Transportador Autônomo de Cargas",
]
_FONTES_USADAS = {f for _, f in _PRINCIPAIS if f}
_RESERVADOS = {"Status", "Erro"}
_ESPERADOS = [
    f for c, f in _PRINCIPAIS
    if c in ("Nome Empresarial", "Simples Nacional", "SIMEI", "IBS", "CBS")
]
_LARGURAS = {
    "CNPJ": 20, "Nome Empresarial": 38, "Status": 16, "Simples Nacional": 30, "SIMEI": 24,
    "IBS": 16, "CBS": 16, "IBS - evento futuro": 28, "CBS - evento futuro": 28,
    "Evento futuro Simples": 30, "Data da consulta": 19, "PDF": 36, "Erro": 30,
}
_LARGURA_DETALHE = 60
_MAX_LINHAS_DETALHE = 6
_ALTURA_MAXIMA = 100
_LARGURA_EXTRA = 30
_NAO_EXISTEM = "não existem"

_VERDE_ESCURO = "1F4E3D"
_VERDE = ("C6EFCE", "006100")
_CINZA = ("E7E6E6", "595959")
_AZUL = ("DDEBF7", "1F4E79")
_AMARELO_FORTE = ("FFC000", "000000")
_AMARELO_SUAVE = ("FFE699", "000000")
_VERMELHO = ("FFC7CE", "9C0006")
_ZEBRA = "F3F7F4"
_BORDA = Border(bottom=Side(style="thin", color="D9D9D9"))


def _formatar_cnpj(cnpj: str) -> str:
    if len(cnpj) != 14 or not cnpj.isalnum():
        return cnpj
    return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"


def _preenchido(valor) -> bool:
    return bool(valor) and str(valor).strip().lower() != _NAO_EXISTEM


def _comeca_com(valor, prefixo: str) -> bool:
    return str(valor or "").strip().lower().startswith(prefixo)


def _cor_da_celula(cabecalho: str, valor):
    """(preenchimento, fonte) conforme o significado do valor, ou None."""
    if cabecalho == "Status":
        return None if valor == "ok" else _VERMELHO
    if cabecalho == "Simples Nacional":
        return _VERDE if _comeca_com(valor, "optante") else (_CINZA if valor else None)
    if cabecalho in ("IBS", "CBS"):
        return _VERDE if _comeca_com(valor, "optante") else None
    if cabecalho == "SIMEI":
        return _AZUL if valor and not _comeca_com(valor, "não") else None
    if cabecalho in ("IBS - evento futuro", "CBS - evento futuro"):
        return _AMARELO_FORTE if _preenchido(valor) else None
    if cabecalho == "Evento futuro Simples":
        return _AMARELO_SUAVE if _preenchido(valor) else None
    return None


def avisos_do_portal(campos: dict[str, str]) -> list[str]:
    """Campos que o portal trouxe e o programa não conhecia, e campos esperados que não vieram."""
    conhecidos = _FONTES_USADAS | set(_DETALHES) | {"CNPJ"}
    novos = [f"campo novo no portal: {n} (coluna adicionada no fim)" for n in campos if n not in conhecidos]
    sumidos = [f"campo esperado não veio do portal: {f}" for f in _ESPERADOS if f not in campos]
    return novos + sumidos


def _linhas_estimadas(texto, largura: int) -> int:
    if not texto:
        return 1
    return sum(max(1, math.ceil(len(seg) / (largura * 1.1))) for seg in str(texto).split("\n"))


class Planilha:
    def __init__(self, caminho: str):
        self.caminho = caminho
        self._extras: list[str] = []
        self._linhas: list[dict[str, str]] = []

    def adicionar(self, cnpj: str, campos: dict[str, str], status: str, erro: str = "") -> None:
        linha = {"CNPJ": _formatar_cnpj(cnpj), "Status": status, "Erro": erro or None}
        for cabecalho, fonte in _PRINCIPAIS:
            if fonte:
                linha[cabecalho] = campos.get(fonte)
        for nome, valor in campos.items():
            if nome == "CNPJ" or nome in _FONTES_USADAS:
                continue
            nome = f"{nome} (portal)" if nome in _RESERVADOS else nome
            if nome not in self._extras:
                self._extras.append(nome)
            linha[nome] = valor
        self._linhas.append(linha)
        self._salvar()

    def _cabecalhos(self) -> list[str]:
        detalhes = [d for d in _DETALHES if d in self._extras]
        outros = [e for e in self._extras if e not in _DETALHES]
        return [c for c, _ in _PRINCIPAIS] + detalhes + outros

    def _salvar(self) -> None:
        wb = Workbook()
        self._aba_consulta(wb.active)
        self._aba_resumo(wb.create_sheet("Resumo"))
        try:
            wb.save(self.caminho)
        except PermissionError as e:
            raise PermissionError(
                f"Não foi possível salvar '{self.caminho}'. Feche o arquivo no Excel e tente de novo."
            ) from e

    def _largura(self, cabecalho: str) -> int:
        if cabecalho in _LARGURAS:
            return _LARGURAS[cabecalho]
        return _LARGURA_DETALHE if cabecalho in _DETALHES else _LARGURA_EXTRA

    def _aba_consulta(self, ws) -> None:
        ws.title = "Consulta"
        cabecalhos = self._cabecalhos()
        ws.append(cabecalhos)
        ws.row_dimensions[1].height = 32
        for i, cab in enumerate(cabecalhos, 1):
            c = ws.cell(row=1, column=i)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=_VERDE_ESCURO)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = self._largura(cab)

        for n, linha in enumerate(self._linhas, 2):
            ws.append([linha.get(c) or None for c in cabecalhos])
            linhas_texto = 1
            for i, cab in enumerate(cabecalhos, 1):
                celula = ws.cell(row=n, column=i)
                celula.alignment = Alignment(wrap_text=True, vertical="top")
                celula.border = _BORDA
                cor = _cor_da_celula(cab, linha.get(cab))
                if cor:
                    celula.fill = PatternFill("solid", fgColor=cor[0])
                    celula.font = Font(color=cor[1], bold=cor[0] == _AMARELO_FORTE[0])
                elif n % 2 == 1:
                    celula.fill = PatternFill("solid", fgColor=_ZEBRA)
                if cab == "PDF" and linha.get(cab):
                    caminho_pdf = os.path.abspath(linha[cab])
                    celula.value = os.path.basename(caminho_pdf)
                    celula.hyperlink = caminho_pdf
                    celula.font = Font(color="0563C1", underline="single")
                texto = celula.value if cab == "PDF" else linha.get(cab)
                estimadas = _linhas_estimadas(texto, self._largura(cab))
                linhas_texto = max(linhas_texto, min(estimadas, _MAX_LINHAS_DETALHE) if cab in _DETALHES else estimadas)
            ws.cell(row=n, column=1).number_format = "@"
            ws.row_dimensions[n].height = min(max(18, 15 * linhas_texto), _ALTURA_MAXIMA)

        ws.freeze_panes = "C2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cabecalhos))}{len(self._linhas) + 1}"

    def _aba_resumo(self, ws) -> None:
        ls = self._linhas
        com_evento = sum(
            _preenchido(l.get("IBS - evento futuro")) or _preenchido(l.get("CBS - evento futuro")) for l in ls
        )
        metricas = [
            ("CNPJs processados", len(ls)),
            ("Consultados com sucesso", sum(l["Status"] == "ok" for l in ls)),
            ("Com erro, inválidos ou não encontrados", sum(l["Status"] != "ok" for l in ls)),
            ("Optantes pelo Simples Nacional", sum(_comeca_com(l.get("Simples Nacional"), "optante") for l in ls)),
            ("Enquadrados no SIMEI", sum(bool(l.get("SIMEI")) and not _comeca_com(l.get("SIMEI"), "não") for l in ls)),
            ("Optantes no regime regular do IBS", sum(_comeca_com(l.get("IBS"), "optante") for l in ls)),
            ("Optantes no regime regular da CBS", sum(_comeca_com(l.get("CBS"), "optante") for l in ls)),
            ("Com evento futuro de IBS/CBS", com_evento),
            ("Com evento futuro no Simples Nacional", sum(_preenchido(l.get("Evento futuro Simples")) for l in ls)),
        ]
        ws["A1"] = "Resumo da consulta"
        ws["A1"].font = Font(bold=True, size=14, color=_VERDE_ESCURO)
        ws.column_dimensions["A"].width = 46
        ws.column_dimensions["B"].width = 14
        for n, (rotulo, valor) in enumerate(metricas, 3):
            ws.cell(row=n, column=1, value=rotulo).border = _BORDA
            v = ws.cell(row=n, column=2, value=valor)
            v.alignment = Alignment(horizontal="center")
            v.border = _BORDA
            if rotulo == "Com evento futuro de IBS/CBS":
                for c in (ws.cell(row=n, column=1), v):
                    c.fill = PatternFill("solid", fgColor=_AMARELO_FORTE[0])
                    c.font = Font(bold=True)
