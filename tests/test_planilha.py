import openpyxl
import pytest
from openpyxl import load_workbook

from planilha import Planilha

PORTAL = {
    "Data da consulta": "30/10/2026 10:55:30",
    "CNPJ": "11.222.333/0001-81",
    "Nome Empresarial": "EMPRESA TESTE LTDA",
    "Situação no Simples Nacional": "Optante pelo Simples Nacional desde 01/01/2026",
    "Situação no SIMEI": "NÃO enquadrado no SIMEI",
    "Situação no Regime Regular do IBS": "Não optante",
    "Situação no Regime Regular da CBS": "Não optante",
    "Períodos Anteriores": "Não Existem",
    "IBS - evento futuro": "Opção pelo regime regular em 01/01/2027",
    "CBS - evento futuro": "Opção pelo regime regular em 01/01/2027",
    "Eventos Futuros (Simples Nacional)": "Não Existem",
}

PRINCIPAIS = [
    "CNPJ", "Nome Empresarial", "Status", "Simples Nacional", "SIMEI", "IBS", "CBS",
    "IBS - evento futuro", "CBS - evento futuro", "Evento futuro Simples", "Data da consulta", "PDF", "Erro",
]


def _ws(caminho, aba="Consulta"):
    return load_workbook(caminho)[aba]


def _tabela(caminho):
    ws = _ws(caminho)
    cab = [c.value for c in ws[1]]
    return cab, [dict(zip(cab, [c.value for c in row])) for row in ws.iter_rows(min_row=2)]


def _cor(celula):
    return celula.fill.fgColor.rgb[-6:]


def _celula(caminho, linha, coluna):
    ws = _ws(caminho)
    cab = [c.value for c in ws[1]]
    return ws.cell(row=linha, column=cab.index(coluna) + 1)


def test_cria_abas_e_colunas_principais_na_ordem(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", PORTAL, "ok")
    assert load_workbook(caminho).sheetnames == ["Consulta", "Resumo"]
    cab, linhas = _tabela(caminho)
    assert cab[: len(PRINCIPAIS)] == PRINCIPAIS
    assert linhas[0]["Nome Empresarial"] == "EMPRESA TESTE LTDA"
    assert linhas[0]["Simples Nacional"] == "Optante pelo Simples Nacional desde 01/01/2026"
    assert linhas[0]["IBS"] == "Não optante"


def test_cnpj_sai_formatado_como_texto_com_zeros(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("00000000000191", {}, "ok")
    assert _tabela(caminho)[1][0]["CNPJ"] == "00.000.000/0001-91"
    assert _celula(caminho, 2, "CNPJ").number_format == "@"


def test_cnpj_alfanumerico_formatado(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("12ABC34501DE35", {}, "ok")
    assert _tabela(caminho)[1][0]["CNPJ"] == "12.ABC.345/01DE-35"


def test_cnpj_do_portal_nao_vira_coluna_e_colisao_de_nome_e_renomeada(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", {"CNPJ": "x", "Status": "y"}, "ok")
    cab, linhas = _tabela(caminho)
    assert cab.count("CNPJ") == 1
    assert linhas[0]["Status"] == "ok"
    assert linhas[0]["Status (portal)"] == "y"


def test_campos_extras_ficam_no_fim_na_ordem_de_chegada(tmp_path):
    caminho = tmp_path / "r.xlsx"
    p = Planilha(str(caminho))
    p.adicionar("11222333000181", {"A": "1"}, "ok")
    p.adicionar("00000000000191", {"A": "2", "B": "3"}, "ok")
    cab, linhas = _tabela(caminho)
    assert cab[-2:] == ["A", "B"]
    assert linhas[0]["B"] is None and linhas[1]["B"] == "3"


def test_salva_a_cada_adicao(tmp_path):
    caminho = tmp_path / "r.xlsx"
    p = Planilha(str(caminho))
    for i in range(1, 4):
        p.adicionar(f"1122233300018{i}", {}, "ok")
        assert len(_tabela(caminho)[1]) == i


def test_linha_de_erro_mantem_cnpj_cru_e_mensagem(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("123", {}, "inválido", "tamanho incorreto")
    linha = _tabela(caminho)[1][0]
    assert (linha["CNPJ"], linha["Status"], linha["Erro"]) == ("123", "inválido", "tamanho incorreto")


def test_cabecalho_estilizado_congelado_e_com_filtro(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", PORTAL, "ok")
    ws = _ws(caminho)
    assert ws["A1"].font.bold is True and ws["A1"].font.color.rgb[-6:] == "FFFFFF"
    assert _cor(ws["A1"]) == "1F4E3D"
    assert ws.freeze_panes == "C2"
    assert ws.auto_filter.ref


def test_larguras_definidas(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", PORTAL, "ok")
    ws = _ws(caminho)
    assert ws.column_dimensions["B"].width >= 30


def test_celulas_com_quebra_de_linha_tem_wrap(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", {"A": "l1\nl2"}, "ok")
    assert _celula(caminho, 2, "A").alignment.wrap_text is True


def test_cores_por_significado(tmp_path):
    caminho = tmp_path / "r.xlsx"
    p = Planilha(str(caminho))
    p.adicionar("11222333000181", PORTAL, "ok")
    p.adicionar("00000000000191", {**PORTAL, "Situação no Simples Nacional": "NÃO optante pelo Simples Nacional",
                                   "IBS - evento futuro": "Não Existem", "CBS - evento futuro": "Não Existem"}, "ok")
    p.adicionar("123", {}, "inválido", "x")
    assert _cor(_celula(caminho, 2, "Simples Nacional")) == "C6EFCE"
    assert _cor(_celula(caminho, 3, "Simples Nacional")) == "E7E6E6"
    assert _cor(_celula(caminho, 2, "IBS - evento futuro")) == "FFC000"
    assert _cor(_celula(caminho, 3, "IBS - evento futuro")) != "FFC000"
    assert _cor(_celula(caminho, 4, "Status")) == "FFC7CE"
    assert _cor(_celula(caminho, 2, "Status")) != "FFC7CE"


def test_simei_enquadrado_fica_azul(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", {**PORTAL, "Situação no SIMEI": "Enquadrado no SIMEI desde 01/01/2024"}, "ok")
    assert _cor(_celula(caminho, 2, "SIMEI")) == "DDEBF7"


def test_resumo_conta_corretamente(tmp_path):
    caminho = tmp_path / "r.xlsx"
    p = Planilha(str(caminho))
    p.adicionar("11222333000181", PORTAL, "ok")
    p.adicionar("00000000000191", {**PORTAL, "Situação no Simples Nacional": "NÃO optante pelo Simples Nacional",
                                   "IBS - evento futuro": "Não Existem", "CBS - evento futuro": "Não Existem"}, "ok")
    p.adicionar("98765432000198", {}, "não encontrado", "x")
    resumo = {r[0].value: r[1].value for r in _ws(caminho, "Resumo").iter_rows(min_row=1) if r[0].value}
    assert resumo["CNPJs processados"] == 3
    assert resumo["Consultados com sucesso"] == 2
    assert resumo["Com erro, inválidos ou não encontrados"] == 1
    assert resumo["Optantes pelo Simples Nacional"] == 1
    assert resumo["Com evento futuro de IBS/CBS"] == 1


def test_permissao_negada_avisa_e_mantem_linhas_em_memoria(tmp_path, monkeypatch):
    caminho = tmp_path / "r.xlsx"
    p = Planilha(str(caminho))
    original = openpyxl.Workbook.save
    monkeypatch.setattr(openpyxl.Workbook, "save", lambda self, c: (_ for _ in ()).throw(PermissionError("x")))
    with pytest.raises(PermissionError, match="Feche o arquivo"):
        p.adicionar("11222333000181", PORTAL, "ok")
    monkeypatch.setattr(openpyxl.Workbook, "save", original)
    p.adicionar("00000000000191", {}, "ok")
    assert len(_tabela(caminho)[1]) == 2


def test_altura_da_linha_nao_explode_com_detalhes_longos(tmp_path):
    caminho = tmp_path / "r.xlsx"
    longo = "\n".join(f"periodo {i} | 01/01/20{i:02d} | Excluída por ato administrativo" for i in range(12))
    Planilha(str(caminho)).adicionar("11222333000181", {**PORTAL, "Períodos Anteriores": longo}, "ok")
    assert _ws(caminho).row_dimensions[2].height <= 100


def test_linha_so_com_campos_curtos_fica_baixa(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", {**PORTAL, "IBS - evento futuro": "Não Existem"}, "ok")
    assert _ws(caminho).row_dimensions[2].height <= 50


def test_pdf_vira_link_clicavel_com_o_nome_do_arquivo(tmp_path):
    caminho = tmp_path / "r.xlsx"
    pdf = str(tmp_path / "PDFs" / "EMPRESA TESTE LTDA.pdf")
    Planilha(str(caminho)).adicionar("11222333000181", {**PORTAL, "PDF": pdf}, "ok")
    c = _celula(caminho, 2, "PDF")
    assert c.value == "EMPRESA TESTE LTDA.pdf"
    assert c.hyperlink is not None and c.hyperlink.target.endswith("EMPRESA TESTE LTDA.pdf")
    assert c.font.underline == "single"


def test_sem_pdf_a_celula_fica_vazia(tmp_path):
    caminho = tmp_path / "r.xlsx"
    Planilha(str(caminho)).adicionar("11222333000181", PORTAL, "ok")
    assert _celula(caminho, 2, "PDF").value is None


def test_avisos_do_portal_campo_novo_e_campo_sumido():
    from planilha import avisos_do_portal
    completo = {
        "Nome Empresarial": "X", "Situação no Simples Nacional": "Optante", "Situação no SIMEI": "Não",
        "Situação no Regime Regular do IBS": "Optante", "Situação no Regime Regular da CBS": "Optante",
        "Períodos Anteriores": "...",
    }
    assert avisos_do_portal(completo) == []
    assert avisos_do_portal({**completo, "Regime Novo": "Sim"}) == ["campo novo no portal: Regime Novo (coluna adicionada no fim)"]
    sem_simei = {k: v for k, v in completo.items() if k != "Situação no SIMEI"}
    assert avisos_do_portal(sem_simei) == ["campo esperado não veio do portal: Situação no SIMEI"]
