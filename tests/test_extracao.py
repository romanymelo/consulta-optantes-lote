from extracao import eventos_regime


def test_evento_futuro_ibs_e_cbs():
    linhas = [
        ["IBS", "Opção pelo regime regular", "01/01/2027"],
        ["CBS", "Opção pelo regime regular", "01/01/2027"],
    ]
    assert eventos_regime(linhas) == {
        "IBS - evento futuro": "Opção pelo regime regular em 01/01/2027",
        "CBS - evento futuro": "Opção pelo regime regular em 01/01/2027",
    }


def test_sem_eventos_marca_nao_existem():
    assert eventos_regime([]) == {
        "IBS - evento futuro": "Não Existem",
        "CBS - evento futuro": "Não Existem",
    }


def test_so_um_regime_com_evento():
    r = eventos_regime([["CBS", "Opção pelo regime regular", "01/01/2027"]])
    assert r["IBS - evento futuro"] == "Não Existem"
    assert r["CBS - evento futuro"] == "Opção pelo regime regular em 01/01/2027"


def test_varios_eventos_no_mesmo_regime_ficam_juntos():
    r = eventos_regime([["IBS", "Opção", "01/01/2027"], ["IBS", "Exclusão", "01/01/2028"]])
    assert r["IBS - evento futuro"] == "Opção em 01/01/2027; Exclusão em 01/01/2028"


def test_coluna_a_mais_ou_linha_curta_nao_quebra():
    r = eventos_regime([["IBS", "Opção", "01/01/2027", "extra"], ["Sem eventos"], ["CBS", "Opção", "01/01/2027"]])
    assert r["IBS - evento futuro"] == "Opção em 01/01/2027"
    assert r["CBS - evento futuro"] == "Opção em 01/01/2027"
