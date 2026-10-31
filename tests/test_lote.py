from consulta import CaptchaTimeout, CnpjNaoEncontrado, PortalBloqueou
from lote import executar_lote


class ConsultorFalso:
    def __init__(self, respostas):
        self.respostas = respostas
        self.chamados = []

    def consultar(self, cnpj, timeout_s=180, pasta_pdf=None):
        self.chamados.append(cnpj)
        self.pastas = getattr(self, 'pastas', []) + [pasta_pdf]
        r = self.respostas[cnpj]
        if isinstance(r, Exception):
            raise r
        return r


class PlanilhaFalsa:
    def __init__(self, erro_ao_salvar=None):
        self.linhas = []
        self.erro_ao_salvar = erro_ao_salvar

    def adicionar(self, cnpj, campos, status, erro=""):
        if self.erro_ao_salvar:
            raise self.erro_ao_salvar
        self.linhas.append((cnpj, campos, status, erro))


def test_registra_invalidos_sem_consultar():
    consultor, planilha = ConsultorFalso({}), PlanilhaFalsa()
    executar_lote([], ["123"], consultor, planilha, lambda m: None)
    assert consultor.chamados == []
    assert planilha.linhas == [("123", {}, "inválido", "CNPJ inválido")]


def test_erro_em_um_cnpj_nao_para_o_lote():
    consultor = ConsultorFalso(
        {"A": {"x": "1"}, "B": CnpjNaoEncontrado("não achei"), "C": {"x": "3"}}
    )
    planilha = PlanilhaFalsa()
    executar_lote(["A", "B", "C"], [], consultor, planilha, lambda m: None)
    assert [(l[0], l[2]) for l in planilha.linhas] == [
        ("A", "ok"),
        ("B", "não encontrado"),
        ("C", "ok"),
    ]
    assert planilha.linhas[1][3] == "não achei"


def test_captcha_timeout_vira_status_e_lote_continua():
    consultor = ConsultorFalso({"A": CaptchaTimeout("tempo"), "B": {"x": "2"}})
    planilha = PlanilhaFalsa()
    executar_lote(["A", "B"], [], consultor, planilha, lambda m: None)
    assert [l[2] for l in planilha.linhas] == ["captcha expirou", "ok"]


def test_portal_bloqueou_interrompe_o_lote():
    consultor = ConsultorFalso({"A": PortalBloqueou("bloqueado"), "B": {"x": "2"}})
    planilha, log = PlanilhaFalsa(), []
    executar_lote(["A", "B"], [], consultor, planilha, log.append)
    assert consultor.chamados == ["A"]
    assert planilha.linhas[0][2] == "bloqueado pelo portal"
    assert any("bloqueou" in m for m in log)


def test_permissao_negada_vai_para_o_log():
    consultor = ConsultorFalso({"A": {"x": "1"}})
    planilha = PlanilhaFalsa(erro_ao_salvar=PermissionError("Feche o arquivo no Excel"))
    log = []
    executar_lote(["A"], [], consultor, planilha, log.append)
    assert any("Feche o arquivo no Excel" in m for m in log)


def test_repassa_a_pasta_de_pdf_ao_consultor():
    consultor = ConsultorFalso({"A": {"x": "1", "PDF": "p.pdf"}})
    executar_lote(["A"], [], consultor, PlanilhaFalsa(), lambda m: None, pasta_pdf="PDFs")
    assert consultor.pastas == ["PDFs"]


def test_avisa_no_log_quando_o_pdf_nao_foi_salvo_mas_mantem_ok():
    consultor = ConsultorFalso({"A": {"x": "1"}})
    planilha, log = PlanilhaFalsa(), []
    executar_lote(["A"], [], consultor, planilha, log.append, pasta_pdf="PDFs")
    assert planilha.linhas[0][2] == "ok"
    assert any("PDF não salvo" in m for m in log)


def test_sem_pasta_de_pdf_nao_avisa_nada():
    consultor = ConsultorFalso({"A": {"x": "1"}})
    log = []
    executar_lote(["A"], [], consultor, PlanilhaFalsa(), log.append)
    assert not any("PDF" in m for m in log)


def test_avisa_campo_novo_so_na_primeira_vez():
    base = {
        "Nome Empresarial": "X", "Situação no Simples Nacional": "Optante", "Situação no SIMEI": "Não",
        "Situação no Regime Regular do IBS": "Optante", "Situação no Regime Regular da CBS": "Optante",
    }
    consultor = ConsultorFalso({"A": {**base, "Novo": "1"}, "B": {**base, "Novo": "2"}})
    log = []
    executar_lote(["A", "B"], [], consultor, PlanilhaFalsa(), log.append)
    avisos = [m for m in log if "campo novo" in m]
    assert avisos == ["[1/2] A: campo novo no portal: Novo (coluna adicionada no fim)"]
