import os

from pdf import caminho_pdf


def _nome(caminho):
    return os.path.basename(caminho)


def test_usa_o_nome_da_empresa(tmp_path):
    assert _nome(caminho_pdf(str(tmp_path), "EMPRESA EXEMPLO LTDA", "11222333000181")) == "EMPRESA EXEMPLO LTDA.pdf"


def test_troca_caracteres_invalidos_do_windows(tmp_path):
    assert _nome(caminho_pdf(str(tmp_path), 'EMPRESA S/A: "X"?', "1")) == "EMPRESA S-A- -X--.pdf"


def test_colapsa_espacos_e_tira_pontos_finais(tmp_path):
    assert _nome(caminho_pdf(str(tmp_path), "  ACME   LTDA.  ", "1")) == "ACME LTDA.pdf"


def test_limita_o_tamanho_do_nome(tmp_path):
    n = _nome(caminho_pdf(str(tmp_path), "A" * 300, "1"))
    assert n == "A" * 120 + ".pdf"


def test_sem_nome_usa_o_cnpj(tmp_path):
    assert _nome(caminho_pdf(str(tmp_path), "", "11222333000181")) == "11222333000181.pdf"
    assert _nome(caminho_pdf(str(tmp_path), None, "11222333000181")) == "11222333000181.pdf"


def test_nome_reservado_do_windows_nao_vira_dispositivo(tmp_path):
    assert _nome(caminho_pdf(str(tmp_path), "CON", "1")) == "CON-.pdf"


def test_repetido_ganha_sufixo_numerico_sem_sobrescrever(tmp_path):
    (tmp_path / "ACME LTDA.pdf").write_bytes(b"x")
    assert _nome(caminho_pdf(str(tmp_path), "ACME LTDA", "1")) == "ACME LTDA (2).pdf"
    (tmp_path / "ACME LTDA (2).pdf").write_bytes(b"x")
    assert _nome(caminho_pdf(str(tmp_path), "ACME LTDA", "1")) == "ACME LTDA (3).pdf"


def test_cria_a_pasta_se_nao_existir(tmp_path):
    pasta = tmp_path / "PDFs"
    caminho_pdf(str(pasta), "ACME", "1")
    assert pasta.is_dir()
