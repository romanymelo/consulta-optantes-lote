from cnpj import normalizar, parse_lista, validar


def test_validar_cnpj_valido():
    assert validar("11222333000181") is True


def test_validar_rejeita_digitos_iguais():
    assert validar("11111111111111") is False


def test_validar_rejeita_digito_verificador_errado():
    assert validar("11222333000182") is False


def test_normalizar_remove_mascara():
    assert normalizar("11.222.333/0001-81") == "11222333000181"


def test_parse_lista_misturada():
    texto = "11.222.333/0001-81\n\n11222333000181; 123 ,11222333000199"
    assert parse_lista(texto) == (["11222333000181"], ["123", "11222333000199"])


def test_parse_lista_preserva_zeros_a_esquerda():
    assert parse_lista("00.000.000/0001-91") == (["00000000000191"], [])


def test_validar_cnpj_alfanumerico():
    assert validar("12ABC34501DE35") is True


def test_normalizar_mantem_letras_em_maiusculo():
    assert normalizar("12.abc.345/01de-35") == "12ABC34501DE35"


def test_parse_lista_aceita_alfanumerico():
    assert parse_lista("12.ABC.345/01DE-35") == (["12ABC34501DE35"], [])
