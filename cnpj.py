import re

_PESOS_1 = [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]
_PESOS_2 = [6] + _PESOS_1


def normalizar(texto: str) -> str:
    return re.sub(r"[^0-9A-Za-z]", "", texto).upper()


def _digito(base: str, pesos: list[int]) -> str:
    resto = sum((ord(d) - 48) * p for d, p in zip(base, pesos)) % 11
    return "0" if resto < 2 else str(11 - resto)


def validar(cnpj: str) -> bool:
    if not re.fullmatch(r"[0-9A-Z]{12}[0-9]{2}", cnpj) or len(set(cnpj)) == 1:
        return False
    d1 = _digito(cnpj[:12], _PESOS_1)
    d2 = _digito(cnpj[:12] + d1, _PESOS_2)
    return cnpj[12:] == d1 + d2


def parse_lista(texto: str) -> tuple[list[str], list[str]]:
    validos: list[str] = []
    invalidos: list[str] = []
    for item in re.split(r"[\s;,]+", texto):
        if not item:
            continue
        cnpj = normalizar(item)
        if validar(cnpj):
            if cnpj not in validos:
                validos.append(cnpj)
        else:
            invalidos.append(item)
    return validos, invalidos
