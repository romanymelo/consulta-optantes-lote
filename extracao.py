NAO_EXISTEM = "Não Existem"


def eventos_regime(linhas: list[list[str]]) -> dict[str, str]:
    """Linhas da tabela 'Eventos Futuros (Regime Regular IBS/CBS)': [regime, descrição, data efeito]."""
    eventos = {"IBS": [], "CBS": []}
    for linha in linhas:
        if len(linha) >= 3 and linha[0] in eventos:
            eventos[linha[0]].append(f"{linha[1]} em {linha[2]}")
    return {f"{r} - evento futuro": "; ".join(e) or NAO_EXISTEM for r, e in eventos.items()}
