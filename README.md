# Consulta Optantes em Lote

Aplicativo desktop (Windows) que consulta em lote a situação de CNPJs no portal do
[Simples Nacional](https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21)
(Receita Federal): optante pelo Simples Nacional, SIMEI e regime regular de IBS e CBS.

## O que faz

- Recebe uma lista de CNPJs (um por linha ou separados por `;` ou `,`).
- Valida os dígitos verificadores, inclusive o **CNPJ alfanumérico**, e remove duplicados.
- Consulta cada CNPJ no portal e extrai os dados, incluindo os eventos futuros de IBS/CBS.
- Grava tudo em uma planilha Excel (aba de consulta e aba de resumo).
- Opcionalmente salva o PDF do relatório de cada consulta, nomeado pela razão social.

## Requisitos

- Windows
- Python 3.10 ou superior
- Google Chrome instalado em `C:\Program Files\Google\Chrome\Application\chrome.exe`
  (caminho definido em `consulta.py`)

## Instalação e uso

```bash
pip install -r requirements.txt
python app.py
```

O app abre o Chrome com um perfil próprio (`.perfil-chrome`, criado ao lado do app).
O portal usa hCaptcha; se ele pedir, resolva o desafio no navegador e a consulta continua.

## Testes

```bash
pytest
```

## Gerar o executável

```bash
pip install pyinstaller
pyinstaller ConsultaOptantes.spec
```

## Aviso

Projeto independente, sem vínculo com a Receita Federal. O app automatiza o uso do portal
público e não contorna o captcha: se o portal bloquear a consulta, o lote é interrompido.
Use com responsabilidade e respeite os termos do portal.

## Licença

[MIT](LICENSE) © Romany Melo
