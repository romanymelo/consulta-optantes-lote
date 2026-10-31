# Consulta de Optantes do Simples Nacional em lote — Design

Data: 2026-10-30

## Objetivo

Receber uma lista de CNPJs colada pelo usuário, consultar cada um no portal
"Consulta Optantes" do Simples Nacional
(https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21)
e gerar uma planilha `.xlsx` com tudo que a página retornar.

## Decisões já tomadas

- Entrada: lista colada diretamente numa janela do programa.
- Dados: todos os campos que a página exibir (situação no Simples Nacional e no
  SIMEI, datas de opção, períodos anteriores, exclusões etc.).
- Captcha: modo semiautomático. O navegador fica visível; o robô preenche o
  CNPJ, o usuário resolve o captcha e envia, e o robô lê o resultado. O captcha
  não é contornado.
- Plataforma: Windows, Python 3.12 (já instalado).

## Abordagem

Python + Playwright (Chrome visível) + janela Tkinter + openpyxl.

## Componentes

| Unidade | Responsabilidade | Depende de |
|---|---|---|
| `cnpj.py` | Limpar a lista colada, remover pontuação e duplicados, validar 14 dígitos e dígitos verificadores | — |
| `consulta.py` | Abrir a página, preencher o CNPJ, esperar o usuário resolver o captcha, extrair todos os campos exibidos | Playwright |
| `planilha.py` | Gravar o `.xlsx` (uma linha por CNPJ, colunas dinâmicas, status e erro), salvando a cada consulta | openpyxl |
| `app.py` | Janela Tkinter: caixa para colar, botão iniciar, progresso, log | os três acima |

## Fluxo de dados

1. Usuário cola a lista e clica em iniciar.
2. `cnpj.py` devolve CNPJs válidos e únicos; inválidos vão direto para a planilha com status "inválido".
3. Para cada CNPJ, `consulta.py` retorna um dicionário campo→valor.
4. `planilha.py` acrescenta a linha e salva em disco.

## Erros

CNPJ inválido, não encontrado, tempo esgotado esperando o captcha e falha de
rede são registrados na coluna de status/erro da linha. A execução continua com
o próximo CNPJ. Como o arquivo é salvo a cada consulta, uma interrupção não
perde o que já foi coletado.

## Teste

- `cnpj.py` e `planilha.py`: testes unitários (TDD).
- `consulta.py`: depende do site real e do captcha; verificação manual
  assistida com 1–2 CNPJs.

## Riscos e pontos em aberto

- O formulário real não foi inspecionado (a busca HTTP trouxe só a moldura do
  portal). Os seletores serão definidos na primeira execução no navegador.
- Se a página bloquear automação, voltaremos à discussão (por exemplo, fonte
  alternativa sem captcha).
- O volume de consultas é limitado pelo tempo de resolução manual do captcha.
