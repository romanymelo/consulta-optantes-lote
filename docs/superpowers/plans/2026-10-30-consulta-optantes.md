# Consulta de Optantes do Simples Nacional em lote — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Colar uma lista de CNPJs numa janela, consultá-los um a um no portal de Optantes do Simples Nacional (com o usuário resolvendo o captcha) e gerar um `.xlsx` com tudo que a página retornar.

**Architecture:** Quatro módulos pequenos: `cnpj.py` (limpeza/validação), `planilha.py` (escrita incremental do xlsx), `consulta.py` (Playwright, Chrome visível) e `app.py` (janela Tkinter que orquestra). A consulta roda em thread separada para não travar a janela.

**Tech Stack:** Python 3.12, Playwright (Chrome), openpyxl, Tkinter, pytest.

**Spec:** `docs/superpowers/specs/2026-10-30-consulta-optantes-design.md`

## Global Constraints

- Windows, Python 3.12.10 (já instalado).
- URL: `https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21`
- Captcha nunca é contornado: o navegador fica visível e o usuário resolve.
- Planilha salva em disco a cada CNPJ consultado.
- Uma linha por CNPJ, com colunas de status e erro; colunas de dados são dinâmicas (tudo que a página retornar).
- A pasta não é repositório git: os passos de commit são omitidos.

## Review Focus

- Lista colada com máscara (`12.345.678/0001-95`), espaços, linhas em branco e separadores `;` ou `,`: todos devem ser aceitos.
- CNPJ repetido na lista: consultado uma vez só.
- CNPJ com zeros à esquerda: deve ficar como texto na planilha, sem perder zeros.
- CNPJ de dígitos iguais (`11111111111111`) ou com dígito verificador errado: marcado como inválido, sem consulta.
- Planilha aberta no Excel durante a execução (`PermissionError` ao salvar): erro claro no log, sem perder os dados em memória.
- Usuário fecha o navegador ou o captcha estoura o tempo: a linha recebe status de erro e o lote continua (ou termina limpo se o navegador morreu).

---

## File Structure

- `requirements.txt` — dependências.
- `cnpj.py` — limpar, validar e deduplicar CNPJs.
- `planilha.py` — gravação incremental do `.xlsx`.
- `consulta.py` — automação do portal.
- `app.py` — janela Tkinter e orquestração.
- `tests/test_cnpj.py`, `tests/test_planilha.py`.

### Task 1: Limpeza e validação de CNPJs

**Files:**
- Create: `requirements.txt`, `cnpj.py`
- Test: `tests/test_cnpj.py`

**Interfaces:**
- Produces:
  - `normalizar(texto: str) -> str` — só os dígitos.
  - `validar(cnpj: str) -> bool` — 14 dígitos, não todos iguais, dígitos verificadores corretos.
  - `parse_lista(texto: str) -> tuple[list[str], list[str]]` — `(validos_unicos_em_ordem, invalidos)`; separadores: quebra de linha, `;`, `,`, espaço; itens vazios ignorados; inválidos preservam o texto original do item.

- [ ] **Step 1: Criar `requirements.txt`** com `playwright`, `openpyxl`, `pytest`; rodar `py -m pip install -r requirements.txt` e `py -m playwright install chromium`.
- [ ] **Step 2: Escrever os testes em `tests/test_cnpj.py`**
  - `test_validar_cnpj_valido`: `validar("11222333000181")` é `True`.
  - `test_validar_rejeita_digitos_iguais`: `validar("11111111111111")` é `False`.
  - `test_validar_rejeita_digito_verificador_errado`: `validar("11222333000182")` é `False`.
  - `test_normalizar_remove_mascara`: `normalizar("11.222.333/0001-81") == "11222333000181"`.
  - `test_parse_lista_misturada`: entrada `"11.222.333/0001-81\n\n11222333000181; 123 ,11222333000199"` retorna `(["11222333000181"], ["123", "11222333000199"])`.
  - `test_parse_lista_preserva_zeros_a_esquerda`: `parse_lista("00.000.000/0001-91")` retorna `(["00000000000191"], [])`.
- [ ] **Step 3: Rodar `py -m pytest tests/test_cnpj.py -v`** — esperado: FAIL (módulo inexistente).
- [ ] **Step 4: Implementar `normalizar`, `validar`, `parse_lista` em `cnpj.py`** — validação pelo algoritmo módulo 11 padrão, pesos `[5,4,3,2,9,8,7,6,5,4,3,2]` e `[6,5,4,3,2,9,8,7,6,5,4,3,2]`.
- [ ] **Step 5: Rodar os testes** — esperado: PASS.

### Task 2: Gravação incremental da planilha

**Files:**
- Create: `planilha.py`
- Test: `tests/test_planilha.py`

**Interfaces:**
- Consumes: nada.
- Produces: `class Planilha` com
  - `__init__(self, caminho: str)`
  - `adicionar(self, cnpj: str, campos: dict[str, str], status: str, erro: str = "") -> None` — acrescenta a linha e grava em disco.
  - Colunas fixas `CNPJ`, `Status`, `Erro`, seguidas das colunas de dados na ordem em que aparecem pela primeira vez. Coluna nova que surge depois é adicionada ao final.
  - Levanta `PermissionError` com mensagem em português se o arquivo estiver aberto, mantendo a linha em memória.

- [ ] **Step 1: Escrever os testes em `tests/test_planilha.py`** (usar `tmp_path` e `openpyxl.load_workbook` para ler de volta)
  - `test_cria_arquivo_com_cabecalho_e_linha`: após um `adicionar`, a linha 1 é `["CNPJ","Status","Erro", <campos>]` e a linha 2 tem os valores.
  - `test_colunas_dinamicas`: segundo CNPJ com campo novo gera coluna nova ao final; a linha do primeiro fica com a célula vazia.
  - `test_cnpj_com_zeros_fica_texto`: `"00000000000191"` lido de volta é a string igual, sem perder zeros.
  - `test_salva_a_cada_adicao`: o arquivo existe e tem N linhas após cada chamada.
  - `test_linha_de_erro_sem_campos`: `adicionar("123", {}, "inválido", "tamanho incorreto")` grava status e erro.
- [ ] **Step 2: Rodar `py -m pytest tests/test_planilha.py -v`** — esperado: FAIL.
- [ ] **Step 3: Implementar `Planilha` em `planilha.py`** — manter as linhas em memória (lista de dicts) e reescrever o arquivo inteiro a cada `adicionar` com openpyxl; gravar CNPJ como string com `number_format = "@"`.
- [ ] **Step 4: Rodar os testes** — esperado: PASS.

### Task 3: Automação do portal (verificação manual)

**Files:**
- Create: `consulta.py`

**Interfaces:**
- Consumes: nada dos outros módulos.
- Produces:
  - `class ConsultaErro(Exception)` — base, com subclasses `CnpjNaoEncontrado` e `CaptchaTimeout`.
  - `class Consultor` com
    - `abrir(self) -> None` — abre Chrome visível na URL do portal.
    - `consultar(self, cnpj: str, timeout_s: int = 180) -> dict[str, str]` — preenche o CNPJ, aguarda o usuário resolver o captcha e enviar, e devolve todos os campos exibidos (rótulo → valor).
    - `fechar(self) -> None`
  - Parâmetro `timeout_s` vem do usuário via `app.py`.

- [ ] **Step 1: Inspecionar a página real** com um script descartável (fora do projeto, no scratchpad): abrir a URL, localizar se o formulário está em `iframe`, anotar seletores do campo CNPJ, do captcha e dos blocos de resultado. Registrar os seletores como constantes no topo de `consulta.py`.
- [ ] **Step 2: Implementar `Consultor`** com `playwright.sync_api`: Chrome com `headless=False`; esperar o resultado com `wait_for_selector(timeout=timeout_s*1000)` convertendo `TimeoutError` em `CaptchaTimeout`; extrair pares rótulo/valor de todos os blocos de resultado; se a página indicar CNPJ sem registro, levantar `CnpjNaoEncontrado`.
- [ ] **Step 3: Verificação manual** com 2 CNPJs reais (um optante, um não): rodar um script curto no scratchpad que chama `abrir`, `consultar` para cada um e imprime o dict. Esperado: campos coerentes com o que a página mostra na tela. Conferir também o que acontece se a janela do Chrome for fechada durante a espera (deve gerar `ConsultaErro`, não travar).

### Task 4: Janela e orquestração

**Files:**
- Create: `app.py`

**Interfaces:**
- Consumes: `cnpj.parse_lista`, `planilha.Planilha`, `consulta.Consultor` e as exceções de `consulta.py`.
- Produces: `py app.py` abre a janela; função `executar_lote(cnpjs: list[str], invalidos: list[str], consultor, planilha, log: Callable[[str], None]) -> None` separada da interface para poder ser exercitada sem Tkinter.

- [ ] **Step 1: Escrever testes de `executar_lote` em `tests/test_lote.py`** com `Consultor` falso:
  - `test_registra_invalidos_sem_consultar`: inválidos vão para a planilha com status `inválido` e o consultor falso não é chamado para eles.
  - `test_erro_em_um_cnpj_nao_para_o_lote`: o falso levanta `CnpjNaoEncontrado` no segundo; o terceiro ainda é consultado e gravado com status `ok`.
  - `test_permissao_negada_vai_para_o_log`: `Planilha` falsa levanta `PermissionError`; a mensagem aparece no `log` e o lote termina sem exceção não tratada.
- [ ] **Step 2: Rodar `py -m pytest tests/test_lote.py -v`** — esperado: FAIL.
- [ ] **Step 3: Implementar `executar_lote`** — para cada CNPJ chama `consultor.consultar`, grava `ok` ou o status/erro da exceção (`não encontrado`, `captcha expirou`, `erro`), e escreve no `log`.
- [ ] **Step 4: Rodar os testes** — esperado: PASS.
- [ ] **Step 5: Implementar a janela Tkinter em `app.py`** — caixa de texto para colar, botão "Iniciar", campo para o caminho do `.xlsx` (padrão `resultado.xlsx` na pasta do projeto), barra/contador de progresso e área de log; o lote roda em `threading.Thread`, atualizando a interface via `root.after`.
- [ ] **Step 6: Verificação manual ponta a ponta** — colar 3 CNPJs (um válido e optante, um inexistente, um com dígito errado), resolver o captcha quando pedido e abrir o `.xlsx`. Esperado: 3 linhas com status `ok`, `não encontrado` e `inválido`; todos os campos retornados pela página aparecem como colunas.
