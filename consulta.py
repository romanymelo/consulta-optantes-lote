import os
import re
import subprocess
import sys
import time

from extracao import eventos_regime
from pdf import caminho_pdf
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

URL = "https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
PORTA = 9222
PASTA = os.path.dirname(sys.executable if getattr(sys, "frozen", False) else os.path.abspath(__file__))
PERFIL = os.path.join(PASTA, ".perfil-chrome")

SEL_CNPJ = "#Cnpj"
SEL_CONSULTAR = "button.h-captcha"
SEL_RESULTADO = "#conteudo"
SEL_MAIS_INFO = "#btnMaisInfo"
SEL_GERAR_PDF = "#GerarPDF"

_JS_PARES = """
() => [...document.querySelectorAll('#conteudo .panel-body .spanValorVerde:not(#maisInfo *)')].map(s => {
    let n = s.previousSibling, rotulo = '';
    while (n && !rotulo.trim()) { rotulo = n.textContent || ''; n = n.previousSibling; }
    return [rotulo.split('\\n').pop().replace(/:\\s*$/, '').trim(), s.textContent.trim()];
})
"""
_JS_PAINEIS_MAIS_INFO = r"""
() => [...document.querySelectorAll('#maisInfo .panel')].map(p => {
    const corpo = p.querySelector('.panel-body').cloneNode(true);
    const linhas = [...corpo.querySelectorAll('tbody tr')].map(tr =>
        [...tr.querySelectorAll('td')].map(td => td.textContent.replace(/\s+/g, ' ').trim()));
    corpo.querySelectorAll('table').forEach(t => {
        const cab = [...t.querySelectorAll('thead th')].map(th => th.textContent.replace(/\s+/g, ' ').trim());
        const todas = (cab.length ? [cab] : []).concat([...t.querySelectorAll('tbody tr')].map(tr =>
            [...tr.querySelectorAll('td')].map(td => td.textContent.replace(/\s+/g, ' ').trim())));
        t.replaceWith(document.createTextNode('\n' + todas.map(l => l.join(' | ')).join('\n') + '\n'));
    });
    return [p.querySelector('.panel-title').textContent.trim(), corpo.textContent, linhas];
})
"""


def _apos(campos: dict[str, str], chave: str, novos: dict[str, str]) -> dict[str, str]:
    """Devolve `campos` com `novos` logo depois de `chave` (ou no fim, se ela não existir)."""
    if chave not in campos:
        return {**campos, **novos}
    saida: dict[str, str] = {}
    for k, v in campos.items():
        saida[k] = v
        if k == chave:
            saida.update(novos)
    return saida


class ConsultaErro(Exception):
    pass


class CnpjNaoEncontrado(ConsultaErro):
    pass


class CaptchaTimeout(ConsultaErro):
    pass


class PortalBloqueou(ConsultaErro):
    pass


class Consultor:
    def __init__(self):
        self._pw = None
        self._browser = None
        self._chrome = None
        self._page = None

    def abrir(self) -> None:
        self._pw = sync_playwright().start()
        try:
            self._browser = self._pw.chromium.connect_over_cdp(f"http://localhost:{PORTA}")
        except PlaywrightError:
            self._chrome = subprocess.Popen(
                [CHROME, f"--remote-debugging-port={PORTA}", f"--user-data-dir={PERFIL}", "about:blank"]
            )
            self._browser = self._conectar_com_espera()
        self._page = self._browser.contexts[0].new_page()

    def _conectar_com_espera(self):
        for _ in range(20):
            time.sleep(1)
            try:
                return self._pw.chromium.connect_over_cdp(f"http://localhost:{PORTA}")
            except PlaywrightError:
                continue
        raise ConsultaErro("Não foi possível conectar ao Chrome.")

    def _frame(self):
        for f in self._page.frames:
            if "consultaoptantes" in f.url and "hcaptcha" not in f.url:
                return f
        raise ConsultaErro("Formulário do portal não encontrado.")

    def consultar(self, cnpj: str, timeout_s: int = 180, pasta_pdf: str | None = None) -> dict[str, str]:
        try:
            return self._consultar(cnpj, timeout_s, pasta_pdf)
        except PlaywrightError as e:
            raise ConsultaErro(f"Navegador indisponível: {e.message.splitlines()[0]}") from e

    def _consultar(self, cnpj: str, timeout_s: int, pasta_pdf: str | None) -> dict[str, str]:
        self._page.goto(URL, wait_until="networkidle")
        self._page.wait_for_selector("iframe#frame")
        inicio = time.time()
        while time.time() - inicio < 30:
            try:
                f = self._frame()
                f.wait_for_selector(SEL_CNPJ, timeout=2000)
                break
            except (ConsultaErro, PlaywrightError):
                continue
        f.fill(SEL_CNPJ, cnpj)
        f.click(SEL_CONSULTAR)

        inicio = time.time()
        while time.time() - inicio < timeout_s:
            time.sleep(1)
            try:
                f = self._frame()
                if f.query_selector(SEL_RESULTADO):
                    campos = self._extrair(f)
                    if pasta_pdf:
                        pdf = self._baixar_pdf(f, campos.get("Nome Empresarial"), cnpj, pasta_pdf)
                        if pdf:
                            campos["PDF"] = pdf
                    return campos
                texto = f.inner_text("body")
                msg = f.query_selector(".validation-summary-errors, .alert, .text-danger:not(:empty)")
                aviso = re.sub(r"[×\s]+", " ", msg.inner_text()).strip() if msg else ""
            except PlaywrightError:
                continue  # o iframe navegou no meio da leitura; tenta de novo
            except ConsultaErro:
                continue
            if "Comportamento de Rob" in texto:
                raise PortalBloqueou("O portal bloqueou a consulta automatizada.")
            if aviso:
                raise CnpjNaoEncontrado(aviso)
        raise CaptchaTimeout("Tempo esgotado esperando o resultado (captcha?).")

    def _baixar_pdf(self, f, nome: str | None, cnpj: str, pasta: str) -> str | None:
        """Clica em 'Gerar PDF' e salva o download como '<nome da empresa>.pdf'; None se falhar."""
        try:
            with self._page.expect_download(timeout=30000) as download:
                f.click(SEL_GERAR_PDF)
            destino = caminho_pdf(pasta, nome, cnpj)
            download.value.save_as(destino)
            return destino
        except (PlaywrightError, OSError):
            return None

    def _extrair(self, f) -> dict[str, str]:
        campos: dict[str, str] = {}
        data = f.query_selector("#conteudo h5 span")
        if data:
            campos["Data da consulta"] = data.inner_text().strip()
        for rotulo, valor in f.evaluate(_JS_PARES):
            campos[rotulo or "Valor"] = valor
        botao = f.query_selector(SEL_MAIS_INFO)
        if botao:
            botao.click()
            f.wait_for_function(
                "() => { const l = document.querySelector('.loader');"
                " return !l || getComputedStyle(l).display === 'none'; }",
                timeout=15000,
            )
            time.sleep(0.5)
            linhas_regime: list[list[str]] = []
            for titulo, texto, linhas in f.evaluate(_JS_PAINEIS_MAIS_INFO):
                campos[titulo] = "\n".join(
                    re.sub(r"[ \t]+", " ", l).strip() for l in texto.splitlines() if l.strip()
                )
                if titulo == "Eventos Futuros (Regime Regular IBS/CBS)":
                    linhas_regime = linhas
            campos = _apos(campos, "Situação no Regime Regular da CBS", eventos_regime(linhas_regime))
        return campos

    def fechar(self) -> None:
        if self._page:
            try:
                self._page.close()
            except PlaywrightError:
                pass
        if self._pw:
            self._pw.stop()
