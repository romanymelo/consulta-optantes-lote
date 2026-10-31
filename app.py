import ctypes
import os
import queue
import re
import threading
import tkinter as tk
from tkinter import filedialog, ttk

from cnpj import parse_lista
from consulta import PASTA, Consultor, ConsultaErro
from lote import executar_lote
from planilha import Planilha

PADRAO = os.path.join(PASTA, "resultado.xlsx")

VERDE = "#2E9C6E"
VERDE_VIVO = "#3FB583"
VERDE_ESCURO = "#237A56"
CABECALHO = "#0E1F18"
FUNDO = "#121816"
SUPERFICIE = "#1B2421"
BORDA = "#2E3B36"
TEXTO = "#E6EEEA"
MUDO = "#8A9A93"
MENTA = "#8FD0B3"
ERRO = "#FF8A80"
LOG_FUNDO = "#0A100E"
LOG_TEXTO = "#C9D6D0"
LOG_OK = "#6FD3A0"
LOG_ERRO = "#FF8A80"
LOG_AVISO = "#FFC000"

FONTE = "Segoe UI"
MONO = "Consolas"

_PROGRESSO = re.compile(r"^\[(\d+)/(\d+)\]")
_MARCAS_ERRO = ("ERRO", "erro", "bloque", "inválido", "não encontrado", "expirou", "interrompido", "não salvo")


def _nitidez_no_windows() -> None:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.fila: queue.Queue = queue.Queue()
        root.title("Consulta Optantes — Simples Nacional")
        root.configure(bg=FUNDO)
        self._estilos()

        topo = tk.Frame(root, bg=CABECALHO)
        topo.pack(fill="x")
        tk.Label(topo, text="Consulta Optantes", bg=CABECALHO, fg=TEXTO, font=(FONTE, 20, "bold")).pack(
            anchor="w", padx=28, pady=(20, 0)
        )
        tk.Label(
            topo, text="Situação no Simples Nacional, SIMEI e regime regular de IBS e CBS",
            bg=CABECALHO, fg=MENTA, font=(FONTE, 10),
        ).pack(anchor="w", padx=28, pady=(2, 18))

        tk.Label(root, text="Desenvolvido por Romany Melo", bg=FUNDO, fg=MUDO, font=(FONTE, 9)).pack(
            side="bottom", pady=(0, 12)
        )
        corpo = tk.Frame(root, bg=FUNDO)
        corpo.pack(fill="both", expand=True, padx=28, pady=(20, 10))

        cab = tk.Frame(corpo, bg=FUNDO)
        cab.pack(fill="x")
        tk.Label(cab, text="CNPJs para consultar", bg=FUNDO, fg=TEXTO, font=(FONTE, 11, "bold")).pack(side="left")
        self.contador = tk.Label(cab, text="", bg=FUNDO, fg=MUDO, font=(FONTE, 10))
        self.contador.pack(side="right")

        self.entrada = tk.Text(
            self._cartao(corpo, fill="x", pady=(8, 16)), height=5, relief="flat", bd=0, wrap="word",
            font=(MONO, 11), fg=TEXTO, bg=SUPERFICIE, insertbackground=VERDE_VIVO, padx=14, pady=12,
            selectbackground="#2E5A47", selectforeground=TEXTO, undo=True,
        )
        self.entrada.pack(fill="both", expand=True)
        self.dica = tk.Label(
            self.entrada, text="Cole os CNPJs aqui, um por linha ou separados por ; ou ,",
            bg=SUPERFICIE, fg=MUDO, font=(FONTE, 10), cursor="xterm",
        )
        self.dica.place(x=16, y=12)
        self.dica.bind("<Button-1>", lambda e: self.entrada.focus_set())
        for evento in ("<KeyRelease>", "<<Paste>>", "<<Cut>>"):
            self.entrada.bind(evento, lambda e: self.root.after_idle(self._atualizar_contador))
        self._atualizar_contador()

        tk.Label(corpo, text="Salvar planilha em", bg=FUNDO, fg=TEXTO, font=(FONTE, 11, "bold")).pack(anchor="w")
        linha = tk.Frame(corpo, bg=FUNDO)
        linha.pack(fill="x", pady=(8, 10))
        self.caminho = tk.StringVar(value=PADRAO)
        ttk.Entry(linha, textvariable=self.caminho).pack(side="left", fill="x", expand=True, ipady=3)
        ttk.Button(linha, text="Escolher…", style="Sec.TButton", command=self._escolher).pack(side="left", padx=(10, 0))

        self.salvar_pdf = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            corpo, text="Salvar o PDF do relatório de cada consulta (pasta PDFs, ao lado da planilha)",
            variable=self.salvar_pdf,
        ).pack(anchor="w", pady=(0, 16))

        acao = tk.Frame(corpo, bg=FUNDO)
        acao.pack(fill="x", pady=(0, 16))
        self.botao = ttk.Button(acao, text="Iniciar consulta", style="Primario.TButton", command=self._iniciar)
        self.botao.pack(side="left")
        self.barra = ttk.Progressbar(acao, style="Verde.Horizontal.TProgressbar", mode="determinate")
        self.barra.pack(side="left", fill="x", expand=True, padx=(18, 12))
        self.andamento = tk.Label(acao, text="", bg=FUNDO, fg=MUDO, font=(FONTE, 10), width=8, anchor="e")
        self.andamento.pack(side="left")

        self.log = tk.Text(
            self._cartao(corpo, fill="both", expand=True, cor=LOG_FUNDO), height=9, relief="flat", bd=0,
            font=(MONO, 10), bg=LOG_FUNDO, fg=LOG_TEXTO, padx=14, pady=10, state="disabled", wrap="word",
            selectbackground=VERDE_VIVO,
        )
        self.log.pack(fill="both", expand=True)
        self.log.tag_configure("ok", foreground=LOG_OK)
        self.log.tag_configure("erro", foreground=LOG_ERRO)
        self.log.tag_configure("fim", foreground=LOG_OK, font=(MONO, 10, "bold"))
        self.log.tag_configure("info", foreground=LOG_AVISO)

        self.root.after(200, self._drenar)

    def _estilos(self) -> None:
        s = ttk.Style()
        s.theme_use("clam")
        s.configure("TCheckbutton", background=FUNDO, foreground=TEXTO, font=(FONTE, 10), focuscolor=FUNDO)
        s.map("TCheckbutton", background=[("active", FUNDO)], indicatorcolor=[("selected", VERDE)])
        s.configure("TCheckbutton", indicatorbackground=SUPERFICIE, indicatorforeground="white")
        s.configure(
            "TEntry", fieldbackground=SUPERFICIE, foreground=TEXTO, bordercolor=BORDA, lightcolor=BORDA,
            darkcolor=BORDA, padding=(10, 4), font=(FONTE, 10),
        )
        s.map("TEntry", bordercolor=[("focus", VERDE)], lightcolor=[("focus", VERDE)], darkcolor=[("focus", VERDE)])
        s.configure(
            "Primario.TButton", background=VERDE, foreground="white", borderwidth=0, focusthickness=0,
            focuscolor=VERDE, padding=(26, 11), font=(FONTE, 11, "bold"),
        )
        s.map(
            "Primario.TButton",
            background=[("disabled", "#3A4A43"), ("pressed", VERDE_ESCURO), ("active", VERDE_VIVO)],
            foreground=[("disabled", MUDO)],
        )
        s.configure(
            "Sec.TButton", background=SUPERFICIE, foreground=MENTA, bordercolor=BORDA, lightcolor=SUPERFICIE,
            darkcolor=SUPERFICIE, focuscolor=SUPERFICIE, padding=(16, 7), font=(FONTE, 10, "bold"),
        )
        s.map("Sec.TButton", background=[("active", "#26332E")], bordercolor=[("active", VERDE)])
        s.configure(
            "Verde.Horizontal.TProgressbar", troughcolor=SUPERFICIE, background=VERDE_VIVO, borderwidth=0,
            lightcolor=VERDE_VIVO, darkcolor=VERDE_VIVO, thickness=8,
        )

    def _cartao(self, pai, cor: str = BORDA, **pack) -> tk.Frame:
        """Moldura de 1px que acende em verde quando o campo de dentro recebe o foco."""
        moldura = tk.Frame(pai, bg=cor, padx=1, pady=1)
        moldura.pack(**pack)
        moldura.bind("<FocusIn>", lambda e: moldura.configure(bg=VERDE))
        moldura.bind("<FocusOut>", lambda e: moldura.configure(bg=cor))
        return moldura

    def _atualizar_contador(self) -> None:
        texto = self.entrada.get("1.0", "end")
        if texto.strip():
            self.dica.place_forget()
        else:
            self.dica.place(x=16, y=12)
        validos, invalidos = parse_lista(texto)
        if not validos and not invalidos:
            self.contador.configure(text="Nenhum CNPJ ainda", fg=MUDO)
        elif invalidos:
            self.contador.configure(text=f"{len(validos)} válido(s), {len(invalidos)} inválido(s)", fg=ERRO)
        else:
            self.contador.configure(text=f"{len(validos)} CNPJ(s) válido(s)", fg=VERDE_VIVO)

    def _escolher(self) -> None:
        c = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if c:
            self.caminho.set(c)

    def _escrever(self, msg: str) -> None:
        if msg.startswith("Concluído"):
            tag = "fim"
        elif msg.endswith(": ok"):
            tag = "ok"
        elif "CNPJ(s) válido(s)" in msg or "campo novo no portal" in msg or "campo esperado não veio" in msg:
            tag = "info"
        elif any(m in msg for m in _MARCAS_ERRO):
            tag = "erro"
        else:
            tag = ""
        self.log.configure(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")
        m = _PROGRESSO.match(msg)
        if m:
            atual, total = int(m.group(1)), int(m.group(2))
            self.barra.configure(maximum=total, value=atual)
            self.andamento.configure(text=f"{atual} de {total}")

    def _drenar(self) -> None:
        while not self.fila.empty():
            msg = self.fila.get()
            if msg is None:
                self.botao.configure(state="normal", text="Iniciar consulta")
            else:
                self._escrever(msg)
        self.root.after(200, self._drenar)

    def _iniciar(self) -> None:
        validos, invalidos = parse_lista(self.entrada.get("1.0", "end"))
        if not validos and not invalidos:
            self._escrever("Cole ao menos um CNPJ.")
            return
        self.botao.configure(state="disabled", text="Consultando…")
        self.barra.configure(value=0)
        self.andamento.configure(text="")
        self._escrever(f"{len(validos)} CNPJ(s) válido(s), {len(invalidos)} inválido(s).")
        threading.Thread(target=self._rodar, args=(validos, invalidos, self.caminho.get(), self.salvar_pdf.get()), daemon=True).start()

    def _rodar(self, validos, invalidos, caminho, com_pdf) -> None:
        consultor = Consultor()
        log = self.fila.put
        try:
            consultor.abrir()
            pasta_pdf = os.path.join(os.path.dirname(os.path.abspath(caminho)), "PDFs") if com_pdf else None
            executar_lote(validos, invalidos, consultor, Planilha(caminho), log, pasta_pdf)
            log(f"Concluído. Planilha: {caminho}")
        except ConsultaErro as e:
            log(f"ERRO: {e}")
        finally:
            consultor.fechar()
            self.fila.put(None)


def centralizar(root: tk.Tk, largura: int, altura: int) -> None:
    """Dimensiona a janela pela escala da tela e a abre no centro, sem piscar no canto."""
    escala = root.winfo_fpixels("1i") / 96
    w, h = int(largura * escala), int(altura * escala)
    w, h = min(w, root.winfo_screenwidth() - 40), min(h, root.winfo_screenheight() - 140)
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2 - 30
    root.geometry(f"{w}x{h}+{max(x, 0)}+{max(y, 0)}")
    root.minsize(int(w * 0.8), int(h * 0.8))


if __name__ == "__main__":
    _nitidez_no_windows()
    raiz = tk.Tk()
    raiz.withdraw()
    App(raiz)
    centralizar(raiz, 860, 760)
    raiz.deiconify()
    raiz.mainloop()
