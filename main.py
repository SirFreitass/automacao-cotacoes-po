"""
main.py
-------
Interface gráfica principal do sistema de análise de cotações e PO.
Execute este arquivo com duplo clique ou: python main.py
"""

import os
import re
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Verifica se as dependências estão instaladas
try:
    import google.genai
    import openpyxl
    import pdfplumber
except ImportError as e:
    import subprocess
    resposta = tk.messagebox.askyesno(
        "Dependências não instaladas",
        f"Módulo ausente: {e}\n\nDeseja instalar as dependências agora?\n"
        "(Requer conexão com a internet)"
    )
    if resposta:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        messagebox.showinfo("Instalação concluída", "Dependências instaladas. Reinicie o programa.")
    sys.exit(0)

from config import GOOGLE_API_KEY
from extractor import extrair_cotacoes, extrair_po
from analyzer import analisar
from excel_exporter import exportar_excel


# =====================================================================
# Helpers
# =====================================================================

def _extrair_req_do_nome(caminho: str):
    """Extrai o Nº ECO REQ do nome do arquivo. Ex: 'Cotação REQ 031326015461' → '031326015461'"""
    nome = os.path.basename(caminho)
    match = re.search(r'REQ\s*(\d{8,})', nome, re.IGNORECASE)
    return match.group(1) if match else None


def _e_po(caminho: str) -> bool:
    """
    Heurística: arquivo é PO se o nome indicar uma Purchase Order.
    Cobre padrões como: 'PO 031326015461', 'PO BRAM - 430581', 'PO 475919 - QT...',
    'Purchase Order', '467694 - ECO' (gerado pelo ECO Requisition).
    """
    nome = os.path.basename(caminho).lower()
    # Começa com "po " ou "po_" ou "po-" (ex: "PO 475919 - QT 2026.007638")
    if re.match(r'^po[\s\-_]', nome):
        return True
    # Outros padrões comuns
    return any(p in nome for p in (
        "purchase order", "ordem de compra",
        "- eco", "_eco",  # ex: "467694 - ECO"
    ))


def _parear_por_req(cotacoes: list, pos: list) -> tuple:
    """
    Pareia cotações com POs pelo nº REQ no nome do arquivo.
    Retorna (pares, nao_pareados_cotacoes, nao_pareados_pos).
    """
    idx_cot = {}  # req → caminho
    for c in cotacoes:
        req = _extrair_req_do_nome(c)
        if req:
            idx_cot[req] = c
        else:
            idx_cot[f"__sem_req_{c}"] = c

    idx_po = {}
    for p in pos:
        req = _extrair_req_do_nome(p)
        if req:
            idx_po[req] = p
        else:
            idx_po[f"__sem_req_{p}"] = p

    pares = []
    usados_cot = set()
    usados_po = set()

    # Pareia por REQ
    for req, cot in idx_cot.items():
        if req in idx_po:
            pares.append((cot, idx_po[req]))
            usados_cot.add(cot)
            usados_po.add(idx_po[req])

    # Arquivos sem REQ no nome: pareia por ordem
    sem_req_cot = [c for c in cotacoes if c not in usados_cot]
    sem_req_po  = [p for p in pos       if p not in usados_po]
    for cot, po in zip(sem_req_cot, sem_req_po):
        pares.append((cot, po))
        usados_cot.add(cot)
        usados_po.add(po)

    nao_cot = [c for c in cotacoes if c not in usados_cot]
    nao_po  = [p for p in pos       if p not in usados_po]
    return pares, nao_cot, nao_po


# =====================================================================
# Verifica configuração da API Key
# =====================================================================

def _verificar_api_key():
    chave_invalida = (
        "COLE-SUA-CHAVE-AQUI" in GOOGLE_API_KEY
        or len(GOOGLE_API_KEY) < 10
        or not GOOGLE_API_KEY.startswith("AIzaSy")
    )
    if chave_invalida:
        messagebox.showwarning(
            "API Key inválida",
            "A chave da API Google Gemini está ausente ou incorreta.\n\n"
            "A chave deve começar com 'AIzaSy...' (39 caracteres).\n\n"
            "Como obter:\n"
            "1. Acesse aistudio.google.com e faça login\n"
            "2. Clique em 'Get API key' → 'Create API key'\n"
            "3. Copie a chave e cole em config.py\n"
            "4. Salve o arquivo e reinicie o programa"
        )
        return False
    return True


# =====================================================================
# Janela principal
# =====================================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Análise de Cotações e PO — ECO Purchasing")
        self.resizable(False, False)
        self.configure(bg="#1F3864")

        self._fila = []  # Lista de (caminho_cotacao, caminho_po)
        self._pasta_saida = tk.StringVar(value=os.path.dirname(os.path.abspath(__file__)))

        self._construir_ui()
        self._centralizar()

    def _centralizar(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _construir_ui(self):
        # --- Cabeçalho ---
        header = tk.Frame(self, bg="#1F3864", pady=16, padx=24)
        header.pack(fill="x")
        tk.Label(header, text="Análise de Cotações & PO",
                 font=("Segoe UI", 18, "bold"), fg="white", bg="#1F3864").pack()
        tk.Label(header, text="ECO Purchasing  |  Powered by Gemini AI",
                 font=("Segoe UI", 9), fg="#A0B4D0", bg="#1F3864").pack()

        # --- Corpo ---
        corpo = tk.Frame(self, bg="white", padx=28, pady=20)
        corpo.pack(fill="both", expand=True)

        # Botões de importação (topo)
        frame_import = tk.Frame(corpo, bg="#F0F4FA", bd=1, relief="solid")
        frame_import.pack(fill="x", pady=(0, 12))

        tk.Label(frame_import, text="Adicionar arquivos:",
                 font=("Segoe UI", 9, "bold"), bg="#F0F4FA", fg="#1F3864"
                 ).pack(side="left", padx=(12, 8), pady=8)

        tk.Button(frame_import, text="📂  Importar pasta",
                  command=self._importar_pasta,
                  font=("Segoe UI", 9), bg="#1F3864", fg="white",
                  activebackground="#2E539E", activeforeground="white",
                  relief="flat", cursor="hand2", padx=12, pady=5
                  ).pack(side="left", padx=4, pady=6)

        tk.Button(frame_import, text="📄  Selecionar múltiplas Cotações + POs",
                  command=self._importar_multiplos,
                  font=("Segoe UI", 9), bg="#2E539E", fg="white",
                  activebackground="#1F3864", activeforeground="white",
                  relief="flat", cursor="hand2", padx=12, pady=5
                  ).pack(side="left", padx=4, pady=6)

        tk.Label(frame_import,
                 text="Pareamento automático por nº REQ no nome do arquivo",
                 font=("Segoe UI", 8), bg="#F0F4FA", fg="#666"
                 ).pack(side="left", padx=(8, 12), pady=6)

        # Label fila
        tk.Label(corpo, text="Fila de análises:",
                 font=("Segoe UI", 10, "bold"), bg="white", anchor="w"
                 ).pack(fill="x", pady=(0, 4))

        # Cabeçalho da tabela
        cab = tk.Frame(corpo, bg="#1F3864")
        cab.pack(fill="x")
        tk.Label(cab, text="#",    width=3,  font=("Segoe UI", 8, "bold"),
                 fg="white", bg="#1F3864", anchor="center").pack(side="left", padx=(4, 0), pady=3)
        tk.Label(cab, text="Cotações (PDF)", width=30, font=("Segoe UI", 8, "bold"),
                 fg="white", bg="#1F3864", anchor="w").pack(side="left", padx=4)
        tk.Label(cab, text="PO (PDF)", width=30, font=("Segoe UI", 8, "bold"),
                 fg="white", bg="#1F3864", anchor="w").pack(side="left", padx=4)

        # Lista com scrollbar
        frame_scroll = tk.Frame(corpo, bg="#EEEEEE", bd=1, relief="solid")
        frame_scroll.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(frame_scroll, orient="vertical")
        self._listbox = tk.Listbox(
            frame_scroll, yscrollcommand=scrollbar.set,
            font=("Segoe UI", 9), height=7,
            selectbackground="#2E539E", selectforeground="white",
            activestyle="none", bg="white", relief="flat",
            highlightthickness=0
        )
        scrollbar.config(command=self._listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self._listbox.pack(side="left", fill="both", expand=True)

        # Botão remover
        frame_fila_btns = tk.Frame(corpo, bg="white")
        frame_fila_btns.pack(fill="x", pady=(4, 0))
        tk.Button(frame_fila_btns, text="🗑  Remover selecionado",
                  command=self._remover_par,
                  font=("Segoe UI", 9), bg="#FFE8E8", fg="#9C0006",
                  relief="flat", cursor="hand2", padx=10, pady=3
                  ).pack(side="left")
        tk.Button(frame_fila_btns, text="✖  Limpar tudo",
                  command=self._limpar_fila,
                  font=("Segoe UI", 9), bg="#FFE8E8", fg="#9C0006",
                  relief="flat", cursor="hand2", padx=10, pady=3
                  ).pack(side="left", padx=(6, 0))

        ttk.Separator(corpo, orient="horizontal").pack(fill="x", pady=12)

        # Pasta de saída
        tk.Label(corpo, text="Salvar Excel em:",
                 font=("Segoe UI", 10), bg="white", anchor="w"
                 ).pack(fill="x", pady=(0, 4))
        frame_saida = tk.Frame(corpo, bg="white")
        frame_saida.pack(fill="x", pady=(0, 14))
        tk.Entry(frame_saida, textvariable=self._pasta_saida,
                 font=("Segoe UI", 9), fg="#333", relief="solid", bd=1,
                 state="readonly", readonlybackground="#F5F5F5", width=56
                 ).pack(side="left", fill="x", expand=True, ipady=4)
        tk.Button(frame_saida, text="Pasta...",
                  command=self._selecionar_pasta_saida,
                  font=("Segoe UI", 9), bg="#E8E8E8", relief="flat",
                  cursor="hand2", padx=8
                  ).pack(side="left", padx=(6, 0))

        # Barra de progresso
        self._progress = ttk.Progressbar(corpo, mode="determinate", length=520)
        self._progress.pack(fill="x", pady=(0, 4))

        # Status
        self._status = tk.StringVar(value="Importe uma pasta ou selecione os PDFs e clique em Analisar.")
        tk.Label(corpo, textvariable=self._status,
                 font=("Segoe UI", 9), bg="white", fg="#555",
                 wraplength=520, justify="left", anchor="w"
                 ).pack(fill="x", pady=(0, 12))

        # Botão principal
        self._btn_analisar = tk.Button(
            corpo, text="  ▶  ANALISAR TODOS",
            command=self._iniciar_analise,
            font=("Segoe UI", 12, "bold"),
            bg="#1F3864", fg="white",
            activebackground="#2E539E", activeforeground="white",
            relief="flat", cursor="hand2", pady=10, padx=24,
        )
        self._btn_analisar.pack(pady=(0, 4))

    # ------------------------------------------------------------------
    # Importação de arquivos
    # ------------------------------------------------------------------

    def _importar_pasta(self):
        """
        Seleciona uma pasta e pareia os PDFs em sequência por data de modificação.
        Os arquivos são salvos em pares consecutivos (PO + cotação ou cotação + PO),
        então basta ordenar por data e parear de dois em dois.
        """
        pasta = filedialog.askdirectory(title="Selecione a pasta com os PDFs de Cotações e POs")
        if not pasta:
            return

        todos = [
            os.path.join(pasta, f)
            for f in os.listdir(pasta)
            if f.lower().endswith(".pdf")
        ]
        if not todos:
            messagebox.showinfo("Nenhum PDF", "Nenhum arquivo PDF encontrado na pasta selecionada.")
            return

        # Ordena por data de modificação (mais antigo primeiro = mesma ordem de salvamento)
        todos.sort(key=lambda f: os.path.getmtime(f))

        pares = []
        nao_pareados = []

        i = 0
        while i + 1 < len(todos):
            f1, f2 = todos[i], todos[i + 1]
            e_po_f1 = _e_po(f1)
            e_po_f2 = _e_po(f2)

            if e_po_f1 and not e_po_f2:
                pares.append((f2, f1))   # (cotação, PO)
            elif e_po_f2 and not e_po_f1:
                pares.append((f1, f2))   # (cotação, PO)
            else:
                # Não foi possível distinguir — assume f1=cotação, f2=PO por ordem
                pares.append((f1, f2))
            i += 2

        # Se total ímpar, sobra um arquivo sem par
        if len(todos) % 2 != 0:
            nao_pareados.append(os.path.basename(todos[-1]))

        if not pares:
            messagebox.showwarning("Sem pares", "Nenhum par pôde ser formado com os PDFs da pasta.")
            return

        for cot, po in pares:
            self._fila.append((cot, po))
            nome_cot = os.path.basename(cot)[:32]
            nome_po  = os.path.basename(po)[:32]
            n = len(self._fila)
            self._listbox.insert("end", f"  {n}   {nome_cot:<32}  {nome_po}")

        msg = f"{len(self._fila)} par(es) na fila."
        if nao_pareados:
            msg += f"  ⚠ Arquivo sem par: {nao_pareados[0]}"
            messagebox.showwarning("Arquivo sem par",
                                   f"O arquivo '{nao_pareados[0]}' ficou sem par.\n"
                                   "Total de PDFs na pasta é ímpar.")
        self._status.set(msg)

    def _importar_multiplos(self):
        """
        Seleção manual em duas etapas:
        1) Seleciona N PDFs de cotações (multi-select)
        2) Seleciona N PDFs de POs     (multi-select)
        Pareia automaticamente por nº REQ ou por ordem.
        """
        cotacoes = filedialog.askopenfilenames(
            title="Selecione os PDFs de COTAÇÕES (pode selecionar vários)",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        if not cotacoes:
            return

        pos = filedialog.askopenfilenames(
            title="Selecione os PDFs de PO (pode selecionar vários)",
            filetypes=[("PDF", "*.pdf"), ("Todos", "*.*")]
        )
        if not pos:
            return

        self._adicionar_pares(list(cotacoes), list(pos), origem="seleção manual")

    def _adicionar_pares(self, cotacoes: list, pos: list, origem: str = ""):
        """Para a lista de cotações e POs, pareia e adiciona à fila."""
        pares, sem_cot, sem_po = _parear_por_req(cotacoes, pos)

        if not pares:
            messagebox.showwarning(
                "Sem pares",
                "Nenhum par Cotação+PO pôde ser formado.\n\n"
                "Verifique se os arquivos têm o mesmo nº REQ no nome."
            )
            return

        for cot, po in pares:
            self._fila.append((cot, po))
            nome_cot = os.path.basename(cot)[:32]
            nome_po  = os.path.basename(po)[:32]
            n = len(self._fila)
            self._listbox.insert("end", f"  {n}   {nome_cot:<32}  {nome_po}")

        msg = f"{len(self._fila)} par(es) na fila."
        avisos = []
        if sem_cot:
            avisos.append(f"{len(sem_cot)} cotação(ões) sem PO correspondente.")
        if sem_po:
            avisos.append(f"{len(sem_po)} PO(s) sem cotação correspondente.")
        if avisos:
            msg += "  ⚠ " + "  ".join(avisos)
        self._status.set(msg)

        if avisos:
            messagebox.showwarning(
                "Arquivos não pareados",
                "\n".join(avisos) + "\n\nVerifique se o nº REQ aparece no nome dos arquivos."
            )

    def _remover_par(self):
        sel = self._listbox.curselection()
        if not sel:
            messagebox.showinfo("Nenhum selecionado", "Clique em um item da lista para selecioná-lo.")
            return
        idx = sel[0]
        self._listbox.delete(idx)
        self._fila.pop(idx)
        self._renumerar_lista()
        self._status.set(f"{len(self._fila)} par(es) na fila.")

    def _limpar_fila(self):
        if not self._fila:
            return
        if messagebox.askyesno("Limpar fila", "Remover todos os pares da fila?"):
            self._fila.clear()
            self._listbox.delete(0, "end")
            self._status.set("Fila limpa. Importe os arquivos para começar.")

    def _renumerar_lista(self):
        itens = list(self._listbox.get(0, "end"))
        self._listbox.delete(0, "end")
        for i, item in enumerate(itens):
            partes = item.split("   ", 1)
            novo = f"  {i + 1}   {partes[1] if len(partes) > 1 else item}"
            self._listbox.insert("end", novo)

    def _selecionar_pasta_saida(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta para salvar o Excel")
        if pasta:
            self._pasta_saida.set(pasta)

    # ------------------------------------------------------------------
    # Análise
    # ------------------------------------------------------------------

    def _iniciar_analise(self):
        if not _verificar_api_key():
            return
        if not self._fila:
            messagebox.showwarning("Fila vazia", "Importe pelo menos um par de Cotação + PO.")
            return

        self._btn_analisar.config(state="disabled")
        self._progress["value"] = 0
        self._progress["maximum"] = len(self._fila)
        self._status.set(f"Iniciando análise de {len(self._fila)} par(es)...")

        thread = threading.Thread(target=self._executar_analise, daemon=True)
        thread.start()

    def _executar_analise(self):
        lote = []
        erros = []

        for i, (caminho_cotacao, caminho_po) in enumerate(self._fila):
            nome = f"Par {i + 1}/{len(self._fila)}"
            try:
                self._atualizar_status(f"[{nome}] Extraindo cotações...")
                dados_cotacao = extrair_cotacoes(caminho_cotacao)

                self._atualizar_status(f"[{nome}] Extraindo PO...")
                dados_po = extrair_po(caminho_po)

                self._atualizar_status(f"[{nome}] Comparando e validando...")
                analise = analisar(dados_cotacao, dados_po)

                req_numero = _extrair_req_do_nome(caminho_cotacao) or _extrair_req_do_nome(caminho_po)
                lote.append({"analise": analise, "req_numero": req_numero})

            except Exception as e:
                erros.append((i + 1, str(e)))

            self.after(0, lambda v=i + 1: self._progress.configure(value=v))

        if lote:
            self._atualizar_status("Gerando arquivo Excel unificado...")
            try:
                caminho_excel = exportar_excel(lote, self._pasta_saida.get())
            except Exception as e:
                erros.append(("Excel", str(e)))
                caminho_excel = None
        else:
            caminho_excel = None

        self.after(0, lambda: self._concluir(caminho_excel, lote, erros))

    def _atualizar_status(self, mensagem: str):
        self.after(0, lambda: self._status.set(mensagem))

    def _concluir(self, caminho_excel, lote: list, erros: list):
        self._progress["value"] = self._progress["maximum"]
        self._btn_analisar.config(state="normal")

        n_ok = len(lote)
        n_err = len(erros)
        linhas = [f"✓ Análise concluída — {n_ok} par(es) processado(s), {n_err} erro(s).\n"]

        for i, entrada in enumerate(lote):
            analise = entrada["analise"]
            po = analise.get("po", {})
            alertas = len(analise.get("alertas_po", []))
            melhor = (analise.get("melhor_preco") or {}).get("nome", "N/A")
            numero_po = po.get("numero_po") or f"Par {i + 1}"
            eco_req = po.get("numero_eco_req") or entrada.get("req_numero") or "—"
            linhas.append(
                f"  PO: {numero_po} | REQ: {eco_req} | {alertas} alerta(s) | Melhor: {melhor}"
            )

        for n, erro in erros:
            linhas.append(f"  Par {n}: ERRO — {erro[:80]}")

        if caminho_excel:
            linhas.append(f"\nArquivo salvo: {os.path.basename(caminho_excel)}")
            self._status.set(f"✓ Concluído! {n_ok} par(es), {n_err} erro(s). Excel salvo.")
        else:
            self._status.set(f"Concluído com erros. {n_err} erro(s).")

        resposta = messagebox.askyesno(
            "Análise concluída",
            "\n".join(linhas) + "\n\nDeseja abrir o arquivo Excel?"
        )
        if resposta and caminho_excel:
            os.startfile(caminho_excel)

    def _erro(self, mensagem: str):
        self._btn_analisar.config(state="normal")
        self._status.set(f"Erro: {mensagem}")
        messagebox.showerror("Erro na análise", mensagem)


# =====================================================================
# Ponto de entrada
# =====================================================================

if __name__ == "__main__":
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    app = App()
    app.mainloop()
