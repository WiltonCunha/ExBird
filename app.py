"""Excel -> Firebird: atualiza colunas de uma tabela usando uma coluna de referência."""
import json
import os
import queue
import threading
import webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import core


APP_NAME = "ExBird - Conversor"
AUTOR = "Wilton Cunha"
LINKEDIN = "https://www.linkedin.com/in/wiltonendril/"

CONFIG_FILE = os.path.join(core.app_dir(), "config.json")
SAVED = ("host", "port", "db", "user", "charset", "table", "path", "sheet")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("880x860")
        self.minsize(780, 700)

        self.headers = []
        self.tables = []
        self.columns = {}  # coluna -> kind
        self.pair_rows = []
        self.q = queue.Queue()
        self.running = False

        self.v = {
            "path": tk.StringVar(), "sheet": tk.StringVar(),
            "host": tk.StringVar(value="localhost"), "port": tk.StringVar(value="3050"),
            "db": tk.StringVar(), "user": tk.StringVar(value="SYSDBA"),
            "password": tk.StringVar(value="masterkey"), "charset": tk.StringVar(value="WIN1252"),
            "table": tk.StringVar(value="ITENS"),
            "ref_x": tk.StringVar(), "ref_db": tk.StringVar(), "ref_pad": tk.StringVar(),
            "dry": tk.BooleanVar(value=True), "write_null": tk.BooleanVar(value=False),
        }
        self._build()
        self._restore()
        self.log("Pronto. Escolha o Excel, preencha o banco e clique em Conectar.")
        self.protocol("WM_DELETE_WINDOW", self._close)
        self.after(100, self._poll)

    # ------------------------------------------------- configurações salvas
    def _restore(self):
        """Reabre com os últimos dados usados (a senha não é salva)."""
        try:
            with open(CONFIG_FILE, encoding="utf-8") as f:
                cfg = json.load(f)
        except Exception:
            return
        for k in SAVED:
            if k in cfg:
                self.v[k].set(cfg[k])
        p = self.v["path"].get()
        if p and os.path.exists(p):
            try:
                sheets = core.excel_sheets(p)
                self.cb_sheet["values"] = sheets
                if self.v["sheet"].get() not in sheets:
                    self.v["sheet"].set(sheets[0] if sheets else "")
                self.headers = core.read_headers(p, self.v["sheet"].get())
                self.lbl_cols.config(text="Colunas encontradas: " + ", ".join(self.headers))
                self.cb_ref_x["values"] = self.headers
                for r in self.pair_rows:
                    r["x"]["values"] = self.headers
            except Exception:
                self.v["path"].set("")

    def _save(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({k: self.v[k].get() for k in SAVED}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # pasta sem permissão de escrita: só não lembra na próxima vez

    def _close(self):
        self._save()
        self.destroy()

    # ------------------------------------------------------------------ UI
    def _build(self):
        pad = {"padx": 6, "pady": 3}

        # rodapé (empacotado primeiro pra ficar sempre visível embaixo)
        rod = ttk.Frame(self)
        rod.pack(side="bottom", fill="x", padx=10, pady=(0, 6))
        ttk.Label(rod, text="Criado por", foreground="#666").pack(side="left")
        link = tk.Label(rod, text=AUTOR, fg="#0a66c2", cursor="hand2", font=("Segoe UI", 9, "underline"))
        link.pack(side="left", padx=(4, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open_new_tab(LINKEDIN))

        # 1. Excel
        f1 = ttk.LabelFrame(self, text="1. Arquivo Excel")
        f1.pack(fill="x", padx=10, pady=(10, 4))
        f1.columnconfigure(1, weight=1)
        ttk.Label(f1, text="Arquivo:").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f1, textvariable=self.v["path"]).grid(row=0, column=1, sticky="ew", **pad)
        ttk.Button(f1, text="Escolher...", command=self.pick_excel).grid(row=0, column=2, **pad)
        ttk.Label(f1, text="Planilha:").grid(row=1, column=0, sticky="w", **pad)
        self.cb_sheet = ttk.Combobox(f1, textvariable=self.v["sheet"], state="readonly", width=30)
        self.cb_sheet.grid(row=1, column=1, sticky="w", **pad)
        self.cb_sheet.bind("<<ComboboxSelected>>", lambda e: self.load_headers())
        self.lbl_cols = ttk.Label(f1, text="", foreground="#555", wraplength=800)
        self.lbl_cols.grid(row=2, column=0, columnspan=3, sticky="w", **pad)

        # 2. Firebird
        f2 = ttk.LabelFrame(self, text="2. Banco Firebird")
        f2.pack(fill="x", padx=10, pady=4)
        f2.columnconfigure(1, weight=1)
        ttk.Label(f2, text="Banco (.fdb):").grid(row=0, column=0, sticky="w", **pad)
        ttk.Entry(f2, textvariable=self.v["db"]).grid(row=0, column=1, columnspan=5, sticky="ew", **pad)
        ttk.Button(f2, text="Escolher...", command=self.pick_db).grid(row=0, column=6, **pad)

        ttk.Label(f2, text="Host:").grid(row=1, column=0, sticky="w", **pad)
        ttk.Entry(f2, textvariable=self.v["host"], width=18).grid(row=1, column=1, sticky="w", **pad)
        ttk.Label(f2, text="Porta:").grid(row=1, column=2, sticky="e", **pad)
        ttk.Entry(f2, textvariable=self.v["port"], width=7).grid(row=1, column=3, sticky="w", **pad)
        ttk.Label(f2, text="Usuário:").grid(row=1, column=4, sticky="e", **pad)
        ttk.Entry(f2, textvariable=self.v["user"], width=12).grid(row=1, column=5, sticky="w", **pad)

        ttk.Label(f2, text="Senha:").grid(row=2, column=0, sticky="w", **pad)
        ttk.Entry(f2, textvariable=self.v["password"], show="*", width=18).grid(row=2, column=1, sticky="w", **pad)
        ttk.Label(f2, text="Charset:").grid(row=2, column=2, sticky="e", **pad)
        ttk.Combobox(f2, textvariable=self.v["charset"], width=10,
                     values=["WIN1252", "UTF8", "ISO8859_1", "NONE"]).grid(row=2, column=3, sticky="w", **pad)
        ttk.Button(f2, text="Conectar e listar tabelas", command=self.connect).grid(
            row=3, column=0, columnspan=3, sticky="w", **pad)

        # 3. Tabela
        f3 = ttk.LabelFrame(self, text="3. Tabela")
        f3.pack(fill="x", padx=10, pady=4)
        ttk.Label(f3, text="Tabela:").grid(row=0, column=0, sticky="w", **pad)
        self.cb_table = ttk.Combobox(f3, textvariable=self.v["table"], width=34)
        self.cb_table.grid(row=0, column=1, sticky="w", **pad)
        self.cb_table.bind("<<ComboboxSelected>>", lambda e: self.load_columns())
        self.cb_table.bind("<Return>", lambda e: self.load_columns())
        ttk.Button(f3, text="Carregar colunas", command=self.load_columns).grid(row=0, column=2, **pad)
        self.lbl_table = ttk.Label(f3, text="Padrão: ITENS (pode digitar ou escolher outra)", foreground="#555")
        self.lbl_table.grid(row=0, column=3, sticky="w", **pad)

        # 4. Referência
        f4 = ttk.LabelFrame(self, text="4. Coluna de referência (precisa bater no Excel e no banco)")
        f4.pack(fill="x", padx=10, pady=4)
        ttk.Label(f4, text="Excel:").grid(row=0, column=0, sticky="w", **pad)
        self.cb_ref_x = ttk.Combobox(f4, textvariable=self.v["ref_x"], state="readonly", width=28)
        self.cb_ref_x.grid(row=0, column=1, **pad)
        ttk.Label(f4, text="  =  Banco:").grid(row=0, column=2, **pad)
        self.cb_ref_db = ttk.Combobox(f4, textvariable=self.v["ref_db"], state="readonly", width=28)
        self.cb_ref_db.grid(row=0, column=3, **pad)
        ttk.Label(f4, text="Zeros à esq. (tam.):").grid(row=0, column=4, **pad)
        ttk.Entry(f4, textvariable=self.v["ref_pad"], width=4).grid(row=0, column=5, **pad)

        # 5. Colunas a atualizar
        f5 = ttk.LabelFrame(self, text="5. Colunas a atualizar (Excel → Banco)  |  \"Zeros\" = completa com zeros à esquerda, só p/ coluna de texto (opcional)")
        f5.pack(fill="x", padx=10, pady=4)
        self.pairs_frame = ttk.Frame(f5)
        self.pairs_frame.pack(fill="x")
        ttk.Button(f5, text="+ Adicionar coluna", command=self.add_pair).pack(anchor="w", **pad)
        self.add_pair()

        # 6. Executar
        f6 = ttk.LabelFrame(self, text="6. Executar")
        f6.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        opts = ttk.Frame(f6)
        opts.pack(fill="x")
        ttk.Checkbutton(opts, text="Simular (não grava nada, só mostra o resultado)",
                        variable=self.v["dry"]).pack(side="left", **pad)
        ttk.Checkbutton(opts, text="Célula vazia grava NULL (senão a linha é pulada)",
                        variable=self.v["write_null"]).pack(side="left", **pad)
        self.btn_run = ttk.Button(opts, text="Executar", command=self.run)
        self.btn_run.pack(side="right", **pad)
        self.bar = ttk.Progressbar(f6, mode="determinate")
        self.bar.pack(fill="x", **pad)
        box = ttk.Frame(f6)
        box.pack(fill="both", expand=True)
        self.txt = tk.Text(box, height=10, state="disabled", wrap="none")
        sb = ttk.Scrollbar(box, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.txt.pack(side="left", fill="both", expand=True)

    # ------------------------------------------------------------- helpers
    def log(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _conn_params(self):
        return {k: self.v[k].get() for k in
                ("host", "port", "db", "user", "password", "charset")}

    def _match(self, name, options):
        for o in options:
            if o.lower() == name.lower():
                return o
        return ""

    # -------------------------------------------------------------- Excel
    def pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if not p:
            return
        self.v["path"].set(p)
        try:
            sheets = core.excel_sheets(p)
        except Exception as e:
            messagebox.showerror("Erro ao abrir Excel", str(e))
            return
        self.cb_sheet["values"] = sheets
        self.v["sheet"].set(sheets[0] if sheets else "")
        self.load_headers()

    def load_headers(self):
        try:
            self.headers = core.read_headers(self.v["path"].get(), self.v["sheet"].get())
        except Exception as e:
            messagebox.showerror("Erro ao ler planilha", str(e))
            return
        self.lbl_cols.config(text="Colunas encontradas: " + ", ".join(self.headers))
        self.cb_ref_x["values"] = self.headers
        for r in self.pair_rows:
            r["x"]["values"] = self.headers
        self._auto_ref()

    # ------------------------------------------------------------ Firebird
    def pick_db(self):
        p = filedialog.askopenfilename(filetypes=[("Firebird", "*.fdb *.gdb"), ("Todos", "*.*")])
        if p:
            self.v["db"].set(p)

    def connect(self):
        try:
            conn = core.open_conn(self._conn_params())
            try:
                self.tables = core.list_tables(conn)
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Erro de conexão", str(e))
            return
        self.cb_table["values"] = self.tables
        self.log(f"Conectado. {len(self.tables)} tabelas encontradas.")
        m = self._match(self.v["table"].get().strip(), self.tables)
        if m:
            self.v["table"].set(m)
            self.load_columns()
        else:
            self.lbl_table.config(text="Tabela padrão não encontrada, escolha uma da lista")

    def load_columns(self):
        if not self.tables:
            messagebox.showinfo("Atenção", "Conecte no banco primeiro.")
            return
        t = self._match(self.v["table"].get().strip(), self.tables)
        if not t:
            messagebox.showerror("Tabela", f"Tabela '{self.v['table'].get()}' não existe no banco.")
            return
        self.v["table"].set(t)
        try:
            conn = core.open_conn(self._conn_params())
            try:
                self.columns = core.list_columns(conn, t)
            finally:
                conn.close()
        except Exception as e:
            messagebox.showerror("Erro ao ler colunas", str(e))
            return
        names = list(self.columns)
        self.cb_ref_db["values"] = names
        for r in self.pair_rows:
            r["d"]["values"] = names
        self.lbl_table.config(text=f"{len(names)} colunas carregadas")
        self._auto_ref()
        for r in self.pair_rows:
            if r["x"].get() and not r["d"].get():
                r["d"].set(self._match(r["x"].get(), names))

    def _auto_ref(self):
        if not self.headers or not self.columns:
            return
        for h in self.headers:
            m = self._match(h, list(self.columns))
            if h.lower() == "id" and m:
                self.v["ref_x"].set(h)
                self.v["ref_db"].set(m)
                return

    # ---------------------------------------------------- colunas a atualizar
    def add_pair(self):
        row = ttk.Frame(self.pairs_frame)
        row.pack(fill="x")
        x = ttk.Combobox(row, state="readonly", width=28, values=self.headers)
        d = ttk.Combobox(row, state="readonly", width=28, values=list(self.columns))
        x.pack(side="left", padx=6, pady=2)
        ttk.Label(row, text="  →  ").pack(side="left")
        d.pack(side="left", padx=6, pady=2)
        ttk.Label(row, text="Zeros:").pack(side="left")
        z = ttk.Entry(row, width=4)
        z.pack(side="left", padx=4)
        item = {"frame": row, "x": x, "d": d, "z": z}
        ttk.Button(row, text="Remover", width=8, command=lambda: self.remove_pair(item)).pack(side="left", padx=6)
        self.pair_rows.append(item)

    def remove_pair(self, item):
        if len(self.pair_rows) > 1:
            item["frame"].destroy()
            self.pair_rows.remove(item)

    # -------------------------------------------------------------- Executar
    def run(self):
        if self.running:
            return
        if not self.headers or not self.columns:
            messagebox.showinfo("Atenção", "Carregue o Excel e as colunas da tabela primeiro.")
            return
        ref_x, ref_db = self.v["ref_x"].get(), self.v["ref_db"].get()
        pairs = [(r["x"].get(), r["d"].get()) for r in self.pair_rows]
        if not ref_x or not ref_db or any(not a or not b for a, b in pairs):
            messagebox.showinfo("Atenção", "Escolha a referência e as colunas (Excel e banco).")
            return
        if any(d == ref_db for _, d in pairs):
            messagebox.showinfo("Atenção", "A coluna de referência não pode ser atualizada.")
            return
        pads = {}
        try:
            for col, txt in [(ref_db, self.v["ref_pad"].get())] + [(r["d"].get(), r["z"].get()) for r in self.pair_rows]:
                if txt.strip():
                    pads[col] = int(txt)
        except ValueError:
            messagebox.showinfo("Atenção", "O campo Zeros precisa ser um número (ex: 8).")
            return
        dry = self.v["dry"].get()
        if not dry and not messagebox.askyesno(
                "Confirmar", "Isso vai GRAVAR no banco. Já fez backup do .fdb?\n\nContinuar?"):
            return
        cfg = {
            "params": self._conn_params(), "path": self.v["path"].get(), "sheet": self.v["sheet"].get(),
            "table": self.v["table"].get(), "ref_x": ref_x, "ref_db": ref_db, "pairs": pairs,
            "kinds": dict(self.columns), "pads": pads, "write_null": self.v["write_null"].get(), "dry": dry,
        }
        self.running = True
        self.btn_run.state(["disabled"])
        self.bar["value"] = 0
        self.log("=" * 50)
        self.log("SIMULAÇÃO (nada será gravado)" if dry else "GRAVANDO no banco...")
        threading.Thread(target=self._worker, args=(cfg,), daemon=True).start()

    def _worker(self, c):
        try:
            headers, rows = core.read_excel(c["path"], c["sheet"])
            self.q.put(("log", f"{len(rows)} linhas lidas do Excel."))
            conn = core.open_conn(c["params"])
            try:
                st = core.run_update(
                    conn, c["table"], c["ref_x"], c["ref_db"], c["pairs"], headers, rows, c["kinds"],
                    write_null=c["write_null"], dry=c["dry"], pads=c["pads"],
                    log=lambda m: self.q.put(("log", m)),
                    progress=lambda a, b: self.q.put(("progress", a, b)),
                )
            finally:
                conn.close()
            self.q.put(("done", st, c["dry"]))
        except Exception as e:
            self.q.put(("fail", str(e)))

    def _poll(self):
        try:
            while True:
                m = self.q.get_nowait()
                if m[0] == "log":
                    self.log(m[1])
                elif m[0] == "progress":
                    self.bar["maximum"] = max(m[2], 1)
                    self.bar["value"] = m[1]
                elif m[0] == "done":
                    st, dry = m[1], m[2]
                    self.log(
                        ("[SIMULAÇÃO] " if dry else "") +
                        f"Total: {st['total']} | Atualizadas: {st['atualizadas']} | "
                        f"Não encontradas: {st['nao_encontradas']} | Puladas: {st['puladas']} | "
                        f"Erros: {st['erros']} | Ref. duplicada: {st['duplicadas']}"
                    )
                    self.log("Nada foi gravado (simulação)." if dry else "Gravado e confirmado no banco.")
                    self._finish()
                elif m[0] == "fail":
                    self.log("ERRO: " + m[1])
                    messagebox.showerror("Erro", m[1])
                    self._finish()
        except queue.Empty:
            pass
        self.after(100, self._poll)

    def _finish(self):
        self.running = False
        self.btn_run.state(["!disabled"])


if __name__ == "__main__":
    App().mainloop()
