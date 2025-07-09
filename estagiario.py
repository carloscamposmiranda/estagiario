import os
import sys
import threading
import subprocess
from tkinter import *
from tkinter import ttk, filedialog, messagebox
from PyPDF2 import PdfMerger
import comtypes.client


def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.abspath("."))
    return os.path.join(base_path, relative_path)


def convert_to_pdf(input_path, output_path):
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)

    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(input_path)
        doc.SaveAs(output_path, FileFormat=17)
        doc.Close()
    except Exception as e:
        raise e
    finally:
        word.Quit()


class EstagiarioApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Estagiário Jurídico - Assistente de Conversão")
        self.root.geometry("1024x700")
        self.root.minsize(800, 600)
        self.root.configure(bg="#f4f4f4")

        # Configuração responsiva principal
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(3, weight=1)  # Linha das abas

        try:
            self.root.iconbitmap(resource_path("icones/estagiario.ico"))
        except:
            pass

        # Carrega ícones com tamanho ajustado
        icon_size = 24  # Tamanho reduzido para melhor visualização
        self.icon_folder = PhotoImage(file=resource_path("icones/icons/folder.png")).subsample(icon_size // 8,
                                                                                               icon_size // 8)
        self.icon_save = PhotoImage(file=resource_path("icones/icons/save.png")).subsample(icon_size // 8,
                                                                                           icon_size // 8)
        self.icon_play = PhotoImage(file=resource_path("icones/icons/play.png")).subsample(icon_size // 8,
                                                                                           icon_size // 8)
        self.icon_pause = PhotoImage(file=resource_path("icones/icons/pause.png")).subsample(icon_size // 8,
                                                                                             icon_size // 8)
        self.icon_reset = PhotoImage(file=resource_path("icones/icons/reset.png")).subsample(icon_size // 8,
                                                                                             icon_size // 8)

        # Variáveis de controle
        self.doc_paths = []
        self.pasta_origem = ""
        self.pasta_destino = ""
        self.remover_word = BooleanVar(value=True)
        self.unificar_pdfs = BooleanVar(value=False)
        self.abrir_destino = BooleanVar(value=True)
        self.pausado = False
        self.executando = False

        self.arquivos_convertidos = []
        self.arquivos_com_erro = []

        self.construir_interface()

    def construir_interface(self):
        """Constrói a interface com layout melhorado"""
        # Frame do cabeçalho
        header_frame = Frame(self.root, bg="#f4f4f4")
        header_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        header_frame.grid_columnconfigure(0, weight=1)

        # Logo e título
        self.logo = PhotoImage(file=resource_path("icones/icons/logo_estagiario.png"))
        Label(header_frame, image=self.logo, bg="#f4f4f4").pack()
        Label(header_frame, text="Estagiário Jurídico", font=("Segoe UI", 22, "bold"),
              bg="#f4f4f4", fg="#3f3f3f").pack(pady=(5, 0))
        Label(header_frame, text="Seu assistente de confiança para tarefas jurídicas do dia a dia.",
              font=("Segoe UI", 11), bg="#f4f4f4", fg="#666").pack(pady=(0, 10))

        # Frame dos botões - agora mais compacto
        buttons_frame = Frame(self.root, bg="#f4f4f4")
        buttons_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=5)

        # Configuração de colunas para os botões
        for i in range(5):
            buttons_frame.grid_columnconfigure(i, weight=1, uniform="buttons")

        btn_config = {
            "compound": LEFT,
            "font": ("Segoe UI", 9, "bold"),
            "relief": FLAT,
            "bd": 0,
            "cursor": "hand2",
            "padx": 8,
            "pady": 6
        }

        # Função para efeito hover
        def add_hover_effect(btn, color):
            btn.original_color = color
            btn.bind("<Enter>", lambda e: btn.config(bg=color))
            btn.bind("<Leave>", lambda e: btn.config(bg=btn.original_color))

        # Botões principais
        buttons = [
            (self.icon_folder, " Origem", self.selecionar_pasta_origem, "#fca311", "#e39300"),
            (self.icon_save, " Destino", self.selecionar_pasta_destino, "#2a9d8f", "#238b80"),
            (self.icon_play, " Iniciar", self.iniciar, "#007bff", "#006be0"),
            (self.icon_pause, " Pausar", self.pausar, "#6c757d", "#5a6268"),
            (self.icon_reset, " Reiniciar", self.reiniciar, "#dc3545", "#c82333")
        ]

        for i, (icon, text, cmd, color, hover) in enumerate(buttons):
            btn = Button(buttons_frame, image=icon, text=text, command=cmd,
                         bg=color, fg="white", **btn_config)
            btn.grid(row=0, column=i, sticky="ew", padx=2)
            add_hover_effect(btn, hover)

        # Frame dos checkboxes
        checks_frame = Frame(self.root, bg="#f4f4f4")
        checks_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=5)
        checks_frame.grid_columnconfigure(0, weight=1)

        # Checkboxes com estilo moderno
        Checkbutton(checks_frame, text="❌ Remover arquivos Word após conversão",
                    variable=self.remover_word, bg="#f4f4f4", font=("Segoe UI", 10),
                    activebackground="#f4f4f4").grid(row=0, column=0, sticky="w")

        Checkbutton(checks_frame, text="🔗 Unificar PDFs convertidos em um só",
                    variable=self.unificar_pdfs, bg="#f4f4f4", font=("Segoe UI", 10),
                    activebackground="#f4f4f4").grid(row=1, column=0, sticky="w")

        Checkbutton(checks_frame, text="📂 Abrir pasta de destino ao finalizar",
                    variable=self.abrir_destino, bg="#f4f4f4", font=("Segoe UI", 10),
                    activebackground="#f4f4f4").grid(row=2, column=0, sticky="w")

        # Notebook (abas)
        self.tabs = ttk.Notebook(self.root)
        self.tabs.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))

        # Aba de Conversão
        self.frame_conv = Frame(self.tabs, bg="#ffffff")
        self.tabs.add(self.frame_conv, text="🧠 CONVERSÃO EM TEMPO REAL")
        self.frame_conv.grid_rowconfigure(0, weight=1)
        self.frame_conv.grid_columnconfigure(0, weight=1)

        self.log_text = Text(self.frame_conv, wrap="word", font=("Consolas", 10), bg="#fefefe")
        self.log_text.grid(row=0, column=0, sticky="nsew")

        scroll_conv = Scrollbar(self.frame_conv, command=self.log_text.yview)
        scroll_conv.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=scroll_conv.set)

        # Aba de Resumo
        self.frame_resumo = Frame(self.tabs, bg="#f9f9f9")
        self.tabs.add(self.frame_resumo, text="📋 RESUMO DA SESSÃO")
        self.frame_resumo.grid_columnconfigure(0, weight=1)

        # Conteúdo da aba de resumo
        Label(self.frame_resumo, text="Resumo da Conversão", font=("Segoe UI", 14, "bold"),
              bg="#f9f9f9").grid(row=0, column=0, sticky="w", padx=20, pady=10)

        # Labels de status
        self.lbl_total = Label(self.frame_resumo, text="🔍 Total de arquivos: 0", bg="#f9f9f9")
        self.lbl_total.grid(row=1, column=0, sticky="w", padx=20)

        self.lbl_convertidos = Label(self.frame_resumo, text="✅ Convertidos: 0", bg="#f9f9f9")
        self.lbl_convertidos.grid(row=2, column=0, sticky="w", padx=20, pady=2)

        self.lbl_erros = Label(self.frame_resumo, text="❌ Erros: 0", bg="#f9f9f9")
        self.lbl_erros.grid(row=3, column=0, sticky="w", padx=20, pady=2)

        # Lista de sucessos
        Label(self.frame_resumo, text="✔️ Arquivos convertidos:",
              bg="#f9f9f9").grid(row=4, column=0, sticky="w", padx=20, pady=(10, 0))

        self.lista_sucesso = Text(self.frame_resumo, height=8, bg="#eaffea")
        self.lista_sucesso.grid(row=5, column=0, sticky="nsew", padx=20)

        scroll_sucesso = Scrollbar(self.frame_resumo, command=self.lista_sucesso.yview)
        scroll_sucesso.grid(row=5, column=1, sticky="ns")
        self.lista_sucesso.config(yscrollcommand=scroll_sucesso.set)

        # Lista de erros
        Label(self.frame_resumo, text="❌ Arquivos com erro:",
              bg="#f9f9f9").grid(row=6, column=0, sticky="w", padx=20, pady=(10, 0))

        self.lista_erros = Text(self.frame_resumo, height=8, bg="#ffeaea")
        self.lista_erros.grid(row=7, column=0, sticky="nsew", padx=20, pady=(0, 20))

        scroll_erros = Scrollbar(self.frame_resumo, command=self.lista_erros.yview)
        scroll_erros.grid(row=7, column=1, sticky="ns")
        self.lista_erros.config(yscrollcommand=scroll_erros.set)

        # Configura weights para os frames de texto
        self.frame_resumo.grid_rowconfigure(5, weight=1)
        self.frame_resumo.grid_rowconfigure(7, weight=1)

    # Os métodos restantes permanecem os mesmos...
    def selecionar_pasta_origem(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de origem")
        if pasta:
            self.pasta_origem = pasta
            self.doc_paths = [
                os.path.join(dp, f)
                for dp, _, files in os.walk(pasta)
                for f in files
                if f.lower().endswith(('.doc', '.docx')) and not f.startswith('~$')
            ]
            self.log(f"📂 Origem: {pasta}")
            self.log(f"🔍 {len(self.doc_paths)} arquivos encontrados.")

    def selecionar_pasta_destino(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta de destino")
        if pasta:
            self.pasta_destino = pasta
            self.log(f"💾 Destino: {pasta}")

    def log(self, msg):
        self.log_text.insert(END, msg + "\n")
        self.log_text.see(END)

    def iniciar(self):
        if not self.pasta_origem or not self.pasta_destino:
            messagebox.showerror("Erro", "Selecione as pastas de origem e destino.")
            return

        if not self.doc_paths:
            messagebox.showerror("Erro", "Nenhum arquivo Word encontrado na pasta de origem.")
            return

        self.arquivos_convertidos.clear()
        self.arquivos_com_erro.clear()
        self.executando = True
        threading.Thread(target=self.executar, daemon=True).start()

    def pausar(self):
        self.pausado = not self.pausado
        estado = "⏸️ Pausado" if self.pausado else "▶️ Continuando"
        self.log(estado)

    def reiniciar(self):
        self.pausado = False
        self.executando = False
        self.doc_paths.clear()
        self.arquivos_convertidos.clear()
        self.arquivos_com_erro.clear()
        self.lista_sucesso.delete(1.0, END)
        self.lista_erros.delete(1.0, END)
        self.log_text.delete(1.0, END)
        self.lbl_total.config(text="🔍 Total de arquivos: 0")
        self.lbl_convertidos.config(text="✅ Convertidos: 0")
        self.lbl_erros.config(text="❌ Erros: 0")
        self.log("🔄 Interface reiniciada.")

    def executar(self):
        total_arquivos = len(self.doc_paths)
        self.log(f"🚀 Iniciando conversão de {total_arquivos} arquivos...")

        for i, caminho in enumerate(self.doc_paths, 1):
            while self.pausado:
                if not self.executando:
                    return
                continue

            if not os.path.exists(caminho):
                erro_msg = f"Arquivo não encontrado: {caminho}"
                self.arquivos_com_erro.append((os.path.basename(caminho), erro_msg))
                self.log(f"⚠️ {erro_msg}")
                continue

            nome = os.path.basename(caminho)
            try:
                caminho = os.path.normpath(caminho)
                pasta_nome = os.path.basename(os.path.dirname(caminho))
                nome_base = os.path.splitext(nome)[0]

                pasta_nome = pasta_nome.replace('%20', ' ').strip()
                nome_base = nome_base.replace('%20', ' ').strip()

                novo_nome = f"{pasta_nome} - {nome_base}.pdf"
                destino = os.path.join(self.pasta_destino, novo_nome)

                os.makedirs(self.pasta_destino, exist_ok=True)

                self.log(f"⏳ Convertendo ({i}/{total_arquivos}): {nome}...")
                convert_to_pdf(caminho, destino)

                if self.remover_word.get() and os.path.exists(caminho):
                    os.remove(caminho)

                self.arquivos_convertidos.append(destino)
                self.log(f"✅ Convertido: {novo_nome}")

            except Exception as e:
                erro_msg = str(e)
                self.arquivos_com_erro.append((nome, erro_msg))
                self.log(f"❌ Erro em {nome}: {erro_msg}")
                continue

        self.atualizar_resumo()

        if self.unificar_pdfs.get() and self.arquivos_convertidos:
            self.unificar_arquivos_pdf()

        if self.abrir_destino.get() and self.pasta_destino:
            try:
                subprocess.Popen(f'explorer "{os.path.normpath(self.pasta_destino)}"')
            except:
                pass

        self.tabs.select(1)
        self.executando = False
        self.log("🎉 Conversão concluída!")

    def atualizar_resumo(self):
        total = len(self.doc_paths)
        convertidos = len(self.arquivos_convertidos)
        erros = len(self.arquivos_com_erro)

        self.lbl_total.config(text=f"🔍 Total de arquivos: {total}")
        self.lbl_convertidos.config(text=f"✅ Convertidos: {convertidos}")
        self.lbl_erros.config(text=f"❌ Erros: {erros}")

        self.lista_sucesso.delete(1.0, END)
        for nome in self.arquivos_convertidos:
            self.lista_sucesso.insert(END, f"✔️ {os.path.basename(nome)}\n")

        self.lista_erros.delete(1.0, END)
        for nome, erro in self.arquivos_com_erro:
            self.lista_erros.insert(END, f"❌ {nome} → {erro}\n")

    def unificar_arquivos_pdf(self):
        try:
            if not self.arquivos_convertidos:
                self.log("⚠️ Nenhum arquivo PDF para unificar.")
                return

            self.log("📎 Unificando PDFs...")

            merger = PdfMerger()
            for pdf in self.arquivos_convertidos:
                if os.path.exists(pdf):
                    merger.append(pdf)

            saida = os.path.join(self.pasta_destino, "Unificado_Estagiario.pdf")
            merger.write(saida)
            merger.close()

            self.log(f"📦 PDF unificado criado: {saida}")
            self.arquivos_convertidos.append(saida)
            self.atualizar_resumo()

        except Exception as e:
            self.log(f"❌ Erro ao unificar PDFs: {e}")


if __name__ == "__main__":
    root = Tk()
    app = EstagiarioApp(root)
    root.mainloop()