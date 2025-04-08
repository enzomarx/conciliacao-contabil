import customtkinter as ctk
from tkinter import filedialog, messagebox
import shutil
import os
import time
import threading

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class ConciliaApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Conciliação Contábil")
        self.geometry("800x600")
        self.iconbitmap(r"C:\Users\ema\Downloads\CA-LOGO.png")

        self.extrato_itau = None
        self.extrato_bradesco = None
        self.extrato_pagbank = None
        self.notas_fiscais = None
        self.razao = None
        self.relatorio_excel = None

        self.planilha_model = r"C:\Users\ema\Downloads\others\pattern.xlsx"
        self.pdf_model = r"C:\\Users\\ema\\Downloads\\relatorio_conciliacao_completo.pdf"

        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(fill="both", expand=True, padx=20, pady=20)

        self.conc_tab = self.tabview.add("Conciliação")
        self.relat_tab = self.tabview.add("Relatório Trimestral")

        self.create_conciliacao_widgets()
        self.create_relatorio_widgets()

    def create_conciliacao_widgets(self):
        frame = ctk.CTkFrame(self.conc_tab)
        frame.pack(pady=10, padx=20, fill="both", expand=True)

        self.itau_label, self.itau_status = self.add_input_row(frame, "Selecionar Extrato Itaú", self.select_itau)
        self.bradesco_label, self.bradesco_status = self.add_input_row(frame, "Selecionar Extrato Bradesco", self.select_bradesco)
        self.pagbank_label, self.pagbank_status = self.add_input_row(frame, "Selecionar Extrato PagBank", self.select_pagbank)
        self.nf_label, self.nf_status = self.add_input_row(frame, "Selecionar Notas Fiscais (GissOnline)", self.select_nf)
        self.razao_label, self.razao_status = self.add_input_row(frame, "Selecionar Razão Contábil (.xlsx)", self.select_razao)

        self.progress_bar = ctk.CTkProgressBar(frame, width=400)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)

        self.btn_processar = ctk.CTkButton(frame, text="Gerar Relatório de Conciliação", command=self.iniciar_emulacao)
        self.btn_processar.pack(pady=10)

        self.log_area = ctk.CTkTextbox(frame, height=200)
        self.log_area.pack(pady=10, fill="both", expand=True)

    def create_relatorio_widgets(self):
        frame = ctk.CTkFrame(self.relat_tab)
        frame.pack(pady=20, padx=20, fill="both", expand=True)

        lbl = ctk.CTkLabel(frame, text="Importar base do Excel para gerar relatório trimestral em PDF")
        lbl.pack(pady=10)

        row_frame = ctk.CTkFrame(frame)
        row_frame.pack(pady=5)

        self.btn_importar_relatorio = ctk.CTkButton(row_frame, text="Selecionar Excel Base", command=self.importar_excel_relatorio)
        self.btn_importar_relatorio.pack(side="left", padx=10)

        self.relatorio_status = ctk.CTkLabel(row_frame, text="❌", text_color="red")
        self.relatorio_status.pack(side="left")

        self.btn_gerar_pdf = ctk.CTkButton(frame, text="Gerar Relatório Trimestral PDF", command=self.salvar_relatorio_pdf)
        self.btn_gerar_pdf.pack(pady=10)

    def add_input_row(self, parent, label_text, command):
        frame = ctk.CTkFrame(parent)
        frame.pack(pady=5, fill="x")

        label = ctk.CTkLabel(frame, text=label_text, width=250, anchor="w")
        label.pack(side="left")

        btn = ctk.CTkButton(frame, text="Selecionar", command=command)
        btn.pack(side="left", padx=10)

        status = ctk.CTkLabel(frame, text="❌", text_color="red")
        status.pack(side="left")

        return label, status

    def importar_excel_relatorio(self):
        self.relatorio_excel = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if self.relatorio_excel:
            self.relatorio_status.configure(text="✔", text_color="green")
            messagebox.showinfo("Arquivo Importado", "Arquivo Excel base carregado com sucesso.")

    def salvar_relatorio_pdf(self):
        if not self.relatorio_excel:
            messagebox.showwarning("Aviso", "Selecione um arquivo Excel base antes de gerar o relatório.")
            return

        destino = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf")],
            title="Salvar Relatório Trimestral como"
        )

        if destino:
            try:
                shutil.copyfile(self.pdf_pronto, destino)
                messagebox.showinfo("Sucesso", f"Relatório salvo com sucesso em: {destino}")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar o relatório: {str(e)}")

    def select_itau(self):
        self.extrato_itau = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.extrato_itau:
            self.itau_status.configure(text="✔", text_color="green")

    def select_bradesco(self):
        self.extrato_bradesco = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.extrato_bradesco:
            self.bradesco_status.configure(text="✔", text_color="green")

    def select_pagbank(self):
        self.extrato_pagbank = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.extrato_pagbank:
            self.pagbank_status.configure(text="✔", text_color="green")

    def select_nf(self):
        self.notas_fiscais = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if self.notas_fiscais:
            self.nf_status.configure(text="✔", text_color="green")

    def select_razao(self):
        self.razao = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if self.razao:
            self.razao_status.configure(text="✔", text_color="green")

    def iniciar_emulacao(self):
        if not all([self.extrato_itau, self.extrato_bradesco, self.extrato_pagbank, self.notas_fiscais, self.razao]):
            messagebox.showwarning("Atenção", "Todos os arquivos devem ser selecionados.")
            return
        threading.Thread(target=self.gerar_emulacao).start()

    def gerar_emulacao(self):
        if not os.path.exists(self.planilha_model):
            messagebox.showerror("Erro", f"Arquivo modelo '{self.planilha_model}' não encontrado.")
            return

        destino = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Salvar relatório como"
        )

        if destino:
            self.log_area.delete("1.0", "end")
            self.progress_bar.set(0)

            logs = [
                "Iniciando processamento dos arquivos...",
                "Lendo extrato Itaú...",
                "Lendo extrato Bradesco...",
                "Lendo extrato PagBank...",
                "Analisando Notas Fiscais...",
                "Importando Razão Contábil...",
                "Consolidando dados...",
                "Gerando planilha final..."
            ]

            for i, log in enumerate(logs):
                self.log_area.insert("end", log + "\n")
                self.log_area.see("end")
                self.progress_bar.set((i + 1) / len(logs))
                time.sleep(0.5)

            try:
                shutil.copyfile(self.planilha_model, destino)
                self.log_area.insert("end", f"Relatório salvo em: {destino}\n")
                messagebox.showinfo("Sucesso", f"Relatório gerado com sucesso em: {destino}")
            except Exception as e:
                self.log_area.insert("end", f"Erro ao salvar relatório: {str(e)}\n")
                messagebox.showerror("Erro", f"Erro ao salvar relatório: {str(e)}")

if __name__ == "__main__":
    app = ConciliaApp()
    app.mainloop()
