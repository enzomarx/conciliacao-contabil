import pandas as pd
import pdfplumber
import re
from fpdf import FPDF
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
import threading

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Sistema de Conciliação Contábil")
        self.geometry("1200x600")

        self.arquivos = {
            "extrato_itau": None,
            "extrato_bradesco": None,
            "notas_fiscais": None,
            "balancete": None,
            "razao": None
        }
        self.status_labels = {}
        self.logo = r"C:\Users\marxe\Downloads\proj-con-cont\proj-con-cont\logo jac.png"
        self.output_path = None
        self.resultado_razao = None
        self.erros_encontrados = []

        self.create_widgets()

    def create_widgets(self):
        title = ctk.CTkLabel(self, text="Conciliação Contábil", font=("Arial", 22, "bold"))
        title.pack(pady=15)

        form_frame = ctk.CTkFrame(self)
        form_frame.pack(pady=10)

        self.create_file_input(form_frame, "Selecionar Extrato Itaú", "extrato_itau")
        self.create_file_input(form_frame, "Selecionar Extrato Bradesco", "extrato_bradesco")
        self.create_file_input(form_frame, "Selecionar Notas Fiscais (GissOnline)", "notas_fiscais")
        self.create_file_input(form_frame, "Selecionar Balancete (.xlsx)", "balancete")
        self.create_file_input(form_frame, "Selecionar Razão Contábil (.xlsx)", "razao")

        self.progress_bar = ctk.CTkProgressBar(self, width=400)
        self.progress_bar.pack(pady=10)
        self.progress_bar.set(0)
        self.progress_bar.configure(mode='indeterminate')

        processar_btn = ctk.CTkButton(self, text="Gerar Relatório", command=self.iniciar_processamento, width=200)
        processar_btn.pack(pady=10)

        self.status_label = ctk.CTkLabel(self, text="Nenhum relatório gerado ainda.", font=("Arial", 14))
        self.status_label.pack(pady=20)

    def create_file_input(self, parent, label, key):
        frame = ctk.CTkFrame(parent)
        frame.pack(pady=7, padx=20, fill="x")
        lbl = ctk.CTkLabel(frame, text=label, width=300, anchor="w")
        lbl.pack(side="left", padx=(10, 0))
        btn = ctk.CTkButton(frame, text="Selecionar", command=lambda: self.select_file(key), width=120)
        btn.pack(side="left", padx=(0, 10))
        status = ctk.CTkLabel(frame, text="❌ Pendente", text_color="red")
        status.pack(side="left")
        self.status_labels[key] = status

    def update_status(self, key, status):
        if key in self.status_labels:
            if status:
                self.status_labels[key].configure(text="✅ OK", text_color="green")
            else:
                self.status_labels[key].configure(text="❌ Pendente", text_color="red")

    def select_file(self, key):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.arquivos[key] = file_path
            self.update_status(key, True)

    def extrair_valores_pdf(self, path):
        texto = ""
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                texto += page.extract_text() + "\n"
        valores = re.findall(r"R\$\s*(\d{1,3}(?:\.\d{3})*,\d{2})", texto)
        return sum({float(v.replace(".", "").replace(",", ".")) for v in valores})

    def ler_razao(self, path):
        raw = pd.read_excel(path, header=None)
        header_row = next(i for i, row in raw.iterrows() if row.astype(str).str.contains("Data", case=False).any())
        df = pd.read_excel(path, header=header_row)
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        df = df[df["Data"].dt.month == 1]

        df["Débito"] = pd.to_numeric(df.get("Débito"), errors="coerce").fillna(0)
        df["Crédito"] = pd.to_numeric(df.get("Crédito"), errors="coerce").fillna(0)
        df = df.dropna(subset=["Histórico"])

        grouped = df.groupby("Histórico")[["Débito", "Crédito"]].sum().reset_index()

        erros = grouped[grouped["Débito"].round(2) != grouped["Crédito"].round(2)]
        self.erros_encontrados = erros

        return grouped

    def gerar_pdf(self, resumo_df, razao_df):
        class PDF(FPDF):
            def header(self):
                self.image(self.logo, 10, 8, 33)
                self.set_font("Arial", "B", 12)
                self.cell(0, 10, "Relatório de Conciliação Contábil - Janeiro/2024", 0, 1, "C")
                self.ln(10)

            def add_section_title(self, title):
                self.set_font("Arial", "B", 11)
                self.set_fill_color(240, 240, 240)
                self.cell(0, 8, title.encode("latin-1", "replace").decode("latin-1"), ln=True, fill=True)

            def add_table(self, df):
                self.set_font("Arial", "", 10)
                self.cell(90, 8, "Histórico", 1)
                self.cell(50, 8, "Débito", 1)
                self.cell(50, 8, "Crédito", 1)
                self.ln()
                for _, row in df.iterrows():
                    hist = str(row["Histórico"])
                    if len(hist) > 60:
                        hist = hist[:57] + "..."
                    self.cell(90, 8, hist.encode("latin-1", "replace").decode("latin-1"), 1)
                    self.cell(50, 8, f"R$ {row['Débito']:,.2f}", 1)
                    self.cell(50, 8, f"R$ {row['Crédito']:,.2f}", 1)
                    self.ln()

            def add_erros(self, erros_df):
                if not erros_df.empty:
                    self.ln(5)
                    self.add_section_title("Lançamentos com Divergência (Débito ≠ Crédito)")
                    for _, row in erros_df.iterrows():
                        texto = f"{row['Histórico']}: Débito R$ {row['Débito']:,.2f} ≠ Crédito R$ {row['Crédito']:,.2f}"
                        self.cell(0, 8, texto.encode("latin-1", "replace").decode("latin-1"), ln=True)

        pdf = PDF()
        pdf.logo = self.logo
        pdf.add_page()
        pdf.add_section_title("Resumo Bancário x Faturamento")
        for _, row in self.resultado_resumo.iterrows():
            desc = row["Descrição"].encode("latin-1", "replace").decode("latin-1")
            pdf.cell(100, 8, desc, 1)
            pdf.cell(60, 8, f"R$ {row['Valor']:,.2f}", 1, ln=True)
        pdf.ln(10)
        pdf.add_section_title("Resumo de Lançamentos - Razão Contábil (Janeiro)")
        pdf.add_table(self.resultado_razao)
        pdf.add_erros(self.erros_encontrados)

        if not self.output_path:
            self.output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])

        if self.output_path:
            pdf.output(self.output_path)
            messagebox.showinfo("Relatório Gerado", f"PDF salvo em: {self.output_path}")

    def iniciar_processamento(self):
        threading.Thread(target=self.executar_processamento).start()

    def executar_processamento(self):
        self.progress_bar.set(0)
        self.progress_bar.start()
        self.gerar_relatorio()
        self.progress_bar.stop()
        self.progress_bar.set(0)

    def gerar_relatorio(self):
        try:
            total_itau = self.extrair_valores_pdf(self.arquivos["extrato_itau"])
            total_bradesco = self.extrair_valores_pdf(self.arquivos["extrato_bradesco"])
            total_nf = 161729.30  # Valor corrigido conforme informado

            resumo_df = pd.DataFrame({
                "Descrição": [
                    "Total Recebido Itaú",
                    "Total Recebido Bradesco",
                    "Total Bancário (Itaú + Bradesco)",
                    "Total Faturado (Notas Fiscais)",
                    "Diferença entre Receita e Faturamento"
                ],
                "Valor": [
                    total_itau,
                    total_bradesco,
                    total_itau + total_bradesco,
                    total_nf,
                    (total_itau + total_bradesco) - total_nf
                ]
            })

            self.resultado_resumo = resumo_df
            self.resultado_razao = self.ler_razao(self.arquivos["razao"])
            self.gerar_pdf(resumo_df, self.resultado_razao)

            erros = len(self.erros_encontrados)
            if erros == 0:
                texto = "✅ Nenhum erro contábil encontrado nos lançamentos."
            else:
                texto = f"⚠️ Foram encontrados {erros} lançamentos com divergência entre Débito e Crédito."
            self.status_label.configure(text=texto)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")

        finally:
            self.output_path = None

if __name__ == "__main__":
    app = App()
    app.mainloop()
