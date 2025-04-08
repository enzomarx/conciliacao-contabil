import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfMerger

def merge_pdfs():
    file_paths = filedialog.askopenfilenames(
        title="Selecione os arquivos PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )

    if not file_paths:
        return

    merger = PdfMerger()
    for pdf in file_paths:
        merger.append(pdf)

    save_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF file", "*.pdf")],
        title="Salvar PDF unido como"
    )

    if save_path:
        try:
            merger.write(save_path)
            merger.close()
            messagebox.showinfo("Sucesso", f"PDFs unidos com sucesso em:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao salvar o PDF: {str(e)}")

root = tk.Tk()
root.title("Unir PDFs")
root.geometry("300x150")

label = tk.Label(root, text="Clique no botão para unir seus PDFs")
label.pack(pady=20)

merge_button = tk.Button(root, text="Selecionar e Unir PDFs", command=merge_pdfs)
merge_button.pack(pady=10)

root.mainloop()
