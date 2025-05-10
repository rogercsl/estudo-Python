import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog, messagebox

def unificar_pdfs():
    pasta = filedialog.askdirectory(title="Selecione a pasta com os arquivos PDF")

    if not pasta:
        messagebox.showinfo("Cancelado", "Nenhuma pasta selecionada.")
        return

    lista_arquivos = os.listdir(pasta)
    lista_arquivos.sort()

    merger = PyPDF2.PdfMerger()
    pdfs_encontrados = False

    for arquivo in lista_arquivos:
        if arquivo.lower().endswith(".pdf"):
            caminho_completo = os.path.join(pasta, arquivo)
            merger.append(caminho_completo)
            pdfs_encontrados = True

    if pdfs_encontrados:
        destino = os.path.join(pasta, "PDF Unificado.pdf")
        merger.write(destino)
        merger.close()
        messagebox.showinfo("Sucesso", f"PDFs unificados com sucesso em:\n{destino}")
    else:
        messagebox.showwarning("Nenhum PDF encontrado", "A pasta selecionada não contém arquivos PDF.")

root = tk.Tk()
root.withdraw()

unificar_pdfs()