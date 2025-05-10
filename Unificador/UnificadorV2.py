import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import PyPDF2

# Configurações iniciais
dark_bg = "#1f1f1f"
ctk.set_appearance_mode("System")  # "Dark", "Light", "System"
ctk.set_default_color_theme("blue")

# Tela inicial
class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Central de Unificações")
        self.geometry("420x320")
        self.resizable(False, False)
        self.configure(fg_color=dark_bg)

        ctk.CTkLabel(self, text="Central de Unificações", font=ctk.CTkFont(size=22, weight="bold")).pack(pady=(40, 10))

        ctk.CTkButton(self, text="Unificar Excel", command=self.abrir_excel, width=200, corner_radius=10).pack(pady=10)
        ctk.CTkButton(self, text="Unificar PDF", command=self.abrir_pdf, width=200, corner_radius=10).pack(pady=10)

        ctk.CTkLabel(self, text="by Róger Cassol & Vitor Fidelis", font=ctk.CTkFont(size=10), text_color="#aaaaaa").pack(side="bottom", pady=10)

    def abrir_excel(self):
        self.destroy()
        ExcelApp().mainloop()

    def abrir_pdf(self):
        self.destroy()
        PDFApp().mainloop()

# Tela de unificação de Excel
class ExcelApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Unificador de Tabelas Excel")
        self.geometry("450x550")
        self.resizable(False, False)
        self.configure(fg_color=dark_bg)

        self.arquivos = []
        self.destino = ""

        ctk.CTkLabel(self, text="Unificação de Arquivos Excel", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        ctk.CTkButton(self, text="Selecionar Arquivos", command=self.selecionarArquivos, corner_radius=8).pack(pady=10, padx=30, fill="x")

        self.listaArquivos = ctk.CTkTextbox(self, height=100)
        self.listaArquivos.pack(pady=5, padx=30, fill="both")
        self.listaArquivos.configure(state="disabled")

        ctk.CTkButton(self, text="Selecionar Destino", command=self.selecionarDestino, corner_radius=8).pack(pady=15, padx=30, fill="x")

        self.pathDestino1 = ctk.CTkTextbox(self, height=60)
        self.pathDestino1.pack(pady=5, padx=30, fill="both")
        self.pathDestino1.configure(state="disabled")

        ctk.CTkButton(self, text="Unificar Excel", command=self.unificarExcel, fg_color="#4CAF50", hover_color="#45A049", corner_radius=8).pack(pady=20, padx=30, fill="x")

        ctk.CTkButton(self, text="Voltar", command=self.voltar, fg_color="#f44336", hover_color="#d32f2f", corner_radius=8).pack(pady=10, padx=30, fill="x")

        ctk.CTkLabel(self, text="by Róger Cassol & Vitor Fidelis", font=ctk.CTkFont(size=10), text_color="#aaaaaa").pack(side="bottom", pady=10)

    def selecionarArquivos(self):
        path = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.arquivos = path
            self.listaArquivos.configure(state="normal")
            self.listaArquivos.delete("1.0", "end")
            for i in self.arquivos:
                self.listaArquivos.insert("end", i + "\n")
            self.listaArquivos.configure(state="disabled")

    def selecionarDestino(self):
        pathDestino = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if pathDestino:
            self.destino = pathDestino
            self.pathDestino1.configure(state="normal")
            self.pathDestino1.delete("1.0", "end")
            self.pathDestino1.insert("end", pathDestino + "\n")
            self.pathDestino1.configure(state="disabled")

    def unificarExcel(self):
        if not self.arquivos:
            messagebox.showerror("Aviso", "Nenhum Arquivo Selecionado")
            return

        if not self.destino:
            messagebox.showerror("Aviso", "Nenhum Destino Selecionado")
            return

        try:
            with pd.ExcelWriter(self.destino, engine="openpyxl") as writer:
                for arquivo in self.arquivos:
                    df = pd.read_excel(arquivo)
                    nomeAba = os.path.splitext(os.path.basename(arquivo))[0][:31]
                    df.to_excel(writer, sheet_name=nomeAba, index=False)
            messagebox.showinfo("Sucesso", f"Arquivo unificado em:\n{self.destino}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu o seguinte erro:\n{str(e)}")

    def voltar(self):
        self.destroy()
        MainApp().mainloop()

# Tela de unificação de PDF
class PDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Unificador de PDFs")
        self.geometry("450x550")
        self.resizable(False, False)
        self.configure(fg_color=dark_bg)

        self.arquivos_pdf = []
        self.destino_pdf = ""

        ctk.CTkLabel(self, text="Unificação de Arquivos PDF", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=20)

        ctk.CTkButton(self, text="Selecionar PDFs", command=self.selecionar_pdfs, corner_radius=8).pack(pady=10, padx=30, fill="x")

        self.lista_pdfs = ctk.CTkTextbox(self, height=100)
        self.lista_pdfs.pack(pady=5, padx=30, fill="both")
        self.lista_pdfs.configure(state="disabled")

        ctk.CTkButton(self, text="Selecionar Destino", command=self.selecionar_destino, corner_radius=8).pack(pady=15, padx=30, fill="x")

        self.caminho_destino = ctk.CTkTextbox(self, height=60)
        self.caminho_destino.pack(pady=5, padx=30, fill="both")
        self.caminho_destino.configure(state="disabled")

        ctk.CTkButton(self, text="Unificar PDFs", command=self.unificar_pdfs, fg_color="#4CAF50", hover_color="#45A049", corner_radius=8).pack(pady=20, padx=30, fill="x")

        ctk.CTkButton(self, text="Voltar", command=self.voltar, fg_color="#f44336", hover_color="#d32f2f", corner_radius=8).pack(pady=10, padx=30, fill="x")

        ctk.CTkLabel(self, text="by Róger Cassol & Vitor Fidelis", font=ctk.CTkFont(size=10), text_color="#aaaaaa").pack(side="bottom", pady=10)

    def selecionar_pdfs(self):
        arquivos = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        if arquivos:
            self.arquivos_pdf = arquivos
            self.lista_pdfs.configure(state="normal")
            self.lista_pdfs.delete("1.0", "end")
            for arq in self.arquivos_pdf:
                self.lista_pdfs.insert("end", arq + "\n")
            self.lista_pdfs.configure(state="disabled")

    def selecionar_destino(self):
        destino = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if destino:
            self.destino_pdf = destino
            self.caminho_destino.configure(state="normal")
            self.caminho_destino.delete("1.0", "end")
            self.caminho_destino.insert("end", destino + "\n")
            self.caminho_destino.configure(state="disabled")

    def unificar_pdfs(self):
        if not self.arquivos_pdf:
            messagebox.showerror("Erro", "Nenhum arquivo PDF selecionado.")
            return

        if not self.destino_pdf:
            messagebox.showerror("Erro", "Nenhum destino selecionado.")
            return

        try:
            merger = PyPDF2.PdfMerger()
            for arq in self.arquivos_pdf:
                merger.append(arq)
            merger.write(self.destino_pdf)
            merger.close()
            messagebox.showinfo("Sucesso", f"PDF unificado salvo em:\n{self.destino_pdf}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao unir os PDFs:\n{str(e)}")

    def voltar(self):
        self.destroy()
        MainApp().mainloop()

# Executar a tela inicial
if __name__ == "__main__":
    app = MainApp()
    app.mainloop()