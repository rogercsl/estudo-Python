import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import PyPDF2

# Configurações iniciais
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# Tela inicial
class MainApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Central de Unificações")
        self.geometry("400x300")
        self.resizable(False, False)  # Bloqueia redimensionamento

        for i in range(4):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self, text="Escolha uma opção", font=ctk.CTkFont(size=18, weight="bold")).grid(row=0, column=0, pady=(40, 10), sticky="n")
        ctk.CTkButton(self, text="Unificar Excel", command=self.abrir_excel).grid(row=1, column=0, pady=10, padx=20, sticky="ew")
        ctk.CTkButton(self, text="Unificar PDF", command=self.abrir_pdf).grid(row=2, column=0, pady=10, padx=20, sticky="ew")
        ctk.CTkLabel(self, text="by Roger Cassol & Vitor Fidelis", font=ctk.CTkFont(size=10)).grid(row=3, column=0, pady=10, sticky="s")

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
        self.geometry("400x500")
        self.resizable(False, False)  # Bloqueia redimensionamento

        self.arquivos = []
        self.destino = ""

        for i in range(9):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(self, text="Selecionar Arquivos", command=self.selecionarArquivos).grid(row=0, column=0, padx=20, pady=5, sticky="ew")

        self.listaArquivos = ctk.CTkTextbox(self, height=100)
        self.listaArquivos.grid(row=1, column=0, padx=20, pady=5, sticky="nsew")
        self.listaArquivos.configure(state="disabled")

        ctk.CTkButton(self, text="Selecionar Destino", command=self.selecionarDestino).grid(row=2, column=0, padx=20, pady=20, sticky="ew")

        self.pathDestino1 = ctk.CTkTextbox(self, height=70)
        self.pathDestino1.grid(row=3, column=0, padx=20, pady=5, sticky="nsew")
        self.pathDestino1.configure(state="disabled")

        ctk.CTkButton(self, text="Unificar Excel", command=self.unificarExcel).grid(row=4, column=0, padx=20, pady=20, sticky="ew")
        ctk.CTkButton(self, text="Voltar", command=self.voltar).grid(row=5, column=0, padx=20, pady=5, sticky="ew")
        ctk.CTkLabel(self, text="by Roger Cassol & Vitor Fidelis").grid(row=8, column=0, pady=10, sticky="s")

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
# Tela de unificação de PDF
class PDFApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Unificador de PDFs")
        self.geometry("400x500")
        self.resizable(False, False)

        self.arquivos_pdf = []
        self.destino_pdf = ""

        for i in range(9):
            self.grid_rowconfigure(i, weight=1)
        self.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(self, text="Selecionar PDFs", command=self.selecionar_pdfs).grid(row=0, column=0, padx=20, pady=5, sticky="ew")

        self.lista_pdfs = ctk.CTkTextbox(self, height=100)
        self.lista_pdfs.grid(row=1, column=0, padx=20, pady=5, sticky="nsew")
        self.lista_pdfs.configure(state="disabled")

        ctk.CTkButton(self, text="Selecionar Destino", command=self.selecionar_destino).grid(row=2, column=0, padx=20, pady=20, sticky="ew")

        self.caminho_destino = ctk.CTkTextbox(self, height=70)
        self.caminho_destino.grid(row=3, column=0, padx=20, pady=5, sticky="nsew")
        self.caminho_destino.configure(state="disabled")

        ctk.CTkButton(self, text="Unificar PDFs", command=self.unificar_pdfs).grid(row=4, column=0, padx=20, pady=20, sticky="ew")
        ctk.CTkButton(self, text="Voltar", command=self.voltar).grid(row=5, column=0, padx=20, pady=5, sticky="ew")

        ctk.CTkLabel(self, text="by Roger Cassol & Vitor Fidelis").grid(row=8, column=0, pady=10, sticky="s")

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