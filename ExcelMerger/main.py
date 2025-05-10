import customtkinter as ctk
from tkinter import filedialog, messagebox
import pandas as pd
import os

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):  # Corrigido init
        super().__init__()

        self.title("Unificador de Tabelas Excel")
        self.geometry("500x400")
        self.arquivos = []
        self.destino = ""

        # Título
        self.label1 = ctk.CTkLabel(self, text="Selecione os Arquivos", fg_color="transparent")
        self.label1.pack(pady=5)

        # Botão Selecionar Arquivos
        self.botao_arquivos = ctk.CTkButton(self, text="Selecionar Arquivos", command=self.selecionarArquivos)
        self.botao_arquivos.pack(pady=5)

        # Caixa de texto para mostrar arquivos
        self.listaArquivos = ctk.CTkTextbox(self, height=100)
        self.listaArquivos.pack(pady=5, padx=10, fill="both")

        # Seleção do destino
        self.label2 = ctk.CTkLabel(self, text="Selecione o Destino", fg_color="transparent")
        self.label2.pack(pady=5)

        self.botao_destino = ctk.CTkButton(self, text="Selecionar Destino", command=self.selecionarDestino)
        self.botao_destino.pack(pady=5)

        # Botão Unificar
        self.botao_unificar = ctk.CTkButton(self, text="Unificar Excel", command=self.unificarExcel)
        self.botao_unificar.pack(pady=20)

        # Créditos
        self.creditos = ctk.CTkLabel(self, text="by Roger Cassol & Vitor Fidelis", fg_color="transparent")
        self.creditos.pack(side="bottom", pady=5)

    def selecionarArquivos(self):
        paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
        if paths:
            self.arquivos = paths
            self.listaArquivos.delete("1.0", "end")
            for caminho in self.arquivos:
                self.listaArquivos.insert("end", caminho + "\n")

    def selecionarDestino(self):
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if path:
            self.destino = path
            messagebox.showinfo("Destino Selecionado", f"Arquivo será salvo em:\n{path}")

    def unificarExcel(self):
        if not self.arquivos:
            messagebox.showwarning("Aviso", "Nenhum arquivo Excel foi selecionado.")
            return

        if not self.destino:
            messagebox.showwarning("Aviso", "Nenhum destino foi definido.")
            return

        try:
            with pd.ExcelWriter(self.destino, engine="openpyxl") as writer:
                for arquivo in self.arquivos:
                    df = pd.read_excel(arquivo)
                    nome_aba = os.path.splitext(os.path.basename(arquivo))[0][:31]  # Máximo 31 caracteres
                    df.to_excel(writer, sheet_name=nome_aba, index=False)

            messagebox.showinfo("Sucesso", f"Arquivo unificado salvo com sucesso:\n{self.destino}")
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro:\n{str(e)}")

if __name__ == "__main__":  # Corrigido o nome do dunder
    app = App()
    app.mainloop()