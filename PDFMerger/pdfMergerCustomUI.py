import customtkinter as ctk
import tkinter.filedialog as fd
import tkinter.messagebox as msg
import PyPDF2
import os

# Configurações da aparência
ctk.set_appearance_mode("System")  # "Dark", "Light", "System"
ctk.set_default_color_theme("blue")  # azul moderno

# Função principal
def unificar_pdfs():
    origem = fd.askdirectory(title="Selecione a pasta com os PDFs")
    if not origem:
        msg.showinfo("Cancelado", "Pasta de origem não selecionada.")
        return

    destino = fd.askdirectory(title="Selecione a pasta de destino")
    if not destino:
        msg.showinfo("Cancelado", "Pasta de destino não selecionada.")
        return

    arquivos = sorted(os.listdir(origem))
    merger = PyPDF2.PdfMerger()
    encontrou_pdf = False

    for arquivo in arquivos:
        if arquivo.lower().endswith(".pdf"):
            merger.append(os.path.join(origem, arquivo))
            encontrou_pdf = True

    if encontrou_pdf:
        caminho_saida = os.path.join(destino, "PDF Unificado.pdf")
        merger.write(caminho_saida)
        merger.close()
        msg.showinfo("Sucesso", f"PDF unificado salvo em:\n{caminho_saida}")
    else:
        msg.showwarning("Nenhum PDF", "Nenhum arquivo PDF foi encontrado na pasta selecionada.")

# Interface principal
app = ctk.CTk()
app.title("Unificador de PDFs")
app.geometry("400x250")

label = ctk.CTkLabel(app, text="Unificador de PDFs", font=ctk.CTkFont(size=20, weight="bold"))
label.pack(pady=20)

btn_unificar = ctk.CTkButton(app, text="Selecionar pastas e unificar PDFs", command=unificar_pdfs)
btn_unificar.pack(pady=10)

creditos = ctk.CTkLabel(app, text="Feito com customtkinter", font=ctk.CTkFont(size=10))
creditos.pack(side="bottom", pady=10)

app.mainloop()