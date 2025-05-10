import customtkinter as ctk
from tkinter import filedialog, messagebox
import subprocess
import os
import threading

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

def escolher_diretorio():
    pasta = filedialog.askdirectory()
    if pasta:
        pasta_entry.configure(state="normal")
        pasta_entry.delete(0, "end")
        pasta_entry.insert(0, pasta)
        pasta_entry.configure(state="disabled")

def iniciar_download_thread():
    threading.Thread(target=baixar_audio, daemon=True).start()

def atualizar_progresso(percent_str):
    try:
        percent = float(percent_str.strip('%')) / 100
        progress_bar.set(percent)
    except:
        pass

def baixar_audio():
    link = link_entry.get()
    destino = pasta_entry.get()

    if not link or not destino:
        messagebox.showerror("Erro", "Preencha todos os campos.")
        return

    try:
        # Definir o comando para baixar o áudio e converter para MP3
        output = os.path.join(destino, "%(title)s.%(ext)s")

        comando = [
            "python", "-m", "yt_dlp",
            "-f", "bestaudio",  # Baixa o melhor áudio disponível
            "--extract-audio",  # Extrai o áudio
            "--audio-format", "mp3",  # Converte para MP3
            "--audio-quality", "0",  # Melhor qualidade de áudio
            "-o", output,  # Caminho do arquivo de saída
        ]

        comando.append(link)

        processo = subprocess.Popen(comando, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)

        for linha in processo.stdout:
            if "[download]" in linha and "%" in linha:
                partes = linha.strip().split()
                for parte in partes:
                    if "%" in parte:
                        atualizar_progresso(parte)
                        break

        processo.wait()
        progress_bar.set(1.0)
        messagebox.showinfo("Sucesso", f"Áudio convertido para MP3 com sucesso em:\n{destino}")

    except subprocess.CalledProcessError as e:
        messagebox.showerror("Erro", f"Erro ao baixar:\n{str(e)}")

# GUI
app = ctk.CTk()
app.title("YouTube MP3 Converter")
app.geometry("500x450")

ctk.CTkLabel(app, text="YouTube MP3 Converter", font=ctk.CTkFont(size=20, weight="bold")).pack(pady=10)

ctk.CTkLabel(app, text="Link do Vídeo:").pack()
link_entry = ctk.CTkEntry(app, width=400)
link_entry.pack(pady=5)

ctk.CTkLabel(app, text="Pasta de Destino:").pack()
frame_pasta = ctk.CTkFrame(app)
frame_pasta.pack(pady=5)

pasta_entry = ctk.CTkEntry(frame_pasta, width=300, state="disabled")
pasta_entry.pack(side="left", padx=(0, 5))
ctk.CTkButton(frame_pasta, text="Escolher", command=escolher_diretorio).pack(side="left")

ctk.CTkButton(app, text="Iniciar Download", command=iniciar_download_thread).pack(pady=20)

progress_bar = ctk.CTkProgressBar(app, width=400)
progress_bar.set(0)
progress_bar.pack(pady=10)

app.mainloop()