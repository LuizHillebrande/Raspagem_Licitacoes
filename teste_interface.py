import customtkinter as ctk
from tkinter import messagebox
from PIL import Image
import openpyxl
import os

# Funções de raspagem (essas são apenas placeholders, substitua pelas suas funções de raspagem reais)
def iniciar_raspagem_pncp():
    messagebox.showinfo("Raspagem Portal Nacional", "Iniciando raspagem de dados no Portal Nacional de Licitações...")

def iniciar_raspagem_site1():
    messagebox.showinfo("Raspagem Site 1", "Iniciando raspagem de dados no Site 1...")

def iniciar_raspagem_site2():
    messagebox.showinfo("Raspagem Site 2", "Iniciando raspagem de dados no Site 2...")

def iniciar_raspagem_site3():
    messagebox.showinfo("Raspagem Site 3", "Iniciando raspagem de dados no Site 3...")

def iniciar_raspagem_site4():
    messagebox.showinfo("Raspagem Site 4", "Iniciando raspagem de dados no Site 4...")

def eliminar_cnpj_repetido():
    # Lista de arquivos que devem ser processados
    arquivos = [
        "dados_fornecedores_1.xlsx", 
        "dados_fornecedores_2.xlsx", 
        "dados_fornecedores_3.xlsx", 
        "dados_fornecedores_4.xlsx", 
        "dados_fornecedores_5.xlsx"
    ]

    # Verifica se ao menos um arquivo existe
    arquivos_existentes = [arquivo for arquivo in arquivos if os.path.exists(arquivo)]

    if len(arquivos_existentes) == 0:
        # Exibe mensagem de erro se nenhum arquivo for encontrado
        messagebox.showwarning("Atenção", "Nenhum arquivo de dados encontrado para eliminar CNPJs duplicados.")
        return

    # Processa os arquivos encontrados
    cnpjs_unicos = carregar_cnpjs_de_arquivos(arquivos_existentes)
    salvar_cnpjs_unicos(cnpjs_unicos)
    messagebox.showinfo("Concluído", "CNPJs duplicados eliminados com sucesso!")

def carregar_cnpjs_de_arquivos(arquivos):
    cnpjs_unicos = set()  # Usando um set para garantir que não haja CNPJs duplicados

    for arquivo in arquivos:
        wb = openpyxl.load_workbook(arquivo)
        ws = wb.active
        
        for row in ws.iter_rows(min_row=2, max_col=2, values_only=True):  # Pular o cabeçalho
            cnpj = row[1] 
            if cnpj:
                cnpjs_unicos.add(cnpj.strip())  # Adiciona o CNPJ ao set (remover espaços extras)

    return cnpjs_unicos

def salvar_cnpjs_unicos(cnpjs_unicos, nome_arquivo="cnpjs_unicos.xlsx"):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["CNPJ"])  # Cabeçalho da planilha

    for cnpj in cnpjs_unicos:
        ws.append([cnpj])

    wb.save(nome_arquivo)
    print(f"CNPJs únicos salvos em: {nome_arquivo}")

def enviar_emails_site1():
    messagebox.showinfo("Enviar E-mails Site 1", "Enviando e-mails para o Site 1...")

def enviar_emails_site2():
    messagebox.showinfo("Enviar E-mails Site 2", "Enviando e-mails para o Site 2...")

def enviar_emails_site3():
    messagebox.showinfo("Enviar E-mails Site 3", "Enviando e-mails para o Site 3...")

def enviar_emails_site4():
    messagebox.showinfo("Enviar E-mails Site 4", "Enviando e-mails para o Site 4...")

# Configuração do tema escuro
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")

# Criação da janela principal
app = ctk.CTk()
app.title("Sistema de Raspagem de Licitações")
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
app.geometry(f"{screen_width}x{screen_height-40}+0+0")

# Título principal
title_label = ctk.CTkLabel(
    app, 
    text="Sistema de Raspagem de Licitações", 
    font=("Helvetica", 30, "bold"), 
    text_color="#ffffff"
)
title_label.pack(pady=20)

# Frame principal que contém todos os módulos
main_frame = ctk.CTkFrame(app, fg_color="transparent")
main_frame.pack(padx=20, pady=20, fill="both", expand=True)

# Frame para os 5 módulos
frame_modulos = ctk.CTkFrame(main_frame)
frame_modulos.pack(pady=30, fill="both", expand=True)

# Módulo 1 - Portal Nacional de Licitações
frame_pncp = ctk.CTkFrame(frame_modulos)
frame_pncp.pack(side="left", padx=20, pady=20, fill="y", expand=True)

# Imagem de exemplo para o módulo
img_pncp = ctk.CTkImage(Image.open("img.png"), size=(300, 100))  # Substitua pelo caminho correto da imagem
label_img_pncp = ctk.CTkLabel(frame_pncp, image=img_pncp, text="")
label_img_pncp.pack(pady=10)

label_pncp = ctk.CTkLabel(frame_pncp, text="Portal Nacional de Licitações", font=("Helvetica", 18, "bold"))
label_pncp.pack(pady=10)

button_pncp = ctk.CTkButton(
    frame_pncp, 
    text="Iniciar Raspagem", 
    command=iniciar_raspagem_pncp,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_pncp.pack(pady=10)

button_enviar_pncp = ctk.CTkButton(
    frame_pncp, 
    text="Enviar E-mails", 
    command=enviar_emails_site1,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#4caf50",  
    hover_color="#388e3c",  
)
button_enviar_pncp.pack(pady=10)

# Módulo 2 - Outro Site de Licitações
frame_site1 = ctk.CTkFrame(frame_modulos)
frame_site1.pack(side="left", padx=20, pady=20, fill="y", expand=True)

# Imagem de exemplo para o módulo
img_site1 = ctk.CTkImage(Image.open("img.png"), size=(300, 100))  # Substitua pelo caminho correto da imagem
label_img_site1 = ctk.CTkLabel(frame_site1, image=img_site1, text="")
label_img_site1.pack(pady=10)

label_site1 = ctk.CTkLabel(frame_site1, text="Outro Site de Licitações 1", font=("Helvetica", 18, "bold"))
label_site1.pack(pady=10)

button_site1 = ctk.CTkButton(
    frame_site1, 
    text="Iniciar Raspagem", 
    command=iniciar_raspagem_site1,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_site1.pack(pady=10)

button_enviar_site1 = ctk.CTkButton(
    frame_site1, 
    text="Enviar E-mails", 
    command=enviar_emails_site2,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#4caf50",  
    hover_color="#388e3c",  
)
button_enviar_site1.pack(pady=10)

# Módulo 3 - Outro Site de Licitações
frame_site2 = ctk.CTkFrame(frame_modulos)
frame_site2.pack(side="left", padx=20, pady=20, fill="y", expand=True)

# Imagem de exemplo para o módulo
img_site2 = ctk.CTkImage(Image.open("img.png"), size=(300, 100))  # Substitua pelo caminho correto da imagem
label_img_site2 = ctk.CTkLabel(frame_site2, image=img_site2, text="")
label_img_site2.pack(pady=10)

label_site2 = ctk.CTkLabel(frame_site2, text="Outro Site de Licitações 2", font=("Helvetica", 18, "bold"))
label_site2.pack(pady=10)

button_site2 = ctk.CTkButton(
    frame_site2, 
    text="Iniciar Raspagem", 
    command=iniciar_raspagem_site2,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_site2.pack(pady=10)

button_enviar_site2 = ctk.CTkButton(
    frame_site2, 
    text="Enviar E-mails", 
    command=enviar_emails_site3,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#4caf50",  
    hover_color="#388e3c",  
)
button_enviar_site2.pack(pady=10)

# Módulo 4 - Outro Site de Licitações
frame_site3 = ctk.CTkFrame(frame_modulos)
frame_site3.pack(side="left", padx=20, pady=20, fill="y", expand=True)

# Imagem de exemplo para o módulo
img_site3 = ctk.CTkImage(Image.open("img.png"), size=(300, 100))  # Substitua pelo caminho correto da imagem
label_img_site3 = ctk.CTkLabel(frame_site3, image=img_site3, text="")
label_img_site3.pack(pady=10)

label_site3 = ctk.CTkLabel(frame_site3, text="Outro Site de Licitações 3", font=("Helvetica", 18, "bold"))
label_site3.pack(pady=10)

button_site3 = ctk.CTkButton(
    frame_site3, 
    text="Iniciar Raspagem", 
    command=iniciar_raspagem_site3,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_site3.pack(pady=10)

button_enviar_site3 = ctk.CTkButton(
    frame_site3, 
    text="Enviar E-mails", 
    command=enviar_emails_site4,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#4caf50",  
    hover_color="#388e3c",  
)
button_enviar_site3.pack(pady=10)

# Módulo 5 - Outro Site de Licitações
frame_site4 = ctk.CTkFrame(frame_modulos)
frame_site4.pack(side="left", padx=20, pady=20, fill="y", expand=True)

# Imagem de exemplo para o módulo
img_site4 = ctk.CTkImage(Image.open("img.png"), size=(300, 100))  # Substitua pelo caminho correto da imagem
label_img_site4 = ctk.CTkLabel(frame_site4, image=img_site4, text="")
label_img_site4.pack(pady=10)

label_site4 = ctk.CTkLabel(frame_site4, text="Outro Site de Licitações 4", font=("Helvetica", 18, "bold"))
label_site4.pack(pady=10)

button_site4 = ctk.CTkButton(
    frame_site4, 
    text="Iniciar Raspagem", 
    command=iniciar_raspagem_site4,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#007acc",  
    hover_color="#005b99",  
)
button_site4.pack(pady=10)

button_enviar_site4 = ctk.CTkButton(
    frame_site4, 
    text="Enviar E-mails", 
    command=enviar_emails_site4,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#4caf50",  
    hover_color="#388e3c",  
)
button_enviar_site4.pack(pady=10)

# Rodapé
footer_frame = ctk.CTkFrame(app, fg_color="transparent")
footer_frame.pack(side="bottom", pady=20)

button_eliminar = ctk.CTkButton(
    footer_frame, 
    text="Eliminar CNPJ Repetido", 
    command=eliminar_cnpj_repetido,
    font=("Helvetica", 14), 
    width=250, 
    height=40, 
    fg_color="#ff5722",  
    hover_color="#d43f00",  
)
button_eliminar.pack(pady=10)

# Iniciar o loop principal da interface gráfica
app.mainloop()
