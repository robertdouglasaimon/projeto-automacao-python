# ------------------------------------------------------------
# main.py
# Sistema de envio automatizado de relatórios via WhatsApp
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------

import os
import tkinter as tk
from tkinter import messagebox
import pandas as pd
import pywhatkit
import pyautogui
import schedule
import datetime
import time

# ------------------------------------------------------------
# Etapa 1 - Tela inicial estilizada para nome e quantidade
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
def interface_inicial():
    def iniciar():
        nome_valor = nome_entry.get().strip()
        qtde_valor = qtde_entry.get().strip()

        if not nome_valor or not qtde_valor.isdigit():
            messagebox.showwarning("Campos obrigatórios", "Preencha corretamente os campos.")
            return

        nonlocal nome_usuario, quantidade
        nome_usuario = nome_valor
        quantidade = int(qtde_valor)
        root.quit()
        
        

    root = tk.Tk()
    root.title("Relatório Diário - Início")
    root.geometry("400x250")
    root.resizable(False, False)

    tk.Label(root, text="Enviar relatório diário", font=("Arial", 14, "bold")).pack(pady=15)

    tk.Label(root, text="Seu nome:", anchor="w").pack(fill="x", padx=30)
    nome_entry = tk.Entry(root)
    nome_entry.pack(padx=30, fill="x")

    tk.Label(root, text="Quantidade de contatos:", anchor="w").pack(pady=(15,0), padx=30, fill="x")
    qtde_entry = tk.Entry(root)
    qtde_entry.pack(padx=30, fill="x")

    tk.Button(root, text="Iniciar", command=iniciar).pack(pady=20)

    nome_usuario = ""
    quantidade = 0
    root.mainloop()
    root.destroy()
    return nome_usuario, quantidade

# ------------------------------------------------------------
# Etapa 2 - Leitura dinâmica da planilha com caminho absoluto
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
def extrair_contatos_do_excel():
    try:
        base_path = os.path.dirname(os.path.abspath(__file__))
        planilha_path = os.path.join(base_path, "..", "assets", "relatorio_diario.xlsx")
        df = pd.read_excel(planilha_path, sheet_name="Contatos", header=1)
        contatos = list(zip(df['Nome'], df['Telefone']))
        return contatos
    except Exception as e:
        messagebox.showerror("Erro ao carregar planilha", f"{e}")
        exit()

# ------------------------------------------------------------
# Etapa 3 - Seleção dos contatos via Listbox interativo
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
def selecionar_contatos_via_lista(contatos, quantidade):
    def confirmar():
        selecionados_idx = listbox.curselection()
        if len(selecionados_idx) != quantidade:
            messagebox.showwarning("Seleção incompleta", f"Selecione exatamente {quantidade} contato(s).")
            return
        for idx in selecionados_idx:
            selecionados.append(contatos[idx])
        janela.quit()

    janela = tk.Toplevel()
    janela.title("Selecione os contatos")
    janela.geometry("400x400")
    janela.resizable(False, False)

    tk.Label(janela, text=f"Selecione {quantidade} contato(s):", font=("Arial", 11)).pack(pady=10)

    listbox = tk.Listbox(janela, selectmode=tk.MULTIPLE, height=15, width=40)
    for contato in contatos:
        listbox.insert(tk.END, contato[0])
    listbox.pack(padx=10, pady=5)

    tk.Button(janela, text="Confirmar", command=confirmar).pack(pady=15)

    selecionados = []
    janela.grab_set()
    janela.mainloop()
    janela.destroy()

    return selecionados

# ------------------------------------------------------------
# Etapa 4 - Envia mensagem no WhatsApp com ENTER automático
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
def enviar_mensagem_whatsapp(numero, mensagem):
    agora = datetime.datetime.now()
    hora = agora.hour
    minuto = (agora.minute + 2) % 60
    if agora.minute >= 58:
        hora = (hora + 1) % 24

    print(f"📤 Agendando envio para +{numero} às {hora:02d}:{minuto:02d}")
    pywhatkit.sendwhatmsg(f"+{numero}", mensagem, hora, minuto, wait_time=15, tab_close=False, close_time=3)

    print("⏳ Aguardando o WhatsApp Web abrir...")
    time.sleep(10)
    print("✅ Pressionando ENTER para enviar!")
    pyautogui.press('enter')

# ------------------------------------------------------------
# Etapa 5 - Agendamento periódico com schedule
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
def agendar_envios(contatos, mensagem, intervalo_minutos=1):
    for nome, telefone in contatos:
        schedule.every(intervalo_minutos).minutes.do(
            lambda nome=nome, telefone=telefone: enviar_mensagem_whatsapp(telefone, mensagem)
        )
    print(f"🕒 Envio programado a cada {intervalo_minutos} min para {len(contatos)} contato(s).")
    while True:
        schedule.run_pending()
        time.sleep(1)

# ------------------------------------------------------------
# Gera a mensagem padrão do relatório
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
def montar_mensagem(nome_operador):
    return (
        "RELATÓRIO DIÁRIO DE ITEM\n\n"
        "Data:\n"
        "Hora:\n"
        "Instalador/auxiliar:\n"
        "Tipo:\n"
        "Largura:\n"
        "Altura:\n"
        "Localização:\n"
        "Já instalado?\n"
        "Se não, o que faltou?\n"
        "Já regulado?\n"
        "Se não, o que faltou?\n"
        "Já calafetado?\n"
        "Se não, o que faltou?\n"
        "Já finalizado?\n"
        "Se não, o que faltou?\n"
        "Faltou algum material/componente?\n"
        "Descrever qual componente ou perfil ou encaminhar a foto de referência\n\n"
        f"Documento enviado por {nome_operador}."
    )

# ------------------------------------------------------------
# Execução principal do sistema
# Desenvolvido por Robert Douglas (https://r-douglas.vercel.app)
# ------------------------------------------------------------
if __name__ == "__main__":
    nome_operador, quantidade = interface_inicial()
    contatos = extrair_contatos_do_excel()
    destinatarios = selecionar_contatos_via_lista(contatos, quantidade)
    mensagem = montar_mensagem(nome_operador)
    agendar_envios(destinatarios, mensagem, intervalo_minutos=1) # Agendar envios com intervalo de 1 minuto
