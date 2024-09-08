import tkinter as tk
from tkinter import messagebox
import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import os
import threading

class WhatsAppBotApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Bot")
        self.running = False

        self.run_button = tk.Button(root, text="Rodar Código", command=self.start_process)
        self.run_button.pack(pady=10)

        self.stop_button = tk.Button(root, text="Cancelar Processo", command=self.stop_process)
        self.stop_button.pack(pady=10)

    def start_process(self):
        self.running = True
        threading.Thread(target=self.send_messages).start()

    def stop_process(self):
        self.running = False
        messagebox.showinfo("Parar", "Processo cancelado.")

    def send_messages(self):
        if not os.path.exists('clientes.xlsx'):
            messagebox.showerror("Erro", "O arquivo 'clientes.xlsx' não foi encontrado.")
            return

        webbrowser.open('https://web.whatsapp.com/')
        sleep(30)

        workbook = openpyxl.load_workbook('clientes.xlsx')
        pagina_clientes = workbook['Sheet1']

        for linha in pagina_clientes.iter_rows(min_row=2):
            if not self.running:
                break

            nome = linha[0].value
            telefone = linha[1].value    
            mensagem = f'Olá {nome}, tudo bem? Seu pacote de limpeza da vitrine expirou, realize o pagamento por gentileza na seguinte chave pix: *44.691.237/0001-31* para que possamos realizar um novo ciclo de atendimento.'

            try:
                link_mensagem_whatsapp = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                webbrowser.open(link_mensagem_whatsapp)
                sleep(10)
                seta = pyautogui.locateCenterOnScreen('seta.png')
                if seta is not None:
                    sleep(5)
                    pyautogui.click(seta[0], seta[1])
                    sleep(5)
                    pyautogui.hotkey('ctrl', 'w')
                    sleep(5)
                else:
                    raise ValueError("Não foi possível localizar a seta na tela.")
            except Exception as e:
                print(f'Não foi possível enviar mensagem para {nome}. Erro: {e}')
                with open('erros.csv', 'a', newline='', encoding='utf-8') as arquivo:
                    arquivo.write(f'{nome},{telefone},{e}{os.linesep}')

        messagebox.showinfo("Concluído", "Processo de envio de mensagens concluído.")

# Inicializar a aplicação Tkinter
root = tk.Tk()
app = WhatsAppBotApp(root)
root.mainloop()
