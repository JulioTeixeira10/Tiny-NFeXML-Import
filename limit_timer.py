import tkinter as tk
from tkinter import messagebox

def create_timer_window():
    def countdown(seconds):
        if seconds >= 0:
            label.config(text=f"Tempo Restante: {seconds} segundos")
            if seconds > 0:
                root.after(1000, countdown, seconds - 1)
            else:
                #messagebox.showinfo("Pronto!", "Clique em OK para continuar a importação dos XMLs.")
                root.destroy()

    def on_closing():
        pass

    root = tk.Tk()
    root.title("Timer")

    # Calcula a posição central da janela
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    window_width = 450
    window_height = 150
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2

    root.geometry(f"{window_width}x{window_height}+{x}+{y}")

    # Mensagem inicial
    message_label = tk.Label(root, text="Limite de importação de XMLs atingido,\n aguarde o temporizador finalizar para continuar a operação.", font=("Helvetica", 12))
    message_label.pack(pady=10)

    # Label do temporizador
    label = tk.Label(root, text="", font=("Helvetica", 18))
    label.pack(pady=10)

    # Define o tempo para o temporizador
    countdown(60)

    # Previne o fechamento da janela
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()

if __name__ == "__main__":
    create_timer_window()