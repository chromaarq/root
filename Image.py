import tkinter as tk
from PIL import Image, ImageTk

# Criar a janela principal
root = tk.Tk()
root.title("Exibição de Imagem")

# Configurar o tamanho da janela
root.geometry("800x600")

# Carregar a imagem
image_path = "Logotipo_Chroma.png"
image = Image.open(image_path)
photo = ImageTk.PhotoImage(image)

# Criar um widget Label para exibir a imagem
label = tk.Label(root, image=photo)
label.image = photo  # Necessário para manter a referência
label.pack()

# Mostrar a janela
root.mainloop()