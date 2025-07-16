import tkinter as tk
from tkinter import messagebox, font, filedialog
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from docx import Document


# Função para ler o template do contrato de um arquivo 
def ler_template_contrato(): 
    with open('template_contrato.txt', 'r', encoding='utf-8') as file: 
        template = file.read() 
    return template
     
# Carrega o template do contrato 
template_contrato = ler_template_contrato()


# Função para gerar o contrato
def gerar_contrato(dados):
    contrato = template_contrato.format(
        nome_cliente=dados['nome_cliente'],
        endereco_cliente=dados['endereco_cliente'],
        nome_empresa=dados['nome_empresa'],
        cnpj_empresa=dados['cnpj_empresa'],
        endereco_empresa=dados['endereco_empresa'],
        descricao_servico=dados['descricao_servico'],
        data_inicio=dados['data_inicio'],
        data_fim=dados['data_fim'],
        valor_servico=dados['valor_servico'],
        forma_pagamento=dados['forma_pagamento'],
        data_contrato=dados['data_contrato']
    )
    return contrato

# Função para coletar os dados da interface gráfica
def coletar_dados():
    dados = {
        'nome_cliente': entry_nome_cliente.get(),
        'endereco_cliente': entry_endereco_cliente.get(),
        'nome_empresa': entry_nome_empresa.get(),
        'cnpj_empresa': entry_cnpj_empresa.get(),
        'endereco_empresa': entry_endereco_empresa.get(),
        'descricao_servico': entry_descricao_servico.get(),
        'data_inicio': entry_data_inicio.get(),
        'data_fim': entry_data_fim.get(),
        'valor_servico': entry_valor_servico.get(),
        'forma_pagamento': entry_forma_pagamento.get(),
        'data_contrato': entry_data_contrato.get()
    }
    
    contrato = gerar_contrato(dados)
    exibir_contrato(contrato)

# Função para exibir o contrato em uma nova janela e salvar em PDF ou DOCX
def exibir_contrato(contrato):
    janela_contrato = tk.Toplevel(root)
    janela_contrato.title("Contrato Gerado")

    texto_contrato = tk.Text(janela_contrato, wrap='word', font=(fonte_var.get(), tamanho_fonte_var.get(), estilo_fonte_var.get()))
    texto_contrato.insert('1.0', contrato)
    texto_contrato.pack(expand=True, fill='both')
    
    def salvar_pdf():
        arquivo_pdf = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if arquivo_pdf:
            c = canvas.Canvas(arquivo_pdf, pagesize=letter)
            largura, altura = letter

            # Adicionar logomarca
            if logomarca_path:
                c.drawImage(logomarca_path, 50, altura - 100, width=100, height=50)

            # Adicionar texto do contrato
            c.drawString(50, altura - 150, contrato)
            c.save()
            messagebox.showinfo("Salvo em PDF", f"Contrato salvo como {arquivo_pdf}")
    
    def salvar_docx():
        arquivo_docx = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if arquivo_docx:
            doc = Document()
            doc.add_paragraph(contrato)
            doc.save(arquivo_docx)
            messagebox.showinfo("Salvo em DOCX", f"Contrato salvo como {arquivo_docx}")
    
    botao_salvar_pdf = tk.Button(janela_contrato, text="Salvar como PDF", command=salvar_pdf)
    botao_salvar_pdf.pack(pady=10)

    botao_salvar_docx = tk.Button(janela_contrato, text="Salvar como DOCX", command=salvar_docx)
    botao_salvar_docx.pack(pady=10)

# Função para carregar a logomarca
def carregar_logomarca():
    global logomarca_path
    logomarca_path = filedialog.askopenfilename(title="Selecione a Logomarca", filetypes=[("Image Files", "*.png;*.jpg;*.jpeg")])
    if logomarca_path:
        label_logomarca.config(text=f"Logomarca Carregada: {logomarca_path}")

# Variável global para o caminho da logomarca
logomarca_path = None

# Configuração da Interface Gráfica
root = tk.Tk()
root.title("Gerador de Contratos")

tk.Label(root, text="Nome do Cliente:").grid(row=0, column=0)
entry_nome_cliente = tk.Entry(root)
entry_nome_cliente.grid(row=0, column=1)

tk.Label(root, text="Endereço do Cliente:").grid(row=1, column=0)
entry_endereco_cliente = tk.Entry(root)
entry_endereco_cliente.grid(row=1, column=1)

tk.Label(root, text="Nome da Empresa:").grid(row=2, column=0)
entry_nome_empresa = tk.Entry(root)
entry_nome_empresa.grid(row=2, column=1)

tk.Label(root, text="CNPJ da Empresa:").grid(row=3, column=0)
entry_cnpj_empresa = tk.Entry(root)
entry_cnpj_empresa.grid(row=3, column=1)

tk.Label(root, text="Endereço da Empresa:").grid(row=4, column=0)
entry_endereco_empresa = tk.Entry(root)
entry_endereco_empresa.grid(row=4, column=1)

tk.Label(root, text="Descrição do Serviço:").grid(row=5, column=0)
entry_descricao_servico = tk.Entry(root)
entry_descricao_servico.grid(row=5, column=1)

tk.Label(root, text="Data de Início:").grid(row=6, column=0)
entry_data_inicio = tk.Entry(root)
entry_data_inicio.grid(row=6, column=1)

tk.Label(root, text="Data de Fim:").grid(row=7, column=0)
entry_data_fim = tk.Entry(root)
entry_data_fim.grid(row=7, column=1)

tk.Label(root, text="Valor do Serviço:").grid(row=8, column=0)
entry_valor_servico = tk.Entry(root)
entry_valor_servico.grid(row=8, column=1)

tk.Label(root, text="Forma de Pagamento:").grid(row=9, column=0)
entry_forma_pagamento = tk.Entry(root)
entry_forma_pagamento.grid(row=9, column=1)

tk.Label(root, text="Data do Contrato:").grid(row=10, column=0)
entry_data_contrato = tk.Entry(root)
entry_data_contrato.grid(row=10, column=1)

# Seleção de Fonte
tk.Label(root, text="Fonte:").grid(row=11, column=0)
fonte_var = tk.StringVar(value="Nunito")
tk.OptionMenu(root, fonte_var,"Nunito", "Arial", "Helvetica", "Times", "Courier", "Verdana").grid(row=11, column=1)

# Seleção de Tamanho de Fonte
tk.Label(root, text="Tamanho da Fonte:").grid(row=12, column=0)
tamanho_fonte_var = tk.IntVar(value=12)
tk.Spinbox(root, from_=8, to=72, textvariable=tamanho_fonte_var).grid(row=12, column=1)

## Seleção de Estilo de Fonte
tk.Label(root, text="Estilo da Fonte:").grid(row=13, column=0)
estilo_fonte_var = tk.StringVar(value="normal")
tk.OptionMenu(root, estilo_fonte_var, "normal", "bold", "italic", "bold italic").grid(row=13, column=1)

tk.Button(root, text="Gerar Contrato", command=coletar_dados).grid(row=14, columnspan=2)

tk.Button(root, text="Carregar Logo", command=carregar_logomarca).grid(row=14, columnspan=2)
label_logomarca = tk.Label(root, text="Logomarca não carregada") 
label_logomarca.grid(row=14, column=1)

root.mainloop()