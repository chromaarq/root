import tkinter as tk
from tkinter import messagebox, font

# Template do contrato
template_contrato = """
CONTRATO DE PRESTAÇÃO DE SERVIÇOS

Pelo presente instrumento particular, de um lado,
{nome_cliente}, residente e domiciliado em {endereco_cliente},
e de outro lado, {nome_empresa}, inscrita no CNPJ sob o nº {cnpj_empresa}, 
estabelecida na {endereco_empresa}, têm entre si justo e acordado o presente
Contrato de Prestação de Serviços, que se regerá pelas cláusulas e condições
seguintes:

1. CLÁUSULA PRIMEIRA - DO OBJETO
1.1. O objeto do presente contrato é a prestação de serviços de {descricao_servico}.

2. CLÁUSULA SEGUNDA - DO PRAZO E DA RESCISÃO
2.1. O presente contrato terá início em {data_inicio} e término em {data_fim}, podendo ser rescindido por qualquer das partes mediante aviso prévio de 30 dias.

3. CLÁUSULA TERCEIRA - DO VALOR E DA FORMA DE PAGAMENTO
3.1. O valor total dos serviços prestados será de R$ {valor_servico}, que será pago da seguinte forma: {forma_pagamento}.

4. CLÁUSULA QUARTA - DAS OBRIGAÇÕES DAS PARTES
4.1. O contratante se obriga a fornecer todas as informações necessárias para a realização do serviço.
4.2. A contratada se obriga a prestar os serviços contratados de forma diligente e profissional.

Assim, por estarem justos e acordados, assinam o presente contrato em duas vias de igual teor e forma, na presença das testemunhas abaixo.

Arapongas, {data_contrato}.

_____________________________________
{nome_cliente}

_____________________________________
{nome_empresa}

TESTEMUNHAS:

_____________________________________
Nome:
CPF:

_____________________________________
Nome:
CPF:
"""

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

# Função para exibir o contrato em uma nova janela
def exibir_contrato(contrato):
    janela_contrato = tk.Toplevel(root)
    janela_contrato.title("Contrato Gerado")

    texto_contrato = tk.Text(janela_contrato, wrap='word', font=(fonte_var.get(), tamanho_fonte_var.get(), estilo_fonte_var.get()))
    texto_contrato.insert('1.0', contrato)
    texto_contrato.pack(expand=True, fill='both')

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
fonte_var = tk.StringVar(value="Arial")
tk.OptionMenu(root, fonte_var,"Nunito", "Arial", "Helvetica", "Times", "Courier", "Verdana").grid(row=11, column=1)

# Seleção de Tamanho de Fonte
tk.Label(root, text="Tamanho da Fonte:").grid(row=12, column=0)
tamanho_fonte_var = tk.IntVar(value=12)
tk.Spinbox(root, from_=8, to=72, textvariable=tamanho_fonte_var).grid(row=12, column=1)

# Seleção de Estilo de Fonte
tk.Label(root, text="Estilo da Fonte:").grid(row=13, column=0)
estilo_fonte_var = tk.StringVar(value="normal")
tk.OptionMenu(root, estilo_fonte_var, "normal", "bold", "italic", "bold italic").grid(row=13, column=1)

tk.Button(root, text="Gerar Contrato", command=coletar_dados).grid(row=14, columnspan=2)

root.mainloop()
