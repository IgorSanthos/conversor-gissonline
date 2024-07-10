import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from PIL import Image, ImageTk
from clientesdata import dados_clientes
from imagem import texto_longo
import base64
from io import BytesIO

def process_csv():
    try:
        filePathIn = filedialog.askopenfilename(filetypes=[("Arquivos GISS CSV", "*.csv")], title="Importar Arquivo Livro - GISONLINE")
        if not filePathIn:
            return
        
        dfGiss = pd.read_csv(filePathIn, encoding='latin1', sep=';', usecols=[14, 25, 24, 6, 7, 1, 19, 4, 3, 23])

        dfGiss.iloc[:, 0:3] = dfGiss.iloc[:, 0:3].astype(str).apply(lambda x: x.str.replace('.', ','))
        dfGiss.iloc[:, 6] = dfGiss.iloc[:, 6].astype(str).apply(lambda x: x.replace('.', ''))

        dia = dfGiss['DIA'].astype(str).str.zfill(2)
        mes = dfGiss['MES_NUM'].astype(str).str.zfill(2)
        dfGiss['TOMA_CPF_CGC'] = dfGiss['TOMA_CPF_CGC'].astype(str).str.zfill(14)
        dfGiss['ATIT_COD_ATIVIDADE'] = dfGiss['ATIT_COD_ATIVIDADE'].str.zfill(16).str.slice(0, 4)
        ano = dfGiss['ANO'].astype(str)
        dfGiss['DIA'] = dia + '/' + mes + '/' + ano + ' 00:00:00'

        dfGiss.rename(columns={'ENFS_NUM_NFS_INI': 'Nº NFS-e', 'DIA': 'Data Hora NFE',
                               'TOMA_CPF_CGC': 'CPF/CNPJ do Prestador', 'VALOR_FATURADO': 'Valor dos Serviços',
                               'ATIT_COD_ATIVIDADE': 'Código do Serviço Prestado na Nota Fiscal',
                               'VALOR_ALIQUOTA': 'Alíquota', 'VALOR_IMPOSTO': 'ISS devido',
                               'TOMA_STA_ESTABELECIDO': 'ISS Retido'}, inplace=True)

        colunasG5 = ['Código de Verificação da NFS-e', 'Tipo de RPS', 'Série do RPS', 'Número do RPS',
                     'Inscrição Municipal do Prestador', 'Tipo do Endereço do Prestador', 'Endereço do Prestador',
                     'Número do Endereço do Prestador', 'Complemento do Endereço do Prestador', 'Bairro do Prestador',
                     'Cidade do Prestador', 'Razão Social do Prestador', 'UF do Prestador', 'CEP do Prestador',
                     'Email do Prestador', 'Data de Cancelamento', 'Campo Reservado', 'Nº da Guia',
                     'Data de Quitação da Guia Vinculada a Nota Fiscal', 'Inscrição Municipal do Tomador',
                     'Inscrição Estadual do Tomador', 'Razão Social do Tomador', 'Tipo do Endereço do Tomador',
                     'Nº NFS-e Consolidada', 'Endereço do Tomador', 'Número do Endereço do Tomador',
                     'Complemento do Endereço do Tomador', 'Bairro do Tomador', 'CEP do Tomador', 'Email do Tomador',
                     'Nº NFS-e Substituta', 'ISS recolhido', 'CPF/CNPJ do Intermediário',
                     'Inscrição Municipal do Intermediário', 'Razão Social do Intermediário', 'Repasse do Plano de Saúde',
                     'Carga tributária: Porcentagem', 'Carga tributária: Fonte', 'Situação do Aceite',
                     'Tipo de Consolidação', 'Discriminação dos Serviços']
        for colunasVazias in colunasG5:
            dfGiss[colunasVazias] = ''

        selected_cliente = combo_cliente.get()
        if selected_cliente in clientes:
            cnpjTomador = clientes[selected_cliente]["cnpj"]
            nomeCidade = clientes[selected_cliente]["cidade"]
        else:
            messagebox.showinfo("Cancelar", "Operação cancelada pelo usuário.")
            return

        dfGiss['Data do Fato Gerador'] = dfGiss['Data Hora NFE'].str.slice(0, 10)
        dfGiss['Tipo de Registro'] = ['2'] * len(dfGiss)
        dfGiss['UF do Tomador'] = ['SP'] * len(dfGiss)
        dfGiss = dfGiss.assign(**{'Indicador de CPF/CNPJ do Prestador': '2'})
        dfGiss = dfGiss.assign(**{'Situação da Nota Fiscal': 'T'})
        dfGiss = dfGiss.assign(**{'Indicador de CPF/CNPJ do Tomador': '2'})
        dfGiss = dfGiss.assign(**{'CPF/CNPJ do Tomador': cnpjTomador})
        dfGiss = dfGiss.assign(**{'Cidade do Tomador': nomeCidade})
        dfGiss = dfGiss.assign(**{'Indicador de CPF/CNPJ do Intermediário': '3'})

        colunazero = ['Opção Pelo Simples', 'Valor das Deduções', 'Valor do Crédito', 'ISS a recolher', 'PIS/PASEP',
                      'COFINS', 'INSS', 'IR', 'CSLL', 'Carga tributária: Valor', 'CEI', 'Matrícula da Obra',
                      'Município Prestação - cód. IBGE', 'Encapsulamento', 'Valor Total Recebido']
        for coluna in colunazero:
            dfGiss[coluna] = '0'

        ordemColunas = ["Tipo de Registro", "Nº NFS-e", "Data Hora NFE", "Código de Verificação da NFS-e", "Tipo de RPS",
                        "Série do RPS", "Número do RPS", "Data do Fato Gerador", "Inscrição Municipal do Prestador",
                        "Indicador de CPF/CNPJ do Prestador", "CPF/CNPJ do Prestador", "Razão Social do Prestador",
                        "Tipo do Endereço do Prestador", "Endereço do Prestador", "Número do Endereço do Prestador",
                        "Complemento do Endereço do Prestador", "Bairro do Prestador", "Cidade do Prestador",
                        "UF do Prestador", "CEP do Prestador", "Email do Prestador", "Opção Pelo Simples",
                        "Situação da Nota Fiscal", "Data de Cancelamento", "Nº da Guia",
                        "Data de Quitação da Guia Vinculada a Nota Fiscal", "Valor dos Serviços", "Valor das Deduções",
                        "Código do Serviço Prestado na Nota Fiscal", "Alíquota", "ISS devido", "Valor do Crédito",
                        "ISS Retido", "Indicador de CPF/CNPJ do Tomador", "CPF/CNPJ do Tomador",
                        "Inscrição Municipal do Tomador", "Inscrição Estadual do Tomador", "Razão Social do Tomador",
                        "Tipo do Endereço do Tomador", "Endereço do Tomador", "Número do Endereço do Tomador",
                        "Complemento do Endereço do Tomador", "Bairro do Tomador", "Cidade do Tomador", "UF do Tomador",
                        "CEP do Tomador", "Email do Tomador", "Nº NFS-e Substituta", "ISS recolhido", "ISS a recolher",
                        "Indicador de CPF/CNPJ do Intermediário", "CPF/CNPJ do Intermediário",
                        "Inscrição Municipal do Intermediário", "Razão Social do Intermediário",
                        "Repasse do Plano de Saúde", "PIS/PASEP", "COFINS", "INSS", "IR", "CSLL",
                        "Carga tributária: Valor", "Carga tributária: Porcentagem", "Carga tributária: Fonte", "CEI",
                        "Matrícula da Obra", "Município Prestação - cód. IBGE", "Situação do Aceite", "Encapsulamento",
                        "Valor Total Recebido", "Tipo de Consolidação", "Nº NFS-e Consolidada", "Campo Reservado",
                        'Discriminação dos Serviços']
        dfGiss = dfGiss[ordemColunas]

        if messagebox.askyesno("Salvar Arquivo", "Deseja salvar o arquivo?"):
            file_path_out = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("Arquivos CSV", "*.csv")])
            if file_path_out:
                dfGiss.to_csv(file_path_out, sep=';', index=False)
                messagebox.showinfo("Sucesso", "Arquivo salvo com sucesso!")
            else:
                messagebox.showinfo("Cancelar", "Operação cancelada pelo usuário.")
        else:
            messagebox.showinfo("Cancelar", "Operação cancelada pelo usuário.")

    except Exception as error:
        messagebox.showerror("Erro", f"Ocorreu o seguinte erro: {error}")



base64_string = texto_longo()
image_data = base64.b64decode(base64_string)

pil_image = Image.open(BytesIO(image_data))

# Cria a janela principal
root = tk.Tk()
root.title("Conversor-GISS Online")
root.geometry("510x600")
root.configure(bg='#4D9D76')

# Cria um Label com a imagem de fundo
bg_image = ImageTk.PhotoImage(pil_image)
background_label = tk.Label(root, image=bg_image)
background_label.place(x=0, y=0, relwidth=1, relheight=1)
background_label.configure(bg='#4D9D76')
# Garantir que a imagem seja mantida em memória pelo tkinter
root.image = bg_image

# Adiciona texto dentro da janela
label = tk.Label(root, text="ConversoR", font=('Helvetica', 70, 'bold'), bg="#4D9D76", fg="white")
label.pack(pady=2)
label = tk.Label(root, text="Converte planilhas do giss online para arquivo importavel do g5phoenix", font=('Helvetica', 10), bg="#4D9D76", fg="#FDD04C")
label.pack(pady=1)

# Seletor de Cliente
clientes = dados_clientes

combo_cliente_label = tk.Label(root, text="Selecione o Cliente:", font=('Helvetica', 12), bg="#4D9D76", fg="white")
combo_cliente_label.pack(pady=5)

# options
style_combobox = ttk.Style()
style_combobox.configure("Custom.TCombobox", background="#FDD04C", foreground="#FF6D55", font=('Helvetica', 14, 'bold'))
# Criar a Combobox usando o estilo personalizado
combo_cliente = ttk.Combobox(root, width=30, style='Custom.TCombobox')
combo_cliente['values'] = list(clientes.keys())  # Define os valores da Combobox
combo_cliente.pack(pady=5)

# Botão
style = ttk.Style()
style.configure("TButton", background="#FDD04C", foreground="#FF6D55", width=20, height=5, font=('Helvetica', 14, 'bold'))
style.map("TButton", background=[('active', 'white')], foreground=[('active', '#FDD04C')])
botao = ttk.Button(root, text="Importar Planilha GISS", command=process_csv, style="TButton")
botao.pack(pady=20)

# Inicia o loop principal da janela
root.mainloop()
