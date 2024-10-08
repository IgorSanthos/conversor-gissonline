"""
Para funcionar tive que ir ao arquivo .sped e subsittuir hiddenimports=[] por hiddenimports=['pkg_resources','pkg_resources.extern']
para que ao rodar o pyinstall o programa funcione.
O PyInstaller pode estar falhando ao incluir o setuptools
e tambem.
icon='logo.ico',  # Adicionando o ícone aqui
"""


# Importe as bibliotecas necessárias
import pandas as pd
import tkinter as tk
import requests
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from clientesdata import dados_clientes
import socket
import threading



# Função para validar o acesso
def validate_access(company):
    SERVER_URL = 'https://giss-backend.onrender.com/validate'  # URL do servidor Flask

    # Obtenha o IP local da máquina
    client_ip = socket.gethostbyname(socket.gethostname())
    
    # Construa a URL com os parâmetros da query string (company e ip)
    url_with_params = f"{SERVER_URL}?company={company}&ip={client_ip}"

    try:
        # Faz uma requisição GET para o servidor Flask com os parâmetros na URL
        response = requests.get(url_with_params)  # Usa o método GET
        return response.json().get("status") == "success"
    except requests.RequestException as e:
        print(f"Erro ao fazer a requisição: {e}")
        return False
    

# Função para criar a interface principal
#    
# Função para processar o CSV



def create_main_window():
    
    def process_csv ():
        def process_csv_task ():
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
                dfGiss['ATIT_COD_ATIVIDADE'] = dfGiss['ATIT_COD_ATIVIDADE'].str.zfill(4).str.slice(0, 4)
                ano = dfGiss['ANO'].astype(str)
                dfGiss['DIA'] = dia + '/' + mes + '/' + ano + ' 00:00:00'

                dfGiss.rename(columns={'ENFS_NUM_NFS_INI': 'Nº NFS-e', 'DIA': 'Data Hora NFE',
                                    'TOMA_CPF_CGC': 'CPF/CNPJ do Prestador', 'VALOR_FATURADO': 'Valor dos Serviços',
                                    'ATIT_COD_ATIVIDADE': 'Codigo do Serviço Prestado na Nota Fiscal',
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
                                "Codigo do Serviço Prestado na Nota Fiscal", "Alíquota", "ISS devido", "Valor do Crédito",
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

        threading.Thread(target=process_csv_task, daemon=True).start()

    # JANELA PRINCIPAL
    root = tk.Tk()
    root.title("Conversor-GISS Online")
    root.geometry("510x600")
    root.configure(bg='#4D9D76')

    #   IMAGEM BACKGROUND
    bg_image_path = 'dist/img.png'
    pil_image = Image.open(bg_image_path)
    bg_image = ImageTk.PhotoImage(pil_image)
    background_label = tk.Label(root, image=bg_image)
    background_label.place(x=0, y=0, relwidth=1, relheight=1)
    background_label.configure(bg='#4D9D76')
    root.image = bg_image

    # Adiciona texto dentro da janela
    label = tk.Label(root, text="ConversoR", font=('Helvetica', 70, 'bold'), bg="#4D9D76", fg="white")
    label.pack(pady=2)
    label = tk.Label(root, text="Converte planilhas do giss online para arquivo importavel do g5phoenix", font=('Helvetica', 10), bg="#4D9D76", fg="#FDD04C")
    label.place(relx=0.1, rely=0.18)

    # Seletor de Cliente
    clientes = dados_clientes
    combo_cliente_label = tk.Label(root, text="Selecione o Cliente", font=('Helvetica', 12), bg="#4D9D76", fg="white")
    combo_cliente_label.place(relx=0.35, rely=0.25)

    # Criar uma lista com os nomes dos clientes
    nomes_clientes = list(clientes.keys())

    # Função para atualizar os valores da Combobox conforme a pesquisa
    def atualizar_combobox(event=None):
        filtro = entry_pesquisa.get().strip().lower()
        if filtro:
            valores_filtrados = [nome for nome in nomes_clientes if filtro in nome.lower()]
        else:
            valores_filtrados = nomes_clientes
        combo_cliente['values'] = valores_filtrados


    # Pesquisa
    entry_pesquisa = ttk.Entry(root, width=10, font=('Helvetica', 12))
    entry_pesquisa.place(relx=0.20, rely=0.30)
    entry_pesquisa.bind('<KeyRelease>', atualizar_combobox)

    # Seletor de clientes
    combo_cliente = ttk.Combobox(root, width=20, font=('Helvetica', 12))
    combo_cliente['values'] = nomes_clientes
    combo_cliente.place(relx=0.40, rely=0.30)

    # Botão
    style = ttk.Style()
    style.configure("TButton", background="#FDD04C", foreground="#FF6D55", width=20, height=5, font=('Helvetica', 14, 'bold'))
    style.map("TButton", background=[('active', 'white')], foreground=[('active', '#FDD04C')])
    botao = ttk.Button(root, text="Selecionar Planilha", command=process_csv, style="TButton")
    botao.place(relx=0.27, rely=0.38)

    # Inicia o loop principal da janela
    root.mainloop()


# Função principal para iniciar o programa
def main():
    company = 'contec'  # Substitua pelo nome da empresa
    if validate_access(company):
        create_main_window()
    else:
        messagebox.showerror("Acesso Negado", "Você não tem permissão para acessar este aplicativo.")
        exit()

if __name__ == "__main__":
    main()