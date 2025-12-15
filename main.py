from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from PyPDF2 import PdfReader 
from openpyxl import Workbook
import re
import os
from datetime import datetime

# --- LÓGICA DE NEGÓCIO (Mantida, mas o salvamento do arquivo usa a data/hora para evitar sobrescrever) ---

def selecionar_pdf():
    caminho = filedialog.askopenfilename(
        title="Selecione um PDF",
        filetypes=[("Arquivos PDF", "*.pdf")]
    )
    
    if caminho:
        label_caminho_arquivo.config(text=caminho, foreground="#33CC33") # Feedback visual
        carregar_pdf(caminho)
    else:
        label_caminho_arquivo.config(text="Nenhum arquivo selecionado.", foreground="yellow")

wb = Workbook()
ws = wb.active
ws.title = "Folha de Pagamento"

# Definição dos cabeçalhos (Mantida)
headers = [
    "NOME", "CARGO", "ADMISSÃO", "SALÁRIO BRUTO", "DIAS V.T", "V.T MÊS", 
    "CESTA", "SAL. FAMILIA", "COMPL.", "$ BRUTO", "ADIANT.", "OUTROS DESC.", 
    "INSS", "FGTS", "VT 6%", "LÍQUIDO"
]
for i, header in enumerate(headers):
    ws.cell(row=1, column=i+1, value=header)

def carregar_pdf(caminho):
    global ws, wb
    
    # Resetar a linha inicial para 2 (logo após o cabeçalho)
    ultima_linha = 2 
    
    label_status.config(text="Extraindo dados do PDF...", foreground="yellow")
    transformador.update_idletasks() # Atualiza a GUI imediatamente

    try:
        reader = PdfReader(caminho)
        texto = ""

        for i, pagina in enumerate(reader.pages):
            texto += pagina.extract_text() + "\n"
            label_status.config(text=f"Processando página {i+1} de {len(reader.pages)}...", foreground="yellow")
            transformador.update_idletasks()

        # Regex Principal (Mantida)
        regex_funcionario = re.compile(
            r'(\n\s*\d{1,2}.*?)\s*'
            r'(Valor FGTS:\s*[\d\.,]+)', 
            re.DOTALL
        )
        
        blocos_funcionarios = regex_funcionario.findall(texto)
        
        # --- 2. Extração de Dados e Preenchimento do Excel ---
        for bloco_principal, trecho_fgts in blocos_funcionarios:
            
            bloco = bloco_principal + trecho_fgts
            dados_linha = {}
            
            # ... (Lógica de Extração de Dados omitida por brevidade, mas mantida a mesma) ...

            # Extração dos campos (mantida a lógica de regex anterior)
            match_nome = re.search(r'([A-Z\s]+?)\s*Empr\.:', bloco)
            dados_linha['NOME'] = match_nome.group(1).strip() if match_nome else 'N/A'
            match_cargo = re.search(r'Cargo:\s*\d*([A-Z\s]+?)\s*[\d\.,]+\s*Salário:', bloco)
            dados_linha['CARGO'] = match_cargo.group(1).strip() if match_cargo else 'N/A'
            match_adm = re.search(r'Empr\.:\s*(\d{2}/\d{2}/\d{4})\s*Adm:', bloco)
            dados_linha['ADMISSAO'] = match_adm.group(1).strip() if match_adm else ''
            match_proventos = re.search(r'Proventos:\s*([\d\.]+,\d{2})', bloco)
            salario_bruto = match_proventos.group(1) if match_proventos else '0,00'
            dados_linha['SALARIO_BRUTO'] = salario_bruto
            dados_linha['BRUTO'] = salario_bruto
            match_vt_valor = re.search(r'202\s*([\d\.,]+)\s*D\s*VALE TRANSPORTE', bloco)
            vt_valor = match_vt_valor.group(1) if match_vt_valor else '0,00'
            dados_linha['VT_MES'] = vt_valor
            dados_linha['VT_6_PERCENT'] = vt_valor
            match_inss = re.search(r'998\s*([\d\.,]+)\s*D\s*P.*?I\.N\.S\.S\.', bloco, re.DOTALL)
            inss_valor = match_inss.group(1) if match_inss else '0,00'
            dados_linha['INSS'] = inss_valor
            match_outros_desc = re.search(r'341\s*([\d\.,]+)\s*D\s*CONTRIB SOCIAL', bloco)
            outros_desc_valor = match_outros_desc.group(1) if match_outros_desc else '0,00'
            dados_linha['OUTROS_DESC'] = outros_desc_valor
            match_adiantamento = re.search(r'ADIANTAMENTO EXTRA\s*([\d\.,]+)', bloco)
            dados_linha['ADIANT'] = match_adiantamento.group(1) if match_adiantamento else '0,00'
            match_fgts = re.search(r'Valor FGTS:\s*([\d\.]+,\d{2})', bloco)
            dados_linha['FGTS'] = match_fgts.group(1) if match_fgts else '0,00'
            match_liquido_final = re.search(r'Informativa Dedutora:\s*\d\s*([\d\.]+,\d{2})', bloco)
            liquido_valor = match_liquido_final.group(1).strip() if match_liquido_final else '0,00'
            dados_linha['LIQUIDO'] = liquido_valor
            dados_linha['DIAS_VT'] = ''
            dados_linha['CESTA'] = ''
            dados_linha['SAL_FAMILIA'] = ''
            dados_linha['COMPL'] = ''

            # Adicionar a linha ao Excel na ordem correta
            ws.append([
                dados_linha['NOME'], dados_linha['CARGO'], dados_linha['ADMISSAO'], 
                dados_linha['SALARIO_BRUTO'], dados_linha['DIAS_VT'], dados_linha['VT_MES'], 
                dados_linha['CESTA'], dados_linha['SAL_FAMILIA'], dados_linha['COMPL'], 
                dados_linha['BRUTO'], dados_linha['ADIANT'], dados_linha['OUTROS_DESC'], 
                dados_linha['INSS'], dados_linha['FGTS'], dados_linha['VT_6_PERCENT'], 
                dados_linha['LIQUIDO']
            ])

            ultima_linha += 1
        
        # --- SALVAMENTO (Adicionando data e hora no nome do arquivo) ---
        
        agora = datetime.now().strftime("%Y%m%d_%H%M%S")
        nome_arquivo_excel = f"Folha_Pagamento_Extraida_{agora}.xlsx"
        
        caminho_downloads = os.path.expanduser('~/Downloads')
        caminho_completo = os.path.join(caminho_downloads, nome_arquivo_excel)
        
        wb.save(caminho_completo)
        
        label_status.config(text=f"SUCESSO! {len(blocos_funcionarios)} funcionários extraídos.", foreground="#33CC33")
        label_caminho_arquivo.config(text=f"Salvo em: {caminho_completo}", foreground="#33CC33")

    except Exception as erro:
        label_status.config(text=f"ERRO: Verifique o console.", foreground="red")
        label_caminho_arquivo.config(text="Falha na extração de dados.", foreground="red")
        print("Erro detalhado na extração:", erro)


# --- MELHORIAS DA INTERFACE GRÁFICA ---

transformador = Tk()
transformador.title('Transformador de PDF para Excel')
# transformador.iconbitmap('PDF.ico') # Removido para evitar erro
transformador.geometry('550x350') # Tamanho ajustado
transformador.resizable(False, False)

# 1. Definir um estilo de tema mais moderno
style = ttk.Style(transformador)
style.theme_use('clam') # Tema 'clam' é mais escuro e limpo que o padrão

# Configurar cores de fundo para o tema 'clam'
transformador.configure(bg='#2e2e2e')
style.configure('TFrame', background='#2e2e2e')
style.configure('TButton', font=('Arial', 10, 'bold'), padding=10, background='#4a4a4a', foreground='white')
style.map('TButton', background=[('active', '#6e6e6e')])

# 2. Frame Principal para centralizar e dar padding
main_frame = ttk.Frame(transformador, padding="20 20 20 20")
main_frame.pack(fill='both', expand=True)

# 3. Título
label_titulo = Label(main_frame, 
                     text='📊 Extrator de Folha de Pagamento',
                     bg='#2e2e2e', fg='white', 
                     font=('Arial', 18, 'bold'))
label_titulo.pack(pady=(0, 30))

# 4. Botão de Seleção
botao = ttk.Button(main_frame, 
                   text="Clique para Selecionar o PDF", 
                   command=selecionar_pdf,
                   cursor="hand2")
botao.pack(pady=20, ipadx=20, ipady=10) # Aumenta o tamanho do botão

# 5. Label de Status do Processo (Progresso/Erro)
label_status = Label(main_frame, 
                      text="Aguardando a seleção do arquivo.",
                      bg='#2e2e2e', fg='yellow', 
                      font=('Arial', 11))
label_status.pack(pady=(10, 5))

# 6. Label para exibir o caminho do arquivo / Resultado final
label_caminho_arquivo = Label(main_frame, 
                              text="Caminho do arquivo aqui.",
                              bg='#2e2e2e', fg='white', 
                              font=('Arial', 9))
label_caminho_arquivo.pack(pady=(0, 20))

# 7. Rodapé informativo (opcional, para créditos ou informações)
label_info = Label(main_frame, 
                   bg='#2e2e2e', fg='#6e6e6e', 
                   font=('Arial', 8))
label_info.pack(side=BOTTOM, pady=(10, 0))

transformador.mainloop()