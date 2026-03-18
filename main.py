import tkinter as tk
from tkinter import filedialog, messagebox
import win32com.client
import os
import logging
import json
import time
import sys

# ==========================================
# CONFIGURAÇÃO DE LOG
# ==========================================
logging.basicConfig(
    filename='automacao_excel_word.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S'
)

def cm_to_points(cm):
    return cm * 28.3464567

# ==========================================
# CONFIGURAÇÃO DE EXECUTÁVEL
# ==========================================

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

# ==========================================
# SUBSTITUIR PLACEHOLDER (VERSÃO ROBUSTA)
# ==========================================
def substituir_placeholder_word(doc, placeholder, func_insercao):
    rng = doc.Content
    find = rng.Find

    find.ClearFormatting()
    find.Text = placeholder
    find.Forward = True
    find.Wrap = 0  # wdFindStop
    find.MatchCase = False

    encontrou = False

    while find.Execute():
        encontrou = True
        
        range_encontrado = rng.Duplicate
        range_encontrado.Text = ""  # Remove placeholder
        
        func_insercao(range_encontrado)

        rng.Start = range_encontrado.End
        rng.End = doc.Content.End

    if encontrou:
        logging.info(f"{placeholder} substituído com sucesso.")
    else:
        logging.warning(f"{placeholder} NÃO encontrado.")


# ==========================================
# INSERIR GRÁFICO VINCULADO
# ==========================================
def inserir_grafico_vinculado(planilha, doc_word, config):
    logging.info(f"Inserindo gráfico: {config['nome_grafico']}")

    grafico = planilha.ChartObjects(config["nome_grafico"])
    grafico.Copy()

    def colar(range_word):
        range_word.PasteSpecial(Link=True)
        shape = doc_word.InlineShapes(doc_word.InlineShapes.Count)

    substituir_placeholder_word(doc_word, config["placeholder_word"], colar)


# ==========================================
# INSERIR TEXTO DO EXCEL (MANTÉM FORMATAÇÃO)
# ==========================================
def inserir_texto_excel(planilha, doc_word, config):
    logging.info(f"Inserindo texto da célula: {config['celula']}")

    celula = planilha.Range(config["celula"])
    valor_formatado = celula.Text  # mantém formato visual do Excel

    def escrever(range_word):
        range_word.Text = str(valor_formatado)

    substituir_placeholder_word(doc_word, config["placeholder_word"], escrever)

# ==========================================
# INSERIR MATRIZ/TABELA DO EXCEL (VERSÃO ROBUSTA)
# ==========================================
def inserir_matriz_excel(planilha, doc_word, config):
    logging.info(f"Inserindo matriz do intervalo: {config['intervalo']}")

    # 1. Selecionar o intervalo exato no Excel e copiar
    intervalo = planilha.Range(config["intervalo"])
    intervalo.Copy()

    # Dá tempo ao Windows para processar a cópia na Área de Transferência
    time.sleep(1) 

    # 2. Definir a ação de colar no Word com sistema de Tentativas (Retry)
    def colar_matriz(range_word):
        tentativas_maximas = 3
        
        for tentativa in range(tentativas_maximas):
            try:
                # Tenta colar o conteúdo (a tabela)
                range_word.Paste()
                
                # Se funcionar, limpamos a formatação extra de parágrafo e saímos do loop
                break 
                
            except Exception as e:
                logging.warning(f"Tentativa {tentativa + 1} falhou. O Word está ocupado. A aguardar...")
                time.sleep(1.5)  # Espera 1.5 segundos antes de tentar novamente
        else:
            # Se o loop terminar e não conseguir colar nenhuma vez
            logging.error(f"Falha ao colar a matriz {config['intervalo']} após {tentativas_maximas} tentativas.")

    # 3. Chamar a nossa função principal para achar o texto e colar a matriz
    substituir_placeholder_word(doc_word, config["placeholder_word"], colar_matriz)
    
# ==========================================
# EXECUTAR AUTOMAÇÃO
# ==========================================
def executar_automacao():
    caminho_excel = lbl_excel.cget("text")
    caminho_word = lbl_word.cget("text")

    if "Nenhum" in caminho_excel or "Nenhum" in caminho_word:
        messagebox.showwarning("Atenção", "Selecione os dois arquivos.")
        return

    caminho_excel = os.path.abspath(caminho_excel)
    caminho_word = os.path.abspath(caminho_word)

    logging.info("---- INICIANDO AUTOMAÇÃO EXCEL -> WORD ----")

    try:
        # with open('configuracoes_word.json', 'r', encoding='utf-8') as f:
        with open(resource_path('configuracoes_word.json'), 'r', encoding='utf-8') as f:
            configuracoes = json.load(f)
    except FileNotFoundError:
        messagebox.showerror("Erro", "Arquivo configuracoes_word.json não encontrado.")
        return
    except json.JSONDecodeError:
        messagebox.showerror("Erro", "Erro de formatação no JSON.")
        return

    excel = None
    word = None

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        word = win32com.client.Dispatch("Word.Application")

        excel.Visible = False
        word.Visible = True

        wb = excel.Workbooks.Open(caminho_excel)
        doc = word.Documents.Open(caminho_word)

        # LOOP ORGANIZADO POR ABAS
        for aba_excel, itens in configuracoes.items():
            logging.info(f"Processando aba: {aba_excel}")

            planilha = wb.Sheets(aba_excel)

            for item in itens:
                if item["tipo"] == "grafico":
                    inserir_grafico_vinculado(planilha, doc, item)

                elif item["tipo"] == "texto":
                    inserir_texto_excel(planilha, doc, item)

                # --- NOVA CONDIÇÃO ADICIONADA AQUI ---
                elif item["tipo"] == "matriz":
                    inserir_matriz_excel(planilha, doc, item)

        # SALVAR NOVA VERSÃO
        nome_base, extensao = os.path.splitext(caminho_word)
        novo_caminho = f"{nome_base}_ComGraficos{extensao}"
        doc.SaveAs(novo_caminho)

        logging.info(f"Documento salvo em: {novo_caminho}")
        messagebox.showinfo("Sucesso", f"Processamento concluído!\nSalvo em:\n{novo_caminho}")

    except Exception as e:
        logging.error(str(e))
        messagebox.showerror("Erro", str(e))


# ==========================================
# INTERFACE TKINTER
# ==========================================
def selecionar_excel():
    caminho = filedialog.askopenfilename(
        title="Selecione a Planilha Excel",
        filetypes=[("Excel", "*.xlsx *.xls *.xlsm")]
    )
    if caminho:
        lbl_excel.config(text=caminho)


def selecionar_word():
    caminho = filedialog.askopenfilename(
        title="Selecione o Documento Word",
        filetypes=[("Word", "*.docx *.docm *.doc")]
    )
    if caminho:
        lbl_word.config(text=caminho)


janela = tk.Tk()
janela.title("Automação Excel → Word (Vinculado)")
janela.geometry("520x320")
janela.eval('tk::PlaceWindow . center')

tk.Label(janela, text="Transferência Excel → Word",
         font=("Arial", 14, "bold")).pack(pady=10)

btn_excel = tk.Button(janela, text="1. Selecionar Planilha Excel",
                      command=selecionar_excel, width=35)
btn_excel.pack(pady=5)

lbl_excel = tk.Label(janela,
                     text="Nenhum arquivo Excel selecionado",
                     fg="gray")
lbl_excel.pack(pady=5)

btn_word = tk.Button(janela, text="2. Selecionar Documento Word",
                     command=selecionar_word, width=35)
btn_word.pack(pady=5)

lbl_word = tk.Label(janela,
                    text="Nenhum arquivo Word selecionado",
                    fg="gray")
lbl_word.pack(pady=5)

tk.Frame(janela, height=2, bd=1, relief=tk.SUNKEN).pack(fill=tk.X,
                                                       padx=20, pady=10)

btn_executar = tk.Button(janela,
                         text="EXECUTAR AUTOMAÇÃO",
                         command=executar_automacao,
                         bg="green", fg="white",
                         font=("Arial", 10, "bold"))
btn_executar.pack(pady=10)

janela.mainloop()