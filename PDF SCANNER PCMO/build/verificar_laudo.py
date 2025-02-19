import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import pandas as pd  # Pandas
import os

# 📌 Definição do checklist
checklist = {
    "Documentação e Estruturação do PCMSO": [
        "PCMSO atualizado", "médico do trabalho", "PPRA", "PGR", "riscos ocupacionais"
    ],
    "Exames Médicos Obrigatórios": [
        "Admissional", "Periódico", "Mudança de função", "Retorno ao trabalho", "Demissional"
    ],
    "Exames Complementares": [
        "Hemograma", "Audiometria", "Espirometria", "Radiografia de tórax", "Exames toxicológicos"
    ],
    "Registros e Gestão de Dados": [
        "ASO", "Relatório anual do PCMSO", "afastamentos", "vacinação"
    ],
    "Medidas Preventivas e Ações Corretivas": [
        "Treinamentos", "promoção da saúde", "exames especializados", "absenteísmo"
    ],
    "Integração com Outras Normas e Programas": [
        "PGR", "eSocial", "Notificação ao INSS"
    ]
}

def extrair_texto(pdf_path):
    """Lê um PDF e retorna o texto extraído."""
    texto_completo = ""
    doc = fitz.open(pdf_path)
    for pagina in doc:
        texto_completo += pagina.get_text("text") + "\n"
    return texto_completo.lower()  # Converte para minúsculas para facilitar a busca

def verificar_checklist(texto):
    """Verifica quais itens do checklist estão presentes e quais estão faltando."""
    resultado = []
    for categoria, itens in checklist.items():
        for item in itens:
            if item.lower() in texto:
                resultado.append([categoria, item, "Presente"])
            else:
                resultado.append([categoria, item, "Faltando"])
    return resultado

def salvar_relatorio(resultado, pdf_path):
    """Salva o relatório de análise em um arquivo Excel com o mesmo nome do PDF."""
    nome_arquivo = os.path.splitext(os.path.basename(pdf_path))[0] + "_relatorio.xlsx"
    df = pd.DataFrame(resultado, columns=["Categoria", "Item", "Status"])
    df.to_excel(nome_arquivo, index=False)
    return nome_arquivo

def selecionar_pdf():
    """Abre o dialogo para selecionar um arquivo PDF"""
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos PDF", "*.pdf")])
    if caminho_arquivo:
        entrada_pdf.delete(0, tk.END)
        entrada_pdf.insert(0, caminho_arquivo)
        rodar_analise(caminho_arquivo)

def rodar_analise(pdf_path):
    """Roda a análise do PDF e exibe o botão para baixar o arquivo Excel"""
    try:
        texto_extraido = extrair_texto(pdf_path)
        resultado = verificar_checklist(texto_extraido)
        nome_arquivo = salvar_relatorio(resultado, pdf_path)
        # Atualiza o botão para permitir o download do arquivo
        botao_baixar["state"] = "normal"
        botao_baixar.config(command=lambda: baixar_relatorio(nome_arquivo))
        messagebox.showinfo("Sucesso", f"Análise concluída. Relatório gerado: {nome_arquivo}")
    except Exception as e:
        messagebox.showerror("Erro", f"❌ Erro ao processar o PDF: {e}")

def baixar_relatorio(nome_arquivo):
    """Permite ao usuário baixar o relatório gerado."""
    try:
        # Oferece para salvar o arquivo Excel
        caminho_arquivo = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")], initialfile=nome_arquivo)
        if caminho_arquivo:
            os.rename(nome_arquivo, caminho_arquivo)  # Renomeia o arquivo gerado para o caminho escolhido
            messagebox.showinfo("Sucesso", f"✅ Relatório salvo como: {caminho_arquivo}")
    except Exception as e:
        messagebox.showerror("Erro", f"❌ Erro ao salvar o relatório: {e}")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Verificador de Laudo PCMSO")

# Layout da interface
frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Label e campo de entrada para o caminho do arquivo PDF
label_pdf = tk.Label(frame, text="Selecione o arquivo PDF:")
label_pdf.grid(row=0, column=0, padx=5, pady=5)

entrada_pdf = tk.Entry(frame, width=50)
entrada_pdf.grid(row=0, column=1, padx=5, pady=5)

botao_selecionar = tk.Button(frame, text="Selecionar PDF", command=selecionar_pdf)
botao_selecionar.grid(row=0, column=2, padx=5, pady=5)

# Botão para baixar o relatório
botao_baixar = tk.Button(frame, text="Baixar Relatório", state="disabled")
botao_baixar.grid(row=1, column=0, columnspan=3, pady=10)

# Rodando a interface
root.mainloop()
