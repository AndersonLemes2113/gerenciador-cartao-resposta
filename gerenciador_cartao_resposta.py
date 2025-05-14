# -*- coding: utf-8 -*-
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch, cm
from reportlab.graphics.barcode import qr
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.pdfgen import canvas
import pandas as pd
import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from PIL import Image
from pdf2image import convert_from_path
import pytesseract
import platform
import re
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime
from pyzbar.pyzbar import decode
import unicodedata
# QR Code - deve ser desenhado diretamente no canvas final
from reportlab.graphics import renderPDF
import sys
import os
sys.path.insert(0, os.path.abspath("pyzbar"))

# Teste alteração git

# Caminho absoluto do Tesseract dentro do seu projeto
tesseract_path = os.path.join(os.getcwd(), "tesseract", "tesseract.exe")
pytesseract.pytesseract.tesseract_cmd = tesseract_path

# Caminho do arquivo JSON com os POPs e gabaritos
ARQUIVO_POP = "pops_gabaritos.json"
ARQUIVO_DADOS = "colaboradores.xlsx"
CAMINHO_ULTIMA_PASTA = "ultima_pasta.json"

def carregar_pops():
    if os.path.exists(ARQUIVO_POP):
        with open(ARQUIVO_POP, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}

def salvar_pops(pops):
    with open(ARQUIVO_POP, 'w', encoding='utf-8') as f:
        json.dump(pops, f, ensure_ascii=False, indent=4)

def obter_ultima_pasta():
    if os.path.exists(CAMINHO_ULTIMA_PASTA):
        with open(CAMINHO_ULTIMA_PASTA, 'r', encoding='utf-8') as f:
            dados = json.load(f)
            return dados.get("ultima_pasta", None)
    return None

def salvar_ultima_pasta(pasta):
    with open(CAMINHO_ULTIMA_PASTA, 'w', encoding='utf-8') as f:
        json.dump({"ultima_pasta": pasta}, f, ensure_ascii=False, indent=4)

def normalizar_nome(nome):
    nome_sem_acento = ''.join(
        c for c in unicodedata.normalize('NFD', nome)
        if unicodedata.category(c) != 'Mn'
    )
    return nome_sem_acento.replace(' ', '_')

def gerar_cartao_resposta():
    matricula = simpledialog.askstring("Matrícula", "Digite a matrícula do colaborador:")
    if not matricula:
        return
    if not os.path.exists(ARQUIVO_DADOS):
        messagebox.showerror("Erro", "Arquivo de dados 'colaboradores.xlsx' não encontrado.")
        return
    try:
        df = pd.read_excel(ARQUIVO_DADOS)
        colaborador = df[df['Matrícula'].astype(str) == matricula]
        if colaborador.empty:
            messagebox.showerror("Erro", f"Matrícula {matricula} não encontrada no arquivo de dados.")
            return
        nome = colaborador.iloc[0]['Nome']
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo de dados: {e}")
        return

    pops = carregar_pops()
    if not pops:
        messagebox.showinfo("Informação", "Nenhum POP cadastrado.")
        return

    janela_pop = tk.Toplevel()
    janela_pop.title("Selecione o POP")
    janela_pop.geometry("600x400")
    janela_pop.minsize(400, 300)
    janela_pop.maxsize(1000, 800)

    tk.Label(janela_pop, text="Selecione o POP para o cartão resposta:").pack(padx=10, pady=5)

    pop_var = tk.StringVar(janela_pop)
    pop_menu = ttk.Combobox(janela_pop, textvariable=pop_var, values=list(pops.keys()), state="readonly")
    pop_menu.pack(padx=10, pady=5, fill='both')

    def gerar():
        pop_escolhido = pop_var.get()
        print("POP selecionado:", pop_escolhido)
        if not pop_escolhido:
            messagebox.showwarning("Aviso", "Selecione um POP.")
            return

        hoje = datetime.now().strftime("%Y-%m-%d")
        nome_formatado = normalizar_nome(nome)
        nome_arquivo_base = f"{matricula}_{nome_formatado}_POP{pop_escolhido}_{hoje}.pdf"

        pasta_destino = obter_ultima_pasta()
        if not pasta_destino:
            pasta_destino = filedialog.askdirectory(title="Selecione onde salvar o cartão de resposta")
            if not pasta_destino:
                return
            salvar_ultima_pasta(pasta_destino)

        nome_arquivo = os.path.join(pasta_destino, nome_arquivo_base)

        c = canvas.Canvas(nome_arquivo, pagesize=letter)

        c.setFont("Helvetica-Bold", 12)
        c.drawString(72, 720, f"Matrícula: {matricula}")
        c.drawString(72, 700, f"Nome: {nome}")
        c.drawString(72, 680, f"Avaliação: {pop_escolhido}")
        c.drawString(72, 660, "Status: Não corrigido")

        # QR Code
        dados_qr = {
        "matricula": matricula,
        "nome": nome,
        "pop": pop_escolhido  # Aqui está o valor correto!
        }
        data_qr = json.dumps(dados_qr, ensure_ascii=False)
        qr_code = qr.QrCodeWidget(data_qr)
        bounds = qr_code.getBounds()
        width = bounds[2] - bounds[0]
        height = bounds[3] - bounds[1]
        d = Drawing(100, 100, transform=[100./width, 0, 0, 100./height, 0, 0])
        d.add(qr_code)
        renderPDF.draw(d, c, 450, 630)

        colunas = ['Pergunta', 'A', 'B', 'C', 'D', 'E']
        start_y = 420
        spacing_y = 28.35
        spacing_x = 56.7
        inicio_x = 150

        c.setFont("Helvetica", 12)
        for i in range(10):
            y = start_y - i * spacing_y
            c.drawString(inicio_x, y + 5, str(i + 1))
            for j in range(1, len(colunas)):
                x = inicio_x + j * spacing_x
                c.rect(x + 8, y + 4, 10, 10)

        c.drawString(72, 100, "Assinatura do Colaborador: ____________________________________")
        c.setFillColor(colors.red)
        c.drawString(72, 80, "Resultado: 0 de 10 (0.0%)")

        c.save()
        messagebox.showinfo("Sucesso", f"Cartão gerado: {nome_arquivo}")
        janela_pop.destroy()

    tk.Button(janela_pop, text="Gerar Cartão", command=gerar, width=25, height=2).pack(pady=10, fill='both')


def corrigir_cartoes_resposta():
    try:
        arquivo_pdf = filedialog.askopenfilename(title="Selecione o cartão-resposta preenchido", filetypes=[("PDF files", "*.pdf")])
        if not arquivo_pdf:
            return

        imagens = convert_from_path(arquivo_pdf)
        qr_data = decode(imagens[0])

        if not qr_data:
            raise ValueError("QR Code não encontrado.")
        conteudo_qr = qr_data[0].data.decode('utf-8')

        try:
            dados = json.loads(conteudo_qr)
        except json.JSONDecodeError:
            raise ValueError(f"QR Code inválido ou não é JSON: {conteudo_qr}")

        matricula = dados.get("matricula")
        nome = dados.get("nome")
        pop = dados.get("pop")

        pops = carregar_pops()
        if pop not in pops:
            raise ValueError("Gabarito do POP não encontrado.")

        raise ValueError("QR Code não encontrado.")

        dados = json.loads(qr_data[0].data.decode('utf-8'))
        matricula = dados.get('matricula')
        nome = dados.get('nome')
        pop = dados.get('pop')

        if not (matricula and nome and pop):
            raise ValueError("QR Code inválido.")

        matricula, nome, pop = dados
        pops = carregar_pops()
        pop = pop.strip()
        print(f"POP extraído: '{pop}'")
        print(f"POPs disponíveis: {list(pops.keys())}")

        if pop not in pops:
            raise ValueError("Gabarito do POP não encontrado.")

        gabarito = pops[pop]
        texto_extraido = pytesseract.image_to_string(imagens[0], lang='por')

        respostas = []
        for i in range(1, 11):
            match = re.search(fr"{i}\.\s*([ABCDE])", texto_extraido)
            respostas.append(match.group(1) if match else "-")

        acertos = sum([1 for r, g in zip(respostas, gabarito) if r == g])
        percentual = (acertos / len(gabarito)) * 100

        novo_nome = os.path.splitext(os.path.basename(arquivo_pdf))[0] + "-CORRIGIDO.pdf"
        pasta = carregar_pasta_salvamento()
        if not pasta:
            pasta = filedialog.askdirectory(title="Escolha onde salvar o PDF corrigido")
            if not pasta:
                return
            salvar_pasta_salvamento(pasta)

        caminho_corrigido = os.path.join(pasta, novo_nome)
        leitor = PdfReader(arquivo_pdf)
        escritor = PdfWriter()

        pagina = leitor.pages[0]
        escritor.add_page(pagina)
        with open("temp.pdf", "wb") as temp:
            escritor.write(temp)

        c = canvas.Canvas(caminho_corrigido, pagesize=letter)
        c.drawString(50, 200, f"Respostas: {' '.join(respostas)}")
        c.drawString(50, 185, f"Gabarito:  {' '.join(gabarito)}")
        c.setFillColor(colors.red)
        c.drawString(50, 170, f"Acertos: {acertos} de {len(gabarito)}  -  {percentual:.1f}%")
        c.setFillColor(colors.black)
        c.showPage()
        c.save()

        with open("temp.pdf", "rb") as temp:
            leitor_corrigido = PdfReader(temp)
            escritor_final = PdfWriter()
            for p in leitor_corrigido.pages:
                escritor_final.add_page(p)

            c_corrigido = PdfReader(caminho_corrigido)
            for p in c_corrigido.pages:
                escritor_final.add_page(p)

            with open(caminho_corrigido, "wb") as f:
                escritor_final.write(f)

        os.remove("temp.pdf")

        messagebox.showinfo("Sucesso", f"Cartão corrigido salvo como: {caminho_corrigido}")

    except Exception as e:
        messagebox.showerror("Erro", str(e))

def gerenciar_pops():
    pops = carregar_pops()

    def salvar_alteracoes():
        try:
            novo_pop = entrada_pop.get()
            if not novo_pop:
                raise ValueError("O nome do POP não pode estar vazio.")
            respostas = [entrada.get().strip().upper() for entrada in entradas_respostas]
            if any(r not in 'ABCDE' for r in respostas):
                raise ValueError("As respostas devem estar entre A e E.")
            pops[novo_pop] = respostas
            salvar_pops(pops)
            messagebox.showinfo("Sucesso", f"Gabarito para POP {novo_pop} salvo com sucesso.")
            janela.destroy()
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    janela = tk.Toplevel()
    janela.title("Gerenciar Gabaritos POP")
    janela.geometry("400x400")

    tk.Label(janela, text="Nome do POP:").pack(pady=5)
    entrada_pop = tk.Entry(janela)
    entrada_pop.pack(pady=5)

    entradas_respostas = []
    for i in range(10):
        frame = tk.Frame(janela)
        frame.pack(pady=2)
        tk.Label(frame, text=f"Pergunta {i+1}:").pack(side=tk.LEFT)
        entrada = tk.Entry(frame, width=2)
        entrada.pack(side=tk.LEFT)
        entradas_respostas.append(entrada)

    tk.Button(janela, text="Salvar Gabarito", command=salvar_alteracoes).pack(pady=10)

def main():
    root = tk.Tk()
    root.title("Sistema de Cartão-Resposta")
    root.geometry("300x250")

    tk.Button(root, text="Gerar Cartão Resposta", command=gerar_cartao_resposta, height=2).pack(fill='x', padx=10, pady=10)
    tk.Button(root, text="Corrigir Cartão Resposta", command=corrigir_cartoes_resposta, height=2).pack(fill='x', padx=10, pady=10)
    tk.Button(root, text="Gerenciar POPs", command=gerenciar_pops, height=2).pack(fill='x', padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
