import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
from tkinter.ttk import Progressbar
from pdfminer.high_level import extract_text
import pytesseract
import fitz  # PyMuPDF
from PIL import Image
import io
import re
import json
import os
import requests
import urllib.parse
import datetime

# Configurar o caminho do executável do Tesseract (no Windows)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Variável global para armazenar os campos de entrada dinâmicos e text_box
search_entries = []
api_url_entry = None
text_box = None

# Função para registrar logs no arquivo de log
def registrar_log(mensagem):
    log_path = os.path.join(os.getcwd(), "execucao.log")
    with open(log_path, 'a') as log_file:
        data_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_file.write(f"[{data_hora}] {mensagem}\n")

# Função para fechar o aplicativo e registrar log de encerramento
def fechar_aplicativo():
    registrar_log("Aplicativo encerrado.")
    root.destroy()  # Fecha a janela principal e encerra o loop do Tkinter
    
    # Função para enviar PDF e deletar após sucesso
def enviar_pdf(file_path):
    try:
        api_url = api_url_entry.get()
        if not api_url:
            messagebox.showwarning("Aviso", "URL da API não configurada.")
            return False
        
        # Abrir o arquivo em modo binário
        with open(file_path, 'rb') as file:
            # Usar um dicionário para o arquivo
            files = {'file': (os.path.basename(file_path), file)}
            
            # Enviar arquivo
            response = requests.post(api_url, files=files)
        
        if response.status_code == 200:
            # Registrar log de envio bem-sucedido
            registrar_log(f"Arquivo '{file_path}' enviado com sucesso.")
            
            # Tentar deletar o arquivo
            try:
                os.remove(file_path)
                registrar_log(f"Arquivo '{file_path}' deletado após envio.")
                return True
            except Exception as e:
                registrar_log(f"Erro ao deletar '{file_path}': {str(e)}")
                return False
        else:
            registrar_log(f"Falha ao enviar '{file_path}'. Status: {response.status_code}")
            return False
            
    except Exception as e:
        registrar_log(f"Erro ao enviar '{file_path}': {str(e)}")
        return False
    
# Função para abrir um arquivo PDF e tentar extrair texto usando PDFMiner e Tesseract
def open_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if not file_path:
        return
    
    try:
        text = extract_text(file_path)
        
        if not text.strip():
            text = extract_text_with_tesseract_pymupdf(file_path)
        
        if text_box:
            text_box.delete(1.0, tk.END)
            text_box.insert(tk.END, text)
    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o PDF: {e}")

# Função para extrair texto usando Tesseract OCR via PyMuPDF
def extract_text_with_tesseract_pymupdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        full_text = ""
        
        for page_num in range(doc.page_count):
            page = doc.load_page(page_num)
            pix = page.get_pixmap()
            image_bytes = pix.tobytes("png")
            image = Image.open(io.BytesIO(image_bytes))
            
            text = pytesseract.image_to_string(image, lang='por')
            full_text += f"\n--- Página {page_num + 1} ---\n" + text
        
        doc.close()
        return full_text
    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar com Tesseract: {e}")
        return ""

# Função para salvar as opções de pesquisa em um arquivo JSON
def save_search_options():
    search_fields = [{"nome": nome.get(), "body": body.get(), "termo": termo.get()} for nome, body, termo in search_entries]
    options = {
        "search_fields": search_fields,
        "api_url": api_url_entry.get()
    }
    
    with open('search_options.json', 'w') as f:
        json.dump(options, f)
    
    messagebox.showinfo("Salvo", "Opções de pesquisa salvas com sucesso.")

# Função para carregar as opções de pesquisa de um arquivo JSON
def load_search_options():
    try:
        with open('search_options.json', 'r') as f:
            options = json.load(f)
        
        clear_search_fields()
        
        for field in options["search_fields"]:
            add_search_field(field["nome"], field["body"], field["termo"])
        
        api_url_entry.delete(0, tk.END)
        api_url_entry.insert(0, options["api_url"])
        
        messagebox.showinfo("Carregado", "Opções de pesquisa carregadas com sucesso.")
    
    except FileNotFoundError:
        messagebox.showwarning("Aviso", "Nenhuma configuração encontrada.")
    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao carregar opções: {e}")

# Função para realizar a pesquisa nos campos selecionados
def search_in_text():
    search_terms = [{"nome": nome.get(), "body": body.get(), "termo": termo.get()} for nome, body, termo in search_entries]
    text_content = text_box.get(1.0, tk.END)
    results = {}

    for entry in search_terms:
        terms = [term.strip() for term in entry["termo"].split(';')]  
        found = False  

        for term in terms:
            pattern = re.compile(rf"{re.escape(term)}[:\s]*(\S+)")  

            match = pattern.search(text_content)

            if match:
                results[entry["body"]] = match.group(1)  
                found = True
                break  

        if entry["nome"].lower() == "dados do tomador de serviço" and not found:
            results[entry["body"]] = "Concrejato"  
        
        elif not found:
            results[entry["body"]] = "Não encontrado"

    result_window = Toplevel(root)
    result_window.title("Resultados da Pesquisa")

    for field_name, info in results.items():
        result_label = tk.Label(result_window, text=f"{field_name}: {info}")
        result_label.pack()

# Função para adicionar um novo campo de pesquisa dinamicamente
def add_search_field(nome_default="", body_default="", termo_default=""):
    global search_entries  

    entry_frame = tk.Frame(frame2)

    nome_entry = tk.Entry(entry_frame, width=30)
    nome_entry.insert(0, nome_default)
    nome_entry.pack(side=tk.LEFT)

    body_entry = tk.Entry(entry_frame, width=30)
    body_entry.insert(0, body_default)
    body_entry.pack(side=tk.LEFT)

    termo_entry = tk.Entry(entry_frame, width=50)
    termo_entry.insert(0, termo_default)
    termo_entry.pack(side=tk.LEFT)

    remove_button = tk.Button(entry_frame, text="Remover", command=lambda: remove_search_field(entry_frame, nome_entry, body_entry, termo_entry))
    remove_button.pack(side=tk.LEFT)

    entry_frame.pack(pady=5)

    search_entries.append((nome_entry, body_entry, termo_entry))

# Função para remover um campo específico
def remove_search_field(frame, nome_entry, body_entry, termo_entry):
    search_entries.remove((nome_entry, body_entry, termo_entry))  
    frame.destroy()  

# Função para limpar todos os campos de pesquisa
def clear_search_fields():
    global search_entries  
    
    for widget in frame2.winfo_children():
        if isinstance(widget, tk.Frame):  
            widget.destroy()
    
    search_entries.clear()

# Função para validar URL
def is_valid_url(url):
    try:
        result = urllib.parse.urlparse(url)
        
        if result.scheme not in ['http', 'https']:
            return False
        
        if not result.netloc:
            return False
        
        url_regex = re.compile(
            r'^https?://'
            r'(?:(?:[A-Z0-9](?:[A-Z0-9-]{0,61}[A-Z0-9])?\.)+[A-Z]{2,6}\.?|'
            r'localhost|'
            r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})'
            r'(?::\d+)?'
            r'(?:/?|[/?]\S+)$', re.IGNORECASE)
        
        return url_regex.match(url) is not None
    
    except ValueError:
        return False

# Função para iniciar a pesquisa em todos os PDFs da pasta
def iniciar_pesquisa_em_pasta():
    folder_path = filedialog.askdirectory()
    
    if not folder_path:
        return
    
    error_log = []
    success_log = []

    try:
        with open('search_options.json', 'r') as f:
            options = json.load(f)
        
        search_terms = options["search_fields"]
        
        api_url = options.get("api_url")

        if not api_url:
            messagebox.showwarning("Aviso", "URL da API não configurada.")
            return
        
        if not is_valid_url(api_url):
            messagebox.showerror("Erro", f"URL inválida: {api_url}")
            return

        pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
        
        total_files = len(pdf_files)
        
        if total_files == 0:
            messagebox.showinfo("Aviso", "Nenhum arquivo PDF encontrado na pasta.")
            return
        
        progress_window = Toplevel(root)
        progress_window.title("Progresso do Envio")
        
        progress_var = tk.DoubleVar()
        
        progress_bar = Progressbar(progress_window, variable=progress_var, maximum=100)
        progress_bar.pack(padx=20, pady=20)

        progress_label = tk.Label(progress_window, text="Enviando arquivos...")
        progress_label.pack()

        for index, filename in enumerate(pdf_files):
            file_path = os.path.join(folder_path, filename)

            text_content = extract_text(file_path).strip()
            
            if not text_content:
                text_content = extract_text_with_tesseract_pymupdf(file_path).strip()
            
            results_body = {}
            for entry in search_terms:
                terms = [term.strip() for term in entry["termo"].split(';')]
                found = False
                for term in terms:
                    pattern = re.compile(rf"{re.escape(term)}[:\s]*(\S+)")
                    match = pattern.search(text_content)
                    if match:
                        results_body[entry["body"]] = match.group(1)
                        found = True
                        break
                if entry["nome"].lower() == "dados do tomador de serviço" and not found:
                    results_body[entry["body"]] = "Concrejato"
                elif not found:
                    results_body[entry["body"]] = "Não encontrado"

            # Converter para lista de dicionários
            results_body = [results_body]

            print("Corpo da requisição:", json.dumps(results_body, indent=4))
  

            try:
                log_dir = os.getcwd()
                log_path = os.path.join(log_dir, "execucao.log")

                # Criar arquivo de log se não existir
                if not os.path.exists(log_path):
                    with open(log_path, 'w', encoding='utf-8') as log_file:
                        log_file.write("Log de Execução Iniciado\n")

                # Registrar início da aplicação
                registrar_log("Aplicativo iniciado.")
                response = requests.post(api_url, json=results_body, timeout=10)
                
                if response.status_code == 400:
                    error_message = f"Erro 400 ao enviar {filename}: Bad Request - Verifique o formato do corpo da requisição"
                    print(error_message)
                    error_log.append(error_message)
                    continue
                
                response.raise_for_status()
                
                print(f"Enviado {filename}: Status {response.status_code}")
                success_log.append(filename)
                
            
            except requests.exceptions.ConnectionError:
                error_message = f"Erro de conexão ao enviar {filename}: Verifique a URL ou conexão de internet"
                print(error_message)
                error_log.append(error_message)
            
            except requests.exceptions.Timeout:
                error_message = f"Tempo limite excedido ao enviar {filename}: A requisição demorou muito"
                print(error_message)
                error_log.append(error_message)
            
            except requests.exceptions.RequestException as e:
                error_message = f"Erro ao enviar {filename}: {str(e)}"
                print(error_message)
                error_log.append(error_message)

            progress_percent_complete = ((index + 1) / total_files) * 100
            progress_var.set(progress_percent_complete)
            progress_window.update_idletasks()
            
        
        progress_label.config(text="Processo concluído!")
        
        if error_log or success_log:
            result_window = Toplevel(root)
            result_window.title("Resultados do Envio")
            
            if success_log:
                success_label_title = tk.Label(result_window, text="Arquivos Enviados com Sucesso:", fg="green")
                success_label_title.pack(pady=10)

                for success_msg in success_log:
                    success_label_msg = tk.Label(result_window, text=success_msg)
                    success_label_msg.pack(padx=10)

            if error_log:
                error_label_title = tk.Label(result_window, text="Erros encontrados durante o envio:", fg="red")
                error_label_title.pack(pady=10)

                for error_msg in error_log:
                    error_label_msg = tk.Label(result_window, text=error_msg)
                    error_label_msg.pack(padx=10)
                    
                    

    
    except FileNotFoundError:
        messagebox.showwarning("Aviso", "Nenhuma configuração de pesquisa encontrada.")
    
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao realizar a pesquisa: {e}")

# Tela inicial com opções "Ajustes" e "Iniciar"
def tela_inicial():
   initial_window=Toplevel(root) 
   initial_window.title("Tela Inicial")

   ajustes_button=tk.Button(initial_window,text="Ajustes ",command=abrir_ajustes) 
   ajustes_button.pack(pady=10)

   iniciar_button=tk.Button(initial_window,text="Iniciar Pesquisa ",command=iniciar_pesquisa_em_pasta) 
   iniciar_button.pack(pady=10)

# Abrir a janela de ajustes com opção de abrir PDF e exibir conteúdo OCR 
def abrir_ajustes(): 
   ajustes_window=Toplevel(root) 

   global frame2

   frame2=tk.Frame(ajustes_window)

   header_frame=tk.Frame(frame2) 
   header_frame.pack(padx=10)

   header_nome_label=tk.Label(header_frame,text="Nome do Campo",width=30) 
   header_nome_label.pack(side=tk.LEFT)

   header_body_label=tk.Label(header_frame,text="Nome no Body",width=30) 
   header_body_label.pack(side=tk.LEFT)

   header_termo_label=tk.Label(header_frame,text="Termo de Pesquisa",width=50) 
   header_termo_label.pack(side=tk.LEFT)

   header_acoes_label=tk.Label(header_frame,text="Ações",width=10) 
   header_acoes_label.pack(side=tk.LEFT)

   frame2.pack(padx=10,pady=10)

   open_pdf_button=tk.Button(frame2,text="Abrir PDF ",command=open_pdf) 
   open_pdf_button.pack(pady=5)

   add_field_button=tk.Button(frame2,text="Adicionar Campo ",command=add_search_field) 
   add_field_button.pack(pady=5)

   search_button=tk.Button(frame2,text="Pesquisar ",command=search_in_text)  
   search_button.pack(pady=5)

   save_button=tk.Button(frame2,text="Salvar Opções ",command=save_search_options) 
   save_button.pack(side=tk.LEFT)

   load_button=tk.Button(frame2,text="Carregar Opções ",command=load_search_options) 
   load_button.pack(side=tk.LEFT)

   
   global api_url_entry 
   api_url_label=tk.Label(ajustes_window,text="URL da API:",width=20) 
   api_url_label.pack(pady=(10 ,0)) 

   api_url_entry=tk.Entry(ajustes_window,width=80) 
   api_url_entry.pack(pady=(0 ,10))

   
   global text_box 
   text_box=tk.Text(ajustes_window,height=20,width=80)  
   text_box.pack(padx=10,pady=10)


# Definir a janela principal root antes do mainloop()
root=tk.Tk() 
root.withdraw() 

# Captura o evento de fechamento da janela para encerrar corretamente e gerar log
root.protocol("WM_DELETE_WINDOW", fechar_aplicativo)

tela_inicial()

root.mainloop()