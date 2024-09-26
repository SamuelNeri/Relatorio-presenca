import tkinter as tk
from tkinter import filedialog, messagebox
from ttkthemes import ThemedTk
from tkinter import ttk
import re
import pandas as pd
from PyPDF2 import PdfReader # Biblioteca para leitura do arquivo PDF
import os
import sys

def extract_text_from_pdf(file_path):
    reader = PdfReader(file_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() + "\n"
    return text
# Campo de leitura e aramzenamento do PDF
def extract_attendance_data(content):
    present_pattern = re.compile(r'ALUNOS PRESENTES(.*?)(?:ALUNOS AUSENTES|$)', re.DOTALL | re.IGNORECASE) 
    present_sections = present_pattern.findall(content)

    data = []
    for section in present_sections:
        student_pattern = re.compile(r'(\w+(?:\s+\w+)*)\s+\d+\s+(\d{2}:\d{2})')
        students = student_pattern.findall(section)
        for student in students:
            data.append({'Aluno': student[0], 'Presença': student[1]})
    
    return data
# Carrega a aplicação
class AttendanceApp:
    def __init__(self, master):
        self.master = master
        self.master.title("Processador de Lista de Presença")
        self.master.geometry("400x300")
        
        # Configurar o ícone
        if getattr(sys, 'frozen', False):
            # Se estiver rodando como executável
            application_path = sys._MEIPASS
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        icon_path = os.path.join(application_path, "sobreposicao.png")
        
        if os.path.exists(icon_path):
            self.master.iconphoto(True, tk.PhotoImage(file=icon_path))
        
        self.create_widgets()
        
        self.pdf_file = None
        self.df = None

    def create_widgets(self):
        # Frame principal
        main_frame = ttk.Frame(self.master, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Botão para importar PDF
        self.import_button = ttk.Button(main_frame, text="Importar PDF", command=self.import_pdf)
        self.import_button.pack(pady=10)

        # Label para mostrar o arquivo selecionado
        self.file_label = ttk.Label(main_frame, text="Nenhum arquivo selecionado")
        self.file_label.pack(pady=5)

        # Botão para processar PDF
        self.process_button = ttk.Button(main_frame, text="Processar PDF", command=self.process_pdf, state=tk.DISABLED)
        self.process_button.pack(pady=10)

        # Botão para salvar Excel
        self.save_button = ttk.Button(main_frame, text="Salvar Excel", command=self.save_excel, state=tk.DISABLED)
        self.save_button.pack(pady=10)

        # Label para mostrar o status
        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(pady=5)

    def import_pdf(self):
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_file = file_path
            self.file_label.config(text=f"Arquivo selecionado: {os.path.basename(file_path)}")
            self.process_button.config(state=tk.NORMAL)
            self.status_label.config(text="")

    def process_pdf(self):
        if self.pdf_file:
            try:
                content = extract_text_from_pdf(self.pdf_file)
                attendance_data = extract_attendance_data(content)
                self.df = pd.DataFrame(attendance_data)
                self.status_label.config(text=f"Processado com sucesso. {len(attendance_data)} registros encontrados.")
                self.save_button.config(state=tk.NORMAL)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao processar o PDF: {str(e)}")
        else:
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo PDF primeiro.")

    def save_excel(self):
        if self.df is not None:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if file_path:
                try:
                    self.df.to_excel(file_path, index=False)
                    messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em {file_path}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao salvar o arquivo: {str(e)}")
        else:
            messagebox.showwarning("Aviso", "Nenhum dado para salvar. Por favor, processe um PDF primeiro.")

def exception_handler(exception_type, exception, traceback):
    # Aqui você pode personalizar como deseja lidar com exceções não tratadas
    error_message = f"{exception_type.__name__}: {exception}"
    messagebox.showerror("Erro Inesperado", error_message)

if __name__ == "__main__":
    # Configurar o tratamento de exceções global
    sys.excepthook = exception_handler

    root = ThemedTk(theme="arc") # Tema
    app = AttendanceApp(root)
    root.mainloop()