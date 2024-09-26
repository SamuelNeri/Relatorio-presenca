# Processador de Lista de Presença - Documentação Técnica

## Visão Geral do Código

Este projeto consiste em uma aplicação GUI para processar listas de presença em arquivos PDF e exportá-las para Excel. O código está estruturado em uma única classe principal `AttendanceApp` e algumas funções auxiliares.

## Dependências

- tkinter: Para a interface gráfica
- ttkthemes: Para temas melhorados na interface
- PyPDF2: Para leitura de arquivos PDF
- pandas: Para manipulação de dados e exportação para Excel
- re: Para expressões regulares usadas na extração de dados

## Estrutura do Código

### Funções Auxiliares

1. `extract_text_from_pdf(file_path)`:
   - Extrai texto de um arquivo PDF.
   - Usa PyPDF2 para ler cada página do PDF.

2. `extract_attendance_data(content)`:
   - Processa o texto extraído do PDF para encontrar informações de presença.
   - Usa expressões regulares para identificar seções de alunos presentes e extrair nomes e horários.

### Classe Principal: AttendanceApp

Esta classe gerencia a interface do usuário e a lógica principal da aplicação.

#### Métodos Principais:

1. `__init__(self, master)`:
   - Inicializa a janela principal e configura o ícone da aplicação.

2. `create_widgets(self)`:
   - Cria e posiciona todos os elementos da interface (botões, labels, etc.).

3. `import_pdf(self)`:
   - Abre um diálogo para selecionar um arquivo PDF.

4. `process_pdf(self)`:
   - Processa o PDF selecionado, extraindo as informações de presença.

5. `save_excel(self)`:
   - Salva os dados processados em um arquivo Excel.

### Tratamento de Exceções

A função `exception_handler` é definida para capturar exceções não tratadas e exibi-las em uma caixa de diálogo, em vez de imprimi-las no console.

## Fluxo de Execução

1. O aplicativo é iniciado, criando a janela principal e a interface.
2. O usuário seleciona um arquivo PDF usando o botão "Importar PDF".
3. Ao clicar em "Processar PDF", o aplicativo extrai o texto do PDF e processa as informações de presença.
4. Os dados processados podem ser salvos em um arquivo Excel usando o botão "Salvar Excel".

## Personalização do Código

### Alterando o Tema da Interface

O tema da interface pode ser alterado modificando a linha:

```python
root = ThemedTk(theme="arc")
```

Substitua "arc" por outro tema disponível no ttkthemes.

### Modificando o Padrão de Extração

As expressões regulares em `extract_attendance_data` podem ser ajustadas para corresponder a diferentes formatos de lista de presença:

```python
present_pattern = re.compile(r'ALUNOS PRESENTES(.*?)(?:ALUNOS AUSENTES|$)', re.DOTALL | re.IGNORECASE)
student_pattern = re.compile(r'(\w+(?:\s+\w+)*)\s+\d+\s+(\d{2}:\d{2})')
```

### Adicionando Funcionalidades

Novas funcionalidades podem ser adicionadas criando novos métodos na classe `AttendanceApp` e adicionando os respectivos widgets na interface.

## Compilação

Para compilar o aplicativo em um executável, use PyInstaller com o seguinte comando:

```
pyinstaller --name="Processador de Lista de Presença" --onefile --windowed --noconsole --add-data "sobreposicao.png:." --icon=sobreposicao.png --clean seu_script.py
```

Certifique-se de que o arquivo `sobreposicao.png` esteja no mesmo diretório que o script ao compilar.

## Contribuindo

Ao contribuir com o código, por favor, mantenha o estilo de codificação consistente e adicione comentários para explicar lógicas complexas. Teste todas as alterações extensivamente antes de submeter um pull request.