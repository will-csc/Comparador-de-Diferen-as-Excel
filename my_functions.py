from tkinter import filedialog
import pandas as pd
from tkinter import messagebox

def select_file(entry):
    file_path = filedialog.askopenfilename(
        title="Selecione um arquivo Excel",
        filetypes=[("Arquivos Excel", "*.xlsx;*.xls")]
    )
    if file_path:
        entry.delete(0, 'end')
        entry.insert(0, file_path)

def comparar_planilhas_por_linha(file1_path, file2_path):
    try:
        df1 = pd.read_excel(file1_path)
        df2 = pd.read_excel(file2_path)
        
        diffs = []
        max_len = max(len(df1), len(df2))
        
        for i in range(max_len):
            if i >= len(df1):
                diffs.append(("", df2.iloc[i].to_list()))  # Linha apenas no segundo arquivo
            elif i >= len(df2):
                diffs.append((df1.iloc[i].to_list(), ""))  # Linha apenas no primeiro arquivo
            else:
                diff_row1 = df1.iloc[i].to_list()
                diff_row2 = df2.iloc[i].to_list()
                if diff_row1 != diff_row2:
                    diffs.append((diff_row1, diff_row2))  # Diferença entre as duas linhas
        
        return diffs
    except Exception as e:
        return f"Erro ao comparar as planilhas: {e}"

def exibir_diferencas_por_linha(diffs, content_rows, columns_rows, rows_rows):
    content_rows.delete(0, 'end')
    columns_rows.delete(0, 'end')
    rows_rows.delete(0, 'end')
    
    if not diffs:
        messagebox.showinfo("Resultado", "As planilhas são idênticas.")
        return

    # Preenche as entradas com as diferenças
    for i, (row1, row2) in enumerate(diffs):
        if row1:  # Preenche conteúdo do arquivo 1
            if i > 5:
                break
            content_rows.insert(i, ', '.join(map(str, row1)))
        if row2:  # Preenche conteúdo do arquivo 2
            if i > 5:
                break
            columns_rows.insert(i, ', '.join(map(str, row2)))
        rows_rows.insert(i, f"Linha {i+1}")  # Exibe a linha em que ocorreu a diferença

def comparar(file1,file2,content_rows,columns_rows,rows_rows):
    file1_path = file1.get()
    file2_path = file2.get()
    
    if file1_path == "..." or file2_path == "...":
        messagebox.showwarning("Aviso", "Por favor, selecione ambos os arquivos.")
        return
    
    diffs = comparar_planilhas_por_linha(file1_path, file2_path)
    exibir_diferencas_por_linha(diffs, content_rows, columns_rows, rows_rows)