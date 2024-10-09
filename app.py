import os
from flask import Flask, render_template, request, redirect
import pandas as pd
import shutil

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return 'No file part', 400
    file = request.files['file']
    if file.filename == '':
        return 'No selected file', 400
    
    # Salvar o arquivo na pasta correta
    upload_folder = 'uploads'
    os.makedirs(upload_folder, exist_ok=True)  # Cria a pasta se não existir
    file.save(os.path.join(upload_folder, file.filename))
    return redirect('/process')

@app.route('/process', methods=['POST'])
def process():
    planilha_caminho = 'planilha.xlsx'  
    pasta_fotos_caminho = 'Fotos'  

    # Verifica se a planilha e a pasta de fotos existem
    if not os.path.exists(planilha_caminho):  
        raise FileNotFoundError(f"Planilha não encontrada: {planilha_caminho}")  

    if not os.path.exists(pasta_fotos_caminho):  
        raise FileNotFoundError(f"Pasta de fotos não encontrada: {pasta_fotos_caminho}")  

    # Lê a planilha
    df = pd.read_excel(planilha_caminho, dtype=str)  

    if 'Foto' not in df.columns or 'Equipamento' not in df.columns:  
        raise ValueError("As colunas 'Foto' e 'Equipamento' não estão na planilha")  

    def generate_next_filename(base_name, existing_files):  
        suffix = 'b'  
        new_filename = f"{base_name} {suffix}.jpg"  
        while new_filename in existing_files:  
            suffix = chr(ord(suffix) + 1)  
            new_filename = f"{base_name} {suffix}.jpg"  
        return new_filename  

    existing_files = set(os.listdir(pasta_fotos_caminho))  

    for index, row in df.iterrows():
        nome_atual = row['Foto']
        nova_numeracao = row['Equipamento']

        if pd.isna(nome_atual) or not isinstance(nome_atual, str) or not nome_atual.strip():
            print(f"Nome atual inválido ou vazio na linha {index + 1}")
            continue

        nome_atual = nome_atual.strip()
        nova_numeracao = nova_numeracao.strip()

        if '...' in nome_atual:
            try:
                start, end = map(int, nome_atual.split('...'))
            except ValueError:
                print(f"Intervalo inválido na linha {index + 1}: {nome_atual}")
                continue
            nome_atual_range = range(start, end + 1)
        else:
            try:
                nome_atual_range = [int(nome_atual)]
            except ValueError:
                print(f"Erro na linha {index + 1}: {nome_atual}")
                continue

        for num in nome_atual_range:
            num_str = str(num)
            encontrado = False
            
            for arquivo in existing_files:
                if num_str in arquivo:
                    caminho_atual = os.path.join(pasta_fotos_caminho, arquivo)
                    novo_nome = f"{nova_numeracao}.jpg"
                    novo_caminho = os.path.join(pasta_fotos_caminho, novo_nome)

                    if os.path.exists(novo_caminho):
                        novo_caminho = os.path.join(pasta_fotos_caminho, generate_next_filename(nova_numeracao, existing_files))

                    shutil.copy2(caminho_atual, novo_caminho)

                    existing_files.add(os.path.basename(novo_caminho))

                    print(f"Duplicado: {caminho_atual} -> {novo_caminho}")
                    encontrado = True
                    break
            
            if not encontrado:
                print(f"Arquivo não encontrado para o número: {num}")

    print("Renomeação concluída.")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))  # Pega a porta do ambiente, ou usa 5000 por padrão
    app.run(host='0.0.0.0', port=port)  # Executa o app na porta especificada
