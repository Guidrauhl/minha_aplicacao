from flask import Flask, request, render_template, flash
import pandas as pd
import os
import shutil
import zipfile

app = Flask(__name__)
app.secret_key = 'supersecretkey'

UPLOAD_FOLDER = 'uploads'
FOTOS_FOLDER = 'uploads/fotos'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def processar_planilha(planilha_path, fotos_folder):
    # Carregar a planilha
    df = pd.read_excel(planilha_path, dtype=str)

    if 'Foto' not in df.columns or 'Equipamento' not in df.columns:
        raise ValueError("As colunas 'Foto' e 'Equipamento' não estão na planilha")

    existing_files = set(os.listdir(fotos_folder))

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
                    caminho_atual = os.path.join(fotos_folder, arquivo)
                    novo_nome = f"{nova_numeracao}.jpg"
                    novo_caminho = os.path.join(fotos_folder, novo_nome)

                    if os.path.exists(novo_caminho):
                        novo_caminho = os.path.join(fotos_folder, generate_next_filename(nova_numeracao, existing_files))

                    shutil.copy2(caminho_atual, novo_caminho)
                    existing_files.add(os.path.basename(novo_caminho))

                    print(f"Duplicado: {caminho_atual} -> {novo_caminho}")
                    encontrado = True
                    break

            if not encontrado:
                print(f"Arquivo não encontrado para o número: {num}")

    print("Renomeação concluída.")

def generate_next_filename(base_name, existing_files):
    suffix = 'b'
    new_filename = f"{base_name} {suffix}.jpg"
    while new_filename in existing_files:
        suffix = chr(ord(suffix) + 1)
        new_filename = f"{base_name} {suffix}.jpg"
    return new_filename

@app.route('/', methods=['GET', 'POST'])
def upload_files():
    if request.method == 'POST':
        planilha = request.files['planilha']
        fotos_zip = request.files['fotos']

        if planilha and fotos_zip:
            # Salvar a planilha e as fotos
            planilha_path = os.path.join(app.config['UPLOAD_FOLDER'], 'planilha.xlsx')
            fotos_zip_path = os.path.join(app.config['UPLOAD_FOLDER'], 'fotos.zip')

            os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
            planilha.save(planilha_path)
            fotos_zip.save(fotos_zip_path)

            # Descompactar a pasta de fotos
            with zipfile.ZipFile(fotos_zip_path, 'r') as zip_ref:
                zip_ref.extractall(FOTOS_FOLDER)

            # Processar a planilha e renomear as fotos
            try:
                processar_planilha(planilha_path, FOTOS_FOLDER)
                flash("Processamento concluído com sucesso!")
            except Exception as e:
                flash(f"Ocorreu um erro: {str(e)}")

        return render_template('upload.html')

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)