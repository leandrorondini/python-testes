from PIL import Image, ImageDraw, ImageFont
import pandas as pd
import os
from datetime import datetime

# Caminho absoluto para o arquivo Excel
caminho_excel = r'C:\Users\leand\OneDrive\Área de Trabalho\automação_certificados\lean\lista_aliunos_certif.xlsx'

# Ler a planilha Excel
df = pd.read_excel(caminho_excel)
print("Colunas na planilha:", df.columns.tolist())
print("Primeiras linhas da planilha:")
print(df.head())

# Definir as fontes
font_nome = ImageFont.truetype(r'C:\Users\leand\OneDrive\Área de Trabalho\automação_certificados\lean\tahoma.ttf', 30)
font_curso = ImageFont.truetype(r'C:\Users\leand\OneDrive\Área de Trabalho\automação_certificados\lean\tahoma.ttf', 30)
font_data = ImageFont.truetype(r'C:\Users\leand\OneDrive\Área de Trabalho\automação_certificados\lean\tahoma.ttf', 25)
font_tipo_participacao = ImageFont.truetype(r'C:\Users\leand\OneDrive\Área de Trabalho\automação_certificados\lean\tahoma.ttf', 25)

# Criar pasta para os certificados se não existir
pasta_certificados = "certificados_gerados"
if not os.path.exists(pasta_certificados):
    os.makedirs(pasta_certificados)

# Função para formatar data
def formatar_data(data):
    try:
        if pd.isna(data):
            return "Data não informada"
        
        # Se for string, tentar converter
        if isinstance(data, str):
            data = pd.to_datetime(data)
        
        # Formatar como dd/mm/aaaa
        return data.strftime('%d/%m/%Y')
    except:
        return "Data inválida"

# Gerar certificados para cada participante
for index, row in df.iterrows():
    try:
        # Extrair dados da linha usando os nomes corretos das colunas
        nome_participante = str(row['Nome do Participante'])
        nome_do_curso = str(row['Nome do Curso'])
        data_emissao = formatar_data(row['Data Emissão do Certificado'])
        tipo_participacao = str(row['Tipo de Participação'])
        
        # Abrir a imagem do certificado
        image = Image.open(r'C:\Users\leand\OneDrive\Área de Trabalho\automação_certificados\lean\certificado.jpg')
        desenhar = ImageDraw.Draw(image)
        
        # Adicionar texto ao certificado
        desenhar.text((135, 210), str(nome_participante), fill='black', font=font_nome)
        desenhar.text((135, 300), str(nome_do_curso), fill='black', font=font_curso)
        desenhar.text((238, 385), str(data_emissao), fill='black', font=font_data)
        desenhar.text((159, 451), str(tipo_participacao), fill='black', font=font_tipo_participacao)
        
        # Salvar o certificado
        nome_arquivo = f"{pasta_certificados}/certificado_{nome_participante.replace(' ', '_').replace('/', '_')}.jpg"
        image.save(nome_arquivo)
        print(f"Certificado gerado para: {nome_participante} - Data: {data_emissao}")
        
    except Exception as e:
        print(f"Erro ao gerar certificado para linha {index + 1}: {e}")

print(f"\nProcesso concluído! Certificados salvos na pasta '{pasta_certificados}'")