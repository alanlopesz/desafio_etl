import requests, pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing.image import Image

#lê a tabela
tabela = pd.read_excel("descricoes.xlsx")

copia_tabela = tabela["descricoes"].copy()

# faz as requisições à IA e gera as imagens
for c, des in tabela["descricoes"].items():
    url = 'https://api.vyro.ai/v1/imagine/api/generations'

    headers = {
            'bearer': 'vk-lVHaHbDZwhyrmBZ5y%2F5wx27VQw0V8RtLtziyZhtKvXo%3D'
    }
    # Using None here allows us to treat the parameters as string
    data = {
        'model_version': (None, '1'),
        'prompt': (None, f'{des}'),
        'style_id': (None, '30'),
    }

    response = requests.post(url, headers=headers, files=data)

    if response.status_code == 200:  # if request is successful
        with open(f'image_{c}.jpg', 'wb') as f:
            f.write(response.content)
    else:
        print("Error:", response.status_code)
#adicionar imagem
tabela_open = load_workbook("descricoes.xlsx")
tabela_ativa = tabela_open.active

for c, des in copia_tabela.items():
    img = openpyxl.drawing.image.Image('image_{}.jpg'.format(c))
    img.height = 100
    img.width = 100
    tabela_ativa.add_image(img, f"B{c+2}")
tabela_open.save("tabela_final.xlsx")
