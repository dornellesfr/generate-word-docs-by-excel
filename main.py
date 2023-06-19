import pandas as pd
import json


def create_json(path_file):
    data_df = pd.read_excel(path_file)
    data_df.to_json('data.json', force_ascii=False, indent=2)


def read_json():
    with open('data.json', 'r') as file:
        data = json.load(file)

        column_empreendimento_index = list(data.keys()).index("EMPREENDIMENTO")
        column_endereco_index = list(data.keys()).index("ENDEREÇO")
        column_area_construida_index = list(data.keys()).index("ÁREA")
        column_multa_index = list(data.keys()).index("VALOR DE MULTA")
        column_valor_proposta_index = list(data.keys()).index("VALOR LTIP")
        column_item_index = list(data.keys()).index("ITEM")

        all_column = list(data.values())[column_item_index]
        total_lines = len(list(all_column.values()))

    for i in range(total_lines):
        all_columns = list(data.values())[column_empreendimento_index]
        text_empreendimento = list(all_columns.values())[i]

        all_columns = list(data.values())[column_endereco_index]
        text_endereco = list(all_columns.values())[i]

        all_columns = list(data.values())[column_area_construida_index]
        text_area = list(all_columns.values())[i]

        all_columns = list(data.values())[column_multa_index]
        text_multa = list(all_columns.values())[i]

        all_columns = list(data.values())[column_valor_proposta_index]
        text_proposta = list(all_columns.values())[i]

        replace_vars(text_empreendimento, text_endereco, text_area, text_multa,
                     text_proposta)


def replace_vars(empreendimento: str,
                 endereco: str, area_construida: str,
                 multa: str,
                 valor_da_proposta: str):
    with open('LTIP_var.txt', 'r') as file:
        content = file.read()
        content = content.replace('#VAR1#', empreendimento)

        content = content.replace('#VAR2#', endereco)

        content = content.replace('#VAR3#', area_construida)

        content = content.replace('#VAR4#', multa)

        content = content.replace('#VAR5#', valor_da_proposta)

    with open(f'output{empreendimento}.rtf', 'w', encoding='utf-8') as file:
        file.write(content)
        file.close()


while True:
    print(20*'-')
    print('[1] - Atualizar a planilha (base de dados)')
    print('[2] - Gerar os arquivos .rtf')
    print('[3] - Fechar o programa')
    input_data = int(input('Digite o que deseja fazer: '))
    print(20*'-')

    if input_data == 3:
        break

    if input_data == 1:
        path = input('Adicione o caminho do arquivo xlsx: ')
        create_json(path)

    if input_data == 2:
        read_json()
