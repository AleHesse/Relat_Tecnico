import pandas as pd

def etapa_final():

    # 1. Carregar o arquivo original
    df = pd.read_excel('relatorios/etapa2.xlsx')

    # 2. Selecionar as colunas que você quer manter
    # colunas_para_salvar = ['Tecnico', 'Ocorrencia','Etiqueta', 'Data_Abertura', 'Data_Inicio', 'Abertura_Datetime', 'Inicio_Datetime', 'Tempo_util_datetime', 'Motivo', 'Tempo_util_formatado']
    colunas_para_salvar = ['Tecnico', 'Ocorrencia','Etiqueta', 'Motivo', 'Data_Abertura', 'Data_Inicio', 'Tempo_util_datetime', 'Tempo_util_formatado']

    
    df['Abertura_Datetime'] = df['Data_Abertura']
    df['Inicio_Datetime'] = df['Data_Inicio']

    # # Converter de string para data time
    # df['Tempo_util_datetime'] = pd.to_datetime(df['Tempo_util_formatado'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

    # 3. Converter datas para string no formato desejado
    df['Data_Abertura'] = df['Data_Abertura'].dt.strftime('%d/%m/%Y %H:%M:%S')
    df['Data_Inicio'] = df['Data_Inicio'].dt.strftime('%d/%m/%Y %H:%M:%S')
    

    # 3. Salvar essas colunas em um novo arquivo Excel
    df[colunas_para_salvar].to_excel('relatorios/etapa3_final.xlsx', index=False)

    print("Finalizando!!!")

