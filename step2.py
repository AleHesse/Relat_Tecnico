import pandas as pd
import holidays
from datetime import datetime, time, timedelta
from step3 import etapa_final

def calcula_tempo():
    # df = pd.read_excel('new_reports/etapa1.xlsx')
    df = pd.read_excel('relatorios/etapa1.xlsx')

    # 2. Converter datas
    df['Data_Abertura'] = pd.to_datetime(df['Data_Abertura'], format='%d/%m/%Y %H:%M:%S', errors='coerce')
    df['Data_Inicio'] = pd.to_datetime(df['Data_Inicio'], format='%d/%m/%Y %H:%M:%S', errors='coerce')

    # 3. Definir horário útil
    hora_inicio = time(8, 0, 0)
    hora_fim = time(17, 0, 0)

    # 4. Feriados nacionais do Brasil
    anos = range(df['Data_Abertura'].dt.year.min(), df['Data_Inicio'].dt.year.max() + 1)
    feriados_br = holidays.Brazil(years=anos)

    # 5. Função para calcular tempo útil com feriados
    def calcular_tempo_util(inicio, fim):
        if pd.isna(inicio) or pd.isna(fim) or fim <= inicio:
            return timedelta(0)

        atual = inicio
        tempo_util = timedelta(0)

        while atual.date() <= fim.date():
            dia = atual.date()
            if atual.weekday() < 5 and dia not in feriados_br:
                inicio_do_dia = datetime.combine(dia, hora_inicio)
                fim_do_dia = datetime.combine(dia, hora_fim)

                periodo_inicio = max(inicio_do_dia, atual)
                periodo_fim = min(fim_do_dia, fim)

                if periodo_inicio < periodo_fim:
                    tempo_util += periodo_fim - periodo_inicio

            atual += timedelta(days=1)
            atual = atual.replace(hour=0, minute=0, second=0)

        return tempo_util

    # 6. Aplicar ao DataFrame
    df['Tempo_util'] = df.apply(lambda row: calcular_tempo_util(row['Data_Abertura'], row['Data_Inicio']), axis=1)

    # 7. Converter para horas decimais e minutos
    df['Tempo_util_horas'] = df['Tempo_util'].dt.total_seconds() / 3600
    df['Tempo_util_minutos'] = df['Tempo_util'].dt.total_seconds() / 60

    # 9. Criar coluna formatada como "X horas e Y minutos"
    def formatar_horas_minutos(td):
        if pd.isna(td):
            return ''
        total_segundos = int(td.total_seconds())
        horas = total_segundos // 3600
        minutos = (total_segundos % 3600) // 60
        return f"{horas} hora{'s' if horas != 1 else ''} e {minutos} minuto{'s' if minutos != 1 else ''}"

    df['Tempo_util_datetime'] = df['Tempo_util'].dt.total_seconds()
    df['Tempo_util_formatado'] = df['Tempo_util'].apply(formatar_horas_minutos)

    # 8. Salvar em novo arquivo
    df.to_excel('relatorios/etapa2.xlsx', index=False)

    etapa_final()

if __name__ == "__main__":

    calcula_tempo()

    



