#######################
# Import libraries
import pandas as pd
import numpy as np
import streamlit as st
import datetime as dt


# Set the directory path

tipo3_file = 'tipo3_30412.XLSX'
dempro_file= 'dempro_30412.XLSX'

# slyle the page
st.set_page_config(
    page_title="Análise de Reservas",
    page_icon=":bar_chart:",
    layout="wide",
    initial_sidebar_state="expanded",
)

@st.cache_data
def load_data():
    return pd.read_excel(tipo3_file), pd.read_excel(dempro_file)
    
tipo3,dempro = load_data()

# adequar tipo de dados da tipo3
tipo3['Data base'] = pd.to_datetime(tipo3['Data base'])
tipo3['Centro custo'] = tipo3['Centro custo'].astype(str)


def analise_rtp3(tipo3, centro_custo,PRAZO_DE_ARMAZENAGEM = 180):
    ## ANALISE DAS RTP3 ##

    # Premissas
    # Prazo de utilização é de 12 meses da data de emissao da DEMPRO ou 6 meses da rtp3
    print('Analise das RTP3')
    print('\nPremissa = Prazo de utilização é de 12 meses da data de emissao da DEMPRO ou 6 meses da rtp3\n')

    analise = tipo3[['Data base','Nome do usuário' ,'Material', 'Texto', 'Centro custo', 'Valor retirado','Com registro final','Item foi eliminado','Tipo de reserva','Motivo da Reserva']]

    analise['Material'] = analise['Material'].astype(str)

    # Filtrar por centro custo, Com registro final, Item foi eliminado, Nome do usuário
    analise = analise[analise['Centro custo'] == centro_custo]
    # Nem foi retirado e nem foi eliminado
    analise = analise[analise['Com registro final'].isnull()]
    analise = analise[analise['Item foi eliminado'].isnull()]
    # WF-BATCH é o usuário que faz as reservas automaticas
    analise = analise[analise['Nome do usuário'] == 'WF-BATCH']
    analise = analise[analise['Tipo de reserva'] == 3]

    analise = analise[['Data base','Nome do usuário' ,'Material', 'Texto', 'Centro custo', 'Valor retirado','Motivo da Reserva']]

    # Data estimada de consumo
    analise['Data estimada de consumo'] = analise['Data base'] + pd.to_timedelta(PRAZO_DE_ARMAZENAGEM, unit='D')

    analise['Dias de reserva vencida'] = (pd.to_datetime('today') - analise['Data estimada de consumo']).dt.days

    # ordenar por dias de reserva vencida
    analise = analise.sort_values(by='Dias de reserva vencida', ascending=False)

    # somar total de reservas
    valor_reservas = analise['Valor retirado'].sum()

    # somar o valor retirado se dias de reserva vencida > 0
    valor_reservas_vencidas = analise[analise['Dias de reserva vencida'] > 0]['Valor retirado'].sum()

    # somar o valor retirado se dias de reserva vencida > 0
    valor_reservas_nao_vencidas = analise[analise['Dias de reserva vencida'] <= 0]['Valor retirado'].sum()

    # porcentagem de reservas vencidas
    porcentagem_reservas_vencidas = (valor_reservas_vencidas / valor_reservas) * 100

    # Top 5 reservas por valor
    top_reservas = analise[analise['Dias de reserva vencida'] > 0][['Material','Texto','Dias de reserva vencida','Valor retirado']].sort_values(by='Valor retirado', ascending=False).head(5)

    table = pd.Series({
        'Valor reservado dentro do prazo': valor_reservas_nao_vencidas.round(2),
        'Valor reservado alem do prazo estimado': valor_reservas_vencidas.round(2)
    })
    
    return table, top_reservas.reset_index(drop=True)

def analise_dempro(dempro, centro_custo=30412):
    ## ANALISE DAS DEMPRO ##

    print('\nAnalise das DEMPRO\n')

    # Converter datas
    dempro['Data base'] = pd.to_datetime(dempro['Data base'], errors='coerce', format='%d/%m/%Y')
    dempro['Data da necessidade'] = pd.to_datetime(dempro['Data da necessidade'], errors='coerce', format='%d/%m/%Y')

    dempro = dempro[dempro['Centro custo'] == centro_custo]
    # filtrar Status DP = Em atendimento ou Parcialmente atendida ou Totalmente atendida
    dempro_filtrada = dempro[dempro['Status DP'].isin(['Em atendimento', 'Parcialmente atendida', 'Totalmente atendida'])]

    # filtrar por Status Item == Saldo total com RTP3
    #dempro_filtrada = dempro[dempro['Status Item'] == 'Saldo total com RTP3']

    # Qual o horizonte de planejamento da DEMPRO?
    # filtrar as demandas que com data de necessidade
    dempro_planejada = dempro_filtrada[dempro_filtrada['Data da necessidade'].notnull()]
    # contar a quantidade e tirar a média da diferença entre a data de necessidade e a data base
    dempro_planejada['Dias de planejamento'] = (dempro_planejada['Data da necessidade'] - dempro_planejada['Data base']).dt.days

    # quantidade de itens total e percentual de itens planejados
    quantidade_total_itens = dempro_filtrada.shape[0]
    per_planejado = dempro_planejada.shape[0] / dempro.shape[0] * 100
    horizonte_planejamento = dempro_planejada["Dias de planejamento"].mean()

    table_data = {
        'Quantidade total de itens': [quantidade_total_itens],
        'Percentual planejado': [per_planejado],
        'Horizonte de planejamento (dias)': [horizonte_planejamento]
    }

    table = pd.DataFrame(table_data).T
    return table


## ANALISE DE RESERVAS com STREAMLIT ##

st.sidebar.markdown('## Filtros')

centro_custo        = st.sidebar.multiselect('Selecione os centros de custo', tipo3['Centro custo'].unique())
print(centro_custo)
prazo_armazenagem = st.sidebar.number_input('Prazo de armazenagem', min_value=1, max_value=365, value=180)
print(prazo_armazenagem)

#if len(centro_custo) == 0: centro_custo = tipo3['Centro custo'].unique()

table, top_reservas = analise_rtp3(tipo3, centro_custo[0], prazo_armazenagem)

print(table)
print(top_reservas)

st.write(f'#### Resumo das reservas - {centro_custo[0]}')


st.write(table)

st.write('#### Top 5 reservas vencidas')
st.write(top_reservas)

#analise_dempro(dempro)