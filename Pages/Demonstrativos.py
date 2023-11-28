import streamlit as st
import plotly.express as px
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials



# Função para conectar ao Google Sheets e ler os dados
def ler_dados_gsheets(spreadsheet_name, worksheet_name):
    scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('gsheets.json', scope)
    client = gspread.authorize(credentials)
    
    hide_st_style = """
                <style>
                #MainMenu {visibility: hidden;}
                footer {display: none !important;}
                header {visibility: hidden;}
                </style>
                """
    st.markdown(hide_st_style, unsafe_allow_html=True)

    page_bg_img = """
    <style>
    [data-testid="stAppViewContainer"] {
    background: linear-gradient(180deg, rgba(137, 137, 137, 0.48), rgba(255, 255, 255, 1), rgba(255,255,255,1), rgba(255, 255, 255, 1), rgba(186, 209, 255, 1))
    }
    </style>
    """
            
    st.markdown(page_bg_img, unsafe_allow_html=True)

    st.image("logovetorizada.png", width=90)

    # Título da aplicação
    st.title(":blue[Gestora Contabilidade]")
        

    try:
        sheet = client.open(spreadsheet_name)
        worksheet = sheet.worksheet(worksheet_name)
        data = worksheet.get_all_values()
        header = data[0]
        rows = data[1:]
        df = pd.DataFrame(rows, columns=header)
        return df

    except gspread.WorksheetNotFound:
        st.error(f"Erro: A guia '{worksheet_name}' não foi encontrada na planilha '{spreadsheet_name}'.")
        return None
    except Exception as e:
        st.error(f"Erro ao ler os dados da guia: {e}")
        return None

# Função para criar o gráfico
def criar_grafico(df, ano_selecionado):
    df['Ano'] = pd.to_datetime(df['Data']).dt.year
    df['Mês'] = pd.to_datetime(df['Data']).dt.month_name()
    atendimentos_por_tecnico_mes = df.groupby(['Ano', 'Mês', 'Tecnico']).size().reset_index(name='Número de Atendimentos')

    atendimentos_por_tecnico_mes_filtrado = atendimentos_por_tecnico_mes[atendimentos_por_tecnico_mes['Ano'] == int(ano_selecionado)]

    fig = px.bar(atendimentos_por_tecnico_mes_filtrado, x='Mês', y='Número de Atendimentos', color='Tecnico',
                    title=f'Número de Atendimentos por Técnico (Por Mês e Ano: {ano_selecionado})',
                    labels={'Número de Atendimentos': 'Atendimentos'}, barmode='stack')

    return fig

def criar_grafico_pizza(df, ano_selecionado):
    df['Ano'] = pd.to_datetime(df['Data']).dt.year
    df['Mês'] = pd.to_datetime(df['Data']).dt.month_name()

    atendimentos_por_tecnico_mes = df[df['Ano'] == int(ano_selecionado)].groupby(['Mês', 'Tecnico']).size().reset_index(name='Número de Atendimentos')

    fig = px.pie(atendimentos_por_tecnico_mes, values='Número de Atendimentos', names='Tecnico',
                 title=f'Totais de Atendimentos por Técnico (Ano: {ano_selecionado})',
                 labels={'Número de Atendimentos': 'Atendimentos'}, hole=0.3)

    return fig

# Função para criar o informativo de totais
def criar_informativo_totais(df, ano_selecionado):
    totais_por_ano_mes = df[df['Ano'] == int(ano_selecionado)].groupby('Mês').size().reset_index(name='Total de Tickets')

    return totais_por_ano_mes

# Layout da aplicação Streamlit
st.set_page_config(page_title="Gestora Contabilidade - Demonstrativos", page_icon="logovetorizada.png", layout="wide")

# Nome da planilha e da guia
spreadsheet_name = "RQ TI 07 - Gestão de Chamados Internos"
worksheet_name = "Apontamentos"

# Ler os dados
df = ler_dados_gsheets(spreadsheet_name, worksheet_name)

if df is not None:
    # Atendimentos por Técnico e Mês
    st.header("Demonstrativos dos Chamados Internos da Gestora")
    
    st.write("___")

    # Filtra por ano com um st.selectbox
    anos_disponiveis = sorted(df['Ano'].unique())
    ano_selecionado = st.selectbox('Selecione o ano:', anos_disponiveis, index=len(anos_disponiveis)-1)

    # Cria o gráfico de barras
    fig = criar_grafico(df, ano_selecionado)

    fig_pizza = criar_grafico_pizza(df, ano_selecionado)

    # Exibe os gráficos lado a lado
    col1, col2 = st.columns(2)
    col1.plotly_chart(fig)
    col2.plotly_chart(fig_pizza)

    # Cria o informativo de totais
    st.header("Totais dos Tickets Atendidos por Mês")
    totais_por_ano_mes = criar_informativo_totais(df, ano_selecionado)
    st.table(totais_por_ano_mes)
    
st.write("___")