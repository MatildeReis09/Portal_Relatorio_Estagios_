import streamlit as st
import pandas as pd
from docx import Document 
from io import BytesIO
from datetime import datetime

# CONFIGURAÇÃO DA PÁGINA
st.set_page_config(page_title="Gerador de Relatórios", layout="wide")

# Link gerado pelas respostas do froms (google sheets) 
URL_CSV_Respostas = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSy4HVfG-ckMI2a0LHV17mq_HWnv5qwOj4R7yhcp7w_srg604hAWBSn6SrebKKI-cqrHKCcEwDo8jjA/pub?gid=1536866325&single=true&output=csv"
# link respetivo de ligação aos pins
URL_Utilizadores = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSy4HVfG-ckMI2a0LHV17mq_HWnv5qwOj4R7yhcp7w_srg604hAWBSn6SrebKKI-cqrHKCcEwDo8jjA/pub?gid=1984584767&single=true&output=csv"

st.title("Portal de Relatórios de Estágio")
st.markdown("---")

def carregar_dados(url):
    try:
        # Lê os dados do url chamado 
        df= pd.read_csv(url)
        df.columns = df.columns.str.strip()# remover espaços nas colunas 
        return df
     
    except:
        st.error("Erro ao carregar dados")
        return None

#criar e formata o documento para donwload
def create_word (dados, nome,titulo): 
    doc= Document()
    doc.add_heading(titulo, 0)
    doc.add_paragraph(f"Estagiário: {nome}")
    doc.add_paragraph(f"Data de Emissão:{datetime.now().strftime('%d/%m/%Y')}")

    #organizado por data
    for _, linha in dados.sort_values ( by = 'Carimbo de data/hora' ).iterrows():
        doc.add_heading(f"Semana de {linha['Semana de Referência']}", level = 1)

        p= doc.add_paragraph()
        p.add_run("Tarefas: ").bold= True
        p.add_run(str(linha['Tarefas Realizadas']))

        p= doc.add_paragraph()
        p.add_run("Aprendizagens:").bold = True
        p.add_run(str(linha['Aprendizagens']))

        p = doc.add_paragraph()
        p.add_run ("Dificultades").bold= True
        p.add_run(str(linha['Dificuldades e Soluções']))

        doc.add_paragraph ("-" *20)
    
    buffer= BytesIO() ##biblioteca de memoria
    doc.save(buffer)
    buffer.seek(0)
    return buffer 


##carregar dados
df_resp = carregar_dados(URL_CSV_Respostas)
df_auth = carregar_dados(URL_Utilizadores)


##interface de streamlit
if df_resp is not None and df_auth is not None:  #as duas verdadeiras
    #sibebar parao login 
    st.sidebar.header(" Autenticação de Utilizador")

    lista_nomes = sorted(df_auth['Utilizadores'].unique().tolist())
    nome_utilizador = st.sidebar.selectbox("Estagiário", lista_nomes)

    #criar o campo para o pin no sidebar e procurar o pin de cada utilizador
    pin = st.sidebar.text_input("Introduz o teu pin", type = "password")
    
    ##ir a tabela de utilizadores e filtar a condição do nome com o pin e devolve nresposta numa lista []
    pin_correct= str(df_auth[df_auth['Utilizadores'] == nome_utilizador]['PIN'].values[0]).strip()
    #values[0]- tira o valor de dentro da lista pegando no unico resoltado encontrado

    acesso_autorizado = False 
    if pin == pin_correct: 
        st.sidebar.success(f"Bem-vindo/a, {nome_utilizador}")
        acesso_autorizado= True
        ##st.write(f"A tentar encontrar: '{nome_utilizador}'")
        ##st.write("Nomes encontrados na folha:", df_resp['Nome'].unique().tolist())

    elif pin!= "":
        st.sidebar.error("Pin Incorreto. Acesso boqueado")
    else:
        st.sidebar.info("Insira o Pin para ter acesso aos seus dados de Relatórios")
        

if acesso_autorizado:

    #st.write("DEBUG: Nome que escolheste no login:", nome_utilizador)
    #st.write("DEBUG: Nomes que existem na folha de respostas:", df_resp['Nome'].unique().tolist())

    #remoção deespaços brancos tanto no inicio como no fim dos nomes das colunas
    df_resp.columns= df_resp.columns.str.strip()
    coluna_data= 'Carimbo de data/hora' 
    if coluna_data in df_resp.columns:
            # Converter a data para o Python entender
            df_resp[coluna_data] = pd.to_datetime(df_resp[coluna_data], dayfirst=False)

            #criar abas de lado 
            tab_semanal, tab_mensal  = st.tabs (["Atividades da semana", "Resumo de Atividades Mensal"])

            with tab_semanal : 
                #temos de escolher uma semana especifica (já preenchida)
                semana_disposicao = df_resp[df_resp['Nome'].str.strip() == nome_utilizador.strip()] ['Semana de Referência'].unique() 

                if len (semana_disposicao) > 0 : 
                    ##caso haja uam pessoa e nao tenha inf o programa nao da erro 
                    semana_sel = st.selectbox("Escolha um semana :", semana_disposicao)
                    dados_semanais = df_resp[(df_resp['Nome'].str.strip() == nome_utilizador.strip()) & (df_resp['Semana de Referência'] == semana_sel)]

                    if not dados_semanais.empty: 
                        buffer_sem = create_word( dados_semanais, nome_utilizador, f"Relatório Semanal")
                        st.download_button( 
                            label = " Descarregar Esta Semana(.docx)", 
                            data = buffer_sem, 
                            file_name = f"Semana_ {nome_utilizador}_{semana_sel.replace('/','-')}.docx",
                            mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    else: 
                        st.warning("Sem dados referentes a esta semana")
                else: 
                    st.info("Ainda não existem registos para este utilizador")
                    ##quando len == 0 / vazia , sem tarefas

            with tab_mensal:
                #escolher o mes (já completo)
                mes_sel = st.slider("Selecione o Mês do Relatório", 1, 12, datetime.now().month)#só ate ao mes atual
                dados_mes = df_resp[(df_resp['Nome'].str.strip() == nome_utilizador.strip()) & (df_resp[coluna_data].dt.month == mes_sel)]

                if not dados_mes.empty:
                    st.write(f"Encontrados {len(dados_mes)} registos para o mês {mes_sel}.")
                    buffer_mes = create_word( dados_mes, nome_utilizador, f"Relatório Mensal - Mês {mes_sel}")
                    st.download_button (
                        label = " Descarregar Relatório Mensal(.docx)",
                        data = buffer_mes, 
                        file_name = f"Mensal{nome_utilizador}_Mes_{mes_sel}.docx",
                        mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )    
                else: 
                    st.warning("Ainda não existem dados para este mês.")
    else: 
        st.error(f"Coluna' {coluna_data}' não encontrada.")



