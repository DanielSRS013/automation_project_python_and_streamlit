import streamlit as st
import pandas as pd
from pathlib import Path
import datetime as dt
from io import BytesIO
#import win32com.client as win32
#outlook = win32.Dispatch('Outlook.Application')



st.set_page_config(
     page_title='AuditFlow',
     page_icon=':mag_right:',
     layout='wide'
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Source+Code+Pro:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"]  {
    font-family: 'Source Code Pro', monospace;
}
</style>
""", unsafe_allow_html=True)

st.title("🔍 AuditFlow")
st.caption("Automated audit for delivery, transit operations, consumption exit, tissue variation")





# Dictionary that maps each branch to its supervisor
supervisor_data = {
                    'LNY BARRA': 'Julyana',
                    'LNY BH': 'CLÁUDIA',
                    'LNY BRASILIA': 'CLÁUDIA',
                    'LNY BUZIOS': 'Julyana',
                    'LNY CAMPINAS': 'CLÁUDIA',
                    'LNY CURITIBA': 'CLÁUDIA',
                    'LNY FASHION MALL': 'Julyana',
                    'LNY IPANEMA': 'Julyana',
                    'LNY GARCIA': 'Julyana',
                    'LNY GAVEA': 'Julyana',
                    'LNY GOIANIA': 'CLÁUDIA',
                    'LNY HIGIENOPOLIS': 'CLÁUDIA',
                    'LNY JARDINS': 'CLÁUDIA',
                    'LNY NITEROI': 'Julyana',
                    'LNY LEBLON': 'Julyana',
                    'LNY OFF CATARINA': 'CLÁUDIA',
                    'LNY OFF LEBLON': 'Julyana',
                    'LNY PORTO ALEGRE': 'CLÁUDIA',
                    'LNY RECIFE': 'Julyana',
                    'LNY RIO SUL': 'Julyana',
                    'LNY SALVADOR': 'Julyana',
                    'LNY SALVADOR SHOPPING': 'Julyana',
                    'LNY SAVASSI': 'CLÁUDIA',
                    'LNY LEBLON': 'Julyana',
                    'LNY TRANCOSO': 'Julyana',
                    'LNY VILLAGE MALL': 'Julyana',
                    'LNY VITORIA': 'CLÁUDIA',
                    'LNY RIBEIRAO PRETO': 'CLÁUDIA',
                    'LNY LOJA INTERNA': 'HERTA',
                    'LNY CONSERTO LOJA': 'ADRIANO/ANDRESSA',
                    'CONSERTO LOJAS': 'ADRIANO/ANDRESSA',
                    'LNY EXPORTAÇÃO': 'ANDRÉ FERNANDES',
                    'LNY DISTRIBUIDORA': 'ANDRÉ FERNANDES',
                    'DISTRIBUIDORA ECOMMERCE': 'THIAGO ALMEIDA',
                    'DIST. EXPORTACAO': 'ANDRÉ FERNANDES',
                    'DISTRIBUIDORA ATACADO': 'ANDRÉ FERNANDES',
                    'MARKETING': 'ANDRÉ FERNANDES',
                    'MOSTRUARIO ATACADO': 'ANDRÉ FERNANDES',
                    'PRONTA ENTREGA': 'ANDRÉ FERNANDES'
                    }
# Dictionary that maps each branch to its state
state_map = {
    'LNY BARRA': 'RJ',
    'LNY BUZIOS': 'RJ',
    'LNY FASHION MALL': 'RJ',
    'LNY IPANEMA': 'RJ',
    'LNY GARCIA': 'RJ',
    'LNY GAVEA': 'RJ',
    'LNY NITEROI': 'RJ',
    'LNY LEBLON': 'RJ',
    'LNY OFF LEBLON': 'RJ',
    'LNY RIO SUL': 'RJ',
    'LNY VILLAGE MALL': 'RJ',
    'LNY BH': 'MG',
    'LNY SAVASSI': 'MG',
    'LNY CAMPINAS': 'SP',
    'LNY HIGIENOPOLIS': 'SP',
    'LNY JARDINS': 'SP',
    'LNY OFF CATARINA': 'SP',
    'LNY CURITIBA': 'PR',
    'LNY PORTO ALEGRE': 'RS',
    'LNY GOIANIA': 'GO',
    'LNY BRASILIA': 'DF',
    'LNY RECIFE': 'PE',
    'LNY SALVADOR': 'BA',
    'LNY SALVADOR SHOPPING': 'BA',
    'LNY VITORIA': 'ES',
    'LNY TRANCOSO': 'BA',
    'LNY RIBEIRAO PRETO': 'SP',
    'LNY ATACADO': 'NA',
    'CONSERTO': 'NA',
    'ESTOQUE BAZAR': 'NA',
    'ESTOQUE DISPONIVEL': 'NA',
    'LNY MATRIZ': 'NA',
    'LNY EXPORTAÇÃO': 'NA',
    'LNY BOTAFOGO': 'NA',
    'LNY LOJAS RJ': 'NA',
    'LNY - DEVOLUCAO': 'NA',
    'LNY 2005': 'NA'
}

# Allowed destination branches for interstate transfe
allowed_destination = {'CONSERTO LOJAS', 'DISTRIBUIDORA ECOMMERCE',
                       'DIST. EXPORTACAO', 'LNY DISTRIBUIDORA'}

# File uploader component (accepts Excel files)
uploaded_data = st.file_uploader('Input the database', type=['xlsx', 'xls'])

if uploaded_data is not None:
    # Read uploaded Excel file into a DataFrame
    df = pd.read_excel(uploaded_data, dtype = {'Numero Reserva': str, 'Numero Nf Retorno': str, 'Numero Nf':str, 'Numero Nf Transferencia':str})
    # Function to identify which type of database was uploaded
    def identify_base(df):
        
        columns = set(df.columns.str.lower())

        # Delivery base identification
        if {'numero reserva', 'nome vendedor'} <= columns:
              return 'delivery'
        # Transit base identification
        elif {'filial origem', 'data saida'} <= columns:
              return 'transito'
        elif {'cm operacao', 'cm desc operacao'} <= columns:
             return 'saida_consumo'
        elif {'material', 'desc material'} <= columns:
             return 'variacao_tecido'
        # Unknown structure
        else:
              return 'desconhecido'
    base_type = identify_base(df)

    
    ###### DELIVERY DATABASE ######
    if base_type == 'delivery':
        
         # Function responsible for cleaning and preparing the delivery database    
        def treat_delivery_database(df):

                #Selecting the columns I want to work with
                df = df[['Filial', 'Codigo Cliente', 'Emissao', 'Numero Reserva', 'Qtde Total', 'Valor Total', 'Cliente Varejo', 'Nome Vendedor', 'Numero Nf Retorno', 'Numero Nf']]
                
                #Renaming columns Numero Nf Retorno e Numero Nf
                df.rename(columns ={'Numero Nf Retorno': 'NF Saida', 'Numero Nf': 'Nf Entrada'}, inplace = True)
                df['Emissao'] = pd.to_datetime(df['Emissao'], errors='coerce').dt.strftime('%d/%m/%Y')

                #Selecting today's date to calc the 'Dias Fora' column
                today = dt.datetime.now()

                days_outside = today-pd.to_datetime(df['Emissao'], format='%d/%m/%Y')

                
                df['Dias Fora'] = days_outside.dt.days
                

                #Codition to select only the delivers that are open using the 'Valor Total' column
                df = df[df['Valor Total']>0].copy()

                #mapping Supervisor by it's branch
                df['Supervisor'] = df['Filial'].map(supervisor_data)

                df['Status1'] = ''
                df['Status2'] = ''

                return df
        df_delivery_treated = treat_delivery_database(df)

        #Function that identify delivery total value divergency
        def identify_delivery_value_divergency(df_delivery_treated):
                df_delivery_treated.loc[df_delivery_treated['Valor Total']>5000, 'Status1'] = 'Valor Acima de R$ 5000'
                return df_delivery_treated
        df_delivery_treated = identify_delivery_value_divergency(df_delivery_treated)

        #Function that identify a entry without the tax invoice entry
        def identify_delivery_without_invoice(df_delivery_treated):
            df_delivery_treated.loc[df_delivery_treated['NF Saida'].isnull(), 'Status2'] = 'SAÍDA SEM NF SAÍDA'

            return df_delivery_treated
        df_delivery_treated = identify_delivery_without_invoice(df_delivery_treated)    
    
        
             
        #Function that user will choose filter by filial or not
        def choose_filial(df_delivery_treated):

          filiais = df_delivery_treated['Filial'].unique()

          filial_selected = st.selectbox(
               'Gostaria de Filtrar uma Filial?',
               filiais,
               index = None,
               placeholder = 'Selecione uma filial...',
          )



          if filial_selected is not None:
               
               #email = outlook.CreateItem(0)
               df_delivery_treated = df_delivery_treated[df_delivery_treated['Filial'] == filial_selected]

               #def set_email_sender(df_delivery_treated):
                    #email.To = 'danielsrs.mkd@gmail.com'
                    #email.Subject = 'Delivery'
                    #email.HTMLBody = f'{df_delivery_treated}'

               st.write('Você selecionou: ', filial_selected)

              
               st.dataframe(df_delivery_treated)
               buffer = BytesIO()
               with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
                    df_delivery_treated.to_excel(writer, index=False, sheet_name='Data')
               st.download_button(label='Download', data=buffer.getvalue(), file_name=f"delivery_{filial_selected}.xlsx")
               
               #st.button('Enviar Email para Loja', on_click = set_email_sender, args=[''])

          else:
               
               
               claudia_supervisor = df_delivery_treated[(df_delivery_treated['Supervisor']=='CLÁUDIA') & (df_delivery_treated['Dias Fora']>5)]
               julyana_supervisor = df_delivery_treated[(df_delivery_treated['Supervisor']=='Julyana') & (df_delivery_treated['Dias Fora']>5)]
            
               def identify_days_out(val):    
                    if val > 5:
                         return 'background-color: red'

               df_delivery_treated_styled = df_delivery_treated.style.applymap(identify_days_out, subset=['Dias Fora'])

               st.dataframe(df_delivery_treated)
               # Export treated data to Excel
               buffer = BytesIO()
               
               with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
                    df_delivery_treated_styled.to_excel(writer, index=False, sheet_name='Delivery')
               st.download_button(label='Download', data=buffer.getvalue(), file_name="data.xlsx")

               with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
                    claudia_supervisor.to_excel(writer, index=False, sheet_name='Delivery')
               st.download_button(label='Download Cláudia',data=buffer.getvalue(), file_name='Claudia.xlsx')

               with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format = 'DD/MM/YYYY') as writer:
                    julyana_supervisor.to_excel(writer, index=False, sheet_name='Delivery')
               st.download_button(label='Download Julyana', data=buffer.getvalue(), file_name='Julyana.xlsx')
        choose_filial(df_delivery_treated)   
    
    ###### TRANSIT DATABASE ######
    elif base_type == 'transito':
         
        # Function responsible for cleaning and preparing the transit database
         def treat_transit_database(df):
            df = df[['Data Saida', 'Filial Origem', 'Filial', 'Qtde Total', 'Romaneio Produto',
                 'Numero Nf Transferencia', 'Romaneio Nf Saida' ]]
            
            # Convert departure date to datetime
            df['Data Saida'] = pd.to_datetime(df['Data Saida'], errors='coerce').dt.strftime('%d/%m/%Y')

            today = dt.datetime.now()

            days_outside = today-pd.to_datetime(df['Data Saida'], format='%d/%m/%Y')

            df['Dias Fora'] = days_outside.dt.days

            # Map supervisor by destination branch
            df['Supervisora'] = df['Filial'].map(supervisor_data)

            return df
        
         df_transit_treated = treat_transit_database(df)

        # Identify improper interstate transfers
         def identify_transit_div(df_transit_treated):
            df_transit_treated['Estado Origem'] = df_transit_treated['Filial Origem'].map(state_map)
            df_transit_treated['Estado Destino'] = df_transit_treated['Filial'].map(state_map)
            df_transit_treated.loc[
                (df_transit_treated['Estado Origem'].notna()) &
                (df_transit_treated['Estado Destino'].notna()) &
                (df_transit_treated['Estado Origem'] != df_transit_treated['Estado Destino']) &
                (~df_transit_treated['Filial'].isin(allowed_destination)),
                'Status'] = 'TRANSFERÊNCIA INDEVIDA'
            df_transit_treated.loc[(df_transit_treated['Status'].isnull()), 'Status'] = ''
            return df_transit_treated
         
         df_transit_treated = identify_transit_div(df_transit_treated)

         # Export treated data to Excel
         buffer = BytesIO()
         with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
            df_transit_treated = df_transit_treated.drop(columns=['Estado Origem', 'Estado Destino'])
            st.dataframe(df_transit_treated)
            df_transit_treated.to_excel(writer, index=False, sheet_name='Data')

         st.download_button(label='Download', data=buffer.getvalue(), file_name="data.xlsx")
    ###### TISSUE VARIATION DATABASE ######
    elif base_type == 'variacao_tecido':
         
         def treat_tissue_variation(df):
            df = df[['Material', 'Desc Material', 'Cor Material', 'Desc Cor Material', 'Qtde Estoque', 'Fabricante', 'Ultimo Custo', 'Ultima Entrada', 'Valor Estoque', 'Classif Fiscal']]
            df['Ultima Entrada']= pd.to_datetime(df['Ultima Entrada'], errors='coerce').dt.strftime('%d/%m/%Y')
            df.insert(0, 'Status', None)

            return df
         df_tissue_variation_treated = treat_tissue_variation(df)

         buffer = BytesIO()
         with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
            st.dataframe(df_tissue_variation_treated)
            df_tissue_variation_treated.to_excel(writer, index=False, sheet_name='Data')

         st.download_button(label='Download', data=buffer.getvalue(), file_name="data.xlsx")
    ###### CONSUMPTION EXIT DATABASE ######
    elif base_type == 'saida_consumo':
         
         def treat_consume_exit(df):
            df = df[['Req Material', 'Emissao', 'Responsavel', 'Destino', 'Requisitante', 'Rateio Centro Custo', 'Desc Rateio Centro Custo', 'Conta Contabil', 'Desc Conta', 'Cm Operacao', 'Cm Desc Operacao']]
            df['Emissao'] = pd.to_datetime(df['Emissao'], errors='coerce').dt.strftime('%d/%m/%Y')
            df.insert(0, 'Status', None)

            return df
        
         df_consumption_exit_treated = treat_consume_exit(df)

         buffer = BytesIO()
         with pd.ExcelWriter(buffer, engine='openpyxl', datetime_format='DD/MM/YYYY') as writer:
            st.dataframe(df_consumption_exit_treated)
            df_consumption_exit_treated.to_excel(writer, index=False, sheet_name='Data')

         st.download_button(label='Download', data=buffer.getvalue(), file_name="data.xlsx")
    else:
         # Case when the uploaded file structure is not recognized
         st.write('Base desconhecida! Adicione a base correta, por favor.')



