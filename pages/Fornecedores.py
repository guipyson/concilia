import streamlit as st
import pandas as pd
from io import BytesIO

def check_sum_difference_between_dfs(df1, col1, df2, col2):
    sum_col1 = df1[col1].sum()
    sum_col2 = df2[col2].sum()
    difference = abs(sum_col1 - sum_col2)
    return difference < 1, difference

def round_small_values(value):
    return 0.0 if -0.01 <= value <= 0.01 else value


st.set_page_config(
    page_title = "Conciliação Fornecedores"
)

uploaded_file = st.file_uploader("Selecione um arquivo")

if uploaded_file is not None:
    
    razao = pd.read_excel(uploaded_file, sheet_name=0)
    posicao = pd.read_excel(uploaded_file, sheet_name=1)
    
    razao = razao.dropna(subset=['Código Empresa'])
    posicao = posicao.dropna(subset=['Empresa'])
    
    #Código-Cliente por aba
    razao_dict = dict(zip(razao["Código Empresa"], razao["Descrição Empresa"]))
    posicao_dict = dict(zip(posicao["Empresa"], posicao["Nome Completo"]))
    
    #Junção do Código-Cliente
    codigo_empresa = {}
    for d in [razao_dict, posicao_dict]:
        for key, value in d.items():
            if key not in codigo_empresa:
                codigo_empresa[key] = value

    #Débito-Crédito por cliente
    dc = razao.groupby('Código Empresa')[['Débito', 'Crédito']].sum().reset_index()
    #Posição por cliente
    posicao_grouped = posicao.groupby('Empresa')['Saldo'].sum().reset_index()
    #Merge dc e posicao
    dc_posicao = pd.merge(posicao_grouped, dc, left_on='Empresa', right_on='Código Empresa', how='outer')
    # Preencher os clientes que estão faltantes em alguma das abas (i.e. c/ movimentação sem saldo)
    dc_posicao.loc[dc_posicao['Empresa'].isna(), 'Empresa'] = dc_posicao['Código Empresa']
  
    #Merge PDD com dc_posicao
    dc_posicao.drop(columns=['Código Empresa'], inplace=True)
    dc_posicao[['Saldo', 'Débito', 'Crédito']] = dc_posicao[['Saldo', 'Débito', 'Crédito']].fillna(0.0)
    #cálculo da verificação
    dc_posicao['Diferença'] = dc_posicao['Saldo'] - (dc_posicao['Débito'] - dc_posicao['Crédito'])
    columns_to_round = ['Saldo', 'Débito', 'Crédito', 'Diferença']
    dc_posicao[columns_to_round] = dc_posicao[columns_to_round].round(2)
    
    codigo_empresa_df = pd.DataFrame(list(codigo_empresa.items()), columns=['Codigo', 'Nome Cliente'])
# Check for duplicate names in 'Nome_Cliente'
    dc_posicao = dc_posicao.merge(codigo_empresa_df, how='left', left_on='Empresa', right_on='Codigo')
    dc_posicao = dc_posicao[["Empresa", "Nome Cliente", "Saldo", "Débito", "Crédito", "Diferença"]]

    dc_posicao["Diferença"] = dc_posicao["Diferença"].apply(round_small_values)


    comparisons = [
    ('Débito', 'Débito', razao),
    ('Crédito', 'Crédito', razao),
    ('Saldo', 'Saldo', posicao),]

    all_match = True
    for col1, col2, df2 in comparisons:
        is_match, diff = check_sum_difference_between_dfs(dc_posicao, col1, df2, col2)
        if not is_match:
            all_match = False
            break

    if all_match:
        output = BytesIO()
        dc_posicao.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        # Write merged_pdd to an Excel file
        st.download_button(
            label="Download",
            data= output,
            file_name='conciliação.xlsx')
    else:
        raise st.error("Valores não batem")
    