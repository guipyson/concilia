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
    page_title = "Falências"
)

uploaded_file = st.file_uploader("Selecione um arquivo")

if uploaded_file is not None:
    
    razao = pd.read_excel(uploaded_file, sheet_name=0)
    posicao = pd.read_excel(uploaded_file, sheet_name=1)
    pdd = pd.read_excel(uploaded_file, sheet_name=2)

    razao = razao.dropna(subset=['Código Empresa'])
    posicao = posicao.dropna(subset=['Empresa'])
    pdd = pdd.dropna(subset=['Empresa'])

    #Código-Cliente por aba
    razao_dict = dict(zip(razao["Código Empresa"], razao["Descrição Empresa"]))
    posicao_dict = dict(zip(posicao["Empresa"], posicao["Nome Completo"]))
    pdd_dict = dict(zip(pdd["Empresa"], pdd["Nome Completo"]))
    
    #Junção do Código-Cliente
    codigo_empresa = {}
    for d in [razao_dict, posicao_dict, pdd_dict]:
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

    #PDD por cliente
    pdd_grouped = pdd.groupby('Empresa')['Saldo'].sum().reset_index()
    pdd_grouped.rename(columns={'Empresa': 'Código', 'Saldo': 'PDD'}, inplace=True)

    #Merge PDD com dc_posicao
    conciliacao = pd.merge(dc_posicao, pdd_grouped, left_on='Empresa', right_on='Código', how='outer')
    conciliacao.loc[dc_posicao['Empresa'].isna(), 'Empresa'] = conciliacao['Código']
    conciliacao.drop(columns=['Código Empresa', 'Código'], inplace=True)
    conciliacao[['Saldo', 'Débito', 'Crédito', 'PDD']] = conciliacao[['Saldo', 'Débito', 'Crédito', 'PDD']].fillna(0.0)
    #cálculo da verificação
    conciliacao['Diferença'] = conciliacao['Saldo'] - (conciliacao['Débito'] - conciliacao['Crédito'] - conciliacao['PDD'])
    columns_to_round = ['Saldo', 'Débito', 'Crédito', 'PDD', 'Diferença']
    conciliacao[columns_to_round] = conciliacao[columns_to_round].round(2)
    
    codigo_empresa_df = pd.DataFrame(list(codigo_empresa.items()), columns=['Codigo', 'Nome Cliente'])
# Check for duplicate names in 'Nome_Cliente'
    conciliacao = conciliacao.merge(codigo_empresa_df, how='left', left_on='Empresa', right_on='Codigo')
    conciliacao = conciliacao[["Empresa", "Nome Cliente", "Saldo", "Débito", "Crédito", "PDD", "Diferença"]]

    conciliacao["Diferença"] = conciliacao["Diferença"].apply(round_small_values)


    comparisons = [
    ('Débito', 'Débito', razao),
    ('Crédito', 'Crédito', razao),
    ('Saldo', 'Saldo', posicao),
    ('PDD', 'Saldo', pdd)]

    all_match = True
    for col1, col2, df2 in comparisons:
        is_match, diff = check_sum_difference_between_dfs(conciliacao, col1, df2, col2)
        if not is_match:
            all_match = False
            break

    if all_match:
        output = BytesIO()
        conciliacao.to_excel(output, index=False, engine='xlsxwriter')
        output.seek(0)
        # Write merged_pdd to an Excel file
        st.download_button(
            label="Download",
            data= output,
            file_name='conciliação.xlsx')
    else:
        raise st.error("Valores não batem")
    