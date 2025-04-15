import streamlit as st
import pandas as pd
from io import BytesIO
from itertools import combinations
from difflib import SequenceMatcher

st.title("Análise de Transações - Correspondência Parcial com Similaridade")

uploaded_file = st.file_uploader("Envie seu arquivo Excel (.xlsx)", type="xlsx")

def similaridade(a, b):
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

if uploaded_file:
    # Ler o Excel e combinar as abas
    excel_data = pd.ExcelFile(uploaded_file)
    combined_df = pd.concat([pd.read_excel(excel_data, sheet_name=sheet) for sheet in excel_data.sheet_names], ignore_index=True)

    # Ajustar cabeçalho e limpar dados
    combined_df.columns = combined_df.iloc[1]
    combined_df = combined_df[2:]

    # Selecionar colunas relevantes
    relevant_columns = ["Nº Externo", "Valor", "Sinal", "Documento", "Dados Adicionais"]
    filtered_df = combined_df[relevant_columns].copy()
    filtered_df["Nº Externo"] = filtered_df["Nº Externo"].astype(str).str.strip()
    filtered_df["Sinal"] = filtered_df["Sinal"].str.strip()
    filtered_df["Valor"] = pd.to_numeric(filtered_df["Valor"], errors="coerce")
    filtered_df["Dados Adicionais"] = filtered_df["Dados Adicionais"].astype(str)

    # Agrupar por Nº Externo e verificar se o somatório é zero (correspondência perfeita)
    saldo_por_transacao = filtered_df.groupby("Nº Externo")["Valor"].sum()
    transacoes_desequilibradas = saldo_por_transacao[abs(saldo_por_transacao) > 1].index
    transacoes_desequilibradas_df = filtered_df[filtered_df["Nº Externo"].isin(transacoes_desequilibradas)]

    st.subheader("Transações Não Correspondentes (Desequilibradas)")
    st.write(transacoes_desequilibradas_df)

    def encontrar_combinacoes(linhas, alvo, tolerancia=50.0):
        resultados = []
        for r in range(2, min(6, len(linhas)+1)):
            for combo in combinations(linhas.index, r):
                soma = linhas.loc[list(combo), "Valor"].sum()
                if abs(abs(soma) - abs(alvo)) <= tolerancia:
                    resultados.append(linhas.loc[list(combo)])
        return resultados

    st.subheader("Possíveis Correspondências Parciais")
    correspondencias_parciais = []

    for trans_id in transacoes_desequilibradas:
        grupo = filtered_df[filtered_df["Nº Externo"] == trans_id]
        creditos = grupo[grupo["Sinal"] == "C"]
        debitos = grupo[grupo["Sinal"] == "D"]

        total_credito = creditos["Valor"].sum()
        total_debito = debitos["Valor"].sum()

        st.write(f"\nTransação: {trans_id} | Crédito: {total_credito} | Débito: {total_debito}")

        descricoes = grupo["Dados Adicionais"].unique()
        agrupamentos_similares = []

        for desc_base in descricoes:
            similares = grupo[grupo["Dados Adicionais"].apply(lambda x: similaridade(x, desc_base) > 0.85)]
            if not similares.empty:
                creditos_sub = similares[similares["Sinal"] == "C"]
                debitos_sub = similares[similares["Sinal"] == "D"]
                total_c = creditos_sub["Valor"].sum()
                total_d = debitos_sub["Valor"].sum()
                if abs(total_c - total_d) <= 50.0:
                    agrupamentos_similares.append(similares.assign(Grupo=trans_id))

        if agrupamentos_similares:
            correspondencias_parciais.extend(agrupamentos_similares)
        else:
            if abs(total_credito) > abs(total_debito):
                combinacoes = encontrar_combinacoes(creditos, total_debito)
            else:
                combinacoes = encontrar_combinacoes(debitos, total_credito)
            for match in combinacoes:
                correspondencias_parciais.append(match.assign(Grupo=trans_id))

    if correspondencias_parciais:
        correspondencias_df = pd.concat(correspondencias_parciais)
        st.write(correspondencias_df)

        output = BytesIO()
        correspondencias_df.to_excel(output, index=False, engine='openpyxl')
        output.seek(0)

        st.download_button(
            label="📥 Baixar Correspondências Parciais",
            data=output,
            file_name="correspondencias_parciais.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nenhuma correspondência parcial encontrada dentro da tolerância definida.")

    output2 = BytesIO()
    transacoes_desequilibradas_df.to_excel(output2, index=False, engine='openpyxl')
    output2.seek(0)

    st.download_button(
        label="📥 Baixar Transações Desequilibradas",
        data=output2,
        file_name="transacoes_desequilibradas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
