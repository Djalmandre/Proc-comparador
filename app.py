import streamlit as st
import pandas as pd
import openpyxl
from io import BytesIO
import base64

# Configuração da página
st.set_page_config(
    page_title="Comparador de Planilhas",
    page_icon="📊",
    layout="wide"
)

# Título e descrição
st.title("📊 Comparador de Planilhas")
st.markdown("""
Esta aplicação permite comparar duas planilhas e identificar itens que se repetem,
similar às funções PROCV e PROCX do Excel.
""")

# Função para carregar planilha
@st.cache_data
def carregar_planilha(arquivo):
    try:
        df = pd.read_excel(arquivo)
        return df
    except Exception as e:
        st.error(f"Erro ao carregar arquivo: {str(e)}")
        return None

# Função para comparar planilhas
def comparar_planilhas(df1, df2, coluna1, coluna2, tipo_comparacao="exata"):
    resultados = []
    contagem = {}
    
    if tipo_comparacao == "exata":
        # Comparação exata
        for idx1, valor1 in df1[coluna1].items():
            if pd.notna(valor1):
                # Procura valores correspondentes na segunda planilha
                matches = df2[df2[coluna2] == valor1]
                
                if not matches.empty:
                    for idx2, row2 in matches.iterrows():
                        resultados.append({
                            'Valor': valor1,
                            'Linha Planilha 1': idx1 + 2,  # +2 porque Excel começa em 1 e tem cabeçalho
                            'Linha Planilha 2': idx2 + 2,
                            'Dados Planilha 1': df1.loc[idx1].to_dict(),
                            'Dados Planilha 2': row2.to_dict()
                        })
                    
                    # Contagem de repetições
                    if valor1 not in contagem:
                        contagem[valor1] = 0
                    contagem[valor1] += len(matches)
    
    elif tipo_comparacao == "parcial":
        # Comparação parcial (contém)
        for idx1, valor1 in df1[coluna1].items():
            if pd.notna(valor1):
                valor1_str = str(valor1).lower()
                
                for idx2, valor2 in df2[coluna2].items():
                    if pd.notna(valor2):
                        valor2_str = str(valor2).lower()
                        
                        if valor1_str in valor2_str or valor2_str in valor1_str:
                            resultados.append({
                                'Valor Planilha 1': valor1,
                                'Valor Planilha 2': valor2,
                                'Linha Planilha 1': idx1 + 2,
                                'Linha Planilha 2': idx2 + 2,
                                'Dados Planilha 1': df1.loc[idx1].to_dict(),
                                'Dados Planilha 2': df2.loc[idx2].to_dict()
                            })
                            
                            if valor1 not in contagem:
                                contagem[valor1] = 0
                            contagem[valor1] += 1
    
    return resultados, contagem

# Função para converter DataFrame para Excel
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Resultados')
    processed_data = output.getvalue()
    return processed_data

# Sidebar para upload de arquivos
st.sidebar.header("📁 Upload de Arquivos")

arquivo1 = st.sidebar.file_uploader(
    "Carregar Planilha 1 (Excel)",
    type=['xlsx', 'xls'],
    key="arquivo1"
)

arquivo2 = st.sidebar.file_uploader(
    "Carregar Planilha 2 (Excel)",
    type=['xlsx', 'xls'],
    key="arquivo2"
)

# Processamento principal
if arquivo1 and arquivo2:
    # Carregar planilhas
    df1 = carregar_planilha(arquivo1)
    df2 = carregar_planilha(arquivo2)
    
    if df1 is not None and df2 is not None:
        # Exibir preview das planilhas
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("📄 Planilha 1")
            st.dataframe(df1.head(), use_container_width=True)
            st.caption(f"Total de linhas: {len(df1)}")
        
        with col2:
            st.subheader("📄 Planilha 2")
            st.dataframe(df2.head(), use_container_width=True)
            st.caption(f"Total de linhas: {len(df2)}")
        
        st.divider()
        
        # Configurações de comparação
        st.header("⚙️ Configurações de Comparação")
        
        col_config1, col_config2, col_config3 = st.columns(3)
        
        with col_config1:
            coluna1 = st.selectbox(
                "Coluna da Planilha 1:",
                options=df1.columns.tolist(),
                key="coluna1"
            )
        
        with col_config2:
            coluna2 = st.selectbox(
                "Coluna da Planilha 2:",
                options=df2.columns.tolist(),
                key="coluna2"
            )
        
        with col_config3:
            tipo_comparacao = st.selectbox(
                "Tipo de Comparação:",
                options=["exata", "parcial"],
                format_func=lambda x: "Exata (=)" if x == "exata" else "Parcial (contém)",
                key="tipo_comp"
            )
        
        # Botão de comparação
        if st.button("🔍 Comparar Planilhas", type="primary", use_container_width=True):
            with st.spinner("Comparando planilhas..."):
                resultados, contagem = comparar_planilhas(
                    df1, df2, coluna1, coluna2, tipo_comparacao
                )
                
                if resultados:
                    st.success(f"✅ Encontradas {len(resultados)} correspondências!")
                    
                    # Estatísticas
                    st.header("📊 Estatísticas")
                    
                    col_stat1, col_stat2, col_stat3 = st.columns(3)
                    
                    with col_stat1:
                        st.metric("Total de Correspondências", len(resultados))
                    
                    with col_stat2:
                        st.metric("Valores Únicos Encontrados", len(contagem))
                    
                    with col_stat3:
                        if contagem:
                            max_repeticoes = max(contagem.values())
                            st.metric("Máximo de Repetições", max_repeticoes)
                    
                    st.divider()
                    
                    # Tabela de contagem
                    st.subheader("🔢 Contagem de Repetições")
                    
                    df_contagem = pd.DataFrame([
                        {'Valor': k, 'Quantidade de Repetições': v}
                        for k, v in sorted(contagem.items(), key=lambda x: x[1], reverse=True)
                    ])
                    
                    st.dataframe(df_contagem, use_container_width=True)
                    
                    st.divider()
                    
                    # Resultados detalhados
                    st.subheader("📋 Resultados Detalhados")
                    
                    # Criar DataFrame de resultados
                    if tipo_comparacao == "exata":
                        df_resultados = pd.DataFrame([
                            {
                                'Valor': r['Valor'],
                                'Linha Planilha 1': r['Linha Planilha 1'],
                                'Linha Planilha 2': r['Linha Planilha 2']
                            }
                            for r in resultados
                        ])
                    else:
                        df_resultados = pd.DataFrame([
                            {
                                'Valor Planilha 1': r['Valor Planilha 1'],
                                'Valor Planilha 2': r['Valor Planilha 2'],
                                'Linha Planilha 1': r['Linha Planilha 1'],
                                'Linha Planilha 2': r['Linha Planilha 2']
                            }
                            for r in resultados
                        ])
                    
                    st.dataframe(df_resultados, use_container_width=True)
                    
                    # Busca específica
                    st.divider()
                    st.subheader("🔎 Buscar Valor Específico")
                    
                    valor_busca = st.text_input(
                        "Digite o valor que deseja procurar:",
                        key="busca"
                    )
                    
                    if valor_busca:
                        resultados_filtrados = [
                            r for r in resultados 
                            if str(valor_busca).lower() in str(r.get('Valor', r.get('Valor Planilha 1', ''))).lower()
                        ]
                        
                        if resultados_filtrados:
                            st.success(f"Encontradas {len(resultados_filtrados)} ocorrências de '{valor_busca}'")
                            
                            for i, r in enumerate(resultados_filtrados, 1):
                                with st.expander(f"Ocorrência {i}"):
                                    col_a, col_b = st.columns(2)
                                    
                                    with col_a:
                                        st.write("**Planilha 1:**")
                                        st.json(r['Dados Planilha 1'])
                                    
                                    with col_b:
                                        st.write("**Planilha 2:**")
                                        st.json(r['Dados Planilha 2'])
                        else:
                            st.warning(f"Nenhuma ocorrência encontrada para '{valor_busca}'")
                    
                    # Download dos resultados
                    st.divider()
                    st.subheader("💾 Exportar Resultados")
                    
                    col_down1, col_down2 = st.columns(2)
                    
                    with col_down1:
                        excel_resultados = to_excel(df_resultados)
                        st.download_button(
                            label="📥 Download Resultados Detalhados (Excel)",
                            data=excel_resultados,
                            file_name="resultados_comparacao.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col_down2:
                        excel_contagem = to_excel(df_contagem)
                        st.download_button(
                            label="📥 Download Contagem (Excel)",
                            data=excel_contagem,
                            file_name="contagem_repeticoes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                
                else:
                    st.warning("⚠️ Nenhuma correspondência encontrada entre as planilhas.")

else:
    # Instruções iniciais
    st.info("""
    ### 📌 Como usar:
    
    1. **Faça upload** das duas planilhas Excel no menu lateral
    2. **Selecione** as colunas que deseja comparar
    3. **Escolha** o tipo de comparação (exata ou parcial)
    4. **Clique** em "Comparar Planilhas"
    5. **Visualize** os resultados e faça download se necessário
    
    ### 💡 Tipos de Comparação:
    
    - **Exata**: Procura valores idênticos (como PROCV)
    - **Parcial**: Procura valores que contêm parte do texto
    """)
    
    # Exemplo visual
    st.subheader("📊 Exemplo Visual")
    
    col_ex1, col_ex2 = st.columns(2)
    
    with col_ex1:
        st.write("**Planilha 1:**")
        exemplo1 = pd.DataFrame({
            'Código': ['A001', 'B002', 'C003'],
            'Produto': ['Notebook', 'Mouse', 'Teclado']
        })
        st.dataframe(exemplo1, use_container_width=True)
    
    with col_ex2:
        st.write("**Planilha 2:**")
        exemplo2 = pd.DataFrame({
            'ID': ['A001', 'D004', 'B002'],
            'Descrição': ['Laptop', 'Monitor', 'Mouse Sem Fio']
        })
        st.dataframe(exemplo2, use_container_width=True)

# Footer
st.divider()
st.caption("Desenvolvido para comparação de planilhas | Djalma A Barbosa 2026 - Petrobras")