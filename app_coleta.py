import streamlit as st
import pandas as pd
import datetime
import os

st.set_page_config(page_title="Monitor de Preços - Royal Canin (JF)", layout="wide")

st.title("🐾 Auditoria de Preços - Royal Canin (Juiz de Fora)")
st.markdown("Ferramenta de campo para monitoramento de preços gôndola vs. preço recomendado.")

ARQUIVO_HISTORICO = "historico_pesquisas.csv"

# 1. Carregar base de clientes de Juiz de Fora
@st.cache_data
def load_clients():
    if os.path.exists('clientes base atualizado 14112026.xlsx'):
        df = pd.read_excel('clientes base atualizado 14112026.xlsx', sheet_name='Pessoas')
    elif os.path.exists('clientes base atualizado 14112026.csv'):
        df = pd.read_csv('clientes base atualizado 14112026.csv')
    else:
        st.error("Arquivo de clientes não encontrado no repositório.")
        return pd.DataFrame()
    return df

df_clientes = load_clients()

# Filtrar apenas Juiz de Fora
if not df_clientes.empty and 'CIDADE' in df_clientes.columns:
    df_jf = df_clientes[df_clientes['CIDADE'].astype(str).str.contains('JUIZ DE FORA', case=False, na=False)].copy()
else:
    df_jf = df_clientes.copy()

# Os 10 códigos principais solicitados
codigos_10 = [21209, 21874, 21878, 21140, 24310, 21873, 21921, 21472, 24077, 21888]

# 2. Carregar tabela de preços recomendados
@st.cache_data
def load_produtos():
    if os.path.exists('preçosrecoroyal.xlsx'):
        try:
            xls = pd.ExcelFile('preçosrecoroyal.xlsx')
            df_p = pd.read_excel('preçosrecoroyal.xlsx', sheet_name=xls.sheet_names[0])
            return df_p
        except Exception as e:
            st.warning(f"Erro ao ler preçosrecoroyal.xlsx: {e}")
            
    # Tabela padrão caso o arquivo não seja encontrado
    data = [
        {"CÓDIGO": "7896029011111", "LINHA": "Feline Care", "PRODUTO": "Royal Canin Hairball Care 1,5kg", "PREÇO": 89.90},
        {"CÓDIGO": "7896029022222", "LINHA": "Feline Care", "PRODUTO": "Royal Canin Urinary Care 1,5kg", "PREÇO": 94.50},
    ]
    return pd.DataFrame(data)

df_produtos = load_produtos()

# --- SIDEBAR: Seleção de Lojas ---
st.sidebar.header("📍 Seleção de Loja (Juiz de Fora)")

modo_selecao = st.sidebar.radio("Modo de Lista de Lojas:", ["10 Lojas Principais", "Todos os Clientes de Juiz de Fora"])

if modo_selecao == "10 Lojas Principais":
    df_lojas_disponiveis = df_jf[df_jf['CÓDIGO'].isin(codigos_10)].copy()
else:
    df_lojas_disponiveis = df_jf.copy()

# Criar rótulo amigável para exibição
if not df_lojas_disponiveis.empty:
    df_lojas_disponiveis['LABEL'] = df_lojas_disponiveis['CÓDIGO'].astype(str) + " - " + df_lojas_disponiveis['NOME'].fillna('') + (
        " (" + df_lojas_disponiveis['NOME FANTASIA'].fillna('') + ")"
    ).str.replace(r' \(\)', '', regex=True)

loja_selecionada = None
if not df_lojas_disponiveis.empty:
    opcao_escolhida = st.sidebar.selectbox("Selecione a Loja:", df_lojas_disponiveis['LABEL'].tolist())
    if opcao_escolhida:
        cod_loja = int(opcao_escolhida.split(" - ")[0])
        loja_selecionada = df_lojas_disponiveis[df_lojas_disponiveis['CÓDIGO'] == cod_loja].iloc[0]
else:
    st.sidebar.warning("Nenhuma loja encontrada.")

# --- ABAS DO APLICATIVO ---
aba_coleta, aba_historico = st.tabs(["📝 Coleta e Auditoria", "📊 Histórico Geral"])

with aba_coleta:
    if loja_selecionada is not None:
        st.success(f"**Loja Ativa:** {loja_selecionada['CÓDIGO']} - {loja_selecionada['NOME']} | Endereço: {loja_selecionada.get('ENDEREÇO', 'N/D')}")
        
        st.markdown("---")
        st.subheader("📋 Catálogo de Produtos e Preços Gôndola")
        
        with st.form("form_auditoria"):
            registros = []
            
            for idx, row in df_produtos.iterrows():
                # Código do produto
                p_cod = str(row.get('CÓDIGO', row.get('CODIGO', row.get('Codigo', row.get('SKU', idx)))))
                # Nome do produto (corrigido com o parêntese fechado corretamente)
                p_nome = str(row.get('PRODUTO', row.get('NOME', row.get('DESCRIÇÃO', row.get('DESC_PRODUTO', f"Produto {idx}")))))
                # Linha / Ramo
                p_linha = str(row.get('RAMO', row.get('LINHA', row.get('CATEGORIA', 'Royal Canin'))))
                # Preço recomendado
                p_rec = float(row.get('PREÇO', row.get('PRECO', row.get('PRECO_RECOMENDADO', row.get('VALOR', 100.0)))))
                
                col_img, col_info, col_preco = st.columns([1, 4, 3])
                
                with col_img:
                    # Buscar foto na pasta imagens/ usando o código do produto
                    img_path = None
                    for ext in ['.jpg', '.png', '.jpeg', '.JPG', '.PNG']:
                        caminho_teste = os.path.join("imagens", f"{p_cod}{ext}")
                        if os.path.exists(caminho_teste):
                            img_path = caminho_teste
                            break
                            
                    if img_path:
                        st.image(img_path, width=70)
                    else:
                        st.markdown("📦 *Sem foto*")
                        
                with col_info:
                    st.markdown(f"**{p_nome}**<br><small style='color:gray;'>Cód: {p_cod} | Linha: {p_linha}</small>", unsafe_allow_html=True)
                    st.text(f"Preço Recomendado: R$ {p_rec:.2f}")
                    
                with col_preco:
                    preco_gondola = st.number_input(
                        f"Preço gôndola ({p_cod})",
                        min_value=0.0,
                        value=0.0,
                        step=0.50,
                        format="%.2f",
                        label_visibility="collapsed"
                    )
                    
                registros.append({
                    "Data/Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "Código Loja": loja_selecionada['CÓDIGO'],
                    "Nome Loja": loja_selecionada['NOME'],
                    "Cidade": loja_selecionada['CIDADE'],
                    "Código Produto": p_cod,
                    "Produto": p_nome,
                    "Preco Recomendado": p_rec,
                    "Preco Praticado": preco_gondola
                })
                st.divider()
                
            enviar = st.form_submit_button("💾 Salvar Auditoria desta Loja")
            
            if enviar:
                df_res = pd.DataFrame(registros)
                df_res = df_res[df_res['Preco Praticado'] > 0]
                
                if df_res.empty:
                    st.warning("Nenhum preço de gôndola foi preenchido.")
                else:
                    df_res['Diferenca (R$)'] = df_res['Preco Praticado'] - df_res['Preco Recomendado']
                    df_res['Variacão (%)'] = (df_res['Diferenca (R$)'] / df_res['Preco Recomendado']) * 100
                    
                    if os.path.exists(ARQUIVO_HISTORICO):
                        df_hist_antigo = pd.read_csv(ARQUIVO_HISTORICO)
                        df_total = pd.concat([df_hist_antigo, df_res], ignore_index=True)
                    else:
                        df_total = df_res
                        
                    df_total.to_csv(ARQUIVO_HISTORICO, index=False)
                    st.success("✅ Auditoria salva com sucesso no histórico!")
                    
                    st.subheader("📊 Resumo Comparativo da Visita")
                    st.dataframe(df_res.style.format({
                        "Preco Recomendado": "R$ {:.2f}",
                        "Preco Praticado": "R$ {:.2f}",
                        "Diferenca (R$)": "R$ {:.2f}",
                        "Variacão (%)": "{:.1f}%"
                    }), use_container_width=True)
    else:
        st.info("👈 Selecione uma loja válida.")

with aba_historico:
    st.header("📂 Histórico Geral de Auditorias")
    if os.path.exists(ARQUIVO_HISTORICO):
        df_h = pd.read_csv(ARQUIVO_HISTORICO)
        st.dataframe(df_h.style.format({
            "Preco Recomendado": "R$ {:.2f}",
            "Preco Praticado": "R$ {:.2f}",
            "Diferenca (R$)": "R$ {:.2f}",
            "Variacão (%)": "{:.1f}%"
        }), use_container_width=True)
        
        csv_bytes = df_h.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Baixar Histórico Completo (Excel/CSV)", data=csv_bytes, file_name="historico_auditoria_jf.xlsx", mime="text/csv")
        
        if st.button("🗑️ Limpar Histórico"):
            os.remove(ARQUIVO_HISTORICO)
            st.rerun()
    else:
        st.info("Nenhum histórico registrado ainda.")
