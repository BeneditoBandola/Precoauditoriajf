import streamlit as st
import pandas as pd
import datetime
import os
import json

st.set_page_config(page_title="Monitor de Preços - Royal Canin (JF)", layout="wide")

st.title("🐾 Auditoria Dinâmica de Preços - Royal Canin (Juiz de Fora)")
st.markdown("Busque os produtos por nome, código ou 5 últimos dígitos do EAN, informe o preço da gôndola e a quantidade.")

ARQUIVO_JSON = "auditorias_persistidas.json"

# 1. Carregar base de clientes de Juiz de Fora
@st.cache_data
def load_clients():
    if os.path.exists('clientes base atualizado 14112026.xlsx'):
        df = pd.read_excel('clientes base atualizado 14112026.xlsx', sheet_name='Pessoas')
    elif os.path.exists('clientes base atualizado 14112026.csv'):
        df = pd.read_csv('clientes base atualizado 14112026.csv')
    else:
        return pd.DataFrame()
    return df

df_clientes = load_clients()

if not df_clientes.empty and 'CIDADE' in df_clientes.columns:
    df_jf = df_clientes[df_clientes['CIDADE'].astype(str).str.contains('JUIZ DE FORA', case=False, na=False)].copy()
else:
    df_jf = df_clientes.copy()

codigos_10 = [21209, 21874, 21878, 21140, 24310, 21873, 21921, 21472, 24077, 21888]

# 2. Carregar base de produtos
@st.cache_data
def load_produtos():
    if os.path.exists('preçosrecoroyal.xlsx'):
        try:
            xls = pd.ExcelFile('preçosrecoroyal.xlsx')
            df_p = pd.read_excel('preçosrecoroyal.xlsx', sheet_name=xls.sheet_names[0])
            return df_p
        except Exception:
            pass
    data = [
        {"CÓDIGO": "7896029011111", "LINHA": "Feline Care", "PRODUTO": "Royal Canin Hairball Care 1,5kg", "PREÇO": 89.90},
        {"CÓDIGO": "7896029022222", "LINHA": "Feline Care", "PRODUTO": "Royal Canin Urinary Care 1,5kg", "PREÇO": 94.50},
    ]
    return pd.DataFrame(data)

df_produtos = load_produtos()

# Padronizar colunas de produtos
df_produtos['CÓDIGO_STR'] = df_produtos.apply(lambda r: str(r.get('CÓDIGO', r.get('CODIGO', r.get('Codigo', r.get('SKU', ''))))), axis=1)
df_produtos['PRODUTO_STR'] = df_produtos.apply(lambda r: str(r.get('PRODUTO', r.get('NOME', r.get('DESCRIÇÃO', r.get('DESC_PRODUTO', ''))))), axis=1)
df_produtos['LINHA_STR'] = df_produtos.apply(lambda r: str(r.get('RAMO', r.get('LINHA', r.get('CATEGORIA', 'Royal Canin')))), axis=1)
df_produtos['PRECO_NUM'] = df_produtos.apply(lambda r: float(r.get('PREÇO', r.get('PRECO', r.get('PRECO_RECOMENDADO', r.get('VALOR', 100.0))))), axis=1)

# 3. Gerenciamento de Persistência via JSON
def carregar_dados_json():
    if os.path.exists(ARQUIVO_JSON):
        try:
            with open(ARQUIVO_JSON, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_dados_json(dados):
    with open(ARQUIVO_JSON, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)

if 'historico_auditoria' not in st.session_state:
    st.session_state['historico_auditoria'] = carregar_dados_json()

# --- SIDEBAR: Seleção de Loja ---
st.sidebar.header("📍 Seleção de Loja (Juiz de Fora)")
modo_selecao = st.sidebar.radio("Modo de Lista de Lojas:", ["10 Lojas Principais", "Todos os Clientes de Juiz de Fora"])

if modo_selecao == "10 Lojas Principais":
    df_lojas_disponiveis = df_jf[df_jf['CÓDIGO'].isin(codigos_10)].copy()
else:
    df_lojas_disponiveis = df_jf.copy()

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

# --- ABAS DO APLICATIVO ---
aba_coleta, aba_historico = st.tabs(["📝 Coleta Dinâmica em Loja", "📊 Histórico e Consultas (JSON)"])

with aba_coleta:
    if loja_selecionada is not None:
        st.success(f"**Loja Ativa:** {loja_selecionada['CÓDIGO']} - {loja_selecionada['NOME']} | Endereço: {loja_selecionada.get('ENDEREÇO', 'N/D')}")
        
        # Inicializar estado temporário da loja atual se necessário
        loja_id_str = str(loja_selecionada['CÓDIGO'])
        if 'loja_atual' not in st.session_state or st.session_state['loja_atual'] != loja_id_str:
            st.session_state['loja_atual'] = loja_id_str
            st.session_state['produtos_coletados_loja'] = []

        st.markdown("---")
        st.subheader("🔍 Buscar e Adicionar Produto na Gôndola")
        
        # Campo de busca flexível (Nome, Código ou 5 últimos dígitos do EAN)
        termo_busca = st.text_input("Digite o nome, código ou os 5 últimos dígitos do EAN/Código do produto:")
        
        if termo_busca:
            termo = termo_busca.strip().lower()
            df_encontrados = df_produtos[
                df_produtos['PRODUTO_STR'].str.lower().str.contains(termo, na=False) |
                df_produtos['CÓDIGO_STR'].str.lower().str.contains(termo, na=False) |
                df_produtos['CÓDIGO_STR'].str.slice(-5).str.lower().str.contains(termo, na=False)
            ]
            
            if not df_encontrados.empty:
                for idx, prod in df_encontrados.iterrows():
                    p_cod = prod['CÓDIGO_STR']
                    p_nome = prod['PRODUTO_STR']
                    p_linha = prod['LINHA_STR']
                    p_rec = prod['PRECO_NUM']
                    
                    with st.container():
                        c_img, c_info, c_preco, c_qtd, c_btn = st.columns([1, 3, 2, 2, 2])
                        
                        with c_img:
                            img_path = None
                            for ext in ['.jpg', '.png', '.jpeg', '.JPG', '.PNG']:
                                teste = os.path.join("imagens", f"{p_cod}{ext}")
                                if os.path.exists(teste):
                                    img_path = teste
                                    break
                            if img_path:
                                st.image(img_path, width=60)
                            else:
                                st.markdown("📦 *Sem foto*")
                                
                        with c_info:
                            st.markdown(f"**{p_nome}**<br><small>Cód: {p_cod}</small>", unsafe_allow_html=True)
                            st.text(f"Rec.: R$ {p_rec:.2f}")
                            
                        with c_preco:
                            preco_prat = st.number_input(f"Preço Gôndola ({p_cod})", min_value=0.0, value=0.0, step=0.50, format="%.2f", key=f"preco_{p_cod}")
                            
                        with c_qtd:
                            qtd_loja = st.number_input(f"Qtd Loja ({p_cod})", min_value=0, value=1, step=1, key=f"qtd_{p_cod}")
                            
                        with c_btn:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("➕ Adicionar", key=f"btn_add_{p_cod}"):
                                if preco_prat <= 0:
                                    st.warning("Informe o preço praticado.")
                                else:
                                    dif = preco_prat - p_rec
                                    var = (dif / p_rec) * 100
                                    
                                    novo_item = {
                                        "Data/Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        "Código Loja": int(loja_selecionada['CÓDIGO']),
                                        "Nome Loja": str(loja_selecionada['NOME']),
                                        "Cidade": str(loja_selecionada['CIDADE']),
                                        "Código Produto": p_cod,
                                        "Produto": p_nome,
                                        "Preco Recomendado": p_rec,
                                        "Preco Praticado": preco_prat,
                                        "Quantidade Loja": int(qtd_loja),
                                        "Diferenca (R$)": round(dif, 2),
                                        "Variação (%)": round(var, 2)
                                    }
                                    
                                    st.session_state['produtos_coletados_loja'].append(novo_item)
                                    st.success(f"Adicionado: {p_nome}")
                        st.divider()
            else:
                st.info("Nenhum produto encontrado com este termo.")

        st.markdown("---")
        st.subheader("🛒 Produtos Já Catalogados Nesta Loja")
        
        if st.session_state['produtos_coletados_loja']:
            df_sessao = pd.DataFrame(st.session_state['produtos_coletados_loja'])
            st.dataframe(df_sessao.style.format({
                "Preco Recomendado": "R$ {:.2f}",
                "Preco Praticado": "R$ {:.2f}",
                "Diferenca (R$)": "R$ {:.2f}",
                "Variação (%)": "{:.1f}%"
            }), use_container_width=True)
            
            if st.button("💾 Salvar Permanentemente no Histórico (JSON)"):
                hist_geral = st.session_state['historico_auditoria']
                hist_geral.extend(st.session_state['produtos_coletados_loja'])
                st.session_state['historico_auditoria'] = hist_geral
                salvar_dados_json(hist_geral)
                st.success("✅ Dados salvos com segurança no arquivo JSON!")
        else:
            st.info("Nenhum produto adicionado ainda nesta visita. Use a busca acima para encontrar e catalogar os produtos.")
    else:
        st.info("👈 Selecione uma loja na barra lateral.")

with aba_historico:
    st.header("📂 Histórico Geral Persistido (JSON)")
    st.markdown("Aqui ficam salvos todos os registros de todas as lojas e visitas.")
    
    if st.session_state['historico_auditoria']:
        df_hist = pd.DataFrame(st.session_state['historico_auditoria'])
        st.metric("Total de Registros Salvos", len(df_hist))
        
        st.dataframe(df_hist.style.format({
            "Preco Recomendado": "R$ {:.2f}",
            "Preco Praticado": "R$ {:.2f}",
            "Diferenca (R$)": "R$ {:.2f}",
            "Variação (%)": "{:.1f}%"
        }), use_container_width=True)
        
        csv_b = df_hist.to_csv(index=False).encode('utf-8')
        st.download_button("📥 Baixar Histórico em CSV/Excel", data=csv_b, file_name="historico_auditoria_completo.csv", mime="text/csv")
        
        if st.button("🗑️ Limpar Todo o Histórico Salvo"):
            st.session_state['historico_auditoria'] = []
            if os.path.exists(ARQUIVO_JSON):
                os.remove(ARQUIVO_JSON)
            st.success("Histórico limpo com sucesso.")
            st.rerun()
    else:
        st.info("Nenhum registro no histórico geral ainda.")
