import streamlit as st
import pandas as pd
import datetime
import os
import json

st.set_page_config(page_title="Monitor de Preços - Royal Canin (JF)", layout="wide")

st.title("🐾 Auditoria Dinâmica de Preços - Royal Canin (Juiz de Fora)")
st.markdown("Busque por parte do nome (colunas D ou E), EAN (coluna F) ou código. As imagens são carregadas automaticamente da pasta `imagens/`.")

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

# 2. Carregar base de produtos da planilha preçosrecoroyal.xlsx
@st.cache_data
def load_produtos():
    if os.path.exists('preçosrecoroyal.xlsx'):
        try:
            xls = pd.ExcelFile('preçosrecoroyal.xlsx')
            df_raw = pd.read_excel('preçosrecoroyal.xlsx', sheet_name=xls.sheet_names[0], header=None)
            
            produtos = []
            for idx, row in df_raw.iterrows():
                if idx == 0 and any(str(v).upper() in ['DESCRIÇÃO', 'PRODUTO', 'NOME', 'EAN', 'PREÇO'] for v in row.values):
                    continue
                
                # Capturar colunas de código/ID (A, B, C)
                cods = []
                for col_idx in range(min(3, len(row))):
                    val = str(row.iloc[col_idx]).strip()
                    if val and val.lower() != 'nan':
                        cods.append(val.replace('.0', ''))
                
                # Coluna D = index 3, Coluna E = index 4, Coluna F = index 5, Coluna I = index 8
                val_d = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                val_e = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
                val_f = str(row.iloc[5]).strip() if len(row) > 5 and pd.notna(row.iloc[5]) else ""
                val_i = row.iloc[8] if len(row) > 8 and pd.notna(row.iloc[8]) else 0.0
                
                try:
                    if isinstance(val_i, (int, float)):
                        preco = float(val_i)
                    else:
                        preco = float(str(val_i).replace('R$', '').replace('.', '').replace(',', '.').strip())
                except:
                    preco = 0.0
                
                if val_d or val_e or val_f or cods:
                    nome_principal = val_d if val_d and val_d.lower() != 'nan' else val_e
                    ean = val_f if val_f and val_f.lower() != 'nan' else ""
                    
                    produtos.append({
                        "CODIGOS": cods,
                        "NOME_D": val_d,
                        "NOME_E": val_e,
                        "PRODUTO": nome_principal,
                        "EAN": ean,
                        "PREÇO": preco,
                        "ROW_INDEX": str(idx)
                    })
            if produtos:
                return pd.DataFrame(produtos)
        except Exception as e:
            pass
            
    data = [
        {"CODIGOS": ["100052"], "NOME_D": "Royal Canin Hairball Care 1,5kg", "NOME_E": "Hairball 1.5", "PRODUTO": "Royal Canin Hairball Care 1,5kg", "EAN": "7896029011111", "PREÇO": 89.90, "ROW_INDEX": "1"},
    ]
    return pd.DataFrame(data)

df_produtos = load_produtos()

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
        
        loja_id_str = str(loja_selecionada['CÓDIGO'])
        if 'loja_atual' not in st.session_state or st.session_state['loja_atual'] != loja_id_str:
            st.session_state['loja_atual'] = loja_id_str
            st.session_state['produtos_coletados_loja'] = []

        st.markdown("---")
        st.subheader("🔍 Buscar Produto por Nome (Col. D ou E), EAN (Col. F) ou Código")
        
        termo_busca = st.text_input("Digite sua busca (ex: 'uri', '1kg', '100052', '78960'):")
        
        if termo_busca:
            termo = termo_busca.strip().lower()
            
            def match_row(row):
                if termo in str(row['NOME_D']).lower(): return True
                if termo in str(row['NOME_E']).lower(): return True
                if termo in str(row['EAN']).lower(): return True
                for c in row['CODIGOS']:
                    if termo in str(c).lower(): return True
                return False

            df_encontrados = df_produtos[df_produtos.apply(match_row, axis=1)]
            
            if not df_encontrados.empty:
                st.info(f"Encontrado(s) {len(df_encontrados)} produto(s):")
                for idx, prod in df_encontrados.iterrows():
                    codigos_lista = prod['CODIGOS']
                    p_ean = str(prod['EAN'])
                    p_nome = str(prod['PRODUTO'])
                    p_rec = float(prod['PREÇO'])
                    row_idx = str(prod['ROW_INDEX'])
                    
                    # Formatar preço com vírgula
                    p_rec_str = f"R$ {p_rec:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    
                    with st.container():
                        c_img, c_info, c_preco, c_qtd, c_btn = st.columns([1, 3, 2, 2, 2])
                        
                        with c_img:
                            img_path = None
                            candidatos = codigos_lista + [p_ean, row_idx]
                            
                            if os.path.exists("imagens"):
                                arquivos_pasta = os.listdir("imagens")
                                for arq in arquivos_pasta:
                                    nome_base, ext = os.path.splitext(arq)
                                    if ext.lower() in ['.png', '.jpg', '.jpeg']:
                                        for cand in candidatos:
                                            if cand and (cand.lower() == nome_base.lower() or cand.lower() in nome_base.lower() or nome_base.lower() in cand.lower()):
                                                img_path = os.path.join("imagens", arq)
                                                break
                                    if img_path:
                                        break
                                        
                            if img_path and os.path.exists(img_path):
                                st.image(img_path, width=70)
                            else:
                                st.markdown("📦 *Sem foto*")
                                
                        with c_info:
                            cod_exibicao = codigos_lista[0] if codigos_lista else p_ean
                            st.markdown(f"**{p_nome}**<br><small>Cód: {cod_exibicao} | EAN: {p_ean}</small>", unsafe_allow_html=True)
                            st.text(f"Preço recomendado: {p_rec_str}")
                            
                        with c_preco:
                            preco_prat = st.number_input(f"Preço Gôndola", min_value=0.0, value=0.0, step=0.50, format="%.2f", key=f"preco_{row_idx}_{idx}")
                            
                        with c_qtd:
                            qtd_loja = st.number_input(f"Qtd Loja", min_value=0, value=1, step=1, key=f"qtd_{row_idx}_{idx}")
                            
                        with c_btn:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("➕ Adicionar", key=f"btn_add_{row_idx}_{idx}"):
                                if preco_prat <= 0:
                                    st.warning("Informe o preço praticado.")
                                else:
                                    dif = preco_prat - p_rec
                                    var = (dif / p_rec) * 100 if p_rec > 0 else 0.0
                                    
                                    novo_item = {
                                        "Data/Hora": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                        "Código Loja": int(loja_selecionada['CÓDIGO']),
                                        "Nome Loja": str(loja_selecionada['NOME']),
                                        "Cidade": str(loja_selecionada['CIDADE']),
                                        "Código Produto": cod_exibicao,
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
                st.info("Nenhum produto encontrado com este termo de busca.")

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
            st.info("Nenhum produto adicionado ainda nesta visita. Use a busca acima.")
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
