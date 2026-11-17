import streamlit as st
import pandas as pd
import datetime
from datetime import timezone, timedelta
import os
import json
import base64
import requests
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="Monitor de Preços - Royal Canin (JF)", layout="wide")

st.title("🐾 Auditoria Dinâmica de Preços - Royal Canin (Juiz de Fora)")
st.markdown("Busque por parte do nome, EAN ou código. Os dados são salvos no GitHub e os e-mails são enviados via Gmail.")

ARQUIVO_JSON = "auditorias_persistidas.json"
CUSTOM_PROD_JSON = "produtos_customizados.json" 

def hora_brasil():
    tz_br = timezone(timedelta(hours=-3))
    return datetime.datetime.now(tz_br).strftime("%d/%m/%Y %H:%M:%S")

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
                
                val_c_raw = row.iloc[2] if len(row) > 2 else None
                if pd.notna(val_c_raw):
                    try:
                        val_c = str(int(float(val_c_raw)))
                    except:
                        val_c = str(val_c_raw).strip()
                else:
                    val_c = ""
                
                val_d = str(row.iloc[3]).strip() if len(row) > 3 and pd.notna(row.iloc[3]) else ""
                val_e = str(row.iloc[4]).strip() if len(row) > 4 and pd.notna(row.iloc[4]) else ""
                
                val_f_raw = row.iloc[5] if len(row) > 5 else None
                if pd.notna(val_f_raw):
                    try:
                        val_f = str(int(float(val_f_raw)))
                    except:
                        val_f = str(val_f_raw).strip()
                else:
                    val_f = ""

                val_i = row.iloc[8] if len(row) > 8 and pd.notna(row.iloc[8]) else 0.0
                
                try:
                    if isinstance(val_i, (int, float)):
                        preco = float(val_i)
                    else:
                        preco = float(str(val_i).replace('R$', '').replace('.', '').replace(',', '.').strip())
                except:
                    preco = 0.0
                
                if val_d or val_e or val_f or val_c:
                    nome_principal = val_d if val_d and val_d.lower() != 'nan' else val_e
                    ean = val_f if val_f and val_f.lower() != 'nan' else f"ITEM_{idx}"
                    
                    produtos.append({
                        "COD_C": val_c,
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
        {"COD_C": "100", "PRODUTO": "Royal Canin Exemplo", "EAN": "7896029011111", "PREÇO": 89.90, "ROW_INDEX": "1"},
    ]
    return pd.DataFrame(data)

df_produtos = load_produtos()

def parse_preco_br(texto):
    if not texto:
        return 0.0
    try:
        limpo = str(texto).strip().replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        return float(limpo)
    except:
        return 0.0

def fmt_brl(val):
    return f"R$ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def carregar_dados_json(nome_arquivo):
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo = st.secrets["GITHUB_REPO"]
            url = f"https://api.github.com/repos/{repo}/contents/{nome_arquivo}"
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
            resp = requests.get(url, headers=headers)
            if resp.status_code == 200:
                file_info = resp.json()
                file_content = base64.b64decode(file_info['content']).decode('utf-8')
                data = json.loads(file_content)
                with open(nome_arquivo, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=4)
                return data
        except Exception:
            pass

    if os.path.exists(nome_arquivo):
        try:
            with open(nome_arquivo, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_dados_json(dados, nome_arquivo):
    with open(nome_arquivo, 'w', encoding='utf-8') as f:
        json.dump(dados, f, ensure_ascii=False, indent=4)
        
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            token = st.secrets["GITHUB_TOKEN"]
            repo = st.secrets["GITHUB_REPO"]
            url = f"https://api.github.com/repos/{repo}/contents/{nome_arquivo}"
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/vnd.github+json"}
            
            resp_get = requests.get(url, headers=headers)
            sha = resp_get.json().get('sha') if resp_get.status_code == 200 else None
            
            content_json_str = json.dumps(dados, ensure_ascii=False, indent=4)
            content_encoded = base64.b64encode(content_json_str.encode('utf-8')).decode('utf-8')
            
            payload = {
                "message": f"Atualização de {nome_arquivo} via app em {hora_brasil()}",
                "content": content_encoded
            }
            if sha:
                payload["sha"] = sha
                
            requests.put(url, headers=headers, json=payload)
        except Exception as e:
            st.warning(f"Aviso: Não foi possível sincronizar o {nome_arquivo} com o GitHub ({e}), mas está salvo localmente.")

if 'historico_auditoria' not in st.session_state:
    st.session_state['historico_auditoria'] = carregar_dados_json(ARQUIVO_JSON)

if 'produtos_customizados' not in st.session_state:
    st.session_state['produtos_customizados'] = carregar_dados_json(CUSTOM_PROD_JSON)

if st.session_state['produtos_customizados']:
    df_customizados = pd.DataFrame(st.session_state['produtos_customizados'])
    df_produtos = pd.concat([df_produtos, df_customizados], ignore_index=True)

def enviar_email_com_anexo(destinatarios_str, assunto, corpo, caminho_pdf):
    remetente = "beneditobandola@gmail.com"
    if "GMAIL_APP_PASSWORD" not in st.secrets:
        return False, "⚠️ GMAIL_APP_PASSWORD não configurado nos Secrets do Streamlit."
    
    senha = st.secrets["GMAIL_APP_PASSWORD"]
    
    try:
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatarios_str
        msg['Subject'] = assunto
        
        msg.attach(MIMEText(corpo, 'plain', 'utf-8'))
        
        if os.path.exists(caminho_pdf):
            with open(caminho_pdf, "rb") as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(caminho_pdf)}"')
            msg.attach(part)
            
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(remetente, senha)
        destinatarios_lista = [d.strip() for d in destinatarios_str.split(',') if d.strip()]
        server.sendmail(remetente, destinatarios_lista, msg.as_string())
        server.quit()
        return True, "E-mail enviado com sucesso!"
    except Exception as e:
        return False, f"Erro ao enviar e-mail: {str(e)}"

def gerar_pdf_loja(nome_loja, cod_loja, cidade_loja, itens):
    pdf_filename = f"auditoria_loja_{cod_loja}.pdf"
    doc = SimpleDocTemplate(pdf_filename, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    story = []
    styles = getSampleStyleSheet()
    
    PRIMARY_COLOR = colors.HexColor("#1A365D")
    SECONDARY_COLOR = colors.HexColor("#2B6CB0")
    TEXT_COLOR = colors.HexColor("#2D3748")
    BG_LIGHT = colors.HexColor("#F7FAFC")
    BORDER_COLOR = colors.HexColor("#E2E8F0")
    
    title_style = ParagraphStyle('DocTitle', parent=styles['Heading1'], fontSize=18, leading=22, textColor=PRIMARY_COLOR, spaceAfter=6, fontName='Helvetica-Bold')
    subtitle_style = ParagraphStyle('DocSubtitle', parent=styles['Normal'], fontSize=10, leading=14, textColor=TEXT_COLOR, spaceAfter=15, fontName='Helvetica')
    section_style = ParagraphStyle('SectionHeading', parent=styles['Heading2'], fontSize=12, leading=16, textColor=SECONDARY_COLOR, spaceBefore=10, spaceAfter=8, fontName='Helvetica-Bold')
    body_style = ParagraphStyle('BodyTextCustom', parent=styles['Normal'], fontSize=9, leading=12, textColor=TEXT_COLOR, fontName='Helvetica')
    footer_info_style = ParagraphStyle('FooterInfo', parent=styles['Normal'], fontSize=8.5, leading=12, textColor=colors.HexColor("#4A5568"), fontName='Helvetica-Oblique')

    story.append(Paragraph("RELATÓRIO DE AUDITORIA DE PREÇOS", title_style))
    story.append(Paragraph(f"<b>Portfólio:</b> Royal Canin &mdash; <b>Loja:</b> {cod_loja} - {nome_loja} | <b>Cidade:</b> {cidade_loja}<br/><b>Data/Hora da Emissão (BR):</b> {hora_brasil()}", subtitle_style))
    story.append(Spacer(1, 5))
    story.append(Paragraph("Detalhamento dos Itens Coletados na Gôndola", section_style))
    
    table_data = [["Produto", "Qtd", "Preço Rec.", "Preço Loja", "Total Rec.", "Total Loja", "Dif. (R$)", "Var (%)"]]
    total_rec_geral = 0.0
    total_loja_geral = 0.0
    
    for item in itens:
        qtd = int(item['Quantidade Loja'])
        p_rec = float(item['Preco Recomendado'])
        p_loja = float(item['Preco Praticado'])
        tot_rec = qtd * p_rec
        tot_loja = qtd * p_loja
        total_rec_geral += tot_rec
        total_loja_geral += tot_loja
        
        table_data.append([
            Paragraph(item['Produto'], body_style),
            str(qtd),
            fmt_brl(p_rec),
            fmt_brl(p_loja),
            fmt_brl(tot_rec),
            fmt_brl(tot_loja),
            fmt_brl(item['Diferenca (R$)']),
            f"{item['Variação (%)']:+.2f}%".replace('.', ',')
        ])
        
    t = Table(table_data, colWidths=[160, 30, 60, 60, 65, 65, 60, 50])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), PRIMARY_COLOR),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('ALIGN', (0,0), (0,-1), 'LEFT'),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('BOTTOMPADDING', (0,0), (-1,0), 6),
        ('TOPPADDING', (0,0), (-1,0), 6),
        ('BACKGROUND', (0,1), (-1,-1), colors.white),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, BG_LIGHT]),
        ('GRID', (0,0), (-1,-1), 0.5, BORDER_COLOR),
        ('FONTSIZE', (0,1), (-1,-1), 8.5),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    
    story.append(t)
    story.append(Spacer(1, 15))
    
    diff_total = total_loja_geral - total_rec_geral
    diff_pct_total = (diff_total / total_rec_geral * 100) if total_rec_geral > 0 else 0.0
    
    story.append(Paragraph("Resumo e Análise Financeira (Informativo)", section_style))
    resumo_texto = f"""
    <b>Total Recomendado (Tabela):</b> {fmt_brl(total_rec_geral)}<br/>
    <b>Total Praticado na Loja:</b> {fmt_brl(total_loja_geral)}<br/>
    <b>Diferença Financeira Total:</b> {fmt_brl(diff_total)} ({diff_pct_total:+.2f}% em relação ao recomendado)<br/><br/>
    <i>Nota Informativa:</i> Esta simulação demonstra o impacto financeiro caso todo o volume auditado fosse adquirido considerando os preços praticados na gôndola em comparação com a tabela de preços recomendados pela marca. Relatório gerado automaticamente para suporte à gestão.
    """
    story.append(Paragraph(resumo_texto, footer_info_style))
    doc.build(story)
    return pdf_filename

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

aba_coleta, aba_historico = st.tabs(["📝 Coleta Dinâmica em Loja", "📊 Histórico, Edição e Exclusão (JSON)"])

with aba_coleta:
    if loja_selecionada is not None:
        st.success(f"**Loja Ativa:** {loja_selecionada['CÓDIGO']} - {loja_selecionada['NOME']} | Endereço: {loja_selecionada.get('ENDEREÇO', 'N/D')}")
        
        cod_loja_atual = int(loja_selecionada['CÓDIGO'])
        itens_loja_salvos = [item for item in st.session_state['historico_auditoria'] if int(item.get('Código Loja', 0)) == cod_loja_atual]

        st.markdown("---")
        st.subheader("🔍 Buscar Produto por Nome, EAN ou Código")
        
        termo_busca = st.text_input("Digite sua busca (ex: 'uri', '100052', '78960'):")
        
        with st.expander("➕ Não encontrou o produto? Cadastre manualmente"):
            with st.form("form_add_produto_custom"):
                st.markdown("Os produtos cadastrados aqui ficarão salvos para sempre no sistema e aparecerão nas buscas.")
                c1, c2, c3 = st.columns(3)
                with c1:
                    novo_nome = st.text_input("Nome do Produto *")
                with c2:
                    novo_codigo = st.text_input("Código na Coluna C ou EAN")
                with c3:
                    novo_preco_rec = st.number_input("Preço Recomendado (R$)", min_value=0.0, format="%.2f")
                
                if st.form_submit_button("Salvar Produto no Sistema"):
                    if not novo_nome.strip():
                        st.error("O nome do produto é obrigatório!")
                    else:
                        _cod = novo_codigo.strip() if novo_codigo.strip() else f"CUSTOM_{datetime.datetime.now().strftime('%H%M%S')}"
                        novo_item_bd = {
                            "PRODUTO": novo_nome.strip().upper(),
                            "EAN": _cod,
                            "COD_C": _cod,
                            "PREÇO": novo_preco_rec,
                            "ROW_INDEX": f"C_{_cod}"
                        }
                        st.session_state['produtos_customizados'].append(novo_item_bd)
                        salvar_dados_json(st.session_state['produtos_customizados'], CUSTOM_PROD_JSON)
                        st.success("Produto cadastrado com sucesso! Já pode buscá-lo acima.")
                        st.rerun()
        
        if termo_busca:
            termo = termo_busca.strip().lower()
            
            def match_row(row):
                if termo in str(row.get('PRODUTO', '')).lower(): return True
                if termo in str(row.get('EAN', '')).lower(): return True
                if termo in str(row.get('COD_C', '')).lower(): return True
                return False

            df_encontrados = df_produtos[df_produtos.apply(match_row, axis=1)]
            
            if not df_encontrados.empty:
                st.info(f"Encontrado(s) {len(df_encontrados)} produto(s):")
                for idx, prod in df_encontrados.iterrows():
                    p_ean = str(prod.get('EAN', ''))
                    p_nome = str(prod.get('PRODUTO', ''))
                    p_rec = float(prod.get('PREÇO', 0.0))
                    row_idx = str(prod.get('ROW_INDEX', ''))
                    
                    p_cod = str(prod.get('COD_C', '')).strip()
                    if not p_cod:
                        p_cod = p_ean
                        
                    p_rec_str = fmt_brl(p_rec)
                    
                    with st.container():
                        c_img, c_info, c_preco, c_qtd, c_btn = st.columns([1, 3, 2, 2, 2])
                        
                        with c_img:
                            img_path = None
                            
                            candidatos = [c for c in [p_cod, p_ean] if c]
                                        
                            if os.path.exists("imagens"):
                                arquivos_pasta = os.listdir("imagens")
                                for arq in arquivos_pasta:
                                    nome_base, ext = os.path.splitext(arq)
                                    if ext.lower() in ['.png', '.jpg', '.jpeg']:
                                        for cand in candidatos:
                                            if cand.lower() == nome_base.lower():
                                                img_path = os.path.join("imagens", arq)
                                                break
                                    if img_path:
                                        break
                                        
                            if img_path and os.path.exists(img_path):
                                st.image(img_path, width=140)
                            else:
                                st.markdown("📦 *Sem foto*")
                                
                        with c_info:
                            st.markdown(f"**{p_nome}**<br><small>Código: {p_cod}</small>", unsafe_allow_html=True)
                            st.text(f"Preço recomendado: {p_rec_str}")
                            
                        with c_preco:
                            preco_prat_txt = st.text_input(f"Preço Gôndola", value="", placeholder="0,00", key=f"preco_{row_idx}_{idx}")
                            preco_prat = parse_preco_br(preco_prat_txt)
                            
                        with c_qtd:
                            qtd_loja = st.number_input(f"Qtd Loja", min_value=0, value=1, step=1, key=f"qtd_{row_idx}_{idx}")
                            
                        with c_btn:
                            st.markdown("<br>", unsafe_allow_html=True)
                            if st.button("➕ Adicionar", key=f"btn_add_{row_idx}_{idx}"):
                                if preco_prat <= 0:
                                    st.warning("Informe o preço praticado.")
                                else:
                                    ja_adicionado = any(str(item.get('Código Produto')) == str(p_cod) for item in itens_loja_salvos)
                                    if ja_adicionado:
                                        st.error(f"⚠️ Este produto ({p_nome}) já está salvo nesta loja!")
                                    else:
                                        dif = preco_prat - p_rec
                                        var = (dif / p_rec) * 100 if p_rec > 0 else 0.0
                                        
                                        novo_item = {
                                            "ID_REGISTRO": f"{cod_loja}_{p_cod}_{datetime.datetime.now().strftime('%H%M%S%f')}",
                                            "Data/Hora": hora_brasil(),
                                            "Código Loja": cod_loja,
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
                                        
                                        st.session_state['historico_auditoria'].append(novo_item)
                                        salvar_dados_json(st.session_state['historico_auditoria'], ARQUIVO_JSON)
                                        st.success(f"Adicionado e sincronizado com o GitHub: {p_nome}")
                                        st.rerun()
                        st.divider()
            else:
                st.info("Nenhum produto encontrado com este termo de busca.")

        st.markdown("---")
        st.subheader("🛒 Produtos Já Catalogados Nesta Loja")
        
        itens_loja_salvos = [item for item in st.session_state['historico_auditoria'] if int(item.get('Código Loja', 0)) == cod_loja_atual]
        
        if itens_loja_salvos:
            for i, item in enumerate(itens_loja_salvos):
                prat_str = fmt_brl(item['Preco Praticado'])
                rec_str = fmt_brl(item['Preco Recomendado'])
                dif_str = fmt_brl(item['Diferenca (R$)'])
                var_val = item['Variação (%)']
                var_str = f"{var_val:+.2f}%".replace('.', ',')
                
                with st.container():
                    st.markdown(f"""
                    **{item['Produto']}**  
                    * **Quantidade:** {item['Quantidade Loja']}  
                    * **Preço Recomendado:** {rec_str}  
                    * **Preço da Loja:** {prat_str}  
                    * **Diferença (R$):** {dif_str}  
                    * **Variação (%):** {var_str}
                    """)
                    if st.button("🗑️ Excluir da Loja", key=f"rm_eterno_{cod_loja_atual}_{i}"):
                        id_alvo = item.get('ID_REGISTRO')
                        st.session_state['historico_auditoria'] = [x for x in st.session_state['historico_auditoria'] if x.get('ID_REGISTRO') != id_alvo]
                        salvar_dados_json(st.session_state['historico_auditoria'], ARQUIVO_JSON)
                        st.success("Item removido e sincronizado!")
                        st.rerun()
                    st.divider()
            
            if st.button("🚪 Salvar e Sair"):
                salvar_dados_json(st.session_state['historico_auditoria'], ARQUIVO_JSON)
                st.success("✅ Progresso salvo!")
                st.rerun()

            st.markdown("---")
            st.subheader("📄 Relatório PDF & Envio por E-mail")
            
            pdf_path = gerar_pdf_loja(loja_selecionada['NOME'], loja_selecionada['CÓDIGO'], loja_selecionada['CIDADE'], itens_loja_salvos)
            
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
                
            st.download_button(
                label="📥 Baixar Relatório em PDF",
                data=pdf_bytes,
                file_name=f"auditoria_{loja_selecionada['CÓDIGO']}.pdf",
                mime="application/pdf"
            )
            
            st.markdown("#### Enviar por E-mail")
            opcao_envio = st.radio("Destinatário:", ["Somente para mim", "Time inteiro"], key="opcao_envio_radio")
            
            if opcao_envio == "Somente para mim":
                email_padrao = "benedito.bandola@minassal.com.br, beneditobandola@gmail.com"
            else:
                email_padrao = "caio.poli@minassal.com.br, daniel.santini@minassal.com.br, rubens.porfirio@minassal.com.br, benedito.bandola@minassal.com.br, beneditobandola@gmail.com"
                
            email_destino = st.text_input("E-mail(s) de destino:", value=email_padrao)
            
            if st.button("📧 Enviar Relatório por E-mail"):
                assunto = f"Relatório de Auditoria de Preços - Loja {loja_selecionada['CÓDIGO']} - {loja_selecionada['NOME']}"
                corpo = f"Olá,\n\nSegue em anexo o relatório emitido em {hora_brasil()}."
                sucesso, mensagem = enviar_email_com_anexo(email_destino, assunto, corpo, pdf_path)
                if sucesso:
                    st.success(f"✅ Enviado com sucesso para: {email_destino}")
                else:
                    st.error(f"❌ Erro: {mensagem}")
        else:
            st.info("Nenhum produto cadastrado nesta loja ainda.")
    else:
        st.info("👈 Selecione uma loja na lateral.")

with aba_historico:
    st.header("📂 Histórico Geral")
    if st.session_state['historico_auditoria']:
        hist = st.session_state['historico_auditoria']
        lojas_hist = list(set([f"{item['Código Loja']} - {item['Nome Loja']}" for item in hist]))
        filtro_l = st.selectbox("Filtrar Histórico:", ["Todas as Lojas"] + lojas_hist)
        
        for i, item in enumerate(hist):
            if filtro_l != "Todas as Lojas" and filtro_l != f"{item['Código Loja']} - {item['Nome Loja']}":
                continue
            with st.expander(f"📌 {item['Nome Loja']} — {item['Produto']} | R$ {item['Preco Praticado']}"):
                with st.form(f"form_edit_{i}"):
                    c_e1, c_e2, c_e3 = st.columns(3)
                    with c_e1:
                        novo_prat = st.number_input("Praticado (R$)", value=float(item['Preco Praticado']), step=0.50)
                    with c_e2:
                        nova_qtd = st.number_input("Qtd", value=int(item['Quantidade Loja']), step=1)
                        
                    col_b1, col_b2 = st.columns(2)
                    with col_b1:
                        if st.form_submit_button("✏️ Atualizar"):
                            rec = float(item['Preco Recomendado'])
                            dif = novo_prat - rec
                            var = (dif / rec) * 100 if rec > 0 else 0.0
                            hist[i].update({'Preco Praticado': float(novo_prat), 'Quantidade Loja': int(nova_qtd), 'Diferenca (R$)': round(dif, 2), 'Variação (%)': round(var, 2)})
                            st.session_state['historico_auditoria'] = hist
                            salvar_dados_json(hist, ARQUIVO_JSON)
                            st.success("Atualizado!")
                            st.rerun()
                    with col_b2:
                        if st.form_submit_button("🗑️ Excluir"):
                            hist.pop(i)
                            st.session_state['historico_auditoria'] = hist
                            salvar_dados_json(hist, ARQUIVO_JSON)
                            st.success("Excluído!")
                            st.rerun()
    else:
        st.info("Nenhum registro.")
