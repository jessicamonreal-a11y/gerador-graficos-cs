import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
import warnings
from pathlib import Path
import base64

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard CS", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 2. FUNÇÃO PARA CARREGAR O LOGO (BASE64) ---
def get_base64_logo(img_path):
    try:
        if Path(img_path).exists():
            with open(img_path, "rb") as f:
                data = f.read()
            return base64.b64encode(data).decode()
    except:
        return None
    return None

# Nome do arquivo de imagem que você deve subir no GitHub
logo_file = "logo.png" 
logo_b64 = get_base64_logo(logo_file)

# --- 3. CABEÇALHO COM LOGO OPCIONAL ---
if logo_b64:
    st.markdown(
        f"""
        <div style="display: flex; align-items: center; gap: 20px;">
            <img src="data:image/png;base64,{logo_b64}" width="80">
            <h1 style="margin: 0;">Gerador de Relatórios CS</h1>
        </div>
        <hr>
        """, 
        unsafe_allow_html=True
    )
else:
    st.title("📊 Gerador de Gráficos CS - Tickets Zendesk")
    st.markdown("Análise comparativa de performance e ocorrências por período.")
    st.divider()

# --- 4. BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    if logo_b64:
        st.markdown(f'<div style="text-align: center;"><img src="data:image/png;base64,{logo_b64}" width="120"></div>', unsafe_allow_html=True)
    
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets  = st.number_input("Mínimo de Tickets (Filtro - quantidade mínima de tickets sobre o tema para que apareça no gráfico. Para aparecer todos, deixe como 1):", value=5, min_value=0)
    
    st.markdown("---")
    st.subheader("Períodos para Análise")
    st.info("Formato: Mês:Valor (Ex: Março:1200, Abril:1500, Maio:1400)")
    dados_envios_input = st.text_area("Quantidade de Envios (digite a quantidade de envios para cada mês que deseja que apareça no gráfico):", "Março:1200, Abril:1500")

# --- 5. CORES E ESTILO ---
AZUL_ESCURO = '#185FA5'
CORES_LISTA = ['#185FA5', '#4A90E2', '#73B2E0', '#A1C9F4', '#BBD9FB', '#CFE2F3', '#DDEEFF']

# --- 6. FUNÇÕES DE GRÁFICOS ---
def gerar_grafico_ocorrencias(df_plot, meses_nomes, totais_por_mes):
    n_ocorrencias = len(df_plot)
    n_meses = len(meses_nomes)
    
    # Ajusta altura do gráfico dinamicamente
    altura_calc = max(6, n_ocorrencias * (0.4 + (n_meses * 0.15)))
    fig, ax = plt.subplots(figsize=(12, altura_calc))
    
    y = np.arange(n_ocorrencias)
    largura_barra = 0.8 / n_meses

    for i, mes in enumerate(meses_nomes):
        # Calcula o deslocamento de cada barra para não sobrepor
        deslocamento = (i - (n_meses - 1) / 2) * largura_barra
        valores = df_plot[mes].values
        
        cor = CORES_LISTA[i % len(CORES_LISTA)]
        bars = ax.barh(y + deslocamento, valores, height=largura_barra, label=mes, color=cor)
        
        # Adiciona rótulos numéricos
        for bar in bars:
            width = bar.get_width()
            if width > 0:
                ax.text(width + 0.2, bar.get_y() + bar.get_height()/2, f'{int(width)}', va='center', fontsize=9)

    ax.set_yticks(y)
    ax.set_yticklabels(df_plot['Ocorrência'].values, fontsize=10)
    ax.invert_yaxis()
    ax.set_xlabel('Quantidade de Tickets')
    ax.legend(title="Meses", loc='lower right')
    plt.tight_layout()
    return fig

def gerar_grafico_taxa(meses_nomes, totais_tickets, envios_map):
    taxas = []
    for m in meses_nomes:
        tix = totais_tickets.get(m, 0)
        env = envios_map.get(m, 1)
        taxas.append(round((tix / env) * 100, 2))

    fig, ax = plt.subplots(figsize=(10, 5))
    x_pos = np.arange(len(meses_nomes))
    
    bars = ax.bar(x_pos, taxas, color=CORES_LISTA[:len(meses_nomes)], width=0.5)
    
    for i, bar in enumerate(bars):
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2, height + 0.1, f'{taxas[i]}%', ha='center', fontweight='bold')

    ax.set_xticks(x_pos)
    ax.set_xticklabels(meses_nomes)
    ax.set_ylabel('Taxa de Contato (%)')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    return fig

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    return buf

# --- 7. LÓGICA PRINCIPAL ---
arquivo = st.file_uploader("📂 Selecione o arquivo Excel do Zendesk", type=["xlsx"])

if arquivo:
    try:
        df = pd.read_excel(arquivo)
        df = df[df.astype(str).ne('SUM').all(axis=1)].copy()
        
        col_ocorr = next((c for c in df.columns if 'Ocorr' in c or 'Motivo' in c), None)
        
        if col_ocorr:
            df.rename(columns={col_ocorr: 'Ocorrência'}, inplace=True)
            
            # 1. Mapeia meses do input
            envios_map = {}
            for item in dados_envios_input.split(','):
                if ':' in item:
                    m, v = item.split(':')
                    envios_map[m.strip()] = int(v.strip())
            
            meses_selecionados = list(envios_map.keys())
            
            # 2. Localiza colunas no Excel
            cols_encontradas = {}
            totais_por_mes = {}
            
            for mes in meses_selecionados:
                col = next((c for c in df.columns if mes.lower() in c.lower() and ('Tickets' in c or 'Qtd' in c)), None)
                if col:
                    cols_encontradas[mes] = col
                    totais_por_mes[mes] = int(df[col].sum())

            if cols_encontradas:
                meses_finais = list(cols_encontradas.keys())
                
                # Agrupamento
                agg = df.groupby('Ocorrência')[list(cols_encontradas.values())].sum().reset_index()
                agg.columns = ['Ocorrência'] + meses_finais
                agg['Total_Soma'] = agg[meses_finais].sum(axis=1)
                agg = agg.sort_values('Total_Soma', ascending=False)
                
                df_plot = agg[agg['Total_Soma'] >= min_tickets].copy()

                # Gráficos
                st.subheader(f"📊 Análise: {nome_cliente}")
                
                fig_oc = gerar_grafico_ocorrencias(df_plot, meses_finais, totais_por_mes)
                st.pyplot(fig_oc)
                st.download_button("💾 Baixar Gráfico de Ocorrências", fig_to_bytes(fig_oc), "ocorrencias.png", "image/png")
                
                st.divider()
                
                st.subheader("📈 Taxa de Contato por Período")
                fig_tx = gerar_grafico_taxa(meses_finais, totais_por_mes, envios_map)
                st.pyplot(fig_tx)
                st.download_button("💾 Baixar Gráfico de Taxa", fig_to_bytes(fig_tx), "taxa.png", "image/png")
            else:
                st.error("Nenhum dos meses digitados foi encontrado nas colunas do Excel.")
        else:
            st.error("Coluna de 'Ocorrência' não identificada.")
            
    except Exception as e:
        st.error(f"Erro no processamento: {e}")
