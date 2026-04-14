import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
import warnings

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard CS", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 2. CABEÇALHO ---
st.title("📊 Gerador de Relatórios CS")
st.markdown("Análise comparativa de performance e ocorrências por período.")
st.divider()

# --- 3. BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets  = st.number_input("Mínimo de Tickets (Filtro)", value=5, min_value=0)
    
    st.markdown("---")
    st.subheader("Períodos para Análise")
    st.info("Formato: Mês:Valor (Ex: Março:1200, Abril:1500)")
    dados_envios = st.text_area("Lista de Envios", "Março:1200, Abril:1500")

# --- 4. CORES ---
AZUL_ESCURO = '#185FA5'
AZUL_CLARO  = '#73B2E0'
CORES = [AZUL_ESCURO, '#4A90E2', AZUL_CLARO, '#A1C9F4', '#BBD9FB', '#CFE2F3']

# --- 5. FUNÇÕES DE GRÁFICOS ---
def gerar_grafico_ocorrencias(df_plot, meses, totais_mes):
    n_rows = len(df_plot)
    y = np.arange(n_rows)
    n_meses = len(meses)
    h = 0.8 / n_meses
    altura = max(6, n_rows * (0.4 + (n_meses * 0.15)))

    fig, ax = plt.subplots(figsize=(14, altura))
    fig.patch.set_facecolor('white')
    ax.set_facecolor('white')

    for i, mes in enumerate(meses):
        offset = (i - (n_meses - 1) / 2) * h
        vals = df_plot[mes].values
        bars = ax.barh(y + offset, vals, height=h, color=CORES[i % len(CORES)], label=mes)
        for bar, val in zip(bars, vals):
            if val > 0:
                ax.text(val + 0.1, bar.get_y() + bar.get_height()/2, str(int(val)), va='center', fontsize=9)

    ax.set_yticks(y)
    ax.set_yticklabels(df_plot['Ocorrência'].values, fontsize=10)
    ax.invert_yaxis()
    ax.set_xlabel('Quantidade de tickets')
    ax.legend(loc='lower right', frameon=False)
    plt.tight_layout()
    return fig

def gerar_grafico_comparativo(meses, totais_mes, envios_map):
    taxas = []
    for m in meses:
        tix = totais_mes.get(m, 0)
        env = envios_map.get(m, 1)
        taxas.append(round((tix / env) * 100, 2))

    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(12, 9), gridspec_kw={'height_ratios': [1, 2]})
    ax_top.axis('off')

    # Cards de Resumo Dinâmicos
    n = len(meses)
    card_w = 0.9 / n
    for i, mes in enumerate(meses):
        x = 0.05 + i * card_w
        env = envios_map.get(mes, 0)
        tix = totais_mes.get(mes, 0)
        pct = taxas[i]
        
        # Correção da linha que gerou o erro (parêntese fechado no final)
        ax_top.add_patch(mpatches.FancyBboxPatch((x, 0.1), card_w-0.02, 0.8, boxstyle="round,pad=0.02", facecolor='#F0F0F0', transform=ax_top.transAxes))
        
        ax_top.text(x + card_w/2, 0.7, mes.upper(), ha='center', fontweight='bold', transform=ax_top.transAxes)
        ax_top.text(x + card_w/2, 0.4, f"{tix} tickets", ha='center', transform=ax_top.transAxes)
        ax_top.text(x + card_w/2, 0.2, f"{pct}%", ha='center', color=AZUL_ESCURO, fontweight='bold', transform=ax_top.transAxes)

    # Gráfico de Barras (Taxa)
    ax_bot.bar(meses, taxas, color=CORES[:n], width=0.5)
    for i, v in enumerate(taxas):
        ax_bot.text(i, v + 0.1, f"{v}%", ha='center', fontweight='bold')
    
    ax_bot.set_ylabel('Taxa de Contato (%)')
    ax_bot.spines['top'].set_visible(False)
    ax_bot.spines['right'].set_visible(False)
    plt.tight_layout()
    return fig

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    return buf

# --- 6. LÓGICA PRINCIPAL ---
arquivo = st.file_uploader("📂 Selecione o arquivo Excel do Zendesk", type=["xlsx"])

if arquivo:
    try:
        df = pd.read_excel(arquivo)
        df = df[df.astype(str).ne('SUM').all(axis=1)].copy()
        
        col_ocorr = next((c for c in df.columns if 'Ocorr' in c or 'Motivo' in c), None)
        
        if col_ocorr:
            df.rename(columns={col_ocorr: 'Ocorrência'}, inplace=True)
            
            envios_map = {item.split(':')[0].strip(): int(item.split(':')[1].strip()) 
                         for item in dados_envios.split(',') if ':' in item}
            meses_list = list(envios_map.keys())
            
            cols_tickets, totais_mes = {}, {}
            for mes in meses_list:
                col = next((c for c in df.columns if mes.lower() in c.lower() and ('Tickets' in c or 'Qtd' in c)), None)
                if col:
                    cols_tickets[mes] = col
                    totais_mes[mes] = int(df[col].sum())

            if cols_tickets:
                agg = df.groupby('Ocorrência')[list(cols_tickets.values())].sum().reset_index()
                agg.columns = ['Ocorrência'] + list(cols_tickets.keys())
                agg['Total'] = agg[list(cols_tickets.keys())].sum(axis=1)
                agg = agg.sort_values('Total', ascending=False)
                df_plot = agg[agg['Total'] >= min_tickets].copy()

                st.subheader("📊 Distribuição de Ocorrências")
                fig_oc = gerar_grafico_ocorrencias(df_plot, list(cols_tickets.keys()), totais_mes)
                st.pyplot(fig_oc)
                st.download_button("💾 Baixar Gráfico de Ocorrências", fig_to_bytes(fig_oc), "ocorrencias.png", "image/png")
                
                st.divider()
                
                st.subheader("📈 Taxa de Contato e Resumo")
                fig_comp = gerar_grafico_comparativo(list(cols_tickets.keys()), totais_mes, envios_map)
                st.pyplot(fig_comp)
                st.download_button("💾 Baixar Gráfico de Taxa", fig_to_bytes(fig_comp), "taxa.png", "image/png")
            else:
                st.error("Não foram encontradas colunas de tickets para os meses informados.")
        else:
            st.error("Não foi possível encontrar a coluna de 'Ocorrência' no arquivo.")
            
    except Exception as e:
        st.error(f"Erro ao processar o arquivo: {e}")
