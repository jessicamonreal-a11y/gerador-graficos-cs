import streamlit as st
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker
import io
import warnings

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard CS", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 2. CABEÇALHO SIMPLES ---
st.title("📊 Gerador de Relatórios CS")
st.markdown("Análise comparativa de performance e ocorrências.")
st.divider()

# --- 3. BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets  = st.number_input("Mínimo de Tickets", value=5, min_value=0)
    
    st.markdown("---")
    st.subheader("Períodos para Análise")
    st.info("Digite os meses e volumes separados por vírgula.")
    dados_envios = st.text_area("Mês:Valor", "Fevereiro:1000, Março:1200, Abril:1500")

# --- 4. DEFINIÇÃO DE CORES (Tons de Azul) ---
AZUL_ESCURO = '#185FA5'
AZUL_CLARO  = '#73B2E0'
CORES       = [AZUL_ESCURO, '#4A90E2', AZUL_CLARO, '#A1C9F4', '#BBD9FB', '#CFE2F3']

# --- 5. FUNÇÕES DE GERAÇÃO DE GRÁFICOS ---
def gerar_grafico_ocorrencias(df_plot, meses, totais_mes):
    n_rows = len(df_plot)
    y = np.arange(n_rows)
    n_meses = len(meses)
    h = 0.85 / n_meses
    altura = max(6, n_rows * (0.4 + (n_meses * 0.15)))

    fig, ax = plt.subplots(figsize=(14, altura))
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    for i, mes in enumerate(meses):
        offset = (i - (n_meses - 1) / 2) * h
        vals = df_plot[mes].values
        bars = ax.barh(y + offset, vals, height=h, color=CORES[i % len(CORES)], label=mes, zorder=3)
        for bar, val in zip(bars, vals):
            if val > 0:
                ax.text(val + 0.15, bar.get_y() + bar.get_height() / 2,
                        str(int(val)), va='center', ha='left', fontsize=9, fontweight='500')

    ax.set_yticks(y)
    ax.set_yticklabels(df_plot['Ocorrência'].values, fontsize=10)
    ax.invert_yaxis()
    ax.set_xlabel('Quantidade de tickets', fontsize=11, color='#888780', labelpad=10)
    for sp in ['top', 'right', 'left']: ax.spines[sp].set_visible(False)
    ax.spines['bottom'].set_color('#D3D1C7')
    ax.xaxis.grid(True, color='#D3D1C7', linewidth=0.5, linestyle='--', zorder=0)
    ax.legend(loc='lower right', frameon=False, fontsize=10)
    plt.tight_layout()
    return fig

def gerar_grafico_comparativo(meses, totais_mes, envios_map):
    n_meses = len(meses)
    rows_cards = 1 if n_meses <= 3 else 2
    
    # Cálculo das taxas antes de gerar o gráfico
    taxas_lista = []
    for m in meses:
        tix = totais_mes.get(m, 0)
        env = envios_map.get(m, 1)
        taxas_lista.append(round((tix / env) * 100, 2))

    fig, (ax_top, ax_bot) = plt.subplots(2, 1
