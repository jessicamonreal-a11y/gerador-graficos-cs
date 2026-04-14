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
from pathlib import Path
import base64

# --- 1. CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Dashboard CS — Mandaê", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 2. FUNÇÕES AUXILIARES ---
def logo_base64(path):
    try:
        if path.exists():
            with open(path, 'rb') as f:
                return base64.b64encode(f.read()).decode()
    except:
        return None
    return None

# --- 3. LOGO E CABEÇALHO ---
logo_path = Path(__file__).parent / 'logo_mandae.png'
b64 = logo_base64(logo_path)
logo_html = f'<img src="data:image/png;base64,{b64}" style="height:60px; margin-bottom:10px;">' if b64 else ""

st.markdown(f"""
<div style="display:flex; align-items:center; gap:20px; padding:10px 0;">
    {logo_html}
    <div>
        <h2 style="margin:0; color:#0D1B2A; font-size:24px; font-weight:700;">Gerador de Relatórios CS</h2>
        <p style="margin:0; color:#888780; font-size:14px;">Análise comparativa de performance e ocorrências.</p>
    </div>
</div>
<hr style="border:none; border-top:1px solid #D3D1C7; margin-bottom:25px;">
""", unsafe_allow_html=True)

# --- 4. BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    if b64:
        st.markdown(f'<div style="text-align: center;"><img src="data:image/png;base64,{b64}" width="150"></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
    
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets  = st.number_input("Mínimo de Tickets", value=5, min_value=0)
    
    st.markdown("---")
    st.subheader("Períodos para Análise")
    st.info("Digite os meses e volumes separados por vírgula.")
    dados_envios = st.text_area("Mês:Valor", "Fevereiro:1000, Março:1200, Abril:1500")

# --- 5. DEFINIÇÃO DE CORES (Tons de Azul) ---
AZUL_ESCURO = '#185FA5'
AZUL_CLARO  = '#73B2E0'
CORES       = [AZUL_ESCURO, '#4A90E2', AZUL_CLARO, '#A1C9F4', '#BBD9FB', '#CFE2F3']

# --- 6. FUNÇÕES DE GERAÇÃO DE GRÁFICOS ---
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
    # Define se os cards ficam em 1 ou 2 linhas
    rows_cards = 1 if n_meses <= 3 else 2
    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(12, 8 if rows_cards == 1 else 10), 
                                          gridspec_kw={'height_ratios': [1 * rows_cards, 2.5]})
    ax_top.axis('off')

    # Cards Dinâmicos
    cards_per_row = n_meses if n_meses <= 3 else int(np.ceil(n_meses / 2))
    card_w = 0.96 / cards_per_row
    
    for i, mes in enumerate(meses):
        env = envios_map.get(mes, 0)
        tix = totais_mes.get(mes, 0)
        pct = (tix / env * 100) if env > 0 else 0
        
        row = i // cards_per_row
        col = i % cards_per_row
        
        x = 0.02 + col * card_w
        y = 0.55 if (rows_cards == 2 and row == 0) else 0.05
        h_box = 0.4 if rows_cards == 2 else 0.8
        
        ax_top.add_patch(mpatches.FancyBboxPatch((x, y), card_w-0.02, h_box, 
                                                boxstyle="round,pad=0.01", facecolor='#F1EFE8', transform=ax_top.transAxes))
        ax_top.text(x + card_w/2, y + h_box*0.75, mes.upper(), ha='center', fontsize=9, fontweight='bold', color='#888780', transform=ax_top.transAxes)
        ax_top.text(x + card_w/2, y + h_box*0.4, f"{tix} tickets", ha='center', fontsize=12, fontweight='bold', transform=ax_top.transAxes)
        ax_top.text(x + card_w/2, y + h_box*0.15, f"Taxa: {pct:.2f}%", ha='center', fontsize=10, color=AZUL_ESCURO, fontweight='bold', transform=ax_top.transAxes)

    # Gráfico de Barras (Taxa de Contato)
    taxas = [round(totais_mes[m] / envios_map[m] * 100, 2) if envios_map.get(m, 0) > 0 else 0 for m in meses]
    ax_bot.bar(meses, taxas, color=CORES[:len(meses)], width=0.4, zorder=3)
    for i, v in enumerate(taxas):
        ax_bot.text(i, v + (max(taxas)*0.02), f'{v:.2f}%'.replace('.', ','), ha='center', fontweight='bold', size=11)
    
    ax_bot.set_ylabel('Taxa de Contato (%)', color='#888780', labelpad=10)
    ax_bot.yaxis.grid(True, color='#D3D1C7', linestyle='--', linewidth=0.5)
    for sp in ['top', 'right', 'left']: ax_bot.spines[sp].set_visible(False)
    plt.tight_layout()
    return fig

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    return buf

# --- 7. LÓGICA DE PROCESSAMENTO ---
arquivo_excel = st.file_uploader("📂 Selecione o Excel do Zendesk", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel)
        df = df[df.astype(str).ne('SUM').all(axis=1)].copy()
        
        col_ocorr = next((c for c in df.columns if 'Ocorr' in c or 'Motivo' in c), None)
        if col_ocorr:
            df.rename(columns={col_ocorr: 'Ocorrência'}, inplace=True)
            envios_map = {item.split(':')[0].strip(): int(item.split(':')[1].strip()) 
                         for item in dados_envios.split(',') if ':' in item}
            meses = list(envios_map.keys())
            
            cols_tickets, totais_mes = {}, {}
            for mes in meses:
                col = next((c for c in df.columns if mes.lower() in c.lower() and ('Tickets' in c or 'Qtd' in c)), None)
                if col:
                    cols_tickets[mes] = col
                    totais_mes[mes] = int(df[col].sum())

            if cols_tickets:
                agg = df.groupby('Ocorrência')[list(cols_tickets.values())].sum().reset_index()
                agg.columns = ['Ocorrência'] + meses
                agg['Total'] = agg[meses].sum(axis=1)
                agg = agg.sort_values('Total', ascending=False)
                df_plot = agg[agg['Total'] >= min_tickets].copy()

                st.divider()
                # Gráfico 1
                st.subheader("📊 Distribuição de Ocorrências")
                fig_oc = gerar_grafico_ocorrencias(df_plot, meses, totais_mes)
                st.pyplot(fig_oc)
                st.download_button("💾 Baixar Gráfico de Ocorrências", fig_to_bytes(fig_oc), f"ocorrencias_{nome_cliente.lower()}.png", "image/png")
                
                st.divider()
                # Gráfico 2
                st.subheader("📈 Taxa de Contato e Resumo")
                fig_comp = gerar_grafico_comparativo(meses, totais_mes, envios_map)
                st.pyplot(fig_comp)
                st.download_button("💾 Baixar Gráfico de Taxa", fig_to_bytes(fig_comp), f"taxa_{nome_cliente.lower()}.png", "image/png")
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
