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

# --- 1. CONFIGURAÇÃO DA PÁGINA (OBRIGATORIAMENTE O PRIMEIRO COMANDO) ---
st.set_page_config(page_title="Dashboard CS — Mandaê", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# --- 2. FUNÇÕES AUXILIARES ---
def logo_base64(path):
    try:
        with open(path, 'rb') as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

# --- 3. LOGO E CABEÇALHO ---
logo_path = Path(__file__).parent / 'logo_mandae.png'
b64 = logo_base64(logo_path)
logo_html = f'<img src="data:image/png;base64,{b64}" style="height:56px; margin-bottom:8px;">' if b64 else ""

st.markdown(f"""
<div style="display:flex; align-items:center; gap:20px; padding:8px 0 16px 0;">
    {logo_html}
    <div>
        <h2 style="margin:0; color:#0D1B2A; font-size:22px; font-weight:600;">Gerador de Relatórios CS</h2>
        <p style="margin:0; color:#888780; font-size:13px;">Preencha os dados na lateral e faça o upload do arquivo.</p>
    </div>
</div>
<hr style="border:none; border-top:1px solid #D3D1C7; margin-bottom:20px;">
""", unsafe_allow_html=True)

# --- 4. BARRA LATERAL (SIDEBAR) ---
with st.sidebar:
    if logo_path.exists():
        st.image(str(logo_path), width=160)
        st.markdown("<hr style='border:none;border-top:1px solid #D3D1C7;margin:8px 0 16px 0'>", unsafe_allow_html=True)
    
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets  = st.number_input("Mínimo de Tickets", value=5, min_value=0)
    
    st.markdown("---")
    st.subheader("Volumes de Envios")
    st.caption("Exemplo: Março:1200, Abril:1500")
    dados_envios = st.text_area("Mês:Valor", "Março:1200, Abril:1500")

# --- 5. DEFINIÇÃO DE CORES ---
AZUL_ESCURO = '#185FA5'
AZUL_CLARO  = '#73B2E0'
CORES       = [AZUL_ESCURO, AZUL_CLARO, '#4A90E2', '#A1C9F4']

# --- 6. FUNÇÕES DE GERAÇÃO DE GRÁFICOS ---
def gerar_grafico_ocorrencias(df_plot, meses, totais_mes):
    n = len(df_plot)
    y = np.arange(n)
    h = 0.8 / len(meses)
    altura = max(6, n * 0.55 + 1.5)

    fig, ax = plt.subplots(figsize=(14, altura))
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    for i, mes in enumerate(meses):
        offset = (i - (len(meses) - 1) / 2) * h
        vals = df_plot[mes].values
        bars = ax.barh(y + offset, vals, height=h, color=CORES[i % 4], label=mes, zorder=3)
        for bar, val in zip(bars, vals):
            if val > 0:
                ax.text(val + 0.15, bar.get_y() + bar.get_height() / 2,
                        str(int(val)), va='center', ha='left', fontsize=9,
                        color='#444441', fontweight='500')

    ax.set_yticks(y)
    ax.set_yticklabels(df_plot['Ocorrência'].values, fontsize=10, color='#333333')
    ax.invert_yaxis()
    ax.set_xlabel('Quantidade de tickets', fontsize=11, color='#888780', labelpad=8)
    for sp in ['top', 'right', 'left']: ax.spines[sp].set_visible(False)
    ax.spines['bottom'].set_color('#D3D1C7')
    ax.xaxis.grid(True, color='#D3D1C7', linewidth=0.5, zorder=0)
    ax.legend(loc='lower right', frameon=False)
    plt.tight_layout(pad=1.5)
    return fig

def gerar_grafico_comparativo(meses, totais_mes, envios_map):
    taxas = [(m, totais_mes[m], envios_map[m], round(totais_mes[m] / envios_map[m] * 100, 2))
             for m in meses if m in envios_map and envios_map[m] > 0]

    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(10, 8), gridspec_kw={'height_ratios': [1, 2.5]})
    fig.patch.set_facecolor('#FFFFFF')
    ax_top.axis('off')

    # Cards dinâmicos de resumo
    n_cards = len(meses) * 2
    card_w = 0.96 / n_cards
    idx = 0
    for mes in meses:
        env = envios_map.get(mes, 0)
        tix = totais_mes.get(mes, 0)
        pct = round(tix / env * 100, 2) if env > 0 else 0
        
        # Card Envios
        x_env = 0.02 + idx * card_w
        ax_top.add_patch(mpatches.FancyBboxPatch((x_env, 0.1), card_w-0.01, 0.8, boxstyle="round,pad=0.02", facecolor='#F1EFE8', transform=ax_top.transAxes))
        ax_top.text(x_env + card_w/2, 0.7, f"Envios {mes}", ha='center', fontsize=8, color='#888780', transform=ax_top.transAxes)
        ax_top.text(x_env + card_w/2, 0.35, f"{env:,}".replace(',','.'), ha='center', fontsize=14, fontweight='bold', transform=ax_top.transAxes)
        idx += 1
        
        # Card Tickets
        x_tix = 0.02 + idx * card_w
        ax_top.add_patch(mpatches.FancyBboxPatch((x_tix, 0.1), card_w-0.01, 0.8, boxstyle="round,pad=0.02", facecolor='#F1EFE8', transform=ax_top.transAxes))
        ax_top.text(x_tix + card_w/2, 0.7, f"Tickets {mes}", ha='center', fontsize=8, color='#888780', transform=ax_top.transAxes)
        ax_top.text(x_tix + card_w*0.35, 0.35, str(tix), ha='center', fontsize=14, fontweight='bold', transform=ax_top.transAxes)
        ax_top.text(x_tix + card_w*0.75, 0.35, f"{pct:.2f}%".replace('.',','), ha='center', fontsize=10, color=AZUL_ESCURO, fontweight='bold', transform=ax_top.transAxes)
        idx += 1

    ax_bot.set_facecolor('#FFFFFF')
    xs = np.arange(len(taxas))
    for i, (mes, tix, env, pct) in enumerate(taxas):
        ax_bot.bar(xs[i], pct, width=0.5, color=CORES[i % 4], zorder=3)
        ax_bot.text(xs[i], pct + (max([t[3] for t in taxas])*0.02), f'{pct:.2f}%'.replace('.', ','), ha='center', va='bottom', fontsize=12, fontweight='bold')

    ax_bot.set_xticks(xs)
    ax_bot.set_xticklabels([t[0] for t in taxas], fontsize=12)
    for sp in ['top', 'right', 'left']: ax_bot.spines[sp].set_visible(False)
    ax_bot.yaxis.grid(True, color='#D3D1C7', linewidth=0.5)
    plt.tight_layout(pad=1.5)
    return fig

def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    return buf

# --- 7. LÓGICA PRINCIPAL ---
arquivo_excel = st.file_uploader("Selecione o arquivo Excel do Zendesk", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel)
        df = df[df.astype(str).ne('SUM').all(axis=1)].copy()

        col_ocorr = next((c for c in df.columns if 'Ocorr' in c or 'Motivo' in c), None)
        if not col_ocorr:
            st.error("Coluna 'Ocorrência' não encontrada no Excel.")
        else:
            df.rename(columns={col_ocorr: 'Ocorrência'}, inplace=True)
            envios_map = {item.split(':')[0].strip(): int(item.split(':')[1].strip()) for item in dados_envios.split(',') if ':' in item}
            meses = list(envios_map.keys())

            cols_tickets, totais_mes = {}, {}
            for mes in meses:
                col = next((c for c in df.columns if mes.lower() in c.lower() and ('Tickets' in c or 'Qtd' in c)), None)
                if col:
                    cols_tickets[mes] = col
                    totais_mes[mes] = int(df[col].sum())

            if not cols_tickets:
                st.warning("Não encontrei colunas para os meses digitados.")
            else:
                agg = df.groupby('Ocorrência')[list(cols_tickets.values())].sum().reset_index()
                agg.columns = ['Ocorrência'] + meses
                agg['Total'] = agg[meses].sum(axis=1)
                agg = agg.sort_values('Total', ascending=False)
                df_plot = agg[agg['Total'] >= min_tickets].copy()

                st.divider()
                c1, c2 = st.columns(2)
                with c1:
                    st.subheader("Ocorrências por Mês")
                    fig1 = gerar_grafico_ocorrencias(df_plot, meses, totais_mes)
                    st.pyplot(fig1)
                    st.download_button("⬇️ Baixar gráfico de ocorrências", fig_to_bytes(fig1), f"ocorrencias_{nome_cliente.lower()}.png", "image/png")

                with c2:
                    st.subheader("Taxa de Contato")
                    fig2 = gerar_grafico_comparativo(meses, totais_mes, envios_map)
                    st.pyplot(fig2)
                    st.download_button("⬇️ Baixar gráfico comparativo", fig_to_bytes(fig2), f"comparativo_{nome_cliente.lower()}.png", "image/png")

    except Exception as e:
        st.error(f"Erro ao processar arquivo: {e}")
