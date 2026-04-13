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

st.set_page_config(page_title="Dashboard CS — Mandaê", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


def logo_base64(path):
    with open(path, 'rb') as f:
        return base64.b64encode(f.read()).decode()


# Tenta carregar o logo (deve estar na mesma pasta que o app)
logo_path = Path(__file__).parent / 'logo_mandae.png'
logo_html = ""
if logo_path.exists():
    b64 = logo_base64(logo_path)
    logo_html = f'<img src="data:image/png;base64,{b64}" style="height:56px; margin-bottom:8px;">'

# Cabeçalho com logo
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

# Sidebar
with st.sidebar:
    if logo_path.exists():
        st.image(str(logo_path), width=160)
        st.markdown("<hr style='border:none;border-top:1px solid #D3D1C7;margin:8px 0 16px 0'>",
                    unsafe_allow_html=True)
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets  = st.number_input("Mínimo de Tickets", value=5, min_value=0)
    dados_envios = st.text_area("Dados de Envios (Mês:Valor)", "Fevereiro:1000, Março:1200")

arquivo_excel = st.file_uploader("Selecione o arquivo Excel do Zendesk", type=["xlsx"])

AZUL_ESCURO = '#185FA5'
AZUL_CLARO  = '#73B2E0'
CORES       = [AZUL_ESCURO, AZUL_CLARO, '#4A90E2', '#A1C9F4']


def gerar_grafico_ocorrencias(df_plot, meses, totais_mes):
    data = []
    for _, row in df_plot.iterrows():
        vals = [int(row[m]) for m in meses]
        data.append((row['Ocorrência'], *vals))
    data.sort(key=lambda x: sum(x[1:]), reverse=True)
    labels = [d[0] for d in data]
    valores = [[d[i+1] for d in data] for i in range(len(meses))]

    n = len(labels)
    y = np.arange(n)
    h = 0.8 / len(meses)
    altura = max(6, n * 0.55 + 1.5)

    fig, ax = plt.subplots(figsize=(14, altura))
    fig.patch.set_facecolor('#FFFFFF')
    ax.set_facecolor('#FFFFFF')

    for i, (mes, vals) in enumerate(zip(meses, valores)):
        offset = (i - (len(meses) - 1) / 2) * h
        bars = ax.barh(y + offset, vals, height=h, color=CORES[i % 4], label=mes, zorder=3)
        for bar, val in zip(bars, vals):
            if val > 0:
                ax.text(val + 0.15, bar.get_y() + bar.get_height() / 2,
                        str(val), va='center', ha='left', fontsize=9,
                        color='#444441', fontweight='500')

    ax.set_yticks(y)
    ax.set_yticklabels(labels, fontsize=10, color='#333333')
    ax.invert_yaxis()
    max_val = max(max(v) for v in valores) if valores else 1
    ax.set_xlim(0, max_val * 1.25 + 2)
    ax.set_xlabel('Quantidade de tickets', fontsize=11, color='#888780', labelpad=8)
    ax.tick_params(axis='x', colors='#888780', labelsize=10)
    ax.tick_params(axis='y', length=0)
    for sp in ['top', 'right', 'left']:
        ax.spines[sp].set_visible(False)
    ax.spines['bottom'].set_color('#D3D1C7')
    ax.xaxis.grid(True, color='#D3D1C7', linewidth=0.5, zorder=0)
    ax.set_axisbelow(True)
    ax.legend(
        handles=[mpatches.Patch(color=CORES[i % 4], label=m) for i, m in enumerate(meses)],
        fontsize=11, frameon=False, loc='lower right'
    )
    plt.tight_layout(pad=1.5)
    return fig


def gerar_grafico_comparativo(meses, totais_mes, envios_map):
    taxas = [(m, totais_mes[m], envios_map[m], round(totais_mes[m] / envios_map[m] * 100, 2))
             for m in meses if m in envios_map and envios_map[m] > 0]

    fig, (ax_top, ax_bot) = plt.subplots(2, 1, figsize=(10, 8),
                                          gridspec_kw={'height_ratios': [1, 2.5]})
    fig.patch.set_facecolor('#FFFFFF')
    ax_top.set_facecolor('#FFFFFF')
    ax_top.axis('off')

    cards = []
    for mes in meses:
        env = envios_map.get(mes, 0)
        tix = totais_mes.get(mes, 0)
        pct = round(tix / env * 100, 2) if env > 0 else 0
        cards.append((f'Envios — {mes.lower()}', f'{env:,}'.replace(',', '.'), None))
        cards.append((f'Tickets — {mes.lower()}', str(tix), f'{pct:.2f}%'.replace('.', ',')))

    n_cards = len(cards)
    card_w = 0.22 if n_cards == 4 else (0.96 / n_cards - 0.01)
    for i, (label, value, pct) in enumerate(cards):
        x = 0.02 + i * (card_w + 0.02)
        ax_top.add_patch(mpatches.FancyBboxPatch(
            (x, 0.05), card_w, 0.88,
            boxstyle='round,pad=0.02', linewidth=0,
            facecolor='#F1EFE8', transform=ax_top.transAxes))
        ax_top.text(x + card_w / 2, 0.72, label,
                    ha='center', va='center', fontsize=9, color='#888780',
                    transform=ax_top.transAxes)
        if pct:
            ax_top.text(x + card_w * 0.35, 0.28, value,
                        ha='center', va='center', fontsize=15,
                        fontweight='500', color='#2C2C2A', transform=ax_top.transAxes)
            c = AZUL_ESCURO if i % 4 == 1 else AZUL_CLARO
            ax_top.text(x + card_w * 0.75, 0.28, pct,
                        ha='center', va='center', fontsize=11,
                        fontweight='500', color=c, transform=ax_top.transAxes)
        else:
            ax_top.text(x + card_w / 2, 0.28, value,
                        ha='center', va='center', fontsize=15,
                        fontweight='500', color='#2C2C2A', transform=ax_top.transAxes)

    ax_bot.set_facecolor('#FFFFFF')
    xs = np.arange(len(taxas))
    max_pct = max(t[3] for t in taxas) if taxas else 1
    for i, (mes, tix, env, pct) in enumerate(taxas):
        ax_bot.bar(xs[i], pct, width=0.5, color=CORES[i % 4], zorder=3)
        ax_bot.text(xs[i], pct + max_pct * 0.03,
                    f'{pct:.2f}%'.replace('.', ','),
                    ha='center', va='bottom', fontsize=12,
                    fontweight='500', color='#444441')

    ax_bot.set_xticks(xs)
    ax_bot.set_xticklabels([t[0] for t in taxas], fontsize=12, color='#444441')
    ax_bot.set_xlim(-0.6, len(taxas) - 0.4)
    ax_bot.set_ylim(0, max_pct * 1.5)
    ax_bot.set_ylabel('Tickets / envios (%)', fontsize=11, color='#888780', labelpad=8)
    ax_bot.yaxis.set_major_formatter(matplotlib.ticker.FuncFormatter(lambda v, _: f'{v:.1f}%'))
    ax_bot.tick_params(axis='y', colors='#888780', labelsize=11)
    ax_bot.tick_params(axis='x', length=0)
    for sp in ['top', 'right', 'left']:
        ax_bot.spines[sp].set_visible(False)
    ax_bot.spines['bottom'].set_color('#D3D1C7')
    ax_bot.yaxis.grid(True, color='#D3D1C7', linewidth=0.5, zorder=0)
    ax_bot.set_axisbelow(True)
    plt.tight_layout(pad=1.5)
    return fig


def fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format='png', dpi=150, bbox_inches='tight', facecolor='white')
    buf.seek(0)
    return buf


if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel)
        df = df[df.astype(str).ne('SUM').all(axis=1)].copy()

        col_ocorr = next((c for c in df.columns if 'Ocorr' in c or 'Motivo' in c), None)
        if not col_ocorr:
            st.error("Coluna 'Ocorrência' não encontrada no Excel.")
        else:
            df.rename(columns={col_ocorr: 'Ocorrência'}, inplace=True)

            envios_map = {item.split(':')[0].strip(): int(item.split(':')[1].strip())
                          for item in dados_envios.split(',') if ':' in item}
            meses = list(envios_map.keys())

            cols_tickets = {}
            totais_mes   = {}
            for mes in meses:
                col = next((c for c in df.columns
                            if mes.lower() in c.lower() and ('Tickets' in c or 'Qtd' in c)), None)
                if col:
                    cols_tickets[mes] = col
                    totais_mes[mes]   = int(df[col].sum())

            if not cols_tickets:
                st.warning("Não encontrei colunas para os meses digitados.")
            else:
                agg = df.groupby('Ocorrência')[list(cols_tickets.values())].sum().reset_index()
                agg.columns = ['Ocorrência'] + list(cols_tickets.keys())
                agg['Total'] = agg[list(cols_tickets.keys())].sum(axis=1)
                agg = agg.sort_values('Total', ascending=False)
                df_plot = agg[agg['Total'] >= min_tickets].copy()

                col1, col2 = st.columns(2)

                with col1:
                    st.subheader("Ocorrências por Mês")
                    fig1 = gerar_grafico_ocorrencias(df_plot, list(cols_tickets.keys()), totais_mes)
                    st.pyplot(fig1)
                    st.download_button(
                        "⬇️ Baixar gráfico de ocorrências",
                        data=fig_to_bytes(fig1),
                        file_name=f"{nome_cliente.lower().replace(' ', '_')}_tickets_ocorrencia.png",
                        mime="image/png"
                    )

                with col2:
                    st.subheader("Taxa de Contato")
                    fig2 = gerar_grafico_comparativo(list(cols_tickets.keys()), totais_mes, envios_map)
                    st.pyplot(fig2)
                    st.download_button(
                        "⬇️ Baixar gráfico comparativo",
                        data=fig_to_bytes(fig2),
                        file_name=f"{nome_cliente.lower().replace(' ', '_')}_comparativo_envios.png",
                        mime="image/png"
                    )

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
