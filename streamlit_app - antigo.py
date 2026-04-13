import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import io
import warnings

# Configurações iniciais da página
st.set_page_config(page_title="Dashboard CS", layout="wide")
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

st.title("📊 Gerador de Relatórios CS")
st.markdown("Preencha os dados na lateral e faça o upload do arquivo.")

# Interface na Barra Lateral
with st.sidebar:
    st.header("⚙️ Configurações")
    nome_cliente = st.text_input("Nome do Cliente", "Cliente Exemplo")
    min_tickets = st.number_input("Mínimo de Tickets", value=5, min_value=0)
    dados_envios = st.text_area("Dados de Envios (Mês:Valor)", "Fevereiro:1000, Março:1200")

# Upload do Arquivo
arquivo_excel = st.file_uploader("Selecione o arquivo Excel do Zendesk", type=["xlsx"])

if arquivo_excel:
    try:
        df = pd.read_excel(arquivo_excel)
        df = df[df.astype(str).ne('SUM').all(axis=1)].copy()
        
        # Identificar coluna de Ocorrência
        col_ocorr = next((c for c in df.columns if 'Ocorr' in c or 'Motivo' in c), None)
        if not col_ocorr:
            st.error("Coluna 'Ocorrência' não encontrada no Excel.")
        else:
            df.rename(columns={col_ocorr: 'Ocorrência'}, inplace=True)
            
            # Processar envios
            envios_map = {item.split(':')[0].strip(): int(item.split(':')[1].strip()) 
                         for item in dados_envios.split(',') if ':' in item}
            
            meses = list(envios_map.keys())
            cols_tickets = {}
            totais_mes = {}
            
            for mes in meses:
                col = next((c for c in df.columns if mes.lower() in c.lower() and ('Tickets' in c or 'Qtd' in c)), None)
                if col:
                    cols_tickets[mes] = col
                    totais_mes[mes] = df[col].sum()

            if not cols_tickets:
                st.warning("Não encontrei colunas para os meses digitados.")
            else:
                # Processamento
                agg = df.groupby('Ocorrência')[list(cols_tickets.values())].sum().reset_index()
                agg.columns = ['Ocorrência'] + meses
                agg['Total'] = agg[meses].sum(axis=1)
                agg = agg.sort_values('Total', ascending=False)
                df_plot = agg[agg['Total'] >= min_tickets]

                # --- Exibição dos Gráficos ---
                col1, col2 = st.columns(2)
                cores_azul = ['#185FA5', '#4A90E2', '#73B2E0', '#A1C9F4']

                with col1:
                    st.subheader("Ocorrências por Mês")
                    fig1, ax1 = plt.subplots(figsize=(10, max(6, len(df_plot)*0.5)))
                    y_pos = np.arange(len(df_plot))
                    largura = 0.8 / len(meses)
                    for i, mes in enumerate(meses):
                        pos = y_pos + (i * largura) - (largura * (len(meses)-1)/2)
                        bars = ax1.barh(pos, df_plot[mes], height=largura, label=mes, color=cores_azul[i%4])
                        for bar in bars:
                            w = bar.get_width()
                            if w > 0:
                                p = (w / totais_mes[mes]) * 100
                                ax1.text(w + 0.1, bar.get_y()+bar.get_height()/2, f'{int(w)} ({p:.1f}%)', va='center', fontsize=8)
                    ax1.set_yticks(y_pos)
                    ax1.set_yticklabels(df_plot['Ocorrência'])
                    ax1.invert_yaxis()
                    ax1.legend()
                    st.pyplot(fig1)

                with col2:
                    st.subheader("Taxa de Contato")
                    taxas = []
                    for mes in meses:
                        tix = agg[mes].sum()
                        env = envios_map[mes]
                        taxas.append({'Mês': mes, 'Tickets': tix, 'Taxa': (tix/env)*100})
                    df_taxa = pd.DataFrame(taxas)
                    fig2, ax2 = plt.subplots()
                    bars2 = ax2.bar(df_taxa['Mês'], df_taxa['Taxa'], color=cores_azul[:len(meses)])
                    for i, b in enumerate(bars2):
                        h = b.get_height()
                        ax2.text(b.get_x()+b.get_width()/2, h + (h*0.02), f"{int(df_taxa.iloc[i]['Tickets'])} ({h:.2f}%)", ha='center', fontweight='bold')
                    ax2.set_ylim(0, df_taxa['Taxa'].max() * 1.3)
                    st.pyplot(fig2)

    except Exception as e:
        st.error(f"Erro ao processar: {e}")
