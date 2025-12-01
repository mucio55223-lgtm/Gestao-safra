from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import io 
import os 
import plotly.express as px
import plotly.graph_objects as go
import numpy as np 
from datetime import datetime, timedelta
import json
import random

app = Flask(__name__)
DATA_FILE = 'data.csv'
CONFIG_FILE = 'config.json'
TIMELINE_FILE = 'timeline_data.csv'

# --- Funções de Gerenciamento de Configuração ---

def load_config():
    """Carrega a configuração do dashboard do config.json, ou cria um arquivo padrão se não existir."""
    default_config = {
        "dashboardTitle": "Dashboard de Recepção de Trigo 🌾",
        "theme": {
            "primary": "#0d6efd", "success": "#198754", "warning": "#ffc107", "danger": "#dc3545"
        },
        "tabNames": {
            "dashboard-tab": "📊 Dashboard & Progresso Individual",
            "avancado-tab": "🚀 Gráficos Avançados",
            "analise-tab": "📈 Análise Detalhada",
            "tabela-tab": "📋 Edição Rápida da Tabela",
            "cadastro-tab": "➕ Cadastro & Opções",
            "slideshow-tab": "🎬 Slideshow"
        },
        "graphTitles": {
            "barGroupedTitle": "1. Previsto vs. Recebido (Barras Agrupadas)",
            "treemapTitle": "2. Proporção do Recebido (Gráfico de Pizza)",
            "sankeyTitle": "3. Fluxo de Volume (Sankey Diagram)"
        }
    }
    
    if not os.path.exists(CONFIG_FILE) or os.path.getsize(CONFIG_FILE) == 0:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=4)
        return default_config
    
    try:
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            for key in default_config:
                if key not in config:
                    config[key] = default_config[key]
            return config
    except Exception:
        return default_config

def save_config(config):
    """Salva a configuração atual no config.json."""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=4)

# --- Funções de Processamento de Dados ---

def load_and_process_data():
    """Carrega, limpa, calcula % e define o Status dos dados."""
    
    raw_cols = ['Cultivar', 'Prev_Colheita', 'Prev_Receb', 'Recepcao', 'Categoria', 'Status']
    
    if not os.path.exists(DATA_FILE) or os.path.getsize(DATA_FILE) == 0:
        df = pd.DataFrame(columns=raw_cols)
        df.to_csv(DATA_FILE, index=False)
    
    try:
        df = pd.read_csv(DATA_FILE)
    except Exception:
        df = pd.DataFrame(columns=raw_cols)

    for col in raw_cols:
        if col not in df.columns:
            if col in ['Prev_Colheita', 'Prev_Receb', 'Recepcao']:
                df[col] = 0.0
            else:
                df[col] = 'N/A'

    for col in ['Prev_Colheita', 'Prev_Receb', 'Recepcao']:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(float)
    
    df['%'] = (df['Recepcao'] / df['Prev_Receb']).where(df['Prev_Receb'] != 0, 0)
    df['%'] = (df['%'] * 100).round(2)

    # SÓ calcula o status automaticamente se a coluna Status não existir ou estiver vazia
    if 'Status' not in df.columns or df['Status'].isna().all():
        def get_status(row):
            if row['Prev_Receb'] == 0:
                return 'Prev. Inválida'
            if row['%'] >= 100: 
                return 'OK'
            elif row['Recepcao'] > 0:
                return 'Em Andamento'
            else:
                return 'Falta Receber'

        df['Status'] = df.apply(get_status, axis=1)
    else:
        # Preenche valores NaN na coluna Status
        df['Status'] = df['Status'].fillna('Em Andamento')
    
    return df

def save_dataframe(df):
    """Função auxiliar para salvar o DataFrame, garantindo que apenas as colunas brutas sejam salvas."""
    raw_cols = ['Cultivar', 'Prev_Colheita', 'Prev_Receb', 'Recepcao', 'Categoria', 'Status']
    df_to_save = df.reindex(columns=raw_cols).fillna(0)
    df_to_save.to_csv(DATA_FILE, index=False)

# --- Funções para Timeline ---

def load_timeline_data():
    """Carrega dados de timeline com datas de recebimento."""
    if not os.path.exists(TIMELINE_FILE):
        # Criar arquivo vazio se não existir
        empty_df = pd.DataFrame(columns=['Data', 'Cultivar', 'Categoria', 'Volume_SC'])
        empty_df.to_csv(TIMELINE_FILE, index=False)
        return empty_df
    
    try:
        df = pd.read_csv(TIMELINE_FILE)
        df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
        return df
    except Exception:
        return pd.DataFrame(columns=['Data', 'Cultivar', 'Categoria', 'Volume_SC'])

def save_timeline_data(df):
    """Salva dados de timeline."""
    df.to_csv(TIMELINE_FILE, index=False)

def process_timeline_data(raw_df):
    """
    Processa dados brutos de timeline no formato:
    Responsável pelo recebimento, Data, Turno, Cultivar, Quantidade (Kg)
    
    Converte para: Data, Cultivar, Categoria, Volume_SC
    """
    # Carregar dados principais para obter categorias
    main_df = load_and_process_data()
    
    # Mapear cultivares para categorias
    cultivar_to_category = dict(zip(main_df['Cultivar'], main_df['Categoria']))
    
    # Processar dados de timeline
    processed_data = []
    
    for _, row in raw_df.iterrows():
        cultivar = row['Cultivar']
        categoria = cultivar_to_category.get(cultivar, 'N/A')  # Usar categoria dos dados principais
        
        # Converter kg para SC (Sacos de Colheita) - assumindo 1 SC = 60 kg
        quantidade_kg = row['Quantidade (Kg)']
        volume_sc = quantidade_kg / 60.0  # Conversão para SC
        
        processed_data.append({
            'Data': row['Data'],
            'Cultivar': cultivar,
            'Categoria': categoria,
            'Volume_SC': volume_sc
        })
    
    return pd.DataFrame(processed_data)

def generate_sample_timeline_data():
    """Gera dados de exemplo para timeline baseado nos dados atuais."""
    main_df = load_and_process_data()
    timeline_df = load_timeline_data()
    
    if timeline_df.empty and not main_df.empty:
        # Gerar dados de exemplo baseados nos dados principais
        sample_data = []
        start_date = datetime.now() - timedelta(days=60)
        
        for _, row in main_df.iterrows():
            if row['Recepcao'] > 0:
                # Distribuir o volume recebido em várias datas
                total_received = row['Recepcao']
                remaining = total_received
                current_date = start_date
                
                while remaining > 0 and current_date <= datetime.now():
                    # Volume aleatório para cada dia (entre 100 e 1000 SC, ou o que restar)
                    daily_volume = min(random.randint(100, 1000), remaining)
                    if daily_volume > 0:
                        sample_data.append({
                            'Data': current_date.strftime('%Y-%m-%d'),
                            'Cultivar': row['Cultivar'],
                            'Categoria': row['Categoria'],
                            'Volume_SC': daily_volume
                        })
                        remaining -= daily_volume
                    current_date += timedelta(days=random.randint(1, 3))
        
        if sample_data:
            sample_df = pd.DataFrame(sample_data)
            save_timeline_data(sample_df)
            return sample_df
    
    return timeline_df

# --- Funções Auxiliares Plotly ---

def get_plotly_colors():
    """Retorna as cores primárias do tema para uso no Plotly, lendo do config.json."""
    config = load_config()
    return {
        'primary': config['theme']['primary'],
        'success': config['theme']['success'],
        'warning': config['theme']['warning'],
        'danger': config['theme']['danger'],
        'text_color': 'black',
        'bg_color': 'white'
    }

# --- FUNÇÕES DE CRIAÇÃO DE GRÁFICOS (para uso no relatório) ---

def create_bar_grouped_graph(df):
    """Cria o gráfico de Barras Agrupadas do Plotly, agora por Cultivar."""
    colors = get_plotly_colors()
    config = load_config()
    
    if df.empty or df['Prev_Receb'].sum() == 0:
        return "<div class='text-center p-5 text-muted'>Adicione dados com previsão para ver o Gráfico de Barras Agrupadas.</div>"

    summary_df = df.groupby('Cultivar').agg(
        Total_Previsto=('Prev_Receb', 'sum'),
        Total_Recebido=('Recepcao', 'sum')
    ).reset_index()

    summary_df = summary_df[summary_df['Total_Previsto'] > 0]
    
    fig = go.Figure(data=[
        go.Bar(
            name='Previsto (SC)', 
            x=summary_df['Cultivar'], 
            y=summary_df['Total_Previsto'], 
            marker_color=colors['primary'],
            hovertemplate = 'Cultivar: %{x}<br>Previsto: %{y:,.0f} SC<extra></extra>'
        ),
        go.Bar(
            name='Recebido (SC)', 
            x=summary_df['Cultivar'], 
            y=summary_df['Total_Recebido'], 
            marker_color=colors['warning'],
            hovertemplate = 'Cultivar: %{x}<br>Recebido: %{y:,.0f} SC<extra></extra>'
        )
    ])

    fig.update_layout(
        title=config['graphTitles']['barGroupedTitle'],
        xaxis_title='Cultivar',
        yaxis_title='Volume (Sacos de Colheita - SC)',
        barmode='group',
        margin=dict(t=40, l=0, r=0, b=100),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color'])
    )
    return fig

def create_pie_chart_recepcao(df):
    """Cria um gráfico de pizza interativo mostrando a proporção do volume recebido por Cultivar."""
    colors = get_plotly_colors()
    config = load_config()
    
    df_recebido = df[df['Recepcao'] > 0].groupby('Cultivar')['Recepcao'].sum().reset_index()

    if df_recebido.empty:
        return "<div class='text-center p-5 text-muted'>Adicione dados com recepção para ver o Gráfico de Pizza.</div>"

    fig = px.pie(
        df_recebido, 
        values='Recepcao', 
        names='Cultivar', 
        title=config['graphTitles']['treemapTitle'],
        hole=0.3,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    
    fig.update_traces(
        textinfo='percent', 
        hovertemplate="<b>%{label}</b><br>Recebido: %{value:,.0f} SC<br>Proporção: %{percent}<extra></extra>"
    )
    
    fig.update_layout(
        margin=dict(t=40, l=0, r=0, b=0),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color'])
    )
    
    return fig

def create_sankey_graph(df):
    """Cria o gráfico Sankey do Plotly."""
    config = load_config()
    
    if df.empty or df['Prev_Receb'].sum() == 0:
         return "<div class='text-center p-5 text-muted'>Adicione dados com previsão para ver o Sankey Diagram (Cultivar -> Status).</div>"

    sankey_df = df[df['Prev_Receb'] > 0].copy()
    if sankey_df.empty:
        return "<div class='text-center p-5 text-muted'>Nenhum dado com Prev. Receb. válido para o Sankey Diagram.</div>"

    cultivares = sankey_df['Cultivar'].unique().tolist()
    statuses = sankey_df['Status'].unique().tolist()
    
    nodes = cultivares + statuses
    node_map = {name: i for i, name in enumerate(nodes)}
    
    sources = []
    targets = []
    values = []
    
    for _, row in sankey_df.iterrows():
        source_index = node_map[row['Cultivar']]
        target_index = node_map[row['Status']]
        flow_value = row['Prev_Receb']
        
        sources.append(source_index)
        targets.append(target_index)
        values.append(flow_value)
    
    node_colors = ['#1f77b4'] * len(cultivares)
    
    colors_map = {
        'OK': get_plotly_colors()['success'],
        'Em Andamento': get_plotly_colors()['warning'],
        'Falta Receber': get_plotly_colors()['danger'],
        'Prev. Inválida': '#6c757d'
    }

    for status in statuses:
        node_colors.append(colors_map.get(status, '#6c757d'))
            
    fig = go.Figure(data=[go.Sankey(
        node = dict(
          pad = 15,
          thickness = 20,
          line = dict(color = "black", width = 0.5),
          label = nodes,
          color = node_colors
        ),
        link = dict(
          source = sources,
          target = targets,
          value = values
        ))])
    
    fig.update_layout(
        title_text=config['graphTitles']['sankeyTitle'],
        font_size=10,
        margin=dict(t=40, l=0, r=0, b=0),
        font=dict(color=get_plotly_colors()['text_color'])
    )
    return fig

def create_diff_bar_graph(df):
    """Gráfico de barras mostrando a diferença entre Recebido e Previsto."""
    colors = get_plotly_colors()
    
    if df.empty:
        return "<div class='text-center p-5 text-muted'>Adicione dados para ver o gráfico de diferenças.</div>"

    # Calcular diferenças
    df['Diferenca'] = df['Recepcao'] - df['Prev_Receb']
    df = df[df['Prev_Receb'] > 0]  # Filtra apenas com previsão válida
    
    # Ordenar por diferença
    df = df.sort_values('Diferenca', ascending=False)
    
    # Criar cores baseadas na diferença
    bar_colors = []
    for diff in df['Diferenca']:
        if diff > 0:
            bar_colors.append(colors['success'])
        elif diff < 0:
            bar_colors.append(colors['danger'])
        else:
            bar_colors.append(colors['warning'])
    
    fig = go.Figure(data=[
        go.Bar(
            name='Diferença (SC)',
            x=df['Cultivar'],
            y=df['Diferenca'],
            marker_color=bar_colors,
            hovertemplate='Cultivar: %{x}<br>Diferença: %{y:,.0f} SC<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="Desempenho vs. Projeção por Cultivar",
        xaxis_title="Cultivar",
        yaxis_title="Diferença (Recebido - Previsto) SC",
        showlegend=False,
        margin=dict(t=40, l=0, r=0, b=100),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color'])
    )
    
    return fig

def create_percent_bar_graph(df):
    """Gráfico de barras ordenado por percentual de atingimento."""
    colors = get_plotly_colors()
    
    if df.empty:
        return "<div class='text-center p-5 text-muted'>Adicione dados para ver o gráfico de percentuais.</div>"

    df = df[df['Prev_Receb'] > 0]  # Filtra apenas com previsão válida
    df = df.sort_values('%', ascending=True)  # Ordena do menor para o maior
    
    # Criar cores baseadas no percentual
    bar_colors = []
    for percent in df['%']:
        if percent >= 100:
            bar_colors.append(colors['success'])
        elif percent >= 50:
            bar_colors.append(colors['warning'])
        else:
            bar_colors.append(colors['danger'])
    
    fig = go.Figure(data=[
        go.Bar(
            name='% de Atingimento',
            x=df['%'],
            y=df['Cultivar'],
            orientation='h',
            marker_color=bar_colors,
            hovertemplate='Cultivar: %{y}<br>Percentual: %{x:.1f}%<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="Percentual de Atingimento por Cultivar",
        xaxis_title="Percentual (%)",
        yaxis_title="Cultivar",
        showlegend=False,
        margin=dict(t=40, l=0, r=0, b=0),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color'])
    )
    
    # Adicionar linha vertical em 100%
    fig.add_vline(x=100, line_dash="dash", line_color=colors['success'])
    
    return fig

def create_category_stacked_graph(df):
    """Gráfico de barras empilhadas por categoria."""
    colors = get_plotly_colors()
    
    if df.empty:
        return "<div class='text-center p-5 text-muted'>Adicione dados para ver o gráfico por categoria.</div>"

    # Agrupar por categoria
    category_summary = df.groupby('Categoria').agg({
        'Prev_Receb': 'sum',
        'Recepcao': 'sum'
    }).reset_index()
    
    fig = go.Figure(data=[
        go.Bar(
            name='Previsto (SC)',
            x=category_summary['Categoria'],
            y=category_summary['Prev_Receb'],
            marker_color=colors['primary'],
            hovertemplate='Categoria: %{x}<br>Previsto: %{y:,.0f} SC<extra></extra>'
        ),
        go.Bar(
            name='Recebido (SC)',
            x=category_summary['Categoria'],
            y=category_summary['Recepcao'],
            marker_color=colors['warning'],
            hovertemplate='Categoria: %{x}<br>Recebido: %{y:,.0f} SC<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="Volume por Categoria (Previsto vs. Recebido)",
        xaxis_title="Categoria",
        yaxis_title="Volume (SC)",
        barmode='group',
        margin=dict(t=40, l=0, r=0, b=0),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color'])
    )
    
    return fig

def create_top_cultivars_graph(df):
    """Top 10 cultivares por volume recebido."""
    colors = get_plotly_colors()
    
    if df.empty:
        return "<div class='text-center p-5 text-muted'>Adicione dados para ver o ranking.</div>"

    # Ordenar por volume recebido e pegar top 10
    top_cultivars = df.nlargest(10, 'Recepcao')
    
    fig = go.Figure(data=[
        go.Bar(
            name='Volume Recebido (SC)',
            x=top_cultivars['Recepcao'],
            y=top_cultivars['Cultivar'],
            orientation='h',
            marker_color=colors['primary'],
            hovertemplate='Cultivar: %{y}<br>Recebido: %{x:,.0f} SC<extra></extra>'
        )
    ])
    
    fig.update_layout(
        title="Top 10 Cultivares por Volume Recebido",
        xaxis_title="Volume Recebido (SC)",
        yaxis_title="Cultivar",
        showlegend=False,
        margin=dict(t=40, l=0, r=0, b=0),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color'])
    )
    
    return fig

def create_timeline_graph():
    """Gráfico de linha do tempo com dados reais de recebimento."""
    colors = get_plotly_colors()
    
    # Carregar dados de timeline
    timeline_df = generate_sample_timeline_data()
    
    if timeline_df.empty:
        return "<div class='text-center p-5 text-muted'>Nenhum dado de timeline disponível. Use a exportação de dados para importar dados de recebimento por data.</div>"

    # Processar dados para timeline acumulada
    timeline_df = timeline_df.sort_values('Data')
    timeline_df['Data'] = pd.to_datetime(timeline_df['Data'])
    
    # Criar dados acumulados por categoria
    categories = timeline_df['Categoria'].unique()
    
    fig = go.Figure()
    
    for category in categories:
        category_data = timeline_df[timeline_df['Categoria'] == category]
        
        # Agrupar por data e calcular acumulado
        daily_data = category_data.groupby('Data').agg({'Volume_SC': 'sum'}).reset_index()
        daily_data = daily_data.sort_values('Data')
        daily_data['Acumulado'] = daily_data['Volume_SC'].cumsum()
        
        fig.add_trace(go.Scatter(
            x=daily_data['Data'],
            y=daily_data['Acumulado'],
            mode='lines+markers',
            name=f'Categoria {category}',
            hovertemplate='Data: %{x|%d/%m/%Y}<br>Acumulado: %{y:,.0f} SC<extra></extra>',
            line=dict(width=3)
        ))
    
    fig.update_layout(
        title="Evolução do Recebimento por Categoria",
        xaxis_title="Data",
        yaxis_title="Volume Acumulado (SC)",
        margin=dict(t=40, l=0, r=0, b=0),
        plot_bgcolor=colors['bg_color'],
        paper_bgcolor=colors['bg_color'],
        font=dict(color=colors['text_color']),
        hovermode='x unified'
    )
    
    return fig

# --- ROTAS PARA GRÁFICOS INDIVIDUAIS (para o dashboard) ---

@app.route('/graph/bar_grouped')
def graph_bar_grouped():
    df = load_and_process_data()
    fig = create_bar_grouped_graph(df)
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/treemap')
def graph_treemap():
    df = load_and_process_data()
    fig = create_pie_chart_recepcao(df) 
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/sankey')
def graph_sankey():
    df = load_and_process_data()
    fig = create_sankey_graph(df)
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/diff_bar')
def graph_diff_bar():
    df = load_and_process_data()
    fig = create_diff_bar_graph(df)
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/percent_bar')
def graph_percent_bar():
    df = load_and_process_data()
    fig = create_percent_bar_graph(df)
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/category_stacked')
def graph_category_stacked():
    df = load_and_process_data()
    fig = create_category_stacked_graph(df)
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/top_cultivars')
def graph_top_cultivars():
    df = load_and_process_data()
    fig = create_top_cultivars_graph(df)
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

@app.route('/graph/timeline')
def graph_timeline():
    fig = create_timeline_graph()
    if isinstance(fig, str): return fig
    return fig.to_html(full_html=False, include_plotlyjs='cdn')

# --- NOVAS ROTAS PARA SLIDESHOW ---

@app.route('/slideshow')
def slideshow():
    """Página de slideshow com todos os gráficos"""
    config = load_config()
    return render_template('slideshow.html', config=config)

@app.route('/slideshow_fullscreen')
def slideshow_fullscreen():
    """Página de slideshow apenas para exibição (sem controles)"""
    config = load_config()
    return render_template('slideshow_fullscreen.html', config=config)

@app.route('/get_slideshow_url')
def get_slideshow_url():
    """Retorna a URL do slideshow para transmissão"""
    base_url = request.host_url.rstrip('/')
    return jsonify({
        'fullscreen_url': f"{base_url}/slideshow_fullscreen",
        'control_url': f"{base_url}/slideshow"
    })

# --- NOVAS ROTAS PARA TIMELINE ---

@app.route('/export_timeline_template')
def export_timeline_template():
    """Exporta template para importação de dados de timeline no formato especificado."""
    template_df = pd.DataFrame(columns=['Responsável pelo recebimento', 'Data', 'Turno', 'Cultivar', 'Quantidade (Kg)'])
    
    # Adicionar exemplos
    example_data = {
        'Responsável pelo recebimento': ['João Silva', 'Maria Santos', 'Pedro Oliveira'],
        'Data': ['2024-01-15', '2024-01-16', '2024-01-17'],
        'Turno': ['Manhã', 'Tarde', 'Noite'],
        'Cultivar': ['TRUNFO', 'TRUNFO', 'VELOZ'],
        'Quantidade (Kg)': [15000, 12000, 8000]
    }
    example_df = pd.DataFrame(example_data)
    template_df = pd.concat([template_df, example_df], ignore_index=True)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        template_df.to_excel(writer, index=False, sheet_name='Template Timeline')
        
        # Adicionar instruções
        workbook = writer.book
        worksheet = writer.sheets['Template Timeline']
        
        instructions = [
            "INSTRUÇÕES PARA TIMELINE:",
            "1. Preencha os dados de recebimento diário",
            "2. Responsável: Nome do responsável pelo recebimento",
            "3. Data: Data do recebimento (formato: AAAA-MM-DD)",
            "4. Turno: Manhã, Tarde ou Noite",
            "5. Cultivar: Nome da cultivar (deve existir nos dados principais)",
            "6. Quantidade (Kg): Volume recebido em quilogramas",
            "7. Mantenha o cabeçalho original",
            "8. O sistema converterá automaticamente Kg para SC (1 SC = 60 kg)",
            "9. Salve e importe no sistema"
        ]
        
        for i, instruction in enumerate(instructions, start=len(template_df) + 3):
            worksheet.write(f'A{i}', instruction)
    
    output.seek(0)
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='template_timeline_recebimento.xlsx'
    )

@app.route('/import_timeline_data', methods=['POST'])
def import_timeline_data():
    """Importa dados de timeline no formato especificado."""
    if 'timeline_file' not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400
    
    file = request.files['timeline_file']
    
    if file.filename == '' or not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"error": "Formato de arquivo inválido. Use .xlsx ou .xls"}), 400
    
    try:
        imported_df = pd.read_excel(file)
        required_cols = ['Responsável pelo recebimento', 'Data', 'Turno', 'Cultivar', 'Quantidade (Kg)']
        
        if not all(col in imported_df.columns for col in required_cols):
            return jsonify({"error": f"O arquivo deve ter as colunas: {', '.join(required_cols)}"}), 400
        
        # Validar dados
        imported_df['Data'] = pd.to_datetime(imported_df['Data'], errors='coerce')
        if imported_df['Data'].isna().any():
            return jsonify({"error": "Datas inválidas no arquivo. Use formato AAAA-MM-DD"}), 400
        
        imported_df['Quantidade (Kg)'] = pd.to_numeric(imported_df['Quantidade (Kg)'], errors='coerce')
        if imported_df['Quantidade (Kg)'].isna().any():
            return jsonify({"error": "Valores de quantidade inválidos no arquivo"}), 400
        
        # Processar dados (converter formato)
        processed_df = process_timeline_data(imported_df)
        
        # Salvar dados processados
        save_timeline_data(processed_df)
        
        return jsonify({
            "message": f"Dados de timeline importados com sucesso! {len(processed_df)} registros processados.",
            "records_processed": len(processed_df)
        }), 200
    
    except Exception as e:
        return jsonify({"error": f"Erro ao processar o arquivo: {str(e)}"}), 500

# --- ROTA DE EXPORTAÇÃO DE RELATÓRIO ---

@app.route('/export_relatorio', methods=['POST'])
def export_relatorio():
    df = load_and_process_data()
    
    custom_titles = request.json.get('titles', {})
    
    # Definir títulos para todos os gráficos
    bar_title = custom_titles.get('barGroupedTitle', 'Volume Previsto vs. Recebido por Cultivar')
    pie_title = custom_titles.get('treemapTitle', 'Proporção do Volume Recebido (SC) por Cultivar')
    sankey_title = custom_titles.get('sankeyTitle', 'Fluxo do Volume Previsto (SC) por Cultivar para o Status Final')
    diff_title = "Desempenho vs. Projeção por Cultivar"
    percent_title = "Percentual de Atingimento por Cultivar"
    category_title = "Volume por Categoria (Previsto vs. Recebido)"
    top_title = "Top 10 Cultivares por Volume Recebido"
    timeline_title = "Evolução do Recebimento por Categoria"
    
    # Calculo dos KPIs
    total_previsto = df['Prev_Receb'].sum()
    total_recebido = df['Recepcao'].sum()
    progresso_geral = (total_recebido / total_previsto * 100).round(2) if total_previsto > 0 else 0
    num_cultivares = len(df)
    
    # Calcular KPIs avançados
    total_excedente = df[df['Recepcao'] > df['Prev_Receb']]['Recepcao'].sum() - df[df['Recepcao'] > df['Prev_Receb']]['Prev_Receb'].sum()
    total_deficit = df[df['Recepcao'] < df['Prev_Receb']]['Prev_Receb'].sum() - df[df['Recepcao'] < df['Prev_Receb']]['Recepcao'].sum()
    cultivares_excedente = len(df[df['Recepcao'] > df['Prev_Receb']])
    
    number_format = lambda x: "{:,.2f}".format(x).replace('.', 'TEMP').replace(',', '.').replace('TEMP', ',')
    
    kpis = {
        'total_previsto': number_format(total_previsto) + ' SC',
        'total_recebido': number_format(total_recebido) + ' SC',
        'progresso_geral': f'{progresso_geral}%',
        'num_cultivares': num_cultivares,
        'total_excedente': number_format(total_excedente) + ' SC',
        'total_deficit': number_format(total_deficit) + ' SC',
        'cultivares_excedente': cultivares_excedente
    }
    
    # Geração de TODOS os gráficos para o relatório
    try:
        bar_fig = create_bar_grouped_graph(df)
        bar_html = bar_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(bar_fig, str) else f"<div>{bar_fig}</div>"
    except Exception as e:
        bar_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Gráfico de Barras: {str(e)}</div>"

    try:
        pie_fig = create_pie_chart_recepcao(df)
        pie_html = pie_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(pie_fig, str) else f"<div>{pie_fig}</div>"
    except Exception as e:
        pie_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Gráfico de Pizza: {str(e)}</div>"

    try:
        sankey_fig = create_sankey_graph(df)
        sankey_html = sankey_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(sankey_fig, str) else f"<div>{sankey_fig}</div>"
    except Exception as e:
        sankey_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Sankey: {str(e)}</div>"

    try:
        diff_fig = create_diff_bar_graph(df)
        diff_html = diff_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(diff_fig, str) else f"<div>{diff_fig}</div>"
    except Exception as e:
        diff_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Gráfico de Diferenças: {str(e)}</div>"

    try:
        percent_fig = create_percent_bar_graph(df)
        percent_html = percent_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(percent_fig, str) else f"<div>{percent_fig}</div>"
    except Exception as e:
        percent_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Gráfico de Percentuais: {str(e)}</div>"

    try:
        category_fig = create_category_stacked_graph(df)
        category_html = category_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(category_fig, str) else f"<div>{category_fig}</div>"
    except Exception as e:
        category_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Gráfico por Categoria: {str(e)}</div>"

    try:
        top_fig = create_top_cultivars_graph(df)
        top_html = top_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(top_fig, str) else f"<div>{top_fig}</div>"
    except Exception as e:
        top_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Ranking: {str(e)}</div>"

    try:
        timeline_fig = create_timeline_graph()
        timeline_html = timeline_fig.to_html(full_html=False, include_plotlyjs='cdn', config={'displayModeBar': False}) if not isinstance(timeline_fig, str) else f"<div>{timeline_fig}</div>"
    except Exception as e:
        timeline_html = f"<div class='text-center p-5 text-muted'>Erro ao gerar Timeline: {str(e)}</div>"
        
    # Geração da Tabela em HTML
    df_export = df.copy()
    df_export['Prev. Colheita'] = df_export['Prev_Colheita'].apply(lambda x: number_format(x))
    df_export['Prev. Receb'] = df_export['Prev_Receb'].apply(lambda x: number_format(x))
    df_export['Recepção (SC)'] = df_export['Recepcao'].apply(lambda x: number_format(x))
    df_export['%'] = df_export['%'].apply(lambda x: f"{x}%")
    df_export['Diferença'] = (df_export['Recepcao'] - df_export['Prev_Receb']).apply(lambda x: number_format(x))
    
    def get_status_class(status):
        return 'status-' + status.lower().replace(' ', '-').replace('.', '')
    
    table_rows = ""
    for index, row in df_export.iterrows():
        status_class = get_status_class(row['Status'])
        diff_class = 'positive-diff' if row['Recepcao'] > row['Prev_Receb'] else 'negative-diff' if row['Recepcao'] < row['Prev_Receb'] else 'neutral-diff'
        
        row_html = f"<tr>"
        for col in ['Cultivar', 'Categoria', 'Prev. Colheita', 'Prev. Receb', 'Recepção (SC)', '%']:
             row_html += f"<td>{row[col]}</td>"
        row_html += f"<td class='{status_class}'>{row['Status']}</td>"
        row_html += f"<td class='{diff_class}'>{row['Diferença']} SC</td>"
        row_html += "</tr>"
        table_rows += row_html
    
    # Renderiza o Template de Relatório com TODOS os gráficos
    report_html = render_template(
        'report_template.html', 
        kpis=kpis,
        table_rows=table_rows, 
        # Gráficos principais
        bar_chart_html=bar_html,
        treemap_chart_html=pie_html, 
        sankey_chart_html=sankey_html,
        # Novos gráficos
        diff_chart_html=diff_html,
        percent_chart_html=percent_html,
        category_chart_html=category_html,
        top_chart_html=top_html,
        timeline_chart_html=timeline_html,
        now=datetime.now(),
        # Títulos
        bar_title=bar_title,
        pie_title=pie_title,
        sankey_title=sankey_title,
        diff_title=diff_title,
        percent_title=percent_title,
        category_title=category_title,
        top_title=top_title,
        timeline_title=timeline_title
    )

    return send_file(
        io.BytesIO(report_html.encode('utf-8')),
        mimetype='text/html',
        as_attachment=True,
        download_name='relatorio_recepcao_trigo.html'
    )

# --- ROTAS PRINCIPAIS ---

@app.route('/')
def index():
    config = load_config()
    return render_template('index.html', config=config)

@app.route('/update_config', methods=['POST'])
def update_config():
    """Rota para o JavaScript salvar configurações no config.json do servidor."""
    try:
        new_settings = request.json
        current_config = load_config()

        if 'dashboardTitle' in new_settings:
            current_config['dashboardTitle'] = new_settings['dashboardTitle']
        if 'tabNames' in new_settings and isinstance(new_settings['tabNames'], dict):
            current_config['tabNames'].update(new_settings['tabNames'])
        if 'graphTitles' in new_settings and isinstance(new_settings['graphTitles'], dict):
            current_config['graphTitles'].update(new_settings['graphTitles'])
        if 'theme' in new_settings and isinstance(new_settings['theme'], dict):
            current_config['theme'].update(new_settings['theme'])
            
        save_config(current_config)
        return jsonify({"message": "Configurações salvas no servidor!"}), 200
        
    except Exception as e:
        return jsonify({"error": f"Erro ao salvar configurações no servidor: {str(e)}"}), 500

# --- Rotas Existentes ---

@app.route('/data')
def get_data():
    df = load_and_process_data()
    data_json = df.to_dict(orient='records')
    return jsonify(data_json)

@app.route('/update_data', methods=['POST'])
def update_data():
    data = request.json
    new_df = pd.DataFrame(data)
    for col in ['Prev_Colheita', 'Prev_Receb', 'Recepcao']:
        new_df[col] = pd.to_numeric(new_df[col], errors='coerce').fillna(0).astype(float)
        
    save_dataframe(new_df)
    
    df_updated = load_and_process_data()
    return jsonify(df_updated.to_dict(orient='records'))

@app.route('/register', methods=['POST'])
def register_cultivar():
    try:
        new_cultivar_data = {
            'Cultivar': request.form.get('cultivar'),
            'Categoria': request.form.get('categoria'),
            'Prev_Colheita': float(request.form.get('prev_colheita', 0)),
            'Prev_Receb': float(request.form.get('prev_receb', 0)),
            'Recepcao': float(request.form.get('recepcao', 0)),
            'Status': 'Em Andamento'  # Status padrão para novos registros
        }

        df = load_and_process_data()
        new_df_row = pd.DataFrame([new_cultivar_data], columns=df.columns)
        df = pd.concat([df, new_df_row], ignore_index=True)
        save_dataframe(df)
        return jsonify({"message": "Cultivar cadastrada com sucesso!"}), 200
    except Exception as e:
        return jsonify({"error": f"Erro ao cadastrar: {str(e)}"}), 500

@app.route('/export_excel')
def export_excel():
    df = load_and_process_data()
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Dados Trigo')
    output.seek(0)
    return send_file(output, 
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True,
                     download_name='dados_trigo_exportado.xlsx')

@app.route('/import_data', methods=['POST'])
def import_data():
    if 'excel_file' not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400
    
    file = request.files['excel_file']
    
    if file.filename == '' or not file.filename.endswith(('.xlsx', '.xls')):
        return jsonify({"error": "Formato de arquivo inválido. Use .xlsx ou .xls"}), 400
    
    try:
        imported_df = pd.read_excel(file)
        required_cols = ['Cultivar', 'Prev_Colheita', 'Prev_Receb', 'Recepcao', 'Categoria']
        
        if not all(col in imported_df.columns for col in required_cols):
             return jsonify({"error": f"O arquivo Excel deve ter as colunas: {', '.join(required_cols)}"}), 400
        
        # Adicionar coluna Status se não existir
        if 'Status' not in imported_df.columns:
            imported_df['Status'] = 'Em Andamento'
        
        save_dataframe(imported_df)
        return jsonify({"message": "Dados importados e compilados com sucesso!"}), 200
    
    except Exception as e:
        return jsonify({"error": f"Erro ao processar o arquivo: {str(e)}"}), 500

if __name__ == '__main__':
    load_and_process_data() 
    app.run(debug=True, host='0.0.0.0')