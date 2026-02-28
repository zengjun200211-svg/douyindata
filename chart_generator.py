import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import pandas as pd
import os

def setup_matplotlib():
    plt.rcParams['axes.unicode_minus'] = False
    
    chinese_fonts = ['Source Han Sans SC', 'Microsoft YaHei', 'SimHei', 'WenQuanYi Micro Hei']
    font_path = None
    for font in chinese_fonts:
        try:
            if os.name == 'nt':
                font_path = f"C:\\Windows\\Fonts\\{font}.ttf"
                if os.path.exists(font_path):
                    break
        except:
            continue
    
    if font_path and os.path.exists(font_path):
        prop = fm.FontProperties(fname=font_path)
        plt.rcParams['font.family'] = prop.get_name()
    else:
        plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial', 'Segoe UI', 'sans-serif']
    
    plt.rcParams['axes.facecolor'] = 'none'
    plt.rcParams['figure.facecolor'] = 'none'

setup_matplotlib()

def save_figure(fig, filename):
    fig.savefig(filename, dpi=300, bbox_inches='tight', transparent=True)
    plt.close(fig)

COLORS = {
    'primary': '#2A6DF4',
    'secondary': '#00C6A7',
    'warning': '#F5A623',
    'text_primary': '#333333',
    'text_secondary': '#666666',
    'bg_light': '#F5F7FA',
    'border': '#EEEEEE'
}

def get_color_variants(base_color, num_variants=5):
    variants = []
    import matplotlib.colors as mcolors
    rgb = mcolors.hex2color(base_color)
    for i in range(num_variants):
        factor = 0.6 + (i / num_variants) * 0.4
        variant = tuple(min(1.0, c * factor) for c in rgb)
        variants.append(mcolors.to_hex(variant))
    return variants

def create_overview_chart(df, output_dir):
    latest_fans = df.sort_values('日期').groupby('账号名称')['粉丝量'].last()
    
    fig, ax = plt.subplots(figsize=(10, 8))
    colors = get_color_variants(COLORS['primary'], len(latest_fans))
    
    wedges, texts, autotexts = ax.pie(
        latest_fans, 
        labels=latest_fans.index, 
        autopct='%1.1f%%',
        colors=colors,
        startangle=90,
        wedgeprops={'width': 0.4, 'edgecolor': 'white', 'linewidth': 1.5}
    )
    
    ax.set_title("各账号粉丝占比", fontsize=16, fontweight='bold', color=COLORS['text_primary'], pad=20)
    
    for text in texts:
        text.set_fontsize(12)
        text.set_color(COLORS['text_secondary'])
    
    for autotext in autotexts:
        autotext.set_fontsize(11)
        autotext.set_fontweight('bold')
    
    save_figure(fig, os.path.join(output_dir, "overview_pie.png"))

def create_account_detail_charts(df, account_name, output_dir):
    account_df = df[df['账号名称'] == account_name].sort_values('日期')
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(14, 10))
    
    ax1.plot(account_df['日期'], account_df['粉丝量'], color=COLORS['primary'], linewidth=3, marker='o', 
             markersize=8, markerfacecolor='white', markeredgewidth=2, markeredgecolor=COLORS['primary'])
    ax1.fill_between(account_df['日期'], account_df['粉丝量'], alpha=0.2, color=COLORS['primary'])
    
    peak_idx = account_df['粉丝量'].idxmax()
    ax1.scatter(account_df.loc[peak_idx, '日期'], account_df.loc[peak_idx, '粉丝量'], 
                color=COLORS['warning'], s=150, zorder=5, edgecolor='white', linewidth=2)
    ax1.text(account_df.loc[peak_idx, '日期'], account_df.loc[peak_idx, '粉丝量'], 
             f"峰值: {account_df.loc[peak_idx, '粉丝量']:,}", 
             ha='center', va='bottom', fontweight='bold', color=COLORS['warning'], fontsize=12)
    
    ax1.set_title(f"{account_name} - 粉丝趋势", fontsize=16, fontweight='bold', color=COLORS['text_primary'], pad=15)
    ax1.set_ylabel('粉丝量', fontsize=12, color=COLORS['text_secondary'])
    ax1.tick_params(axis='x', rotation=45, labelsize=11)
    ax1.tick_params(axis='y', labelsize=11)
    ax1.grid(True, linestyle='--', color=COLORS['border'], alpha=0.7)
    ax1.set_axisbelow(True)
    
    ax2.bar(account_df['日期'], account_df['互动数'], color=COLORS['primary'], alpha=0.8, edgecolor='white', linewidth=1,
            width=0.6)
    ax2.set_title(f"{account_name} - 每日互动", fontsize=16, fontweight='bold', color=COLORS['text_primary'], pad=15)
    ax2.set_ylabel('互动数', fontsize=12, color=COLORS['text_secondary'])
    ax2.tick_params(axis='x', rotation=45, labelsize=11)
    ax2.tick_params(axis='y', labelsize=11)
    ax2.grid(True, linestyle='--', color=COLORS['border'], alpha=0.7, axis='y')
    ax2.set_axisbelow(True)
    
    plt.tight_layout(pad=2.0)
    save_figure(fig, os.path.join(output_dir, f"detail_{account_name}.png"))

def create_top_posts_chart(df, output_dir):
    top_posts = df.sort_values('互动数', ascending=False).head(10)
    
    fig, ax = plt.subplots(figsize=(14, 7))
    colors = [COLORS['primary'] if i == 0 else COLORS['secondary'] if i < 3 else get_color_variants(COLORS['primary'])[2] for i in range(len(top_posts))]
    bars = ax.barh(range(len(top_posts)), top_posts['互动数'], color=colors, edgecolor='white', linewidth=1, height=0.7)
    
    ax.set_yticks(range(len(top_posts)))
    ax.set_yticklabels([title[:35] + '...' if len(title) > 35 else title for title in top_posts['作品标题']], fontsize=11)
    ax.invert_yaxis()
    ax.set_xlabel('互动数', fontsize=12, color=COLORS['text_secondary'])
    ax.set_title('Top 10 爆款作品（按互动量）', fontsize=18, fontweight='bold', color=COLORS['text_primary'], pad=20)
    
    for i, bar in enumerate(bars):
        width = bar.get_width()
        ax.text(width, bar.get_y() + bar.get_height()/2, 
                f'{int(width):,}', ha='left', va='center', fontweight='bold', fontsize=11)
    
    ax.grid(True, linestyle='--', color=COLORS['border'], alpha=0.7, axis='x')
    ax.set_axisbelow(True)
    
    plt.tight_layout()
    save_figure(fig, os.path.join(output_dir, "top_posts.png"))

def create_comparison_charts(df, output_dir):
    latest_data = df.sort_values('日期').groupby('账号名称').last().reset_index()
    metrics = ['涨粉量', '互动率', '播放量', '粉丝量']
    titles = ['各账号涨粉对比', '各账号互动率对比', '各账号播放量对比', '各账号粉丝总量对比']
    
    fig, axes = plt.subplots(2, 2, figsize=(18, 12))
    axes = axes.flatten()
    colors = get_color_variants(COLORS['primary'], len(latest_data))
    
    for i, (metric, title) in enumerate(zip(metrics, titles)):
        ax = axes[i]
        bars = ax.bar(latest_data['账号名称'], latest_data[metric], color=colors, edgecolor='white', linewidth=1, width=0.6)
        ax.set_title(title, fontsize=16, fontweight='bold', color=COLORS['text_primary'], pad=15)
        ax.tick_params(axis='x', rotation=45, labelsize=11)
        ax.tick_params(axis='y', labelsize=11)
        ax.set_xlabel('')
        ax.grid(True, linestyle='--', color=COLORS['border'], alpha=0.7, axis='y')
        ax.set_axisbelow(True)
        
        for bar in bars:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., height,
                    f'{int(height):,}' if metric != '互动率' else f'{height:.2%}',
                    ha='center', va='bottom', fontweight='bold', fontsize=10)
    
    plt.tight_layout(pad=2.0)
    save_figure(fig, os.path.join(output_dir, "comparison.png"))
