import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
import pandas as pd
import os

def setup_matplotlib():
    plt.rcParams['axes.unicode_minus'] = False
    
    chinese_fonts = ['Microsoft YaHei', 'SimHei', 'WenQuanYi Micro Hei']
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
        plt.rcParams['font.family'] = ['DejaVu Sans', 'Arial', 'sans-serif']
    
    sns.set_style("whitegrid")
    sns.set_palette(sns.color_palette(["#0077B6", "#00B4D8", "#90E0EF", "#CAF0F8", "#48CAE4"]))

setup_matplotlib()

def save_figure(fig, filename):
    fig.savefig(filename, dpi=300, bbox_inches='tight')
    plt.close(fig)

def create_overview_chart(df, output_dir):
    latest_fans = df.sort_values('日期').groupby('账号名称')['粉丝量'].last()
    fig, ax = plt.subplots(figsize=(8, 8))
    colors = ["#0077B6", "#00B4D8", "#90E0EF", "#CAF0F8", "#48CAE4", "#03045E"]
    wedges, texts, autotexts = ax.pie(latest_fans, labels=latest_fans.index, autopct='%1.1f%%', 
                                         colors=colors, startangle=90)
    ax.set_title("各账号粉丝占比", fontsize=16, fontweight='bold')
    plt.setp(texts, fontsize=10)
    plt.setp(autotexts, fontsize=9, weight='bold')
    save_figure(fig, os.path.join(output_dir, "overview_pie.png"))

def create_account_detail_charts(df, account_name, output_dir):
    account_df = df[df['账号名称'] == account_name].sort_values('日期')
    
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
    
    ax1.plot(account_df['日期'], account_df['粉丝量'], color='#0077B6', linewidth=2)
    peak_idx = account_df['粉丝量'].idxmax()
    ax1.scatter(account_df.loc[peak_idx, '日期'], account_df.loc[peak_idx, '粉丝量'], 
                color='#FF0000', s=100, zorder=5)
    ax1.text(account_df.loc[peak_idx, '日期'], account_df.loc[peak_idx, '粉丝量'], 
             f"峰值: {account_df.loc[peak_idx, '粉丝量']:,}", 
             ha='center', va='bottom', fontweight='bold')
    ax1.set_title(f"{account_name} - 粉丝趋势", fontsize=14, fontweight='bold')
    ax1.set_ylabel('粉丝量')
    ax1.tick_params(axis='x', rotation=45)
    
    ax2.plot(account_df['日期'], account_df['互动数'], color='#00B4D8', linewidth=2)
    ax2.set_title(f"{account_name} - 每日互动", fontsize=14, fontweight='bold')
    ax2.set_ylabel('互动数')
    ax2.tick_params(axis='x', rotation=45)
    
    plt.tight_layout()
    save_figure(fig, os.path.join(output_dir, f"detail_{account_name}.png"))

def create_top_posts_chart(df, output_dir):
    top_posts = df.sort_values('互动数', ascending=False).head(10)
    
    fig, ax = plt.subplots(figsize=(12, 6))
    bars = ax.barh(range(len(top_posts)), top_posts['互动数'], color='#0077B6')
    ax.set_yticks(range(len(top_posts)))
    ax.set_yticklabels([title[:30] + '...' if len(title) > 30 else title for title in top_posts['作品标题']])
    ax.invert_yaxis()
    ax.set_xlabel('互动数')
    ax.set_title('Top 10 爆款作品（按互动量）', fontsize=16, fontweight='bold')
    
    for i, bar in enumerate(bars):
        width = bar.get_width()
        ax.text(width, bar.get_y() + bar.get_height()/2, 
                f'{int(width):,}', ha='left', va='center')
    
    plt.tight_layout()
    save_figure(fig, os.path.join(output_dir, "top_posts.png"))

def create_comparison_charts(df, output_dir):
    latest_data = df.sort_values('日期').groupby('账号名称').last().reset_index()
    metrics = ['涨粉量', '互动率', '播放量', '粉丝量']
    titles = ['各账号涨粉对比', '各账号互动率对比', '各账号播放量对比', '各账号粉丝总量对比']
    
    fig, axes = plt.subplots(2, 2, figsize=(16, 12))
    axes = axes.flatten()
    
    for i, (metric, title) in enumerate(zip(metrics, titles)):
        ax = axes[i]
        sns.barplot(x='账号名称', y=metric, data=latest_data, ax=ax, palette=["#0077B6", "#00B4D8", "#90E0EF", "#CAF0F8", "#48CAE4", "#03045E"])
        ax.set_title(title, fontsize=14, fontweight='bold')
        ax.tick_params(axis='x', rotation=45)
        ax.set_xlabel('')
    
    plt.tight_layout()
    save_figure(fig, os.path.join(output_dir, "comparison.png"))
