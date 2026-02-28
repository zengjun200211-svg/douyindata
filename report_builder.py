from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from docx import Document
from docx.shared import Inches as DocInches
import pandas as pd
import os

def build_ppt(df, output_dir, output_file="report.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = "抖音运营月度分析报告"
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 119, 182)
    
    date_range = f"{df['日期'].min()} 至 {df['日期'].max()}"
    subtitle.text = date_range
    subtitle.text_frame.paragraphs[0].font.size = Pt(18)
    
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = "目录"
    
    left = Inches(1)
    top = Inches(2)
    width = Inches(8)
    height = Inches(4)
    content = slide.shapes.add_textbox(left, top, width, height)
    text_frame = content.text_frame
    text_frame.text = "1. 整体概览\n2. 账号详情\n3. 爆款作品\n4. 账号对比\n5. 建议与总结"
    
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = "整体概览"
    
    kpis = [
        ("总粉丝", f"{df.sort_values('日期').groupby('账号名称')['粉丝量'].last().sum():,}"),
        ("总涨粉", f"{df['涨粉量'].sum():,}"),
        ("总点赞", f"{df['点赞数'].sum():,}"),
        ("总评论", f"{df['评论数'].sum():,}"),
        ("总收藏", f"{df['收藏数'].sum():,}"),
        ("总播放", f"{df['播放量'].sum():,}")
    ]
    
    for i, (label, value) in enumerate(kpis):
        col = i % 3
        row = i // 3
        left = Inches(0.5 + col * 3)
        top = Inches(1.5 + row * 1.5)
        width = Inches(2.8)
        height = Inches(1.2)
        
        shape = slide.shapes.add_shape(1, left, top, width, height)
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(0, 119, 182)
        
        text_frame = shape.text_frame
        p = text_frame.paragraphs[0]
        p.text = f"{label}\n{value}"
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.bold = True
    
    img_path = os.path.join(output_dir, "overview_pie.png")
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(3.5), Inches(4.5), height=Inches(2.5))
    
    accounts = df['账号名称'].unique()
    for account in accounts:
        slide_layout = prs.slide_layouts[5]
        slide = prs.slides.add_slide(slide_layout)
        title_shape = slide.shapes.title
        title_shape.text = f"账号详情 - {account}"
        
        img_path = os.path.join(output_dir, f"detail_{account}.png")
        if os.path.exists(img_path):
            slide.shapes.add_picture(img_path, Inches(0.5), Inches(1.5), height=Inches(5.5))
    
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = "爆款作品"
    
    img_path = os.path.join(output_dir, "top_posts.png")
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(0.5), Inches(1.5), height=Inches(3))
    
    top_posts = df.sort_values('互动数', ascending=False).head(10)
    table = slide.shapes.add_table(11, 4, Inches(0.5), Inches(4.8), Inches(9), Inches(2.2)).table
    table.columns[0].width = Inches(4)
    table.columns[1].width = Inches(1.5)
    table.columns[2].width = Inches(1.5)
    table.columns[3].width = Inches(2)
    
    headers = ["作品标题", "账号", "互动数", "互动率"]
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(0, 119, 182)
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.bold = True
    
    for idx, (_, post) in enumerate(top_posts.iterrows()):
        row = table.rows[idx + 1]
        row.cells[0].text = post['作品标题'][:30] + "..."
        row.cells[1].text = post['账号名称']
        row.cells[2].text = f"{post['互动数']:,}"
        row.cells[3].text = f"{post['互动率']:.2%}"
    
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = "账号对比"
    
    img_path = os.path.join(output_dir, "comparison.png")
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(0.5), Inches(1.5), height=Inches(5.5))
    
    slide_layout = prs.slide_layouts[5]
    slide = prs.slides.add_slide(slide_layout)
    title_shape = slide.shapes.title
    title_shape.text = "建议与总结"
    
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(9)
    height = Inches(5.5)
    content = slide.shapes.add_textbox(left, top, width, height)
    text_frame = content.text_frame
    
    summary_text = """
【亮点】
1. 整体粉丝增长趋势良好，各账号均有稳定表现
2. 爆款作品互动率突出，内容质量得到用户认可
3. 账号矩阵布局合理，覆盖多个垂直领域

【问题】
1. 部分账号涨粉波动较大，稳定性有待提升
2. 评论互动率相对较低，需加强用户引导
3. 内容发布频率不均衡，建议优化发布策略

【建议】
1. 针对爆款作品内容特点，持续产出同类型优质内容
2. 增加评论区互动，积极回复用户留言
3. 制定固定发布计划，保持内容更新频率
    """
    text_frame.text = summary_text.strip()
    
    output_path = os.path.join(output_dir, output_file)
    prs.save(output_path)
    return output_path

def build_word(df, output_dir, output_file="report.docx"):
    doc = Document()
    doc.add_heading('抖音运营月度分析报告', 0)
    doc.add_paragraph(f"报告期间：{df['日期'].min()} 至 {df['日期'].max()}")
    
    doc.add_heading('整体概览', level=1)
    total_fans = df.sort_values('日期').groupby('账号名称')['粉丝量'].last().sum()
    doc.add_paragraph(f"总粉丝数：{total_fans:,}")
    doc.add_paragraph(f"总涨粉：{df['涨粉量'].sum():,}")
    doc.add_paragraph(f"总点赞：{df['点赞数'].sum():,}")
    doc.add_paragraph(f"总评论：{df['评论数'].sum():,}")
    doc.add_paragraph(f"总收藏：{df['收藏数'].sum():,}")
    doc.add_paragraph(f"总播放：{df['播放量'].sum():,}")
    
    doc.add_heading('爆款作品', level=1)
    top_posts = df.sort_values('互动数', ascending=False).head(10)
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = '作品标题'
    hdr_cells[1].text = '账号'
    hdr_cells[2].text = '互动数'
    hdr_cells[3].text = '互动率'
    
    for _, post in top_posts.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = post['作品标题']
        row_cells[1].text = post['账号名称']
        row_cells[2].text = f"{post['互动数']:,}"
        row_cells[3].text = f"{post['互动率']:.2%}"
    
    doc.add_heading('建议与总结', level=1)
    summary_text = """
【亮点】
1. 整体粉丝增长趋势良好，各账号均有稳定表现
2. 爆款作品互动率突出，内容质量得到用户认可
3. 账号矩阵布局合理，覆盖多个垂直领域

【问题】
1. 部分账号涨粉波动较大，稳定性有待提升
2. 评论互动率相对较低，需加强用户引导
3. 内容发布频率不均衡，建议优化发布策略

【建议】
1. 针对爆款作品内容特点，持续产出同类型优质内容
2. 增加评论区互动，积极回复用户留言
3. 制定固定发布计划，保持内容更新频率
    """
    doc.add_paragraph(summary_text.strip())
    
    output_path = os.path.join(output_dir, output_file)
    doc.save(output_path)
    return output_path
