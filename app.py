import streamlit as st
import pandas as pd
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.data_processor import generate_sample_data, load_data, map_columns
from src.chart_generator import (
    create_overview_chart, 
    create_account_detail_charts, 
    create_top_posts_chart, 
    create_comparison_charts
)
from src.report_builder import build_ppt, build_word

st.set_page_config(
    page_title="抖音运营分析报告生成器",
    page_icon="📊",
    layout="wide"
)

st.title("📊 多账号抖音运营全方位分析报告生成器")
st.markdown("---")

st.sidebar.header("数据输入")
use_sample = st.sidebar.button("📋 使用示例数据")
uploaded_file = st.sidebar.file_uploader("上传数据文件 (.xlsx 或 .csv)", type=['xlsx', 'csv'])

if 'df' not in st.session_state:
    st.session_state.df = None
if 'output_dir' not in st.session_state:
    st.session_state.output_dir = None

if use_sample:
    st.session_state.df = generate_sample_data()
    st.sidebar.success("✅ 示例数据已加载！")

if uploaded_file is not None and not use_sample:
    try:
        df_uploaded = load_data(uploaded_file)
        st.session_state.df = df_uploaded
        
        standard_cols = ["账号名称", "日期", "作品标题", "粉丝量", "涨粉量", 
                         "点赞数", "评论数", "分享数", "收藏数", "播放量"]
        missing_cols = [col for col in standard_cols if col not in df_uploaded.columns]
        
        if missing_cols:
            st.sidebar.warning("⚠️ 检测到列名不标准，请进行列映射")
            column_mapping = {}
            for std_col in standard_cols:
                if std_col not in df_uploaded.columns:
                    column_mapping[st.sidebar.selectbox(f"将哪一列映射为 '{std_col}'", df_uploaded.columns)] = std_col
            
            if st.sidebar.button("确认映射"):
                st.session_state.df = map_columns(df_uploaded, column_mapping)
                st.sidebar.success("✅ 列映射完成！")
        
    except Exception as e:
        st.sidebar.error(f"❌ 加载文件失败: {str(e)}")

if st.session_state.df is not None:
    df = st.session_state.df
    
    st.subheader("📅 日期范围选择")
    df['日期'] = pd.to_datetime(df['日期'])
    min_date = df['日期'].min()
    max_date = df['日期'].max()
    
    start_date, end_date = st.slider(
        "选择报告日期范围",
        min_value=min_date.to_pydatetime(),
        max_value=max_date.to_pydatetime(),
        value=(min_date.to_pydatetime(), max_date.to_pydatetime())
    )
    
    df_filtered = df[(df['日期'] >= start_date) & (df['日期'] <= end_date)].copy()
    df_filtered['日期'] = df_filtered['日期'].dt.strftime('%Y-%m-%d')
    
    with st.expander("📋 数据预览"):
        st.dataframe(df_filtered.head(10))
    
    if st.button("🚀 开始生成报告"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            if st.session_state.output_dir is None:
                st.session_state.output_dir = tempfile.mkdtemp()
            
            status_text.text("Processing...")
            progress_bar.progress(10)
            
            status_text.text("Generating Charts...")
            create_overview_chart(df_filtered, st.session_state.output_dir)
            progress_bar.progress(20)
            
            accounts = df_filtered['账号名称'].unique()
            for i, account in enumerate(accounts):
                create_account_detail_charts(df_filtered, account, st.session_state.output_dir)
                progress_bar.progress(20 + (i + 1) * 8)
            
            create_top_posts_chart(df_filtered, st.session_state.output_dir)
            progress_bar.progress(75)
            
            create_comparison_charts(df_filtered, st.session_state.output_dir)
            progress_bar.progress(85)
            
            status_text.text("Building PPT...")
            ppt_path = build_ppt(df_filtered, st.session_state.output_dir)
            progress_bar.progress(92)
            
            word_path = build_word(df_filtered, st.session_state.output_dir)
            progress_bar.progress(100)
            
            status_text.text("✅ 报告生成完成！")
            st.success("🎉 报告生成成功！")
            
            st.markdown("---")
            st.subheader("📥 下载报告")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                with open(ppt_path, "rb") as file:
                    st.download_button(
                        label="📥 下载 PPTX",
                        data=file,
                        file_name="douyin_report.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
            
            with col2:
                with open(word_path, "rb") as file:
                    st.download_button(
                        label="📥 下载 Word",
                        data=file,
                        file_name="douyin_report.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            
            with col3:
                st.info("PDF 下载功能需要在本地安装 Microsoft PowerPoint，暂不支持在云端直接生成")
        
        except Exception as e:
            st.error(f"❌ 生成报告失败: {str(e)}")
            import traceback
            st.error(traceback.format_exc())

else:
    st.info("👈 请从左侧侧边栏上传数据文件，或点击「使用示例数据」开始！")
    st.markdown("---")
    st.subheader("📊 数据格式要求")
    st.markdown("""
    请确保您的数据文件包含以下列：
    - **账号名称** (文本)
    - **日期** (YYYY-MM-DD 格式)
    - **作品标题** (文本)
    - **粉丝量** (整数)
    - **涨粉量** (整数)
    - **点赞数** (整数)
    - **评论数** (整数)
    - **分享数** (整数)
    - **收藏数** (整数)
    - **播放量** (整数)
    """)
