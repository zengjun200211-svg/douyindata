# 抖音运营分析报告生成器

一个功能强大的 Streamlit 应用，用于生成抖音多账号运营数据的专业分析报告。

## 功能特点

- 📊 **零本地依赖**：无需安装 Python，通过 Streamlit Cloud 部署后即可直接使用
- 🎯 **即插即用**：内置 6 个账号、30 天完整的模拟数据，一键查看效果
- 📁 **支持自定义数据**：支持上传 .xlsx 和 .csv 格式文件，灵活的列名映射功能
- 📈 **自动计算字段**：自动计算互动数和互动率
- 🎨 **商务风格可视化**：蓝绿色渐变配色，完美支持中文字体显示
- 📄 **自动生成报告**：一键生成专业的 PPTX 和 Word 报告
- 📅 **日期范围选择**：灵活选择报告的时间范围

## 数据格式要求

如需上传自己的数据，请确保包含以下列：

| 列名 | 类型 | 说明 |
|------|------|------|
| 账号名称 | 文本 | 账号的名称 |
| 日期 | 日期 | YYYY-MM-DD 格式 |
| 作品标题 | 文本 | 作品的标题 |
| 粉丝量 | 整数 | 当前粉丝数量 |
| 涨粉量 | 整数 | 当日涨粉数量 |
| 点赞数 | 整数 | 当日点赞数量 |
| 评论数 | 整数 | 当日评论数量 |
| 分享数 | 整数 | 当日分享数量 |
| 收藏数 | 整数 | 当日收藏数量 |
| 播放量 | 整数 | 当日播放数量 |

## 部署到 Streamlit Cloud

### 1. 创建 GitHub 仓库

1. 访问 https://github.com/new
2. 仓库名称：`douyin-report-generator`（或其他您喜欢的名字）
3. 设置为 Public 或 Private
4. 点击 **Create repository**

### 2. 上传代码

在项目目录下执行以下命令（需先安装 Git）：

```bash
git init
git add .
git commit -m "Initial commit: Douyin Report Generator"
git branch -M main
git remote add origin https://github.com/您的用户名/douyin-report-generator.git
git push -u origin main
```

### 3. 在 Streamlit Cloud 上部署

1. 访问 https://streamlit.io/cloud 并登录
2. 点击 **New app**
3. 选择您的 GitHub 仓库和 `main` 分支
4. Main file path 选择 `app.py`
5. 点击 **Deploy!**

部署完成后，您将获得一个可直接访问的公共链接！

## 本地运行

如需在本地运行测试：

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 项目结构

```
douyin-report-generator/
├── .gitignore
├── README.md
├── requirements.txt
├── app.py                      # Streamlit 主应用
└── src/
    ├── __init__.py
    ├── data_processor.py       # 数据处理和示例数据生成
    ├── chart_generator.py      # 图表生成
    └── report_builder.py       # PPT/Word 报告生成
```

## 许可证

MIT License
