import pandas as pd
import numpy as np
from datetime import datetime, timedelta

def generate_sample_data():
    account_names = [
        "美食探店达人", "旅行记录家", "职场小能手", 
        "宠物日记", "健身教练小王", "科技测评官"
    ]
    start_date = datetime.now() - timedelta(days=30)
    dates = [start_date + timedelta(days=i) for i in range(30)]
    
    data = []
    for account in account_names:
        base_followers = np.random.randint(50000, 500000)
        for date in dates:
            daily_followers = base_followers + np.random.randint(-500, 2000)
            daily_growth = np.random.randint(-300, 1500)
            likes = np.random.randint(100, 50000)
            comments = np.random.randint(10, 5000)
            shares = np.random.randint(5, 2000)
            favorites = np.random.randint(20, 10000)
            views = np.random.randint(1000, 1000000)
            
            titles = [
                f"{account}的精彩内容 - 第{np.random.randint(1, 100)}期",
                f"今日分享：{np.random.choice(['美食', '旅行', '职场', '宠物', '健身', '科技'])}小技巧",
                f"{np.random.choice(['超实用', '必看', '干货满满'])}！{account}教你一招"
            ]
            title = np.random.choice(titles)
            
            data.append({
                "账号名称": account,
                "日期": date.strftime("%Y-%m-%d"),
                "作品标题": title,
                "粉丝量": daily_followers,
                "涨粉量": daily_growth,
                "点赞数": likes,
                "评论数": comments,
                "分享数": shares,
                "收藏数": favorites,
                "播放量": views
            })
            base_followers = daily_followers
    
    df = pd.DataFrame(data)
    df["互动数"] = df["点赞数"] + df["评论数"] + df["收藏数"]
    df["互动率"] = df["互动数"] / df["播放量"].replace(0, 1)
    return df

def load_data(file):
    if file.name.endswith('.xlsx'):
        df = pd.read_excel(file)
    elif file.name.endswith('.csv'):
        df = pd.read_csv(file)
    else:
        raise ValueError("只支持 .xlsx 和 .csv 格式")
    return df

def map_columns(df, column_mapping):
    df = df.rename(columns=column_mapping)
    required_cols = ["账号名称", "日期", "作品标题", "粉丝量", "涨粉量", 
                     "点赞数", "评论数", "分享数", "收藏数", "播放量"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"缺少必要列: {col}")
    df["互动数"] = df["点赞数"] + df["评论数"] + df["收藏数"]
    df["互动率"] = df["互动数"] / df["播放量"].replace(0, 1)
    return df
