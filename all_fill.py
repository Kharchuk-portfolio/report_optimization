import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np

# change every week
week = 38
time_list = ['Sep 18, 2023', 'Sep 19, 2023', 'Sep 20, 2023', 'Sep 21, 2023', 'Sep 22, 2023', 'Sep 23, 2023', 'Sep 24, 2023']

# write dataframes
videos = pd.read_csv('files/videos_youtube.csv')
streams = pd.read_csv('files/streams_youtube.csv')
shorts = pd.read_csv('files/shorts_youtube.csv')

# function for update the dataframe in the required format
def df_update(df, df_name, week, time_list):

    df = df.iloc[1:, 1:]
    df['Перегляди'] = df['Перегляди'].fillna(0)
    df['Покази'] = df['Покази'].fillna(0)

    df['Перегляди'].astype(int)
    df['Покази'].astype(int)

    df['Середній відсоток перегляду відео (%)'].astype(float)
    df['Середня тривалість перегляду'] = pd.to_datetime(df['Середня тривалість перегляду'], format='%H:%M:%S').dt.time

    # creating new columns
    df.insert(loc=0, column="Тиждень", value=week)

    if df_name == "videos":
        df.insert(loc=1, column="Тип трансляції", value="Відео")
    elif df_name == "streams":
        df.insert(loc=1, column="Тип трансляції", value="Пряма трансляція")
    elif df_name == "shorts":
        df.insert(loc=1, column="Тип трансляції", value="Short")

    df.insert(loc=2, column="Час завантаження", value = 0)
    condition = df['Час публікації відео'].isin(time_list)
    df['Час завантаження'] = np.where(condition, 'Поточний тиждень', 'Доперегляди')
    
    # formatting for the general table
    df_upd = df.copy()   
    df_upd['time_in_seconds'] = (df_upd['Середня тривалість перегляду'].apply(lambda x: x.hour * 3600 + x.minute * 60 + x.second))/(df_upd['Середній відсоток перегляду відео (%)']/100)

    df_upd['hours'] = df_upd['time_in_seconds'] // 3600
    df_upd['minutes'] = (df_upd['time_in_seconds'] % 3600) // 60
    df_upd['seconds'] = df_upd['time_in_seconds'] % 60

    df_upd['Довжина'] = df_upd.apply(lambda row: f"{int(row['hours']) if pd.notnull(row['hours']) else 0:02d}:{int(row['minutes']) if pd.notnull(row['minutes']) else 0:02d}:{int(row['seconds']) if pd.notnull(row['seconds']) else 0:02d}", axis=1)

    df_upd['Довжина група'] = pd.cut(df_upd['time_in_seconds'], bins=[0, 60, 480, 1200, 1800, 3600, float('inf')], 
                                    labels=["до 1 хв", "від 1 до 8 хв", "від 8 до 20 хв", "від 20 до 30 хв", 
                                            "від 30 до 60 хв", "більше 1 години"])
    
    df_upd = df_upd.drop(['time_in_seconds', 'hours', 'minutes', 'seconds'], axis=1)

    df_upd['Перегляди група'] = pd.cut(df_upd['Перегляди'], bins=[0, 1000, 10000, 50000, 100000, 500000, float('inf')], 
                                        labels=["до 1 тис", "від 1 до 10 тис", "від 10 до 50 тис", "від 50 до 100 тис", 
                                        "від 100 до 500 тис", "більше 500 тис"])

    df_upd['Довжина група'] = df_upd['Довжина група'].fillna("до 1 хв")
    df_upd['Перегляди група'] = df_upd['Перегляди група'].fillna("до 1 тис")

    return df, df_upd

# function for writing dataframes to a file
def df_write(df, df_all, df_name):
    # open a file
    book = openpyxl.load_workbook("База Ютюб дохід-.xlsx")

    if df_name == "videos":
        sheet1 = book["Відео"]
    elif df_name == "streams":
        sheet1 = book["Прямі трансляції"]
    elif df_name == "shorts":
        sheet1 = book["Shorts"]

    sheet2 = book["Загалом"]

    # write to a file
    data_rows = list(dataframe_to_rows(df.iloc[:-1, :], index=False, header=False))
    data_rows_all = list(dataframe_to_rows(df_all.iloc[:-1, :], index=False, header=False))

    for row_data in data_rows:
        sheet1.append(row_data)
    
    for row_data_all in data_rows_all:
        sheet2.append(row_data_all)

    book.save("База Ютюб дохід-.xlsx")

videos, videos_all = df_update(videos, "videos", week, time_list)
streams, streams_all = df_update(streams, "streams", week, time_list)
shorts, shorts_all = df_update(shorts, "shorts", week, time_list)

df_write(videos, videos_all, "videos")
df_write(streams, streams_all, "streams")
df_write(shorts, shorts_all, "shorts")