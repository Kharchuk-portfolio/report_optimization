import pandas as pd
import openpyxl

# change every week
week = 38

df = pd.read_excel('videos.xlsx')

# pivot table (all types)
count_all = pd.pivot_table(df, values=['id', 'view_count'],
                           index='platform_channel_title',
                           aggfunc={'id': 'count', 'view_count': sum})

count_all = count_all.sort_values(by='id', ascending=False)

print(count_all['id'])
print(count_all['view_count'].sort_values(ascending=False))

# pivot table (all types)
count = pd.pivot_table(df, values=['id', 'view_count'],
                           index='platform_channel_title',
                           columns='type',
                           aggfunc={'id': 'count', 'view_count': sum})

count = count.drop('24 Канал онлайн')
count = count.fillna(0)

count_shorts = count.iloc[:, 0].sort_values(ascending=False)
count_streams = count.iloc[:, 1].sort_values(ascending=False)
count_videos = count.iloc[:, 2].sort_values(ascending=False)

count_shorts_views = count.iloc[:, 3].sort_values(ascending=False)
count_streams_views = count.iloc[:, 4].sort_values(ascending=False)
count_videos_views = count.iloc[:, 5].sort_values(ascending=False)

book = openpyxl.load_workbook('You Tube TV_w.xlsx')

videos = book['Відео']
streams = book['Прямі трансляції']
shorts = book['Shorts']

videos['C6'] = week
videos['G6'] = week

streams['C6'] = week
streams['G6'] = week

shorts['C6'] = week
shorts['G6'] = week

for i in range(len(count_videos)):
    videos['B'+str(7+i)] = str(count_videos_views.index[i])
    videos['C'+str(7+i)] = int(count_videos_views.iloc[i]) 

    videos['F'+str(7+i)] = str(count_videos.index[i])
    videos['G'+str(7+i)] = int(count_videos.iloc[i]) 

    streams['B'+str(7+i)] = str(count_streams_views.index[i])
    streams['C'+str(7+i)] = int(count_streams_views.iloc[i]) 

    streams['F'+str(7+i)] = str(count_streams.index[i])
    streams['G'+str(7+i)] = int(count_streams.iloc[i]) 

    shorts['B'+str(7+i)] = str(count_shorts_views.index[i])
    shorts['C'+str(7+i)] = int(count_shorts_views.iloc[i]) 

    shorts['F'+str(7+i)] = str(count_shorts.index[i])
    shorts['G'+str(7+i)] = int(count_shorts.iloc[i]) 

book.save('You Tube TV_w.xlsx')
