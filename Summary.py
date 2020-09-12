from tkinter import Tk, filedialog

import numpy as np
import pandas as pd

root = Tk()
root.withdraw()


def sort_resetIndex(df, Transaction_Type):
    df_sort = df[df["Transaction Type"].isin([Transaction_Type])]
    data = df_sort.reset_index(drop=True)
    return data


def cal_time(data_old, data_new, column_name):
    times_Cal = []
    for i in data_old.index.values:
        data_new_Isin = data_new[
            (data_new['Order Number'].isin([data_old['Order Number'].iloc[i]]))
            & (data_new['Line'].isin([data_old['Line'].iloc[i]]))]
        if not data_new_Isin.empty:
            time_new_old = (data_new_Isin['Date'] -
                            data_old['Date'].iloc[i]).values / np.timedelta64(
                                1, 'h')
            time_OrderNUmber = data_new_Isin['Order Number'].values
            time_Line = data_new_Isin["Line"].values
            time_Item = data_new_Isin["Item"].values
            time_New_Date = data_new_Isin["Date"].values
            time_Old_Date = data_old['Date'].iloc[i]
            time_Cal = [
                time_Item[0], time_OrderNUmber[0], time_Line[0], time_Old_Date,
                time_New_Date[0], time_new_old[0]
            ]
            times_Cal.append(time_Cal)
    timedf = pd.DataFrame(times_Cal, columns=column_name)
    print(timedf)
    return timedf


# excel所在地址
excel_file = filedialog.askopenfilename(title="Select the file",
                                        filetypes=[("All files", "*")])
# 打开excel
df = pd.read_excel(excel_file)
# 去掉重复行
df = df.drop_duplicates(subset=['Transaction Type', 'Order Number', 'Line'],
                        keep='last')
# 去掉空行
df = df.dropna(axis=0, how='any')
# 将Date列由字符串转换为时间
df['Date'] = pd.to_datetime(df['Date'])
# 排序date列
df.sort_values(by=['Date'],
               axis=0,
               ascending=True,
               inplace=True,
               kind='quicksort',
               na_position='last')
# 筛选Transaction Type列中包含Receive的数据并重新索引
data_Receive = sort_resetIndex(df, "Receive")
# 筛选Transaction Type列中包含Accept的数据并重新索引
data_Accept = sort_resetIndex(df, "Accept")
# # 筛选Transaction Type列中包含Deliver的数据并重新索引
data_Deliver = sort_resetIndex(df, "Deliver")
# 计算Receive到Accept时间
accept_Receive_Name = [
    'Item', 'Order Number', 'Line', "Receive Date", "Accept Date",
    'Receive to accept time'
]
timedf_Accept_Receive = cal_time(data_Receive, data_Accept,
                                 accept_Receive_Name)
# 计算Accept到Deliver时间
column_name = [
    'Item', 'Order Number', 'Line', "Accept Date", "Deliver Date",
    'Accept to deliver time'
]
timedf_Deliver_Accept = cal_time(data_Accept, data_Deliver, column_name)
# 写入Excel
writer = pd.ExcelWriter('Summary.xlsx')
df.to_excel(writer, sheet_name='1', index=False)
timedf_Accept_Receive.to_excel(writer,
                               sheet_name='Receive to accept',
                               index=False)
timedf_Deliver_Accept.to_excel(writer,
                               sheet_name='Accept to deliver',
                               index=False)
writer.save()
