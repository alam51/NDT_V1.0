import pandas as pd
import matplotlib.pyplot as plt
from dateutil.relativedelta import relativedelta
from pandas import ExcelWriter

path = 'F:\\IMD\Data\\Generation\\On_requirement\\generation_ghora_4.csv'
df = pd.read_csv(path, dayfirst=True, index_col=0, parse_dates=True)
df['MW'] = pd.to_numeric(df['MW'], errors='coerce')
df.fillna(0)

running_month = pd.to_datetime(df.index[0])
running_month_last_hour = running_month.replace(day=running_month.days_in_month, hour=23)
df_list = []  # declares an empty dataframe list

with ExcelWriter('g4.xlsx') as writer:
    # while running_month <= pd.to_datetime(df.index[-1]).month:
    df1 = pd.DataFrame(index=list(range(0, 23 + 1)),
                       columns=list(range(1, running_month_last_hour.day + 1)))  # declares
    # an empty dataframe
    for date_times in df.index:
        dt = pd.to_datetime(date_times)
        del_day = (dt - running_month_last_hour).total_seconds()

        if del_day <= 0:
            df1.loc[dt.hour, dt.day] = df.loc[date_times, 'MW']
        else:  # Data entry of current month is finished. Time to write them in a excel sheet
            df1.sort_index(ascending=True, inplace=True)
            for i in df1.index:
                df1.loc[i, 'time'] = str(i) + ':00'
            df1.set_index('time', inplace=True)
            df1.to_excel(writer, sheet_name=str(running_month_last_hour.month_name()) + '_' +
                                            str(running_month_last_hour.year), na_rep='--')
            running_month_last_hour += relativedelta(months=+1)
            running_month_last_hour = running_month_last_hour.replace(day=running_month_last_hour.days_in_month,
                                                                      hour=23)
            df1 = pd.DataFrame(index=list(range(0, 23 + 1)), columns=list(range(1, running_month_last_hour.day + 1)))
            df1.loc[dt.hour, dt.day] = df.loc[date_times, df.columns[0]]

    df1.sort_index(ascending=True, inplace=True)  # for the last month data that does not go to else block
    for i in df1.index:
        df1.loc[i, 'time'] = str(i) + ':00'
    df1.set_index('time', inplace=True)
    # df2.rename_axis('Time', axis=0, inplace=True)
    df1.to_excel(writer, sheet_name=str(running_month_last_hour.month_name()) + '_' +
                                    str(running_month_last_hour.year), na_rep='--')

b = 5
df.plot()
plt.show()
