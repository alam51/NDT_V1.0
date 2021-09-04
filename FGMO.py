import pandas as pd
import re
import matplotlib.pyplot as plt


def fgmo():
    print('*** Make Sure that the .csv file is in correct format***')
    file_path = re.escape(input('File path: '))
    file_name = input('.csv File name without extension: ')
    absolute_file_path = file_path + '\\' + file_name + '.csv'
    print('First portion of the data frame looks like:')
    df_1 = pd.read_csv(absolute_file_path, index_col=None, header=None, nrows=5)
    print(df_1.to_string())
    day_first = input('Day first in date time column?[T/F] : ')
    while True:
        if 'T' in day_first.upper():
            day_first = True
            break
        elif 'F' in day_first.upper():
            day_first = False
            break
        else:
            print("Neither 'T' or 'F'. Enter Again:  ")
    # df = pd.read_csv(absolute_file_path, index_col=None, header=None)
    skip_row = input('Row to skip : ')
    if skip_row == '':
        skip_row = None
    else:
        skip_row = [int(skip_row)]
    df = pd.read_csv(absolute_file_path, index_col=0, parse_dates=True, infer_datetime_format=True,
                     low_memory=False, dayfirst=day_first, skiprows=skip_row)
    df.dropna(axis=1, thresh=5, inplace=True)
    df.dropna(axis=0, thresh=2, inplace=True)

    columns = []
    for i, column in enumerate(df.columns):
        columns.append(str(i) + ' :' + str(column))
    df.columns = columns
    print('Processed data frame: ')
    print(df.head(3).to_string())
    fig, ax = plt.subplots()
    # gen_unit_name = input('Generation Unit Name: ')
    freq_col = int(input('Frequency column no.:  '))
    power_col = int(input('Power column no.:  '))
    ax.set_title(df.columns[power_col][3:] + ' Generation Curve on ' + f'{df.index[0]} to {df.index[-1]}',
                 fontsize=17)
    ax.set_xlabel('Time', fontsize=17)
    ax.set_ylabel('Frequency [Hz]', fontsize=17)
    ax1 = ax.twinx()  # create secondary axis
    ax1.set_ylabel('GEN ACTIVE POWER [MW]')
    df1 = df.iloc[:, freq_col]
    # df1.plot(ax=ax, style='g--', x_compat=True, label='Freq [Hz]')
    df1.plot(ax=ax, style='r-', x_compat=True, lw=0.8)  # plot freq
    df2 = df.iloc[:, power_col]
    # df2.plot(ax=ax1, style='b-', x_compat=True, label='Gen[MW]')
    df2.plot(ax=ax1, style='b-', x_compat=True, lw=0.8)  # plot power
    leg = ax.legend(loc='lower left', frameon=False)
    leg1 = ax1.legend(loc='lower right', frameon=False)

    plt.grid(True)
    plt.show()

