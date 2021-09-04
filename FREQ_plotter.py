import datetime
import re
from dateutil.relativedelta import relativedelta
from matplotlib.backends.backend_pdf import PdfPages
import matplotlib.pyplot as plt
import pandas as pd
import utils


def freq_plotter():
    print("*****Don't proceed unless data file is in correct form*******")
    n = int(input("Number of data files:  "))
    df_raw = pd.DataFrame()
    pd_td_1min = pd.to_timedelta('1 minute')
    for i in range(1, n + 1):
        absolute_file_path = re.escape(input(f'File #{i} path:  ')) + '\\' + \
                             input(f'.csv File #{i} name [w/o extension]:  ') + '.csv'
        # df = pd.read_csv(absolute_file_path, index_col=None, header=None)
        df_raw_in = pd.read_csv(absolute_file_path, index_col=0, parse_dates=True, infer_datetime_format=True,
                                low_memory=False, dayfirst=False)
        df_raw_in.dropna(axis=1, inplace=True, thresh=3)
        if i == 1:
            df_raw = df_raw_in
        else:
            # assuming df_raw_in is newer,
            # df_raw to be placed after df_raw_in
            if df_raw_in.index[0] > df_raw.index[0]:
                if df_raw_in.index[0] > df_raw.index[-1]:
                    # df_raw.join(df_raw_in, how='outer', sort=False)
                    df_raw = pd.concat([df_raw, df_raw_in], join='outer')
                else:
                    # df_raw.join(df_raw_in[df_raw.index[-1] + pd_td_1min:])
                    df_raw = pd.concat([df_raw, df_raw_in[df_raw.index[-1] + pd_td_1min:]], join='outer')
            # if df_raw_in is older
            # df_raw to be placed before df_raw_in
            elif df_raw_in.index[0] < df_raw.index[0]:
                if df_raw_in.index[-1] < df_raw.index[0]:
                    # df_raw.merge(df_raw_in, how='outer', sort=False)
                    df_raw = pd.concat([df_raw_in, df_raw], join='outer', sort=False)
                else:
                    df_raw = pd.concat([df_raw_in[: df_raw.index[0] - pd_td_1min], df_raw], join='outer')
                    # df_raw.join(df_raw_in[: df_raw.index[0] - pd_td_1min])
            # if df_raw and df_raw_in starts with same time
            else:
                if df_raw_in.index[-1] > df_raw.index[-1]:  # df_raw_in has more data
                    df_raw = df_raw_in

    print('Which column to plot?')
    for i, col in enumerate(df_raw.columns):
        print(f'{i} : {col}')
    col = str(df_raw.columns[int(input('column number: '))])
    print(f'Data available from {df_raw.index[0]} to {df_raw.index[-1]}')
    # Time limits
    T1_str = input(f'Start time [yyyy-mm-dd] \n[blank for starting from first row of of database]:')
    T2_str = input(f'End time [yyyy-mm-dd] \n[blank for ending at last row of of database]:')
    time_period_str = input('Time Period of plotting:  ')
    if 'month' in time_period_str.lower():
        time_period = relativedelta(months=int(time_period_str[0]))
    elif 'y' in time_period_str.lower():
        time_period = relativedelta(years=int(time_period_str[0]), )
    else:
        time_period = pd.to_timedelta(time_period_str)
    if T1_str == '':
        T1 = df_raw.index[0]
    else:
        T1 = pd.to_datetime(T1_str, dayfirst=False, yearfirst=True, exact=False)

    if T2_str == '':
        T2 = df_raw.index[-1]
    else:
        T2 = pd.to_datetime(T2_str, dayfirst=False, yearfirst=True, exact=False)
    df = pd.DataFrame()
    df[col] = df_raw.loc[T1:T2, col]
    t1 = T1
    t2 = t1 + time_period

    with PdfPages('one.pdf') as pdf:
        condition = True
        while condition:
            if t2 > T2:
                t2 = df.index[-1]
                condition = False
            df1 = df[str(t1): str(t2)]

            # bar_y = df2.iloc[:, 0]

            # plotting part
            # if LaTeX is not installed or error caught, change to `False`
            plt.rcParams['text.usetex'] = False
            plt.rc('figure', figsize=(11.69, 8.27))
            # plt.figure(figsize=(8, 6))
            fig, ax = plt.subplots()
            df1.plot(ax=ax, style='b', lw=0.6)
            ax.set_ylim(47.5, 52)
            ax.set_xlabel('[hh:mm]')
            # ax.format_xdata = mdates.DateFormatter('%Y-%m-%d %H-%M')
            # ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d %H-%M'))
            ax.xaxis.labelpad = -5
            ax.set_ylabel('[Hz]')
            ax.legend(loc='upper right', frameon=True)
            plt.title('System Frequency Curve on ' + str(t1)[:10], fontsize=20)
            # plt.title('System Frequency Curve in 2020', fontsize=20)
            plt.grid(True)
            pdf.savefig()  # saves the current figure into a pdf page
            plt.close()

            # *******for bar chart********
            bar_x = ['<48.86Hz', '48.86-\n49.00Hz', '49.01-\n49.50Hz', '49.51-\n50.00Hz',
                     '50.01-\n50.50Hz', '50.51-\n51.00Hz', '>51.00Hz', 'else']
            bins = [48.86, 49.00, 49.50, 50.00, 50.50, 51.00]
            bar_y = [0, 0, 0, 0, 0, 0, 0, 0]
            for i in df1[col]:
                if i < 48.86:
                    bar_y[0] += 1
                elif 48.86 <= i < 49.0:
                    bar_y[1] += 1
                elif 49.0 <= i < 49.5:
                    bar_y[2] += 1
                elif 49.5 <= i < 50.0:
                    bar_y[3] += 1
                elif 50.0 <= i < 50.5:
                    bar_y[4] += 1
                elif 50.5 <= i < 51.0:
                    bar_y[5] += 1
                elif 51.0 <= i:
                    bar_y[6] += 1
                else:
                    bar_y[7] += 1

            for i, val in enumerate(bar_y):
                bar_y[i] = round(bar_y[i] / 60, 2)

            # df2 = df1['Freq'].value_counts(bins=bins, sort=False)
            plt.rcParams['text.usetex'] = False
            plt.rc('figure', figsize=(11.69, 8.27))
            # df3 = pd.DataFrame({'bar_x': bar_x, 'bar_y': bar_y})
            # ax = df3.plot.bar(x='bar_x', y='bar_y', rot=0)
            fig, ax = plt.subplots()
            # ax.inset_axes([x0, y0, w, h])
            inset_ax = ax.inset_axes([0.1, 1 - 0.4, .3, .3])
            ax.bar(x=bar_x, height=bar_y, label='Freq range')
            height_sum = 1
            for p in ax.patches:
                height_sum += p.get_height()
            for p in ax.patches:
                ax.annotate(str(p.get_height()) + ' (' + str(round(p.get_height() / height_sum * 100, 2)) +
                            '%)', (p.get_x() * 1, p.get_height() * 1.008))
            ax.set_xlabel('Frequency range', fontsize=15)
            ax.xaxis.labelpad = 0
            ax.set_ylabel('Hours', fontsize=15)
            ax.legend(loc='upper right', frameon=True)
            plt.suptitle('System Frequency Distribution on ' + str(t1)[:10], fontsize=20)
            # plt.title('System Frequency Distribution Bar Chart on 2020', fontsize=20)
            # plt.grid(True)
            in_range = sum(bar_y[3:4 + 1])
            out_range = sum(bar_y) - in_range
            values = [in_range, out_range]
            labels = ['Within\n49.5-50.5Hz', 'Outside\n49.5-50Hz']
            inset_ax.pie(values, labels=labels,
                         autopct=lambda x: f'{round(x * sum(values) / 100, 2)}\n({round(x, 2)}%)',
                         startangle=-90, shadow=True)
            pdf.savefig()  # saves the current figure into a pdf page
            plt.close()

            if t2 >= T2:
                break  # End has been reached. Time to rest. Actually here t2 can at best touch T2
            t1 = t2
            t2 += time_period

        # We can also set the file's metadata via the PdfPages object:
        d = pdf.infodict()
        d['Title'] = 'Multipage PDF Example'
        d['Author'] = 'Sharif Shamsul Alam (AE, IMD)'
        d['Subject'] = 'IMD, PGCB Report'
        d['Keywords'] = 'PdfPages multipage keywords author title subject'
        d['CreationDate'] = datetime.datetime(2021, 5, 9)
        d['ModDate'] = datetime.datetime.today()
    pdf_file_path = re.escape(input('Output pdf file path : ')) + '\\' + 'freq_rep_' + \
                    str(T1)[:16].replace(':', '') + \
                    ' to ' + str(T2)[:16].replace(':', '') + '.pdf'
    utils.pdf_decorator('one.pdf', pdf_file_path)
    print('successfully saved at' + pdf_file_path)
