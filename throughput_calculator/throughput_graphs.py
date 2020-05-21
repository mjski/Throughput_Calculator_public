############################################### -- throughput_graphs -- #################################################
#                                                   author: Morgan D
# Version Python: 3.7
# Created on 28 Jan 2020
# Purpose: create graphs with all data frames in the throughput_report
#
########################################################################################################################

import matplotlib.pyplot as plt
import numpy as np
from matplotlib import rc
from matplotlib.font_manager import FontProperties


class ThroughputGraphs:

    def __init__(self):
        pass

    def make_army_graphs(self, df, chart_cat):
        '''
            This will use inputs to create specific graphs using the Army dataframe
        :param df: army_only_df
        :param chart_cat: a string input that will determine which chart to create (Active Army, National Guard, Reserves)
        :return: charts (get into spreadsheet later)
        '''
        input = str(chart_cat).title()
        temp_df = df.filter(regex=input, axis=1)
        fig, (ax1, ax2) = plt.subplots(ncols=2, nrows=1, figsize=(18, 11), dpi=90)
        plt.suptitle('Army Quotas', fontsize=16)
        rc('font', weight='bold')

        # if temp_df.columns > 1 and temp_df.columns < 3:

        fontP = FontProperties()
        fontP.set_size('x-small')


        bins = list(df.index.values)
        bars1 = list(df.iloc[:, 0])
        bars2 = list(df.iloc[:, 1])
        bars3 = list(df.iloc[:, 2])

        # The position of the bars on the x-axis
        r = list(range(len(bins)))

        barWidth = 0.5

        # Heights of bars1 + bars2
        bars = np.add(bars1, bars2).tolist()
        # Create brown bars
        bar1 = plt.bar(r, bars1, color='palegreen', edgecolor='white', width=barWidth)
        # Create green bars (middle), on top of the first ones
        bar2 = plt.bar(r, bars2, bottom=bars1, color='gold', edgecolor='white', width=barWidth)
        # Create green bars (top)
        bar3 = plt.bar(r, bars3, bottom=bars, color='coral', edgecolor='white', width=barWidth)

        # Custom X axis
        plt.xticks(r, bins, fontweight='bold', fontProperties=fontP)
        plt.xlabel("Months")
        plt.ylabel("Quota Totals")
        #axs = plt.gca()
        ax2.set_ylim([0, 145])
        plt.legend((bar1, bar2, bar3), (temp_df.columns.values),
                   prop=fontP, loc='upper center', ncol=3, bbox_to_anchor=(0.5, 1.00), fancybox=True)
        plt.title('Cumulative Stacked Bar Chart')

        ############################################################################

        ax1.set_ylabel('Quota Totals')
        ax1.set_title('Line Chart').set_position([.5, 1.05])
        ax1.set_xticklabels(bins, fontproperties=fontP, rotation='0')
        ax1.plot(bins, bars1, color='dodgerblue')
        for i, j in zip(bins, bars1):
            ax1.annotate(str(j), xy=(i, j - 0.5))
        ax1.plot(bins, bars2, color='r', linestyle='dashed')
        for i, j in zip(bins, bars2):
            ax1.annotate(str(j), xy=(i, j - 0.5))
        ax1.plot(bins, bars3, color='g', linestyle='dashdot')
        for i, j in zip(bins, bars3):
            ax1.annotate(str(j), xy=(i, j - 0.5))

        ax1.legend(temp_df.columns)

        ###################################################################################
        '''
        r = [0, 1, 2, 3]
        width = 0.13  # the width of the bars
        r1 = np.arange(len(r))  # the label locations
        r2 = [x + width for x in r1]
        r3 = [x + width for x in r2]
        r4 = [x + width for x in r3]
        r5 = [x + width for x in r4]
        r6 = [x + width for x in r5]
        r7 = [x + width for x in r6]

        rects1 = axs[1, 0].barh(r2, bars1, width, label='Safety Features', color='palegreen')
        rects2 = axs[1, 0].barh(r3, bars2, width, label='Maintenance Cost', color='gold')
        rects3 = axs[1, 0].barh(r4, bars3, width, label='Price Point', color='coral')
        rects4 = axs[1, 0].barh(r5, bars4, width, label='Insurance', color='purple')
        rects5 = axs[1, 0].barh(r6, bars5, width, label='Fuel Economy', color='deeppink')
        rects6 = axs[1, 0].barh(r7, bars6, width, label='Resale Value', color='deepskyblue')
        rects7 = axs[1, 0].barh(r1, group_size, width, label='Cumulative Total', color='black')

        # Add some text for labels, title and custom x-axis tick labels, etc.
        axs[1, 0].set_ylabel('Vehicles')
        axs[1, 0].set_xlim([0, 130])
        axs[1, 0].set_title('Horizontal Grouped Bar Chart')
        axs[1, 0].set_yticks([w + width * 3 for w in range(len(bars1))])
        axs[1, 0].set_yticklabels(names)
        axs[1, 0].set_xlabel("Cumulative Total")
        axs[1, 0].legend((rects6, rects5, rects4, rects3, rects2, rects1, rects7), ('Resale Value', 'Fuel Economy',
                                                                                    'Insurance', 'Price Point',
                                                                                    'Maintenance Cost',
                                                                                    'Safety Features',
                                                                                    'Cumulative Total'),
                         prop=fontP, loc='upper right')


        '''
        plt.savefig("ArmyGraph" + "_" + str(chart_cat) + ".png")
        plt.close()



    def singular_graph(self, df, chart_cat):
        '''
            This will use inputs to create specific graphs using the Non-Army dataframe
        :param df: non_army_df ONLY
        :param chart_cat: a string input that will determine which chart to create (e.g. Air Force)
        :return: charts (get into spreadsheet later)
        '''

        input = str(chart_cat).title()
        temp_df = df.filter(regex=input, axis=1)
        fig, ax1 = plt.subplots(ncols=1, nrows=1, figsize=(17, 10), dpi=80)
        plt.suptitle(str(chart_cat) + ' Quotas', fontsize=16)
        rc('font', weight='bold')

        fontP = FontProperties()
        fontP.set_size('x-small')

        bins = list(df.index.values)
        bars1 = list(df.iloc[:, 0])
        # bars2 = list(df.iloc[:, 1])
        # bars3 = list(df.iloc[:, 2])

        # The position of the bars on the x-axis
        r = list(range(len(bins)))

        barWidth = 0.5

        # Heights of bars1 + bars2
        #bars = np.add(bars1, bars2).tolist()
        # Create brown bars
        bar1 = plt.bar(r, bars1, color='palegreen', edgecolor='white', width=barWidth)
        # Create green bars (middle), on top of the first ones
        #bar2 = plt.bar(r, bars2, bottom=bars1, color='gold', edgecolor='white', width=barWidth)
        # Create green bars (top)
        #bar3 = plt.bar(r, bars3, bottom=bars, color='coral', edgecolor='white', width=barWidth)

        # Custom X axis
        plt.xticks(r, bins, fontweight='bold', fontProperties=fontP)
        plt.xlabel("Months")
        plt.ylabel("Quota Totals")
        # axs = plt.gca()
        #axs[1, 1].set_ylim([0, 145])
        #plt.legend(bar1, (temp_df.columns.values),
                   #prop=fontP, loc='upper center', ncol=3, bbox_to_anchor=(0.5, 1.00), fancybox=True)
        #plt.title('Cumulative Stacked Bar Chart')

        ############################################################################

        ax1.set_ylabel('Quota Totals')
        #ax1.set_title( + 'Line Chart').set_position([.5, 1.05])
        ax1.set_xticklabels(bins, fontproperties=fontP, rotation='0')
        ax1.plot(bins, bars1, color='dodgerblue')
        for i, j in zip(bins, bars1):
            ax1.annotate(str(j), xy=(i, j - 0.5))
        # ax2[0, 1].plot(bins, bars2, color='r', linestyle='dashed')
        # for i, j in zip(bins, bars2):
        #     ax2[0, 1].annotate(str(j), xy=(i, j - 0.5))
        # ax2[0, 1].plot(bins, bars3, color='g', linestyle='dashdot')
        # for i, j in zip(bins, bars3):
        #     ax2[0, 1].annotate(str(j), xy=(i, j - 0.5))

        #ax1.legend(temp_df.columns)
        plt.savefig("SingularGraph" + "_" + str(chart_cat) + ".png")
        plt.close()


    def all_together_now(self, df, title):  # first input must be army dataframe
        '''
            This creates a grouped bar graph for all services. Then uses the column totals to make the graph
        :param df: a dataframe with all services
        :return: filename(saved as title_AllServices.png)
        '''
        # Grouped Bar Chart
        fig, ax = plt.subplots(figsize=(17, 7), dpi=80)
        names = list(df.columns)

        bars1 = list(df['Air Force'])
        bars2 = list(df['Army'])
        bars3 = list(df['Marines'])
        bars4 = list(df['Navy'])
        bars5 = list(df['Coast Guard'])

        r = list(range(12))
        width = 0.15  # the width of the bars
        r1 = np.arange(len(r))  # the label locations
        r2 = [x + width for x in r1]
        r3 = [x + width for x in r2]
        r4 = [x + width for x in r3]
        r5 = [x + width for x in r4]

        rects1 = ax.bar(r1, bars1, width, label=names[0], color='dodgerblue')
        rects2 = ax.bar(r2, bars2, width, label=names[1], color='darkorange')
        rects3 = ax.bar(r3, bars3, width, label=names[2], color='gray')
        rects4 = ax.bar(r4, bars4, width, label=names[3], color='gold')
        rects5 = ax.bar(r5, bars5, width, label=names[4], color='mediumblue')

        # Add some text for labels, title and custom x-axis tick labels, etc.
        ax.set_ylabel('Quota Totals')
        #ax.set_ylim([0, 175])
        ax.set_title(str(title) + ' - All Services')
        ax.set_xticks([w + width*1.9 for w in range(len(bars1))])
        ax.set_xticklabels(list(df.index))
        ax.set_xlabel('Months')
        ax.legend()

        def autolabel(rects):
            """Attach a text label above each bar in *rects*, displaying its height."""
            for rect in rects:
                height = rect.get_height()
                ax.annotate('{}'.format(height),
                            xy=(rect.get_x() + rect.get_width() / 2, height),
                            xytext=(0, 3),  # 3 points vertical offset
                            textcoords="offset points",
                            ha='center', va='bottom')

        autolabel(rects1)
        autolabel(rects2)
        autolabel(rects3)
        autolabel(rects4)
        autolabel(rects5)

        plt.savefig(str(title) + "_AllServices.png")
        plt.close()


        # workbook = openpyxl.Workbook('Throughput Report')
        # worksheet = workbook.active  # worksheet = writer.create_sheet['All Services']
        # worksheet.add_chart('A1', 'AllServices.png')


    def army_component_charts(self, df, title):
        '''
            This creates a grouped bar graph for each component of the Army. Trying to make the above work for
            all inputs.
        :param df: inputted dataframe
        :param title: Name to save chart for filepath
        :return: filepath for chart (saved as title_ArmyCompo.png)
        '''
        # Grouped Bar Chart
        fig, ax = plt.subplots(figsize=(17, 7), dpi=80)
        names = list(df.columns)

        bars1 = list(df[:, 0])
        bars2 = list(df[:, 1])
        bars3 = list(df[:, 2])


        r = list(range(12))
        width = 0.15  # the width of the bars
        r1 = np.arange(len(r))  # the label locations
        r2 = [x + width for x in r1]
        r3 = [x + width for x in r2]
        r4 = [x + width for x in r3]
        r5 = [x + width for x in r4]

        rects1 = ax.bar(r1, bars1, width, label=names[0], color='dodgerblue')
        rects2 = ax.bar(r2, bars2, width, label=names[1], color='darkorange')
        rects3 = ax.bar(r3, bars3, width, label=names[2], color='gray')
        # rects4 = ax.bar(r4, bars4, width, label=names[3], color='gold')
        # rects5 = ax.bar(r5, bars5, width, label=names[4], color='mediumblue')

        # Add some text for labels, title and custom x-axis tick labels, etc.
        ax.set_ylabel('Quota Totals')
        # ax.set_ylim([0, 175])
        ax.set_title(str(title) + ' - All Services')
        ax.set_xticks([w + width * 1.9 for w in range(len(bars1))])
        ax.set_xticklabels(list(df.index))
        ax.set_xlabel('Months')
        ax.legend()

        def autolabel(rects):
            """Attach a text label above each bar in *rects*, displaying its height."""
            for rect in rects:
                height = rect.get_height()
                ax.annotate('{}'.format(height),
                            xy=(rect.get_x() + rect.get_width() / 2, height),
                            xytext=(0, 3),  # 3 points vertical offset
                            textcoords="offset points",
                            ha='center', va='bottom')

        autolabel(rects1)
        autolabel(rects2)
        autolabel(rects3)
        # autolabel(rects4)
        # autolabel(rects5)

        plt.savefig(str(title) + "_AllServices.png")
        plt.close()

