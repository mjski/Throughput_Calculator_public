################################ -- Throughput Report 1_07 (Version 7) -- ##############################################
#                                           Author: Morgan D
#
# Updates to previous editions since last release:
#       * FIXED: Student numbers were inflated. Now the program calculates the row first, then the maxes for each column.
#                Error per quota within 5 (max total numbers correct)
#       * FIXED: All quotas now calculate by report date (requested via det)
#       * Created new function and class (group_quota_sum.py) that finds components and creates
#         dataframe gendered_a_df. This then is written to a third sheet of the excel file.
#       * Created throughput_graphs
#       * Program now recognizes IET courses (BTVEM and MCF, MCF +4) if submitted but will run without IET input
#       * Program now recognizes MOS-I and MOS-T numbers
#       * Program now calculates by gender and entry-level/cross-trained for IET students
#           ***Did not hard code the placement numbers. This is likely to change and would need consistent adjustment.
# Edition PURPOSE:
#       * FIXED: fix charts then add all graphs to excel file 'Throughput Report'
#                       Only using All Service Charts. Adding more compromised efficiency.
#       * FIXED: Some service numbers were showing up in Non-Service Specific Quotas sheets in IET Totals.
#
#  Version Python: 3.7
#  Created on 25 Feb 2020
########################################################################################################################

from pandas import DataFrame, read_excel, to_datetime, to_numeric, ExcelWriter, concat, options
from dateutil.rrule import MONTHLY, rrule
from datetime import datetime
from collections import defaultdict
from tkinter import filedialog, messagebox, Tk, Label, Entry, Button, X
import throughput_calculator.fiscal_year_str as fys
import re
from throughput_calculator.group_quota_sum import GroupQuotasSum as gqs
from throughput_calculator.throughput_graphs import ThroughputGraphs as tg
from xlrd import open_workbook


files = []  # Declaring a global files list because its the easiest way to get a return from a Tkinter button.
courseNumbers = []
totalcourseDict = {}
QuotaFinal = DataFrame()


def GO_button():  # Keeps the GUI root box open, checks for numeric input within a 4 digit range with error messages included
    global txt    # Global variables are needed because the GUI does not allow for passing variables
    global fyyear
    global files
    fyyear = txt.get()  # This retrieves the input from the GUI & user into the fyyear variable
    if files:
        if fyyear.isnumeric():
            fyyear = int(fyyear)
            if 2000 <= fyyear <= 3000:  # makes sure the input is in this range
                root.destroy()  # closes GUI window when all inputs are correct
            else:
                messagebox.showinfo('Error', "That was not a valid year. Please try again.", icon='warning')  # if not in range, error message pops up
        else:
            messagebox.showinfo('Error', "That was not a valid year. Please try again.", icon='warning')  # if input is not a number, error message pops up
    else:
        messagebox.showinfo('Error', "Please select Excel file(s).", icon='warning')  # If no files selected, error message pops up


def getexcel():  # Gets the file dialog, gets excel files (can select multiple files)
    global files
    files = filedialog.askopenfilenames(initialdir='%HOMEPATH', filetypes=(("Excel Files", "*.xlsx"), ("Excel Files", "*.xls"), ('All Files', '*.*')))


def formatframe(file):    # Takes a data frame as input. Formats it to strip bad info.
    AllQuotaNames = []
    df = read_excel(file, drop=False)  # converts each excel file to dataframe
    df.rename(columns={df.columns[0]: "New"}, inplace=True)  # renames first column as New
    findindex = list(df['New'].str.match('Class'))  # finds string "Class" then makes list of booleans
    for i in findindex:  # If the item is True, this prints the index of the True
        if i is True:
            new = int(findindex.index(i))
    try:  # Checks that the correct spreadsheet has been inputted
        new_header = df.iloc[new]  # makes header of df row that contains "Class" in excel
    except UnboundLocalError:  # When spreadsheet is incorrect this error message pops up
        messagebox.showinfo('Error', "Improperly formatted spreadsheet.", icon='warning')
        exit()
    df = df[new+1:]  # Using only cells from header and below
    df = df[:-1]
    df.columns = new_header  # names all columns to the cell names in the header
    df = df.loc[:, df.columns.notnull()]  # Does not include columns that are null (last few columns are null of df)
    df.fillna(0, inplace=True)  # Makes all NaN zeros
    df.drop(columns=['Location'], axis=1, inplace=True)
    df = df.where(df['Class Flag'] != 'C')  # Does not include rows where classes have been cancelled
    df = df.where(df['Class Flag'] != 'N')  # Does not include rows where classes have been non-conducted
    df.dropna(axis=0, how='all', inplace=True)
    index = df.set_index('Class', drop=False)
    QuotaNames = list(df.columns[8:])
    dict1 = df.T.to_dict()  # Makes df into dictionary with nested keys and values
    AllQuotaNames += QuotaNames
    AllQuotaNames = set(AllQuotaNames)
    return dict1, df, index, AllQuotaNames


def formatdict(dict1):      # Formats the dict to have useful info in datetime format
    # FIXME See note below
    '''
        Pulls monthly max numbers from Allocated column
    :param dict1: dict1 is a dictionary as input here  (FIX THIS LATER, WORKS FOR NOW)
    :return: dict1 with only the below columns
    '''
    for k, v in dict1.items():
        tempDict = {}
        tempDict['Start Date'] = datetime.strptime(v.get('Start Date'), '%Y-%m-%d')  # makes these datetime objects
        tempDict['End Date'] = datetime.strptime(v.get('End Date'), '%Y-%m-%d')  # makes these datetime objects
        tempDict['Allocated'] = v.get('Allocated')
        dict1[k] = tempDict  # Makes a nested Dictionary Class on the outside. Start / End / Allocated
    return dict1


def formatFields(df):  # Formats the Report, Start, End dates into datetime objects, makes all other columns numeric
    IntColNames = list(df.columns[6:])
    df['Report Date'] = to_datetime((df['Report Date']))
    df['Start Date'] = to_datetime((df['Start Date']))
    df['End Date'] = to_datetime((df['End Date']))
    df[IntColNames] = df[IntColNames].apply(to_numeric, errors='coerce')
    return df


def domath(dict1):
    templist = []
    for k, v in dict1.items():  # Lists the months a course is running, creates a tuple list of month/Allocated seats
        for dt in rrule(freq=MONTHLY, dtstart=v.get('Start Date'), until=v.get('End Date')):
            templist.append([datetime.strftime(dt, '%b-%Y'), int(v.get('Allocated'))])

    new_dict = defaultdict(list)
    for k, v in templist:  # Makes templist a defaultdict
        new_dict[k].append(v)  # Appends Allocated number of seats to each month the course runs
    final_dict = dict(new_dict)  # defaultdict into standard dictionary
    total_dict = {k: sum(map(int, v)) for k, v in final_dict.items()}  # sums the Allocated seats for each month
    return total_dict


def Join(totaldict):  # Takes the dictionary from domath and makes a long list of all of the classes for each iteration
    for item in totaldict:
        courseNumbers.append(item)
    return courseNumbers  # must be global variable to append each dictionary to


def Convert(courseNumbers):  # Takes the list from Join and makes it a sorted dictionary
    for a, b in courseNumbers:
        totalcourseDict.setdefault(a, []).append(b)
    return totalcourseDict


def ClassSum(totalcourseDict):  # Sums all values for the totalcourseDict
    FinalNumbers = {k: sum(map(int, v)) for k, v in totalcourseDict.items()}
    return FinalNumbers


def Get_FY_keys(MonthSum):  # MonthSum is a dictionary
    FYmonthSum = []
    global fyyear
    fiscal_year = fys.FiscalYearStr(fyyear)
    c = fiscal_year.get_cal()
    months = fiscal_year.get_fy(c)

    #months = [('Oct-' + str(fyyear - 1)), ('Nov-' + str(fyyear - 1)), ('Dec-' + str(fyyear - 1)), ('Jan-' + str(fyyear)), ('Feb-'+ str(fyyear)), ('Mar-' + str(fyyear)),
              #('Apr-' + str(fyyear)), ('May-' + str(fyyear)), ('Jun-' + str(fyyear)), ('Jul-' + str(fyyear)), ('Aug-' + str(fyyear)), ('Sep-' + str(fyyear))]
    for k, v in MonthSum.items():
        for i in months:
            if k == i:
                pair = (datetime.strptime(k, '%b-%Y').strftime('%Y-%m'), v)  # Doing this made it possible to sort by date down below.
                FYmonthSum.append(pair)
    FYmonthSum = sorted(FYmonthSum)  # Sorts the months in the list from Oct-Sept
    return(FYmonthSum)


def Add_Sub_Quotas(df):   # number of quotas increases for start times
    QNames = list(df.columns[8:])  # Makes list of the dataframe column names from column 8 to the end
    starts = DataFrame()
    starts["date"] = df['Report Date']  # Makes dataframe of only report dates, quotas, and column names
    for i in QNames:
        starts[str(i)] = df[str(i)]

    finishes = DataFrame()
    finishes["date"] = df['End Date']
    for i in QNames:
        finishes[str(i)] = -df[str(i)]  # Makes dataframe of only end dates, subtracts quotas, and column names
    return starts, finishes


def Finished_Quotas(QuotaFinal):
    QuotaFinal.fillna(0, inplace=True)  # Any quota that doesn't exist in the second df will hold a zero value
    QuotaFinal = QuotaFinal.groupby(['date']).sum()  # Any quota that has the same date will add quotas together
    #QuotaFinal = QuotaFinal.sort_values(by=['date'])
    QuotaFinal = QuotaFinal.cumsum()  # Cumulative Sum for all quotas by date
    QuotaFinal2 = QuotaFinal  # New dataframe in case I need to find max by day later
    QuotaFinal2.index = QuotaFinal.index.strftime('%Y-%m')  # Changes indexed date into year-month format
    # QuotaFinal2 --- This is where Army quotas need to be calculate by row

    army_only_df = QuotaFinal2.sort_index()  # copy of QuotaFinal2 sorted by index
    col_list = list(army_only_df.columns)    # made a list of all column names

    q = ('W', 'M', 'P', 'T', 'N')    # list of first initial of all Army quota sources
    not_army_list = [x for x in col_list if not x.startswith(q)]  # list comp finding all column names that are not Army

    army_only_df.drop(columns=not_army_list, inplace=True)   # drop all columns that are not Army
    army_only_df['Total'] = army_only_df.sum(axis=1)   # creates 'Total' column that is the sum of all values in each row
    idx = army_only_df.groupby([army_only_df.index])['Total'].transform(max) == army_only_df['Total']   # groups by index, then finds max 'Total' within that group
    army_only_df = army_only_df[idx]  #  only returns the rows that had a max total for each index (date)
    army_only_df = army_only_df.loc[~army_only_df.index.duplicated(keep='first')]  # Doing this for now, keeping first index if multiple indexes - might want to see if I can find the ceiling(mean) of both rows

    QuotaFinal2 = QuotaFinal2.groupby([QuotaFinal2.index]).agg(['max'])  # Groups all dates by date index month, then aggregates the max of each row in the group
    ''' The aggregate max is finding the highest number in each column by month. This is why the numbers appear higher each month. 
        We need to find out if this is how we want the numbers to calculate or find the max altogether for a row and print that line.
        This would get much more complicated as some quotas we want the column max (such as UR, US, UM) but Army we do not.'''

    startdate = to_datetime(str(fyyear-1)+"-9-01").date()  # did the month prior so that October prints
    enddate = to_datetime(str(fyyear)+"-9-30").date()
    army_only_df = army_only_df[str(startdate):str(enddate)]  # only need fiscal year for army quotas
    QuotaFinal2 = QuotaFinal2[str(startdate):str(enddate)]  # Only contains the months of the selected fiscal year
    if QuotaFinal2.empty:
        messagebox.showinfo('Error', 'No classes in the given fiscal year.', icon='warning')
        exit()
    return QuotaFinal2, army_only_df, not_army_list


def new_quota_final(QuotaFinal2, army_only_df, not_army_list):
    no_army_df = QuotaFinal2[not_army_list]  # only the non-Army columns
    no_army_df.columns = no_army_df.columns.droplevel(-1)  # drops MAX on second level header
    abs_final = no_army_df.join(army_only_df)  # joins the non-Army columns to the Army only columns
    abs_final_df = abs_final.iloc[:, :-1]  # drops the total column from the army_only data frame after the join
    return abs_final_df


def GetFY_QuarterSums(MonthTotal):
    global fyyear
    try:
        sub0 = str(fyyear - 1)
        MonthTotal['Indexes1'] = MonthTotal["Month"].str.find(sub0)
        A_FYQ1 = MonthTotal.loc[MonthTotal['Indexes1'] == 0]
        sub1 = str(fyyear) + '-01'
        sub2 = str(fyyear) + '-02'
        sub3 = str(fyyear) + '-03'
        sub4 = str(fyyear) + '-04'
        sub5 = str(fyyear) + '-05'
        sub6 = str(fyyear) + '-06'
        sub7 = str(fyyear) + '-07'
        sub8 = str(fyyear) + '-08'
        sub9 = str(fyyear) + '-09'
        MonthTotal['Indexes2'] = (MonthTotal['Month'].str.find(sub1)) & (
                    MonthTotal['Month'].str.find(sub2) & (MonthTotal['Month'].str.find(sub3)))
        A_FYQ2 = MonthTotal.loc[MonthTotal['Indexes2'] == 0]
        MonthTotal['Indexes3'] = (MonthTotal['Month'].str.find(sub4)) & (
                    MonthTotal['Month'].str.find(sub5) & (MonthTotal['Month'].str.find(sub6)))
        A_FYQ3 = MonthTotal.loc[MonthTotal['Indexes3'] == 0]
        MonthTotal['Indexes4'] = (MonthTotal['Month'].str.find(sub7)) & (
                    MonthTotal['Month'].str.find(sub8) & (MonthTotal['Month'].str.find(sub9)))
        A_FYQ4 = MonthTotal.loc[MonthTotal['Indexes4'] == 0]
        FYQ1 = A_FYQ1['Total'].max()  # Max for only the first quarter Total in MonthTotal dataframe
        FYQ2 = A_FYQ2['Total'].max()  # Max for only the second quartet Total in MonthTotal dataframe
        FYQ3 = A_FYQ3['Total'].max()  # Max for only the third quarter Total in MonthTotal dataframe
        FYQ4 = A_FYQ4['Total'].max()  # Max for only the fourth quarter Total in MonthTotal dataframe
        #  The calculation bug for the first page of excel spreadsheet has been fixed. Will calculate even if all months are not used.

        MonthTotal.drop(['Indexes1', 'Indexes2', 'Indexes3', 'Indexes4'], axis=1, inplace=True)
    except ValueError:
        messagebox.showinfo('Error', 'There were not enough classes inputted to calculate the fiscal year.', icon='warning')
        exit()
    return FYQ1, FYQ2, FYQ3, FYQ4


def Get_Quota_QuarterSums(QuotaFinal2):  # Need to fix so that this works even if all months aren't filled.
    #QuotaFinal2.columns = QuotaFinal2.columns.droplevel(-1)  # dropped header line of all 'max' NO LONGER NEEDED WITH ARMY JOIN TABLE
    QuotaQ1 = QuotaFinal2.iloc[0:3, :].cummax()              # Finds Max for all columns from rows 1-3
    Q1QuotaMax = QuotaQ1.tail(1).reset_index(drop=True)      # Returns only the last row (total sum; on lines 168, 170, 172, 174)
    QuotaQ2 = QuotaFinal2.iloc[3:6, :].cummax()              # Finds Max for all columns from rows 3-6
    Q2QuotaMax = QuotaQ2.tail(1).reset_index(drop=True)
    QuotaQ3 = QuotaFinal2.iloc[6:9, :].cummax()              # Finds Max for all columns from rows 6-9
    Q3QuotaMax = QuotaQ3.tail(1).reset_index(drop=True)
    QuotaQ4 = QuotaFinal2.iloc[9:12, :].cummax()             # Finds Max for all columns from rows 9-12
    Q4QuotaMax = QuotaQ4.tail(1).reset_index(drop=True)
    return Q1QuotaMax, Q2QuotaMax, Q3QuotaMax, Q4QuotaMax


def gender_count(abs_final_df):  # counts Army quotas by gender (will print to excel on sheet 3 named '___')
    col_list = list(abs_final_df.columns)
    nat_g_male = []
    nat_g_female = []
    nat_g_ungender = []
    active_male = []
    active_female = []
    active_ungender = []
    reserve_male = []
    reserve_female = []
    reserve_ungendered = []
    air_force = []
    navy = []
    marines = []
    coasties = []
    intl = []
    civ = []
    dinfos = []
    other = []

    for i in col_list:
        '''
            Below is using re library (REGEX) to create lists of 
            column names that fit the requirements.
        '''
        nat_g_male += re.findall(r'^N[JLN]', i)   # Using regex (starts with N and contain J, L, B after N
        nat_g_female += re.findall(r'^N[KMP]', i)
        nat_g_ungender += re.findall(r'^N[^KMPJLN]', i)
        active_male += re.findall(r'^W[JLN]', i)
        active_female += re.findall(r'^W[KMP]', i)
        active_ungender += re.findall(r'^W[^KMPJLN]', i)
        reserve_male += re.findall(r'^M[JLN]', i)
        reserve_female += re.findall(r'^M[KMP]', i)
        reserve_ungendered += re.findall(r'^M[^KMPJLN]', i)
        reserve_ungendered += re.findall(r'^[PT][CU]', i)
        air_force += re.findall(r'^U[E]', i)
        navy += re.findall(r'^U[MN]', i)
        marines += re.findall(r'^U[R]', i)
        coasties += re.findall(r'^U[S]', i)
        intl += re.findall(r'^Z[A]', i)
        civ += re.findall(r'^K[A-Z]', i)
        dinfos += re.findall(r'^0[4]', i)
        other += re.findall(r'^[^NWMPTUZK0][A-Z]', i)

    gqs().quota_col_sums('Army - Active Male', abs_final_df, active_male)
    gqs().quota_col_sums('Army - Active Female', abs_final_df, active_female)
    gqs().quota_col_sums('Army - Active Ungendered', abs_final_df, active_ungender)
    gqs().quota_col_sums('Army - National Guard Male', abs_final_df, nat_g_male)
    gqs().quota_col_sums('Army - National Guard Female', abs_final_df, nat_g_female)
    gqs().quota_col_sums('Army - National Guard Ungendered', abs_final_df, nat_g_ungender)
    gqs().quota_col_sums('Army - Reserves Male', abs_final_df, reserve_male)
    gqs().quota_col_sums('Army - Reserves Female', abs_final_df, reserve_female)
    gqs().quota_col_sums('Army - Reserves Ungendered', abs_final_df, reserve_ungendered)
    gqs().quota_col_sums('Air Force', abs_final_df, air_force)
    gqs().quota_col_sums('Navy', abs_final_df, navy)
    gqs().quota_col_sums('Marines', abs_final_df, marines)
    gqs().quota_col_sums('Coast Guard', abs_final_df, coasties)
    gqs().quota_col_sums('International', abs_final_df, intl)
    gqs().quota_col_sums('Civilians', abs_final_df, civ)
    gqs().quota_col_sums('DINFOS', abs_final_df, dinfos)
    gqs().quota_col_sums('Other', abs_final_df, other)


    #army_cols = [col for col in abs_final_df.columns if 'Army' in col]
    # national_cols = [col for col in abs_final_df.columns if 'National' in col]
    # reserve_cols = [col for col in abs_final_df.columns if 'Reserve' in col]
    army_cols = [col for col in abs_final_df.columns if re.match(r'Army', col)]
    if len(army_cols) > 0:
        gendered_army_df = abs_final_df.loc[:, army_cols[0]:army_cols[-1]]  # might need +1
        #abs_final_df = abs_final_df.loc[:, :army_cols[0]]
    else:
        gendered_army_df = DataFrame
    #if len(active_cols) > 0:

    other_cols = [col for col in abs_final_df.columns if re.match(r'^((?!Army).)*$', col) and re.match(r'[A-Z]+[a-z]', col)]
    if len(other_cols) > 0:
        non_army_df = abs_final_df.loc[:, other_cols[0]:]  # might need + 1
    else:
        non_army_df = DataFrame

    quota_list = [col for col in abs_final_df.columns if re.match(r'[A-Z][A-Z]', col)]
    abs_final_df = abs_final_df.loc[:, quota_list[0]:quota_list[-1]]

    return gendered_army_df, non_army_df, abs_final_df, army_cols


def get_IET_list(InputFileList):
    '''
        The idea is to produce a list of only the inputted files that contain initial entry training (IET) students.
    :param InputFileList: The list of all files inputted in the beginning
    :return: IET_List: the list of all files that contain IET students
    '''
    IET_list = []
    for i in InputFileList:
        workbook = open_workbook(filename=i)
        ws = workbook.sheet_by_index(0)

        for cell in ws.row(3):
            if re.match(r'DINFOS-BTVEM', cell.value) or re.match(r'DINFOS-MCF', cell.value):
                IET_list.append(i)
    return IET_list


def make_charts(df1, df2, army_list, string):
    '''
        Now have the problem of filenames overwriting the charts...need to think about this, Maybe add strings to inputs
        should be able to use the function multiple times for DRY
    :param gendered_army_df: This will be the Army dataframe
    :param non_army_df: This will be the non-Army dataframe
    :return: all services combined dataframe and each chart filepath
    '''

    all_services_df = df2.copy()
    if 'Total' in df1.columns:
        pass
    else:
        df1['Total'] = df1.sum(axis=1)
    all_services_df.insert(1, 'Army', df1['Total'])
    all_services_df = all_services_df.astype(int)

    '''
    if re.match('IET.*', string.upper()):
        for x in army_list:
            if re.findall('Reserve.*', x):
                tg().make_army_graphs(df1, str(string) + "_reserve")
            elif re.findall('Active.*', x):
                tg().make_army_graphs(df1, str(string) + "_active")
            elif re.findall('Nation.*', x):
                tg().make_army_graphs(df1, str(string) + "_national")
            else:
                pass
    '''
    '''
    for x in df2.columns:
        if re.findall('Air.*', x):
            tg().singular_graph(df2, str(string) + 'Air Force')
        elif re.findall('Navy.*', x):
            tg().singular_graph(df2, str(string) + 'Navy')
        elif re.findall('Marine.*', x):
            tg().singular_graph(df2, str(string) + 'Marines')
        elif re.findall('Coast.*', x):
            tg().singular_graph(df2, str(string) + 'Coast Guard')
    '''

    tg().all_together_now(all_services_df, string)

    return all_services_df


def get_IET_count(IET_list):
    '''
        Take that list and do all the same calculations as the InputFileList but with the IET_List
    :param IET_list: list passed from get_IET_list function
    :return: a dataframe with counts
    '''
    IETQuotaFinal = DataFrame()
    find_non_IET = []
    MOS_T_df = DataFrame()
    if len(IET_list) == 0:
        pass
    else:
        for i in IET_list:
            IETdict1, IETdf, IETindex, IETAllQuotaNames = formatframe(i)
            IETdf = formatFields(IETdf)
            IETstarts, IETfinishes = Add_Sub_Quotas(IETdf)  # splits into two dataframes
            IETresult = concat([IETstarts, IETfinishes])  # merges the starts and the finishes dataframes
            IETresult = IETresult.sort_values(by=['date'])  # sorts the date values
            IETresult = IETresult.groupby(['date'],
                                    as_index=False).sum()  # If there is several starts or finishes at the same time, they are summed
            IETQuotaFinal = concat([IETQuotaFinal, IETresult], axis=0, ignore_index=True,
                                sort=False)  # merges IETQuotaFinal and result dataframes
            IETQuotaFinal.fillna(0, inplace=True)
            IETQuotaFinal.sort_values('date', inplace=True)

        NAQuotaFinal2, IET_df, IET_not_army_list = Finished_Quotas(IETQuotaFinal)
        for quota in IET_df.columns:
            find_non_IET += re.findall(r'^[MNPTW][BCDFLMU]', quota)
        IET_abs_final_df = new_quota_final(NAQuotaFinal2, IET_df, IET_not_army_list)
        IET_gendered_army_df, IET_non_army_df, IET_abs_final_df, IET_army_cols = gender_count(IET_abs_final_df)
        IET_gendered_army_df['Total'] = IET_gendered_army_df.sum(axis=1)
        for x in find_non_IET:
            MOS_T_df[str(x)] = IET_df.loc[:, x]
        #MOS_T_df['Total'] = MOS_T_df.sum(axis=1)
        MOS_I_df = IET_df.drop((IET_df.loc[:, find_non_IET].columns), axis=1)
        MOS_I_df.drop(['Total'], axis=1, inplace=True)

        MOS_T_gendered_df, MOS_T_non_army, MOS_T_not_needed, MOS_T_army_cols = gender_count(MOS_T_df)
        MOS_I_gendered_df, MOS_I_non_army, MOS_I_not_needed, MOS_I_army_cols = gender_count(MOS_I_df)
        quota_list = [col for col in MOS_I_df.columns if re.match(r'[A-Z][A-Z]', col)]
        MOS_I_df = MOS_I_df.loc[:, quota_list[0]:quota_list[-1]]
        T_quota_list = [col for col in MOS_T_df.columns if re.match(r'[A-Z][A-Z]', col)]
        MOS_T_df = MOS_T_df.loc[:, T_quota_list[0]:T_quota_list[-1]]
        MOS_T_df['Total'] = MOS_T_df.sum(axis=1)
        MOS_I_gendered_df['Total'] = MOS_I_gendered_df.sum(axis=1)
        MOS_T_gendered_df['Total'] = MOS_T_gendered_df.sum(axis=1)
        MOS_I_df['Total'] = MOS_I_df.sum(axis=1)

        return IET_df, IET_non_army_df, IET_gendered_army_df, MOS_T_df, MOS_I_df, MOS_T_gendered_df, MOS_I_gendered_df, IET_army_cols


def gendered_split(df):
    '''
        This takes a dataframe and returns the totals of each gender (with ungendered) in a new dataframe.
    :param df: Any dataframe that has been sorted by gender
    :return: Dataframe of gender totals
    '''
    if df.empty == False:
        options.mode.chained_assignment = None
        temp_list = []
        # if dataframe exists then function will find all columns with 'Male, Female, Ungendered' in the col names
        g_split_male = [col for col in df.columns if re.findall('Male.*', col)]
        g_split_female = [col for col in df.columns if re.findall('Female.*', col)]
        g_split_ungendered = [col for col in df.columns if re.findall('Ungendered.*', col)]

        # Checks that the lists are not empty, makes new df, sums in 'Total' column
        if len(g_split_male) > 0:
            g_split_m = df[g_split_male]
            g_split_m['Male'] = g_split_m.sum(axis=1)
            temp_list.append(g_split_m['Male'])
        else:
            pass

        if len(g_split_female) > 0:
            g_split_f = df[g_split_female]
            g_split_f['Female'] = g_split_f.sum(axis=1)
            temp_list.append(g_split_f['Female'])
        else:
            pass

        if len(g_split_ungendered) > 0:
            g_split_u = df[g_split_ungendered]
            g_split_u['Ungendered'] = g_split_u.sum(axis=1)
            temp_list.append(g_split_u['Ungendered'])
        else:
            pass

        # Takes totals for each gender and makes a new dataframe
        g_split_df = concat(temp_list, axis=1)
        return g_split_df


def WriteToExcel(FYQ1, FYQ2, FYQ3, FYQ4, Q1QuotaMax, Q2QuotaMax, Q3QuotaMax, Q4QuotaMax, abs_final_df, MonthTotal,
                 gendered_army_df, non_army_df, all_services_df, IET_df=DataFrame, IET_not_army_df=DataFrame,
                 IET_gendered_army_df=DataFrame, MOS_T_df=DataFrame, MOS_I_df=DataFrame, MOS_T_gendered_df=DataFrame,
                 MOS_I_gendered_df=DataFrame, IET_gender_split=DataFrame, MOS_I_gender_split=DataFrame,
                 MOS_T_gender_split=DataFrame, IET_all_services_df=DataFrame):
    global fyyear
    #non_army_df = non_army_df.drop(['Army'], axis=1)
    FYQs = {'Q1 Max': [str(FYQ1)], 'Q2 Max': [str(FYQ2)], 'Q3 Max': [str(FYQ3)], 'Q4 Max': [str(FYQ4)]}  # Made into dictionary to convert to dataframe quickly
    FYTotals = DataFrame(FYQs)  # Convert above dictionary to dataframe
    text1 = 'Q1 Quota Maxes'
    text2 = 'Q2 Quota Maxes'
    text3 = 'Q3 Quota Maxes'
    text4 = 'Q4 Quota Maxes'
    text5 = 'Max Quotas per Month:'
    text6 = 'Total Students per Quarter:'
    text7 = 'Total Students per Month:'
    text8 = 'Army Max Quotas per Month and Gender (calculated by Report Date):'
    text9 = 'Max Quotas per Month (calculated by Report Date):'
    writer = ExcelWriter("Throughput Report " + str(fyyear) + ".xlsx", engine='xlsxwriter')  # Writes everything to the Throughput Report.xlsx file
    FYTotals.to_excel(writer, sheet_name='Month Totals by Start Date', startrow=1, header=True, index=False)  # Writes FYTotals to Month Totals sheet
    MonthTotal.to_excel(writer, sheet_name='Month Totals by Start Date', startrow=5, header=True, index=False)  # Writes MonthTotal to Month Totals sheet
    worksheet = writer.sheets['Month Totals by Start Date']
    worksheet.write(0, 0, text6)  # Inserts text to this location in the Month Totals sheet (lines 191 & 192)
    worksheet.write(4, 0, text7)
    Q1QuotaMax.to_excel(writer, sheet_name='FY Quota Totals', startrow=1, header=True, index=False)  # Writes Q1QuotaMax to FY Quota Totals sheet (lines 193-197)
    Q2QuotaMax.to_excel(writer, sheet_name='FY Quota Totals', startrow=5, header=True, index=False)
    Q3QuotaMax.to_excel(writer, sheet_name='FY Quota Totals', startrow=9, header=True, index=False)
    Q4QuotaMax.to_excel(writer, sheet_name='FY Quota Totals', startrow=13, header=True, index=False)
    abs_final_df.to_excel(writer, sheet_name='FY Quota Totals', startrow=17, header=True, index=True)
    worksheet = writer.sheets['FY Quota Totals']
    worksheet.write(0, 0, text1)  # Inserts text to this location in the FY Quota Totals sheet (lines 199 - 203)
    worksheet.write(4, 0, text2)
    worksheet.write(8, 0, text3)
    worksheet.write(12, 0, text4)
    worksheet.write(16, 0, text5)
    gendered_army_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=2, startcol=0, header=True, index=True)
    worksheet = writer.sheets['Army Only Quotas']
    worksheet.write(0, 0, text8)

    if IET_df.empty == False:
        worksheet = writer.sheets['Army Only Quotas']
        IET_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=18, startcol=0, header=True, index=True)
        text10 = 'IET Army Totals:'
        text11 = 'MOS-I Totals:'
        text12 = 'MOS-T Totals:'
        text13 = 'IET Gendered Totals:'
        text14 = 'MOS-I Gendered Totals:'
        text15 = 'MOS-T Gendered Totals:'
        text16 = 'IET by gender:'
        text17 = 'MOS-I by gender:'
        text18 = 'MOS-T by gender:'
        text19 = 'IET - All Services:'
        text20 = 'DINFOS - All Services & Courses:'
        worksheet.write(17, 0, text10)
        IET_gendered_army_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=34, startcol=0, header=True, index=True)
        worksheet.write(33, 0, text13)
        IET_gender_split.to_excel(writer, sheet_name='Army Only Quotas', startrow=34, startcol=12, header=True, index=True)
        worksheet.write(33, 12, text16)
        MOS_I_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=51, startcol=0, header=True, index=True)
        worksheet.write(50, 0, text11)
        MOS_I_gendered_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=67, startcol=0, header=True, index=True)
        worksheet.write(66, 0, text14)
        MOS_I_gender_split.to_excel(writer, sheet_name='Army Only Quotas', startrow=67, startcol=12, header=True, index=True)
        worksheet.write(66, 12, text17)
        MOS_T_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=83, header=True, index=True)
        worksheet.write(82, 0, text12)
        MOS_T_gendered_df.to_excel(writer, sheet_name='Army Only Quotas', startrow=99, header=True, index=True)
        worksheet.write(98, 0, text15)
        MOS_T_gender_split.to_excel(writer, sheet_name='Army Only Quotas', startrow=99, startcol=12, header=True, index=True)
        worksheet.write(98, 12, text18)

        IET_not_army_df.to_excel(writer, sheet_name='Non-Army Quotas', startrow=18, header=True, index=True)
        worksheet = writer.sheets['Non-Army Quotas']
        text11 = 'IET Totals:'
        worksheet.write(17, 0, text11)

        IET_all_services_df.to_excel(writer, sheet_name='All Services & Graphs', startrow=51, header=True, index=True)
        worksheet = writer.sheets['All Services & Graphs']
        worksheet.write(50, 0, text19)
        worksheet.insert_image('A66', 'IET_AllServices.png')
    non_army_df.to_excel(writer, sheet_name='Non-Army Quotas', startrow=2, header=True, index=True)
    worksheet = writer.sheets['Non-Army Quotas']
    worksheet.write(1, 0, text9)

    all_services_df.to_excel(writer, sheet_name='All Services & Graphs', startrow=1, header=True, index=True)
    worksheet = writer.sheets['All Services & Graphs']
    worksheet.write(0, 0, text20)
    worksheet.insert_image('A16', 'DINFOS_AllServices.png')

    try:
        writer.save()  # this checks to see if the excel file is open
    except PermissionError:
        messagebox.showinfo('Error', 'Please close the Throughput Report Excel Spreadsheet.', icon='warning')  # If file is open an error message pops up and ends program
        exit()
    writer.close()


# Basically the main funcs under here. Builds GUI and calls other functions.
root = Tk()
root.title('Throughput Report Calculator')                     #title of main root window
lbl = Label(root, text='What fiscal year (YYYY)? ', font=('helvetica', 12))  # Text request for year input
lbl.pack(fill=X, padx=50, pady=10)                  # the location of where the label is in the root window
txt = Entry(root, width=4, font=('helvetica', 12))  # input box on root window
txt.pack(padx=30, pady=5, ipadx=4, ipady=5)            # location of input box
btn = Button(root, text='Select Excel Files', command=getexcel, bg='DeepPink3',
                               fg='white', font=('helvetica', 12, 'bold'))  # creates button that calls the getexcel function
btn.pack(fill=X, padx=50, pady=20)  # Placement of button
go_btn = Button(root, text='Calculate', command=GO_button, bg='SpringGreen3',
                               fg='white', font=('helvetica', 12, 'bold'))  # makes button that calls the GO_button function
go_btn.pack(fill=X, padx=50, pady=15)  # Placement of button
root.mainloop()

InputFileList = root.splitlist(files)  # splits the list into separate elements
#print('Files = ', InputFileList)


for i in InputFileList:             # Call functions for each imported file.
    dict1, df, index, AllQuotaNames = formatframe(i)   # formats each data frame. Returns a dict
    dict1 = formatdict(dict1)       # Formats the dict to contain useful info. Returns a dict.
    df = formatFields(df)           # Formats dataframe fields
    totaldict = domath(dict1)       # does math. Returns a dict of months and total students.
    totaldict = list(totaldict.items())  # turns dictionary into a list
    Join(totaldict)                 # calls Join function
    starts, finishes = Add_Sub_Quotas(df)  # splits into two dataframes
    result = concat([starts, finishes])  # merges the starts and the finishes dataframes
    result = result.sort_values(by=['date'])  # sorts the date values
    result = result.groupby(['date'], as_index=False).sum() # If there is several starts or finishes at the same time, they are summed
    QuotaFinal = concat([QuotaFinal, result], axis=0, ignore_index=True, sort=False) #merges QuotaFinal and result dataframes
    QuotaFinal.fillna(0, inplace=True)
    QuotaFinal.sort_values('date', inplace=True)


def main():
    QuotaFinal2, army_only_df, not_army_list = Finished_Quotas(QuotaFinal)
    Convert(courseNumbers)
    MonthSum = ClassSum(totalcourseDict)
    FY_keys = Get_FY_keys(MonthSum)
    MonthTotal = DataFrame(FY_keys, columns=['Month', 'Total'])  # dataframe of only the Month and Total columns
    FYQ1, FYQ2, FYQ3, FYQ4 = GetFY_QuarterSums(MonthTotal)
    abs_final_df = new_quota_final(QuotaFinal2, army_only_df, not_army_list)
    Q1QuotaMax, Q2QuotaMax, Q3QuotaMax, Q4QuotaMax = Get_Quota_QuarterSums(abs_final_df)
    gendered_army_df, non_army_df, abs_final_df, army_cols = gender_count(abs_final_df)

    IET_list = get_IET_list(InputFileList)
    string1 = 'DINFOS'
    all_services_df = make_charts(gendered_army_df, non_army_df, army_cols, string1)
    if len(IET_list) != 0:
        IET_df, IET_not_army_df, IET_gendered_army_df, MOS_T_df, MOS_I_df, MOS_T_gendered_df, MOS_I_gendered_df, IET_army_cols = get_IET_count(IET_list)
        IET_gender_split = gendered_split(IET_gendered_army_df)
        MOS_T_gender_split = gendered_split(MOS_T_gendered_df)
        MOS_I_gender_split = gendered_split(MOS_I_gendered_df)
        string2 = 'IET'
        IET_all_services_df = make_charts(IET_df, IET_not_army_df, IET_army_cols, string2)
        #IET_not_army_df.drop('Army')

        WriteToExcel(FYQ1, FYQ2, FYQ3, FYQ4, Q1QuotaMax, Q2QuotaMax, Q3QuotaMax, Q4QuotaMax, abs_final_df, MonthTotal,
                     gendered_army_df, non_army_df, all_services_df, IET_df, IET_not_army_df, IET_gendered_army_df,
                     MOS_T_df, MOS_I_df, MOS_T_gendered_df, MOS_I_gendered_df, IET_gender_split, MOS_I_gender_split,
                     MOS_T_gender_split, IET_all_services_df)
    else:
        WriteToExcel(FYQ1, FYQ2, FYQ3, FYQ4, Q1QuotaMax, Q2QuotaMax, Q3QuotaMax, Q4QuotaMax, abs_final_df, MonthTotal,
                     gendered_army_df, non_army_df, all_services_df)

    messagebox.showinfo('Complete!', 'The Throughput Report has been calculated and ready in the file!')  # Successful completion message box


main()



