Throughput Report Calculator

Throughput Report 1_07.exe = 78 MB
Throughput Report 1_07 = 7.8 MB
Throughput Report 1_07 Compressed File = 33 MB

The executable file was created using PyInstaller in Python. There are two builds for the executable program. First, one executable file that contains all libraries used (below). 
Second option is a compressed file with each library in the file including the executable program. The singular file takes longer to run than the second file.
Each file size is due to the packaging of the python libraries required for running the program. These libraries are completely contained within the file/executable program. 
The Python language is not required on a machine to run the programs.


Python libraries required to run program:
-tkinter
-pandas
-dateutil.rrule
-datetime
-collections
-xlrd
-xlsxwriter
-re
-calendar
-matplotlib
-numpy


This project calculates the following by inputting excel spreadsheets:
-Max number of students expected per month
-Max number of students expected per FY quarter
-Max number of students per quota per month
-Max number of students per quota per FY quarter
-Max number of students per service
	- Student numbers by component, entry-level/cross-training students, gender
	- Max number for IET courses


Bug Fixes:
- One service's numbers were inflated. Modified the calculation algorithm.
- All quotas are calculated by report date
- Fixed IET Services tables were adding other category of numbers 
- Fixed how charts were imported into the excel spreadsheet
- Fixed: IET courses are not required for the program to run
- Getting fiscal year string using fiscal_year module inside throughput_calculator package
- Add fiscal year to filepath of Throughput Report.xlsx
- Does not needs full year of class data to tabulate Quarter Totals


Additonal Features:
- The program now recognizes if any IET courses have been inputted
- Categorizes specific services numbers by component and gender
- Added Gender numbers for IET students 
- Added 3 sheets to excel: A Only Quotas, Non-A Quotas, All Services & Graphs
- Creates two graphs: All Services Total Numbers, IET All Services Total Numbers


New Dependent Modules in Package:
- group_quota_sum
- throughput_graphs
- fiscal_year


Installation requirements:
-OS: Windows 8
	-Dependent on Visual C++ Redistributable 2015 install
-OS:Windows 10

Run Program:
1. Double click file
2. A window will open. Type in the fiscal year needed for calculations. Must be in the format of YYYY. If format is incorrect, you will be notified and can reinput the year.
3. Click 'Select Excel Files'. File explorer will open. Select excel spreadsheets required for calculations ("shift + file", or "ctrl + file" to select multiple files)
4. Click 'Calculate'. If program runs successfully, a pop-up message will say the program was successful.
5. An excel file ("Throughput Report 20XX.xlsx") will be created in the same location as the program.
6. Open excel file for calculation totals.
7. File must be renamed prior to running subsequent times. If reran prior to saving, the program will write over the previous file.


Limitation:
Needs full year of data to tabulate monthly totals.
