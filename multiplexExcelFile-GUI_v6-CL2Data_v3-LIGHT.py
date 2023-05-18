# multiplex code
# GUI version of the code

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from pathlib import Path
# import warnings
# # Ignore future warnings:
# warnings.simplefilter('ignore')

''' Constants'''
# Constants used for the GUI colors
Blue = '#87B7E1'
Red = '#D0281A'
# skiprows = Rowstart
Rowstart = 6
# if spreadsheet starts with column labeled A, make Columnstart = 1, 
# else if spreadsheet starts with E make Columnstart = 2
Columnstart = 1
# change sheettitle depending on which data I am working with... 0.1 primary antibody or 1.0 primary antibody
sheettitle = 'Sheet1'

root=tk.Tk(screenName="Ian's window")
frame = tk.Frame(root, bg='black')
frame.pack()

# Attempt to change tkinter theme
style = ttk.Style(root)
style.theme_use('classic')  # I like the classic theme

# print(Path.home()) This gives the home directory
homedir = Path.home()

# declaring string variable for the file path and answers to questions
url = tk.StringVar()
savedir = tk.StringVar()
ispostprint_answer = tk.IntVar()
exposure_answer = tk.IntVar()
RowNum_answer = tk.IntVar()

nameList = [0]  # list initiated with a value so the rest of my code can work...
nameListdir = [0]
saveList = [0]
savePathlist = [0,0]
RNlist = [0]
RunList = []
dflist = [0]
df2list = [0]

'''GUI code'''
# Function for opening the file explorer window
def browseFiles():
    filename = filedialog.askopenfilename(initialdir = "c:\\",
                                          title = "Select a File",
                                          filetypes = (("Excel files",
                                                        "*.xlsx*"),
                                                        ))
    label_file_explorer.configure(text="File Opened: "+filename)
    nameList.append(filename)

# setting for the window bg color
root.config(background = "#D6DADE")

def submit():

    loadfile = str(nameList[-1])
    # loadfile = '.' + loadfile[17:]
    print("Your excel file path is : " + loadfile)
    # PostPrintData = IPPlist[-1]

    df = pd.DataFrame(pd.read_excel(loadfile, skiprows = Rowstart))
    # print (df.head(50))

    df = pd.DataFrame(pd.read_excel(loadfile, skiprows = Rowstart, usecols=lambda x: 'ID' not in x))
    # usecols=lambda x: 'ID' not in x is used to remove the first column of excel doc from dataframe, the one with strings

    # print (f'df after getting rid of ID: \n {df.head(50)}')

    jlist = []
    count = 0

    # Trying to automatically figure out limits of the data
    for i in df:    # prints out only the column names for some reason 
        for j in df[i]:   # j = Each column of data; the 108 was specific for the Bright Spark project, I think...
            count += 1
    
            if isinstance(j, str) == True: # The isinstance() function returns True if, in this case, i is not an integer
                jlist.append(count) 

    print(f'Number of rows in data = {jlist[0]}')     # This should be the nrows

    # So I need to make the df again with the nrows included
    df = pd.DataFrame(pd.read_excel(loadfile, skiprows = Rowstart, usecols=lambda x: 'ID' not in x, nrows=(jlist[0] - 1)))

    df2 = df[1:]        # This gets rid of the irrelevant number at the top of the column (Eg. 5, 6, 7, 8) which interferes with the average result
    dflist.append(df)
    df2list.append(df2)

'''End of Submit function'''
    

def save_inp():

    savePath = savedir.get()    # used to get the user input
    saveList.append(savePath)
    savePathlist.append(str(homedir) + f'\{saveList[-1]}.xlsx')      # saves the path to the savePathlist
    print(f"Your save name is : {savePath}")
    print(f'Your save path is : {savePathlist[-1]}')


def RunProgram():
    # Multiplex may not output data with multiple expsure lengths... 
    # May need to remove that from this code if that is the case...

    # Constants
    saveFile = savePathlist[-1]

    df2 = df2list[-1]
    # troubleshooting
    # print(f'{df2}')

    spotconc0 = 0           # 0.2 data

    spotlist = []
    spotlist2 = []
            
    # 0, 0 is the first spot in the first row, df2.iloc[12, 15] is the last spot in the last row.
    # This is to set all the concentration spots to the same data list

    # Rows 1 and 2 are T2(light)
    # Rows 3 and 4 are T3(Dark)
    # Rows 5 and 6 are T2(light)
    # Rows 7 and 8 are T3(Dark)
    # I need to figure out how to specify this data...

    # Need to define the appropriate for loop for light vs dark data

    # print(df2.head(25))

        # T1 (Light)
    # spotconc0 = 2   # T1 (Dark)  
    for i in range(2):   
        # it's [row, column]
        # Skipping the two blank spaces at the top of the excel document before the data by starting at row 2...

        # 0.2 data
        
  
        spotlist.append(df2.iloc[ 2 , spotconc0 ]) 
        spotlist.append(df2.iloc[ 3 , spotconc0 ])
        spotlist.append(df2.iloc[ 4 , spotconc0 ])
        spotlist.append(df2.iloc[ 5 , spotconc0 ])
        spotlist.append(df2.iloc[ 6 , spotconc0 ])
        spotlist.append(df2.iloc[ 7 , spotconc0 ])
        spotlist.append(df2.iloc[ 8 , spotconc0 ])
        spotlist.append(df2.iloc[ 9 , spotconc0 ])
        spotlist.append(df2.iloc[ 10, spotconc0 ])
        spotlist.append(df2.iloc[ 11, spotconc0 ])
        spotlist.append(df2.iloc[ 12, spotconc0 ])
        spotlist.append(df2.iloc[ 13, spotconc0 ])
        # 0.1 data
        spotlist.append(df2.iloc[ 14 , spotconc0 ])
        spotlist.append(df2.iloc[ 15 , spotconc0 ])
        spotlist.append(df2.iloc[ 16 , spotconc0 ])
        spotlist.append(df2.iloc[ 17 , spotconc0 ])
        spotlist.append(df2.iloc[ 18 , spotconc0 ])
        spotlist.append(df2.iloc[ 19 , spotconc0 ])
        spotlist.append(df2.iloc[ 20 , spotconc0 ])
        spotlist.append(df2.iloc[ 21 , spotconc0 ])
        spotlist.append(df2.iloc[ 22,  spotconc0 ])
        spotlist.append(df2.iloc[ 23,  spotconc0 ])
        spotlist.append(df2.iloc[ 24,  spotconc0 ])
        spotlist.append(df2.iloc[ 25,  spotconc0 ])
        # 0.05 data
        spotlist.append(df2.iloc[ 26 , spotconc0 ])
        spotlist.append(df2.iloc[ 27 , spotconc0 ])
        spotlist.append(df2.iloc[ 28 , spotconc0 ])
        spotlist.append(df2.iloc[ 29 , spotconc0 ])
        spotlist.append(df2.iloc[ 30 , spotconc0 ])
        spotlist.append(df2.iloc[ 31 , spotconc0 ])
        spotlist.append(df2.iloc[ 32 , spotconc0 ])
        spotlist.append(df2.iloc[ 33 , spotconc0 ])
        spotlist.append(df2.iloc[ 34,  spotconc0 ])
        spotlist.append(df2.iloc[ 35,  spotconc0 ])
        spotlist.append(df2.iloc[ 36,  spotconc0 ])
        spotlist.append(df2.iloc[ 37,  spotconc0 ])
        # 0.01 data
        spotlist.append(df2.iloc[ 38 , spotconc0 ])
        spotlist.append(df2.iloc[ 39 , spotconc0 ])
        spotlist.append(df2.iloc[ 40 , spotconc0 ])
        spotlist.append(df2.iloc[ 41 , spotconc0 ])
        spotlist.append(df2.iloc[ 42 , spotconc0 ])
        spotlist.append(df2.iloc[ 43 , spotconc0 ])
        spotlist.append(df2.iloc[ 44 , spotconc0 ])
        spotlist.append(df2.iloc[ 45 , spotconc0 ])
        spotlist.append(df2.iloc[ 46,  spotconc0 ])
        spotlist.append(df2.iloc[ 47,  spotconc0 ])
        spotlist.append(df2.iloc[ 48,  spotconc0 ])
        spotlist.append(df2.iloc[ 49,  spotconc0 ])
        # 0.00 data
        spotlist.append(df2.iloc[ 50 , spotconc0 ])
        spotlist.append(df2.iloc[ 51 , spotconc0 ])
        spotlist.append(df2.iloc[ 52 , spotconc0 ])
        spotlist.append(df2.iloc[ 53 , spotconc0 ])
        spotlist.append(df2.iloc[ 54 , spotconc0 ])
        spotlist.append(df2.iloc[ 55 , spotconc0 ])
        spotlist.append(df2.iloc[ 56 , spotconc0 ])
        spotlist.append(df2.iloc[ 57 , spotconc0 ])
        spotlist.append(df2.iloc[ 58,  spotconc0 ])
        spotlist.append(df2.iloc[ 59,  spotconc0 ])
        spotlist.append(df2.iloc[ 60,  spotconc0 ])
        spotlist.append(df2.iloc[ 61,  spotconc0 ])

        spotconc0 += 1
    
    spotconc0 = 4   # T1 (Light)
    for i in range(2):   
        spotlist.append(df2.iloc[ 2 , spotconc0 ]) 
        spotlist.append(df2.iloc[ 3 , spotconc0 ])
        spotlist.append(df2.iloc[ 4 , spotconc0 ])
        spotlist.append(df2.iloc[ 5 , spotconc0 ])
        spotlist.append(df2.iloc[ 6 , spotconc0 ])
        spotlist.append(df2.iloc[ 7 , spotconc0 ])
        spotlist.append(df2.iloc[ 8 , spotconc0 ])
        spotlist.append(df2.iloc[ 9 , spotconc0 ])
        spotlist.append(df2.iloc[ 10, spotconc0 ])
        spotlist.append(df2.iloc[ 11, spotconc0 ])
        spotlist.append(df2.iloc[ 12, spotconc0 ])
        spotlist.append(df2.iloc[ 13, spotconc0 ])
        # 0.1 data
        spotlist.append(df2.iloc[ 14 , spotconc0 ])
        spotlist.append(df2.iloc[ 15 , spotconc0 ])
        spotlist.append(df2.iloc[ 16 , spotconc0 ])
        spotlist.append(df2.iloc[ 17 , spotconc0 ])
        spotlist.append(df2.iloc[ 18 , spotconc0 ])
        spotlist.append(df2.iloc[ 19 , spotconc0 ])
        spotlist.append(df2.iloc[ 20 , spotconc0 ])
        spotlist.append(df2.iloc[ 21 , spotconc0 ])
        spotlist.append(df2.iloc[ 22,  spotconc0 ])
        spotlist.append(df2.iloc[ 23,  spotconc0 ])
        spotlist.append(df2.iloc[ 24,  spotconc0 ])
        spotlist.append(df2.iloc[ 25,  spotconc0 ])
        # 0.05 data
        spotlist.append(df2.iloc[ 26 , spotconc0 ])
        spotlist.append(df2.iloc[ 27 , spotconc0 ])
        spotlist.append(df2.iloc[ 28 , spotconc0 ])
        spotlist.append(df2.iloc[ 29 , spotconc0 ])
        spotlist.append(df2.iloc[ 30 , spotconc0 ])
        spotlist.append(df2.iloc[ 31 , spotconc0 ])
        spotlist.append(df2.iloc[ 32 , spotconc0 ])
        spotlist.append(df2.iloc[ 33 , spotconc0 ])
        spotlist.append(df2.iloc[ 34,  spotconc0 ])
        spotlist.append(df2.iloc[ 35,  spotconc0 ])
        spotlist.append(df2.iloc[ 36,  spotconc0 ])
        spotlist.append(df2.iloc[ 37,  spotconc0 ])
        # 0.01 data
        spotlist.append(df2.iloc[ 38 , spotconc0 ])
        spotlist.append(df2.iloc[ 39 , spotconc0 ])
        spotlist.append(df2.iloc[ 40 , spotconc0 ])
        spotlist.append(df2.iloc[ 41 , spotconc0 ])
        spotlist.append(df2.iloc[ 42 , spotconc0 ])
        spotlist.append(df2.iloc[ 43 , spotconc0 ])
        spotlist.append(df2.iloc[ 44 , spotconc0 ])
        spotlist.append(df2.iloc[ 45 , spotconc0 ])
        spotlist.append(df2.iloc[ 46,  spotconc0 ])
        spotlist.append(df2.iloc[ 47,  spotconc0 ])
        spotlist.append(df2.iloc[ 48,  spotconc0 ])
        spotlist.append(df2.iloc[ 49,  spotconc0 ])
        # 0.00 data
        spotlist.append(df2.iloc[ 50 , spotconc0 ])
        spotlist.append(df2.iloc[ 51 , spotconc0 ])
        spotlist.append(df2.iloc[ 52 , spotconc0 ])
        spotlist.append(df2.iloc[ 53 , spotconc0 ])
        spotlist.append(df2.iloc[ 54 , spotconc0 ])
        spotlist.append(df2.iloc[ 55 , spotconc0 ])
        spotlist.append(df2.iloc[ 56 , spotconc0 ])
        spotlist.append(df2.iloc[ 57 , spotconc0 ])
        spotlist.append(df2.iloc[ 58,  spotconc0 ])
        spotlist.append(df2.iloc[ 59,  spotconc0 ])
        spotlist.append(df2.iloc[ 60,  spotconc0 ])
        spotlist.append(df2.iloc[ 61,  spotconc0 ])

        spotconc0 += 1
    '''End of T1 strip plates''' 

    spotconc0 = 8   # T2 (Light)
    for i in range(2):     
        spotlist2.append(df2.iloc[ 2 , spotconc0 ]) 
        spotlist2.append(df2.iloc[ 3 , spotconc0 ])
        spotlist2.append(df2.iloc[ 4 , spotconc0 ])
        spotlist2.append(df2.iloc[ 5 , spotconc0 ])
        spotlist2.append(df2.iloc[ 6 , spotconc0 ])
        spotlist2.append(df2.iloc[ 7 , spotconc0 ])
        spotlist2.append(df2.iloc[ 8 , spotconc0 ])
        spotlist2.append(df2.iloc[ 9 , spotconc0 ])
        spotlist2.append(df2.iloc[ 10, spotconc0 ])
        spotlist2.append(df2.iloc[ 11, spotconc0 ])
        spotlist2.append(df2.iloc[ 12, spotconc0 ])
        spotlist2.append(df2.iloc[ 13, spotconc0 ])
        # 0.1 data
        spotlist2.append(df2.iloc[ 14 , spotconc0 ])
        spotlist2.append(df2.iloc[ 15 , spotconc0 ])
        spotlist2.append(df2.iloc[ 16 , spotconc0 ])
        spotlist2.append(df2.iloc[ 17 , spotconc0 ])
        spotlist2.append(df2.iloc[ 18 , spotconc0 ])
        spotlist2.append(df2.iloc[ 19 , spotconc0 ])
        spotlist2.append(df2.iloc[ 20 , spotconc0 ])
        spotlist2.append(df2.iloc[ 21 , spotconc0 ])
        spotlist2.append(df2.iloc[ 22,  spotconc0 ])
        spotlist2.append(df2.iloc[ 23,  spotconc0 ])
        spotlist2.append(df2.iloc[ 24,  spotconc0 ])
        spotlist2.append(df2.iloc[ 25,  spotconc0 ])
        # 0.05 data
        spotlist2.append(df2.iloc[ 26 , spotconc0 ])
        spotlist2.append(df2.iloc[ 27 , spotconc0 ])
        spotlist2.append(df2.iloc[ 28 , spotconc0 ])
        spotlist2.append(df2.iloc[ 29 , spotconc0 ])
        spotlist2.append(df2.iloc[ 30 , spotconc0 ])
        spotlist2.append(df2.iloc[ 31 , spotconc0 ])
        spotlist2.append(df2.iloc[ 32 , spotconc0 ])
        spotlist2.append(df2.iloc[ 33 , spotconc0 ])
        spotlist2.append(df2.iloc[ 34,  spotconc0 ])
        spotlist2.append(df2.iloc[ 35,  spotconc0 ])
        spotlist2.append(df2.iloc[ 36,  spotconc0 ])
        spotlist2.append(df2.iloc[ 37,  spotconc0 ])
        # 0.01 data
        spotlist2.append(df2.iloc[ 38 , spotconc0 ])
        spotlist2.append(df2.iloc[ 39 , spotconc0 ])
        spotlist2.append(df2.iloc[ 40 , spotconc0 ])
        spotlist2.append(df2.iloc[ 41 , spotconc0 ])
        spotlist2.append(df2.iloc[ 42 , spotconc0 ])
        spotlist2.append(df2.iloc[ 43 , spotconc0 ])
        spotlist2.append(df2.iloc[ 44 , spotconc0 ])
        spotlist2.append(df2.iloc[ 45 , spotconc0 ])
        spotlist2.append(df2.iloc[ 46,  spotconc0 ])
        spotlist2.append(df2.iloc[ 47,  spotconc0 ])
        spotlist2.append(df2.iloc[ 48,  spotconc0 ])
        spotlist2.append(df2.iloc[ 49,  spotconc0 ])
        # 0.00 data
        spotlist2.append(df2.iloc[ 50 , spotconc0 ])
        spotlist2.append(df2.iloc[ 51 , spotconc0 ])
        spotlist2.append(df2.iloc[ 52 , spotconc0 ])
        spotlist2.append(df2.iloc[ 53 , spotconc0 ])
        spotlist2.append(df2.iloc[ 54 , spotconc0 ])
        spotlist2.append(df2.iloc[ 55 , spotconc0 ])
        spotlist2.append(df2.iloc[ 56 , spotconc0 ])
        spotlist2.append(df2.iloc[ 57 , spotconc0 ])
        spotlist2.append(df2.iloc[ 58,  spotconc0 ])
        spotlist2.append(df2.iloc[ 59,  spotconc0 ])
        spotlist2.append(df2.iloc[ 60,  spotconc0 ])
        spotlist2.append(df2.iloc[ 61,  spotconc0 ])

        spotconc0 += 1
        

    spotconc0 = 12   # T2 (Light)
    for i in range(14, 16):
        spotlist2.append(df2.iloc[ 2 , spotconc0 ]) 
        spotlist2.append(df2.iloc[ 3 , spotconc0 ])
        spotlist2.append(df2.iloc[ 4 , spotconc0 ])
        spotlist2.append(df2.iloc[ 5 , spotconc0 ])
        spotlist2.append(df2.iloc[ 6 , spotconc0 ])
        spotlist2.append(df2.iloc[ 7 , spotconc0 ])
        spotlist2.append(df2.iloc[ 8 , spotconc0 ])
        spotlist2.append(df2.iloc[ 9 , spotconc0 ])
        spotlist2.append(df2.iloc[ 10, spotconc0 ])
        spotlist2.append(df2.iloc[ 11, spotconc0 ])
        spotlist2.append(df2.iloc[ 12, spotconc0 ])
        spotlist2.append(df2.iloc[ 13, spotconc0 ])
        # 0.1 data
        spotlist2.append(df2.iloc[ 14 , spotconc0 ])
        spotlist2.append(df2.iloc[ 15 , spotconc0 ])
        spotlist2.append(df2.iloc[ 16 , spotconc0 ])
        spotlist2.append(df2.iloc[ 17 , spotconc0 ])
        spotlist2.append(df2.iloc[ 18 , spotconc0 ])
        spotlist2.append(df2.iloc[ 19 , spotconc0 ])
        spotlist2.append(df2.iloc[ 20 , spotconc0 ])
        spotlist2.append(df2.iloc[ 21 , spotconc0 ])
        spotlist2.append(df2.iloc[ 22,  spotconc0 ])
        spotlist2.append(df2.iloc[ 23,  spotconc0 ])
        spotlist2.append(df2.iloc[ 24,  spotconc0 ])
        spotlist2.append(df2.iloc[ 25,  spotconc0 ])
        # 0.05 data
        spotlist2.append(df2.iloc[ 26 , spotconc0 ])
        spotlist2.append(df2.iloc[ 27 , spotconc0 ])
        spotlist2.append(df2.iloc[ 28 , spotconc0 ])
        spotlist2.append(df2.iloc[ 29 , spotconc0 ])
        spotlist2.append(df2.iloc[ 30 , spotconc0 ])
        spotlist2.append(df2.iloc[ 31 , spotconc0 ])
        spotlist2.append(df2.iloc[ 32 , spotconc0 ])
        spotlist2.append(df2.iloc[ 33 , spotconc0 ])
        spotlist2.append(df2.iloc[ 34,  spotconc0 ])
        spotlist2.append(df2.iloc[ 35,  spotconc0 ])
        spotlist2.append(df2.iloc[ 36,  spotconc0 ])
        spotlist2.append(df2.iloc[ 37,  spotconc0 ])
        # 0.01 data
        spotlist2.append(df2.iloc[ 38 , spotconc0 ])
        spotlist2.append(df2.iloc[ 39 , spotconc0 ])
        spotlist2.append(df2.iloc[ 40 , spotconc0 ])
        spotlist2.append(df2.iloc[ 41 , spotconc0 ])
        spotlist2.append(df2.iloc[ 42 , spotconc0 ])
        spotlist2.append(df2.iloc[ 43 , spotconc0 ])
        spotlist2.append(df2.iloc[ 44 , spotconc0 ])
        spotlist2.append(df2.iloc[ 45 , spotconc0 ])
        spotlist2.append(df2.iloc[ 46,  spotconc0 ])
        spotlist2.append(df2.iloc[ 47,  spotconc0 ])
        spotlist2.append(df2.iloc[ 48,  spotconc0 ])
        spotlist2.append(df2.iloc[ 49,  spotconc0 ])
        # 0.00 data
        spotlist2.append(df2.iloc[ 50 , spotconc0 ])
        spotlist2.append(df2.iloc[ 51 , spotconc0 ])
        spotlist2.append(df2.iloc[ 52 , spotconc0 ])
        spotlist2.append(df2.iloc[ 53 , spotconc0 ])
        spotlist2.append(df2.iloc[ 54 , spotconc0 ])
        spotlist2.append(df2.iloc[ 55 , spotconc0 ])
        spotlist2.append(df2.iloc[ 56 , spotconc0 ])
        spotlist2.append(df2.iloc[ 57 , spotconc0 ])
        spotlist2.append(df2.iloc[ 58,  spotconc0 ])
        spotlist2.append(df2.iloc[ 59,  spotconc0 ])
        spotlist2.append(df2.iloc[ 60,  spotconc0 ])
        spotlist2.append(df2.iloc[ 61,  spotconc0 ])

        spotconc0 += 1
    '''End of T2 strip plates''' 

    spotlistdf = pd.DataFrame(spotlist)
    spotlistdf2 = pd.DataFrame(spotlist2)
        
    # Using writer because I wish to write to more than one sheet in the workbook so I have to use ExcelWriter
    with pd.ExcelWriter(saveFile, engine='xlsxwriter') as writer:          

        '''Moving data to a table format with labels'''

        title01 = pd.DataFrame({"1:1000 2nd T1 plate, sciColor T2"})
        title02 = pd.DataFrame({"1:5000 2nd T1 plate, sciColor T2"})
        title03 = pd.DataFrame({"1:1000 2nd T2 plate, sciColor T2"})
        title04 = pd.DataFrame({"1:5000 2nd T2 plate, sciColor T2"})
        title05 = pd.DataFrame({"1:1000 2nd T1 plate, sciColor T3"})
        title06 = pd.DataFrame({"1:5000 2nd T1 plate, sciColor T3"})
        title07 = pd.DataFrame({"1:1000 2nd T2 plate, sciColor T3"})
        title08 = pd.DataFrame({"1:5000 2nd T2 plate, sciColor T3"})

        title1 = pd.DataFrame({"0.2 ug_ul"})
        title2 = pd.DataFrame({"0.1 ug_ul"})
        title3 = pd.DataFrame({"0.05 ug_ul"})
        title4 = pd.DataFrame({"0.01 ug_ul"})
        title5 = pd.DataFrame({"0.000 ug_ul"})


        # Getting dataframe from submit button
        df = dflist[-1]

        # Data will be printed in different places due to the for loop, not starting at top right but that will not impact the accuracy of the data.
        if Columnstart == 1:
            A = 'A'
            B = 'B'
            C = 'C'
            D = 'D'
        elif Columnstart == 2:
            A = 'E'
            B = 'F'
            C = 'G'
            D = 'H'


        # Comment out the below code if I don't need the titles anymore...
        title01.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 3) # 1:1000, Light, T1
        title02.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 5) # 1:5000, Light, T1
        title03.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 7) # 1:1000, Light, T2
        title04.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 9) # 1:5000, Light, T2
        title05.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 11) # 1:1000, dark, T1
        title06.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 13) # 1:5000, dark, T1
        title07.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 15) # 1:1000, dark, T2
        title08.to_excel(writer, sheet_name=sheettitle, startrow = 0, startcol = 17) # 1:5000, dark, T2
        # I may not be able to get all the data on one sheet initially, but I can do so later so this is a goo layout to begin with

        
        title1.to_excel(writer, sheet_name=sheettitle, startrow = 3, startcol = 0) # title1 = 0.2 ug/ul
        title2.to_excel(writer, sheet_name=sheettitle, startrow = 15, startcol = 0) # 0.1 ug/ul
        title3.to_excel(writer, sheet_name=sheettitle, startrow = 27, startcol = 0) # 0.05 ug/ul
        title4.to_excel(writer, sheet_name=sheettitle, startrow = 39, startcol = 0) # 0.01 ug/ul
        title5.to_excel(writer, sheet_name=sheettitle, startrow = 51, startcol = 0) # 0.00 ug/ul
        
        # The spot data goes down in one column of the data table EX. 12 spots of 0.2 go down each column...
        
        # Writing the spot data to the excel document in Table format...
        spotlistdf.to_excel(writer, sheet_name=sheettitle, startrow = 3, startcol = 3 )
        spotlistdf2.to_excel(writer, sheet_name=sheettitle, startrow = 3, startcol = 5 )
        
    

    print('done')
    # print(hconclist)
    # print(hconcmean)

# Create a File Explorer label
label_file_explorer = tk.Label(frame, text = "File Explorer (program made by Ian V.)", width = 100, height = 1,fg = "white", bg = 'black')
button_explore = tk.Button(frame, text = "Browse Files", command = browseFiles)
# button_exit = tk.Button(frame, text = "Exit", command = exit)

Message = tk.Label(root, text = 'This program assumes you printed in a 96 well plate in portrait orientation with A1 at the bottom of the plate.',bg = '#D6DADE', fg = "black", font=('calibre',10,'bold'))

# Button that will call the Submit function for the file browser
Enter = tk.Button(root, text = 'Enter Excel File', command = submit)

save_text = tk.Label(root, text = "What do you want to call your save file?",bg = Blue ,fg = "black", font=('calibre',10,'normal'))
Save_entry = tk.Entry(root, textvariable = savedir, font = ('calibre',10,'normal'), bg="white", fg = 'black')
save_note = tk.Label(root, text = "The save directory is automatically set to your home directory \n If you want it to go to a different directory you must specify that path from the home dir.",bg = Blue, fg = "black", font=('calibre',10,'normal'))
# Button that will call the Enter function for the save file name
Enter2 = tk.Button(root, text = 'Enter Save Name', command = save_inp)

Run_button = tk.Button(root, text = 'Run Program', command = RunProgram)

label_file_explorer.grid(row = 0, column=0)
button_explore.grid(row = 1, column=0)
Enter.pack()
Message.pack(pady=10)

save_text.pack()
Save_entry.pack()
save_note.pack()
Enter2.pack()

Run_button.pack(pady=20)

root.mainloop()
