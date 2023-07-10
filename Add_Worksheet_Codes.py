#!/usr/bin/env python
# coding: utf-8

# In[5]:


import os
#import win32com
import win32com.client

def check_user_input_str(msg, answerlist):
    msg =msg.strip()+" \n"
    while True:
        print ("-"*70)
        answer = input (msg).lower()
        if answer in answerlist:
            return answer
            break
        else:
            print ('Not a valid answer.')

def add_worksheet_code():
    print ('-'*70)
    print ('Starting to add macros...\n')
    #folder path which contains xlsm files
    folder_path=os.getcwd()

    # Get a list of all the xlsm files in the folder
    xlsm_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".xlsm") and "~" not in file]

    # print the total number of files to be changed
    print ('Number of worksheets to be changed is:', len(xlsm_files),'\n')

    #import the code in  ThisWorkbook module
    with open(folder_path+"\\ThisWorkbook.txt", "r") as f:
        # Read the text from the txt file
        ThisWorkbookCode = f.read()

    # get path for code in Module
    code_path = folder_path+"\\ModuleCode.bas"

    # Create an instance of the Excel Application object
    app = win32com.client.Dispatch("Excel.Application")

    #loop through all files and do the following:
                # delete old code in ThisWorkbook module
                # add new code in ThisWorkbook module
                # delete all old modules
                # add new module
    #count of worksheets to be modified
    counter=0
    for file in xlsm_files:
        counter = counter+1
        #open file
        xlwb = app.Workbooks.Open(file)

        #**************__ThisWorkbook module__**************
        #get total lines of code
        total_lines = xlwb.VBProject.VBComponents('ThisWorkbook').CodeModule.CountOfLines
        #delete all lines of code
        xlwb.VBProject.VBComponents('ThisWorkbook').CodeModule.DeleteLines(1,total_lines)
        #add new code
        xlwb.VBProject.VBComponents('ThisWorkbook').CodeModule.AddFromString(ThisWorkbookCode)

        #**************__All other modules__**************
        #delete all modules
        for i in xlwb.VBProject.VBComponents:        
            xlmodule = xlwb.VBProject.VBComponents(i.Name)
            if xlmodule.Type in [1, 2, 3]:            
                xlwb.VBProject.VBComponents.Remove(xlmodule)

        #import new module
        xlwb.VBProject.VBComponents.Import(code_path)

        #print the name of modified worksheet
        last_slash = file.rfind("\\")
        filename = file[last_slash+1:]
        print ("project {}: {} is updated".format(counter,filename))

        #close file
        xlwb.Close(True)

    #close app
    app.Quit;
    print ('\nAll worksheets are successfuly updated with new macros.'+'\n')
    print ('-'*70+'\n')

def main():
    while True:
        print ('Simple tool to replace existing macros external macros in excel-enabled-macro worksheet.')
        print (70*"-")
        
        answer = check_user_input_str("Would you like to start adding codes? [y,n]: ",['y','n'])
        if answer=='y':
            print ('\nImportant instructions:')
            print ('All workbooks should be (.xlsm) and in the current directory.')
            print ('The name of the file containing macros which to be inserted in (ThisWorkbook) module should be: ThisWorkbook.txt')
            print ('The name of the file containg macros which to be inserted in any other modules should be: ModuleCode.bas')
            answer2 = check_user_input_str("Press y when you are ready to start adding macros, for exit please press n:",['y','n'])
            print (answer2)
            if answer2=='y':
                add_worksheet_code()
            elif answer2=='n':
                break
        elif answer=='n':
            break
            
if __name__ == '__main__':
    try:
        main()
    except Exception as ex:
        print ('Error: {}!, try to restart the script'.format(ex))


# In[ ]:




