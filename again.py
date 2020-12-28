import pandas as pd
from openpyxl import load_workbook

import xlsxwriter

import time
from time import perf_counter
import datetime

import os

import seaborn as sns

from xlutils.copy import copy as xl_copy
import openpyxl as xl

import random

### User information ###

name = input("Player's Name:- ")
if(len(name)==0 or name[0]==" "):
    print("Enter name correctly")
    name = input("Player's Name:- ")

now = datetime.datetime.now()

print(" ")

time.sleep(1)

print("***********************************START****************************************")
print(" ")
print("\t\t\t\tWELCOME",name,"!")
print("\t\t\t\t Let's Play!!!")
print("\t\t\t\t   HANGMAN")
print(" ")

print("Start guessing...\n\n(Hint: Start with vowels")


time.sleep(0.5)
t1_start = perf_counter()  # Starting the stopwatch
list_of_names = ['flower','inspire','cactus','android',
                 'samsung','pumpkin','automate',
                 'printer','inkjet','register']
word = random.choice(list_of_names)
length=len(word)
guesses = ''
life = 5
turns = life

while turns > 0:         
    failed = 0             # Counter = 0          
    for char in word:      
        if char in guesses:    
            print(char)
        else:
            
            print ("_")     
            failed += 1
            
    if failed == 0:        
        print("\t\t\t\tYou won")
        s="PASS"
        s_value=1
        break              

    print("\t\t\t\t\t\t", + turns,"lives")
    guess =input("guess an alphabet:")
    len_guess = len(guess)
    
    if guess in guesses:
        print("Already tried!!!Try again ")
    elif(len_guess != 1):
        print("Enter a single character")
    elif(guess.isupper()):
        print("Enter the character in lowercase")
    elif(guess.isalpha()==False):
        print("Entered character should be an alphabet")
    else:
        guesses += guess                    

        if guess not in word:  
            turns -= 1        
            print ("Wrong guess")    

            if turns == 0:           
                print ("\t\t\t\tYou Lose")
                s="FAIL"
                s_value=0

t1_stop = perf_counter()     # Stop the stopwatch 
total_t = t1_stop-t1_start   # Total elapsed time

chances=len(guesses)         # Number of gussess used
no_of_letters=length-failed   # Number of correct guesses

# Score calculation algorithm

if(s_value==1):
    score = (no_of_letters*10/length) + turns - (chances/length) - (total_t/100) + 2
else:    
    score = (no_of_letters*10/length) + turns - (chances/length) - (total_t/100) - 0.5
print('You took', chances,'guesses')

print("********************************************************************************")

### Database formation ###

# Checking if certain file/directory exists
if(os.path.exists('./Database.xlsx')==False): 
    df = pd.DataFrame({'Name': [name],  
                       'Date': [now.strftime("%Y-%m-%d %H:%M:%S")],
                       'Start time':[t1_start],'Stop time':[t1_stop],
                       'Elapsed time':[total_t],'Lives used':[life-turns],
                       'Lives remaining':[turns],'Guesses':[chances],
                       'word completion':[((no_of_letters/length)*100)],
                       'Score':[score],'Status':[s]})

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter('Database.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='data', index=False)
    
    # Create xlsxwriter workbook object. 
    workbook_object = writer.book 
       
    # Create xlsxwriter worksheet object. 
    worksheet_object = writer.sheets['data']

    # Set the column width and format. 
    worksheet_object.set_column('A:B', 20)
    worksheet_object.set_column('C:G', 15)
    worksheet_object.set_column('H:H', 10)
    worksheet_object.set_column('I:J', 15)
    worksheet_object.set_column('K:K', 10)
    
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

else:
    
    #new dataframe with same columns
    df = pd.DataFrame({'Name': [name], 
                       'Date': [now.strftime("%Y-%m-%d %H:%M:%S")],
                       'Start time':[t1_start],'Stop time':[t1_stop],
                       'Elapsed time':[total_t],'Lives used':[life-turns],
                       'Lives remaining':[turns],'Guesses':[chances],
                       'word completion':[((no_of_letters/length)*100)],
                       'Score':[score],'Status':[s]})

    writer = pd.ExcelWriter('Database.xlsx', engine='openpyxl')
    
    # try to open an existing workbook
    writer.book = load_workbook('Database.xlsx')
    
    # copy existing sheets
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    
    # read existing file
    reader = pd.read_excel(r'Database.xlsx')
    
    # write out the new sheet
    df.to_excel(writer,sheet_name='data',
                index=False,header=False,startrow=len(reader)+1)
    
    
    writer.sheets['data'].column_dimensions['A'].width = 20
    writer.sheets['data'].column_dimensions['B'].width = 20
    writer.sheets['data'].column_dimensions['C'].width = 15
    writer.sheets['data'].column_dimensions['D'].width = 15
    writer.sheets['data'].column_dimensions['E'].width = 15
    writer.sheets['data'].column_dimensions['F'].width = 15
    writer.sheets['data'].column_dimensions['G'].width = 15
    writer.sheets['data'].column_dimensions['H'].width = 10
    writer.sheets['data'].column_dimensions['I'].width = 15
    writer.sheets['data'].column_dimensions['J'].width = 15
    writer.sheets['data'].column_dimensions['K'].width = 10
    
    writer.save()
    

    ### Sorting based on Score ###
    
    # read existing file
    ready=pd.read_excel('Database.xlsx',sheet_name='data')

    df_pr = ready.sort_values(by='Score',ascending=False)
    
    writer = pd.ExcelWriter('processing.xlsx')

    # Intermediate sheet formed
    df_pr.to_excel(writer,sheet_name='processing',index=False)
    
    writer.save()


    ### Extractng Data from processing sheet ###
    
    df2 =df_pr.Name

    # Dictionary of all the unique names with their final score
    asdf={} #{Name:Final Score}

    # Dictionary of all the unique names with number of wins
    zxcv={} #{Name:Wins}

    # Dictionary of all the unique names with number of losses
    qwer={} #{Name:Losses}
    
    unique_names = set(list(df2))
    
    for name in unique_names:
        
        asdf[str(name)]=((sum(list(df_pr[df_pr.Name==name].Score)))/10)
        zxcv[str(name)]=len([x for x in (list(df_pr[df_pr.Name==name].Status)) if x == 'PASS'])
        qwer[str(name)] = len([x for x in (list(df_pr[df_pr.Name==name].Status)) if x == 'FAIL'])



    ### Leaderboard formation ###

    df_lb = pd.DataFrame({'Name':list(asdf.keys()),
                          'Wins':list(zxcv.values()),'Losses':list(qwer.values()),
                          'Score':list(asdf.values())})
    writer = pd.ExcelWriter('Leaderboard.xlsx', engine='xlsxwriter')

    # Convert the dataframe to an XlsxWriter Excel object.
    df_lb.to_excel(writer, sheet_name='Leaderboard', index=False)
    
    # Create xlsxwriter workbook object . 
    workbook_object = writer.book 
       
    # Create xlsxwriter worksheet object 
    worksheet_object = writer.sheets['Leaderboard']

    # Set the column width and format. 
    worksheet_object.set_column('A:D', 15)
    
    writer.save()


    ### Sorting of Data ###

    read_lb=pd.read_excel('Leaderboard.xlsx',sheet_name='Leaderboard')
    df_lb = read_lb.sort_values(by='Score',ascending=False)
    writer = pd.ExcelWriter('Leaderboard.xlsx')
    df_lb.to_excel(writer,sheet_name='Leaderboard',index=False)
    workbook_object = writer.book
    worksheet_object = writer.sheets['Leaderboard']
    worksheet_object.set_column('A:D', 15) 
    writer.save()
    
    



