import pandas as pd
import numpy as np 
import tkinter as tk
from tkinter import filedialog

def filter(input_path, output_path):#this is for yes for filtering out the keyord which one want to find and refine

    try:#as we dont know how many column i want to refine  so it is asking
        #as we can get error if it ios not an integer so using try-except
        ask = input("how many numbers columns do you want to refine:").strip()
        #here i got to get rid of the white spaces so strip  is used and then turned to int
        ask = int(ask)
    except ValueError or TypeError:#handling error
        print(f"pls enter number not any words or string")
    i = 0
    df = []
    words = []
    #obviously this is for looping the times as user ask for
    while( i < ask):
        try :
            data = pd.read_excel(input_path, header = None, nrows= 5 )
            #here data is storing the the excel we chose as datframe with no header and for onnly up to 5 rows 
            word = input("Enter the word which you want to extract the data of in column:")
            #then we are here for the word which i want to refine and get the data
            words.append(word)
            header_index = 0
            #now using iterrows which go inside dataframe amd extract index and row of every data i.e cell it go through
            for index, row in data.iterrows() :
                #as iterrows give two value one for index and another for row so we're using two variable
                if any(word.lower().strip() in str(cell).lower().strip() for cell in row  ):
                    # now here we are making the input word key low and white space free 
                    #as well as we are goind through each row and its elements.
                    # we are storing that elemnts in cell and checking that cell contains the word user is searching for
                    header_index = index #if we got that word then we will having that word's indx in that data frame
                    break#after  getting that word lest break the loop
            if header_index is None:#this is for when the word is not in the data frame
                    #remember the word youre searching should be in the first 5 rows as we have called only five of them 
                    print(f"no any {word} is found ")
                    i += 1
                    continue
            else:
                print(f"found the '{word}' in {header_index + 1} row ") 
            #now again we are reading that input file from top to bottom   
            final_data = pd.read_excel(input_path, header = header_index )
            #Now this word_col is searching column of final_data and in that column it is searching the given word and
            #when the given word is found then next syntax terminate the loop 
            #and here col for col....... syntax is cool as col in the front of the for loop remeber or store value of the col in every loop
            word_col = next((col for col in final_data.columns if word.lower().strip() in str(col).lower().strip()), None)
            if not word_col:#if the the given word is not found then this will run
                word_col = final_data.columns[0]#if no word is found as column then it is making the fisrt column as word_col
                i += 1
                continue
            #now creating another variable to stored the extracted column
            extracted_data = final_data[[word_col]].copy()
            extracted_data.columns = [word]#renaming the column 
            confirm = input("is it only numeric data or not. if yes then tyoe 'y' and if no , type 'n'")
            
            if confirm.lower().strip() == 'y':
                extracted_data[word] = pd.to_numeric(extracted_data[word], errors= 'coerce')
            #this is the syntax in which the pandas choose the value which has numeric and this error = 'coerce' is used for changing the non-numeric value to Nan forcibly
            
            extracted_data.dropna(subset = [word], inplace=True)
            #THIS .DROPNA IS USED TO REMOVE THE nan value
            extracted_data.reset_index( drop = True ,inplace=True)
            #this reset_index helps us to reset the index so that the nan value is covered up to a new index and there will be no any unusual gaps
            df.append(extracted_data) 
            print(f"Successfully extracted data to: {output_path}")
        except Exception as e :
            print(f"unexpected error : {e}")
        i += 1

    if df:
        Final_df = pd.concat(df,axis=1,ignore_index=True)
        Final_df.columns = words


        Final_df.to_excel(output_path, index = False)#obvisouly it is for converting that data frame into excel
            
        
    
def choose():
    try:
        asking = input("How many excel file data do you want to clean:").strip()
        asking = int(asking)
    except ValueError:
        print("pls only type numbers")
    j = 0
    while (j < asking): #here while choosing the file you should be careful because you need to choose different file so that you will extract data from panda in future without messing too much
    
    
        root = tk.Tk()
        root.withdraw()
        input_path = filedialog.askopenfilename(
            title = "Select the Excel file",
            filetypes = [("Excel file",".xlsx")]
        )
        output_path = filedialog.asksaveasfilename(
            title = "Select Location to save",
            filetypes = [("Excel file",".xlsx")]

        )
        if input_path and output_path:
            filter(input_path, output_path)
        j += 1
if __name__ == "__main__" :
    choose()

    
        

