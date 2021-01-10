import pandas as pd
import numpy
import xlrd

def strip_punctuation(file):
    
    text = []
    punctuation_chars = ["'", '"', ",", ".", "!", ":", ";", "#", "@", "-", "(", ")", "_"]
    
    for list in file:
        for  z in list:
             for y in punctuation_chars:
                if y in z:
                     z = z.replace(y, '')
             text.append(z)

    tx = " ".join(text)
    return tx.lower()



def get_pos(file):
    count = 0
    positive_word = []
   
    main_file = strip_punctuation(file)
    with open(r"C:\Users\ankit\OneDrive\Desktop\PYTHON\PROJECT_SENTIMENT_ANALYSIS\positivewords.txt") as posi:
        for lin in posi:
            if lin[0] != ';' and lin[0] != '\n':
                positive_word.append(lin.strip())    
        for i in main_file.split(" "):
            for j in positive_word:
                if (j==i):
                    count+=1
    return count




def get_neg(file):
    count = 0
    negative_word = []
    main_file = strip_punctuation(file)
   
    with open(r"C:\Users\ankit\OneDrive\Desktop\PYTHON\PROJECT_SENTIMENT_ANALYSIS\negativewords.txt") as negi:
        for lin in negi:
            if lin[0] != ';' and lin[0] != '\n':
                negative_word.append(lin.strip())
        for i in main_file.split(" "):  
            for j in negative_word:              
                if j==i:
                    count+=1
    return count







def run(data):
    df = pd.DataFrame(columns=["Positive Score", "Negative Score", "Net Score"])
    i = 1
    for sheet in data.sheets():
        list_of_lists=[]
        for row in range(sheet.nrows): 
            for column in range (sheet.ncols):
                k = str(sheet.cell(row, column).value) 
                line_list = k.split()
                list_of_lists.append(line_list)
                
        #Positive Score
        positive_score = get_pos(list_of_lists)
        
        #Negative Score
        negative_score = get_neg(list_of_lists)
        
        #Net Score
        net_score = positive_score - negative_score
        df.loc[i, ['Positive Score']] = positive_score
        df.loc[i, ['Negative Score']] = negative_score
        df.loc[i, ['Net Score']] = net_score 
        i = i+1
    print(df)
           
if _name_ == "_main_":
    
    data = xlrd.open_workbook(r"C:\Users\ankit\OneDrive\Desktop\PYTHON\PROJECT_SENTIMENT_ANALYSIS\sample_movie_data.xlsx")
    run(data)