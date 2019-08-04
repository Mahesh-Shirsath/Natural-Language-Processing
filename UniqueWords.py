import re
import xlsxwriter
import pandas as pd

def unique_file(input_filename):
    inputFile = open(input_filename, 'r', encoding="utf8")
    fileContents = inputFile.read()
    inputFile.close()
    fileContents = fileContents.replace("\n",' ')  # For removing all the Enter keys from the text
    duplicates = []
    wordList = re.split(" |\.|\,|\।|\?|\:|\!",fileContents)  # For spliting fileContents at various punctuations
    workbook = xlsxwriter.Workbook('output.xlsx') 
    worksheet = workbook.add_worksheet() 
    row = col = uniqueWords = tokens = 0
    worksheet.write(row,0,"Word")
    worksheet.write(row,1,"Occurences")
    worksheet.write(row,2,"Rank")
    worksheet.write(row,3,"f.r")  # Making the headers for our output excel file
    for word in wordList:
        if word != "" and word != "		": # For removing unwanted tokens
            tokens += 1
            if word not in duplicates:
                count = 0
                uniqueWords += 1
                duplicates.append(word)
                files = open(input_filename, 'r', encoding="utf8")
                for line in files:
                    words = re.split(" |\.|\,|\।|\?|\:|\!", line)
                    for i in words:
                        if(i==word):
                            count=count+1
                row += 1
                worksheet.write(row, col, str(word)) 
                worksheet.write(row, col + 1, count) # Adding word and its coount into excel file
    workbook.close()
    sort(uniqueWords)
    print("Unique words :" + str(uniqueWords))
    print("Token words:" + str(tokens))
    print("TTR = ",uniqueWords/tokens)

def sort(uniqueWords):
    xl = pd.ExcelFile("output.xlsx")
    df = xl.parse("Sheet1")
    df = df.sort_values("Occurences", ascending=False) #Sorting out output in decending order
    for i in range(0,uniqueWords):
        df.iloc[i,2] = i+1             # Giving rank to sorted output
    for i in range(0,uniqueWords):
        df.iloc[i,3] = df.iloc[i,2]*df.iloc[i,1]  # finding f.r
    writer = pd.ExcelWriter('output.xlsx') 
    df.to_excel(writer,'Sheet1',index=False) # Entering this output into the file 
    writer.save()

unique_file('corpus.txt')
