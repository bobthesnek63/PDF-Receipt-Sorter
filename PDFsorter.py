# This script is an asset of InqueryAI
# Code developed by Pranav Vyas

import PyPDF2
from nltk.tokenize import word_tokenize
from nltk.corpus import stopwords
import re
import xlsxwriter
import pandas as pd


def findNumString(string):
    if any(char.isdigit() for char in string) and any(char1.isdigit() == False for char1 in string):
        for i in range(len(string)):
            if string[i].isdigit() != string[0].isdigit():
                return i
    else:
        return 0


filename = "Nov-Dec'19.pdf"
numFile = open("Nov-Dec'19.pdf", "rb")

pdfReader = PyPDF2.PdfFileReader(numFile)

numPages = pdfReader.numPages
counter = 1
text = ""

# The while loop will read each page.
while counter < numPages:
    pageObj = pdfReader.getPage(counter)
    counter += 1
    text += pageObj.extractText()

# This if statement exists to check if the above library returned words. It's done because PyPDF2 cannot read scanned files.
if text != "":
    text = text

# The word_tokenize() function will break our text phrases into individual words.
tokens = word_tokenize(text)

# We'll create a new list that contains punctuation we wish to clean.
punctuations = ['(', ')', ';', ':', '[', ']', ',', '.', 'If', 'if']

# We initialize the stopwords variable, which is a list of words like "The," "I," "and," etc. that don't hold much value as keywords.
stop_words = stopwords.words('english')

# We create a list comprehension that only returns a list of words that are NOT IN stop_words and NOT IN punctuations.
keywords = [word for word in tokens if not word in stop_words and not word in punctuations]

table = []
count = 0
for i in keywords:
    if i == '4500':
        count += 1

    if count == 1:
        table += [i]

    if i == 'find':
        count += 1

# print(findNumString(table[3]))

betterTable =[]
for i in table:
    location = findNumString(i)
    if location != 0:
        betterTable += [i[:location]]
        betterTable += [i[location:]]
    else:
        betterTable += [i]

lastTable = []
for i in range(len(betterTable)):
    decSplit = []
    if betterTable[i].find('.') != -1 and betterTable[i][0] != 'W':
        decSplit += [str(s) for s in re.findall(r'-?\d+\.?\d*', betterTable[i])]
        splitword = betterTable[i].split(decSplit[0])
        lastTable += [decSplit[0]]
        lastTable += [splitword[1]]
    elif betterTable[i][0] == 'O' and betterTable[i][1] == 'N':
        lastTable += ["ON"]
        lastTable += [betterTable[i][2:]]
    elif betterTable[i].find('Personal') != -1 and betterTable[i][0] != 'P':
        splitword = betterTable[i].split('Personal')
        lastTable += [splitword[0]]
        lastTable += ['Personal']
    else:
        lastTable += [betterTable[i]]

lastTable.pop()
lastTable.pop()
lastTable = lastTable[4:]
print(lastTable)

months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
provinces = ['NL', 'PE', 'NS', 'NB', 'QC', 'ON', 'MB', 'SK', 'ON', 'BC', 'YT', 'NT', 'NU']

orgTable = []
miniTable = []
nums = 0
for i in lastTable:
    if i in months:
        nums += 1

    if i in months and nums%2 == 1:
        orgTable += [miniTable]
        miniTable = []
        miniTable += [i]
    else:
        miniTable += [i]

orgTable = orgTable[1:]
print()
print(orgTable)



workbook = xlsxwriter.Workbook('recepit2.xlsx')
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'Date')
worksheet.write('B1', 'Name')
worksheet.write('C1', 'Location')
worksheet.write('D1', 'Gas')
worksheet.write('E1', 'Groceries')
worksheet.write('F1', 'Clothes')
worksheet.write('D1', 'Restaurants')
worksheet.write('D1', 'Gifts')
worksheet.write('H1', 'Amount Spent')

# data = pd.read_csv('cgn_on_csv_eng.csv')
# data = data[["Geographical Name"]]
# print(data.head())

column = 0
row = 1
for i in range(len(orgTable)):
    date = ''
    name = ''
    locationStr = ''
    for j in range(2):
        date += str(orgTable[i][j])
        date += ' '
        name += str(orgTable[i][j + 4]) + ' '

    worksheet.write(row, column, date)
    worksheet.write(row, column + 1, name)

    for z in range(len(orgTable[i])):
        if orgTable[i][z] in provinces:
            locationStr += str(orgTable[i][z - 1])
            locationStr += ' '
            locationStr += str(orgTable[i][z])
        elif 'WWW' in orgTable[i][z]:
            locationStr += str(orgTable[i][z])

    worksheet.write(row, column + 2, locationStr)
    worksheet.write(row, column + 3, str(orgTable[i][-2]))
    worksheet.write(row, column + 4, str(orgTable[i][-1]))
    row += 1



workbook.close()



