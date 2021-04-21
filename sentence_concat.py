# Reading an excel file using Python 
import xlrd 
import re
import string
import csv

# Give the location of the file 
loc = ("air2.xlsx") 

# To open Workbook 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

# For row 0 and column 0 
sheet.cell_value(1, 1) 

# Program to extract number 
# of rows using Python 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 

# Extracting number of rows 
print('Totall Row = ',sheet.nrows) 
# Program to extract number of 
# columns in Python 

wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

# For row 0 and column 0 
sheet.cell_value(0, 0) 

# Extracting number of columns 
print('Totall col = ',sheet.ncols) 

#check End of Sentence 
def checkEndOfSentence(i):
    endOfSentence = sheet.cell_value(i, 2)
    if(endOfSentence == 1):
        return False
    else:
        return True 
#clean msg
def clean_msg(msg):
    # remove text in <> 
    msg = re.sub(r'<.*?>','', msg)
    
    # remove hashtag
    msg = re.sub(r'#','',msg)
    
    # remove  (punctuation)
    for c in string.punctuation:
        msg = re.sub(r'\{}'.format(c),'',msg)
    # remove separator ex \n \t
    msg = ' '.join(msg.split())
    # remove emoji
    emoji_pattern = re.compile("["
        u"\U0001F600-\U0001F64F"  # emoticons
        u"\U0001F300-\U0001F5FF"  # symbols & pictographs
        u"\U0001F680-\U0001F6FF"  # transport & map symbols
        u"\U0001F1E0-\U0001F1FF"  # flags (iOS)
                           "]+", flags=re.UNICODE)
    msg=emoji_pattern.sub(r'', msg)
    
    return msg
#concat contence 
 def concatSentence():
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    concat = ''
    list= []
    totallRow = sheet.nrows
    firstRow = 1
    for i in range(1,totallRow):
        concat = ''
        listSentence = []
#         print(firstRow)
        for j in range(firstRow,totallRow): 
            if(checkEndOfSentence(j)):
                concat+=sheet.cell_value(j, 1)
            else:
                concat+=sheet.cell_value(j, 1)
                #clean data
                concat=clean_msg(concat)
                listSentence.append(concat)
                #Type Question form column 3 
                typeQ = sheet.cell_value(j, 3)
                listSentence.append(typeQ)
                #change first row focus
                firstRow = j+1
                break
        list.append(listSentence)
    i=0
    return list

list=concatSentence()

##write to csv 
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
totallRow = sheet.nrows
list=concatSentence()
csvName = 'sentence_concat.csv'
with open(csvName, mode='w') as sentence_concat_file:
    sentence_concat_writer = csv.writer(sentence_concat_file, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL)
    for i in range(0,len(list)):
        sentence_concat_writer.writerow(list[i])
        #example data csv
        #sentence_concat_writer.writerow(['Erica Meyers', 'IT', 'March'])

