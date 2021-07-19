#import xlrd module to retrieve information from the xls file
import xlrd
d = {}

loc = (r"french_dictionary.xls")#read the french dictionary 
fin_loc=(r"t8.shakespeare.txt")#read the input file conaining ~1 lask rows
fout_loc=(r"t8.shakespeare.translated.txt")#output file with translated words


wb = xlrd.open_workbook(loc)
sh = wb.sheet_by_index(0)

fin = open(fin_loc, "rt")#input file in read mode
fout=open(fout_loc,"wt")#output file in write mode

for i in range(1000): #converting excel sheet to dictionary
    cell_value_class = sh.cell(i,0).value
    cell_value_id = sh.cell(i,1).value
    d[cell_value_class] = cell_value_id

def convert(s): #to handle words with punctuations
    flag=False
    starting='' #string of special character before the actual word
    ending='' #string of special character after the actual word
    temp=''
    for i in s: #iterating over the string
        if i.isalpha():
            flag=True #flag to mark the start of actual word
            temp+=i
        else:
            if flag==False:
                starting+=i
            elif flag:
                ending+=i
    if temp.lower() in d.keys(): #check if temp is a word to be replaced
        temp=d[temp.lower()]
        temp=starting+temp+ending
        return temp
    else:
        return s


for line in fin:

    for word in line.split():
        if word.isalpha(): #check if word has punctuations in it
            if word.lower() in d.keys():
                line=line.replace(word,d[word.lower()])
        else:
            temp=convert(word)
            line = line.replace(word, temp)


    fout.write(line)
