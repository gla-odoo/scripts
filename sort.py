import openpyxl
import random
sheet = openpyxl.load_workbook('/home/geof/lucie/liste.xlsx')['Sheet1']
column = sheet['A']
words = []
book = openpyxl.Workbook()
new_sheet = book.active
for row in column:
    words.append(row.value)
random.shuffle(words)
words_copy = words[:]
print(words_copy)
new_lists = []
i = 0
while len(new_lists) < 30 :
    next_list = []  
    while len(next_list) < 6:
#        print(i, ": ", words[0])
        next_list.append(words_copy.pop(0))
        if len(words_copy) == 0:
            words.append(words.pop(0))           
            words_copy = words[:]
            new_sheet.append(words)
            print(words_copy)
#    print(next_list)
    new_lists.append(next_list)



book.save('new.xlsx')

