import openpyxl
import random
NUMBER_LISTS = 30
LENGTH_LISTS = 6
WORDS_LIST_PATH = './luce.xlsx'


sheet = openpyxl.load_workbook(WORDS_LIST_PATH).active
column = sheet['A']
book = openpyxl.Workbook()
new_sheet = book.active


def reset_words():
    words = {}
    for row in column:
        words[row.value] = []
    return words

lists = []
words = reset_words()

while len(lists) < NUMBER_LISTS:
    copy = list(words.keys())
    random.shuffle(copy)
    new_list = []
    out = False
    for i in range (0, LENGTH_LISTS):
        out_of_bounds = 0
        while i in words[copy[0]] and out_of_bounds < NUMBER_LISTS:
            copy.append(copy.pop(0))
            out_of_bounds += 1
        if out_of_bounds == NUMBER_LISTS:
            out = True
        words[copy[0]].append(i)
        new_list.append(copy.pop(0))

    lists.append(new_list)
    if out:
        print('!!')
        words = reset_words()
        lists = []

for i in range(0, LENGTH_LISTS):
    line = []
    for l in lists:
        line.append(l[i])
    new_sheet.append(line)

        
book.save('lists.xlsx')
