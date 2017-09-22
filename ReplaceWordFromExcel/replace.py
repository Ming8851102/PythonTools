# encoding=utf-8
import time
import sys
import os
import io
import csv
import string
import docx
import xlrd
import win32com.client
from win32com.client import constants
import fnmatch
'''
#Purpose
Find docx's keyword that format is [xxxxxx] and using this keyword to transfer chinese from excel file   

#Prepare
1.Need import docx .Please isntall it ->  pip install python-docx 

#Note
1.Please check global parameters
2.Please close all of word before run replace.py
'''

# Global parameters
Target_Excel = "Anatomy Index v1.xlsx"
Excel_Sheetpage = 2  # select sheet page -> 0,1,2....
Find_Str_col = 1  # Select find string in which column
Replace_Str_cols = 2  # Select replace string in which column


def area_replace(area, find_text, replace_text):
    area.Find.Execute(FindText=find_text,
                      ReplaceWith=replace_text,
                      Replace=constants.wdReplaceAll,
                      MatchCase=True,
                      MatchWildcards=False)


def word_replace_all(target_file, find_text, replace_text):
    local_encoding = sys.getfilesystemencoding()
    appdir = os.path.abspath(os.path.dirname(
        sys.argv[0])).decode(local_encoding)
    os.chdir(appdir)

    filename = target_file

    word = win32com.client.gencache.EnsureDispatch("Word.Application")
    word.Visible = False
    word.Documents.Open(os.path.abspath(filename))

    for story in word.ActiveDocument.StoryRanges:
        area_replace(story, find_text, replace_text)
        while(True):
            if story.NextStoryRange:
                story = story.NextStoryRange
                area_replace(story, find_text, replace_text)
            else:
                break
    word.ActiveDocument.Save()
    word.ActiveDocument.Close(False)
    word.Quit()


def findkeyword(keyword):

    ad_wb = xlrd.open_workbook(Target_Excel)

    row_data = ad_wb.sheets()[0]

    # select sheet page -> 0,1,2....
    sheet_0 = ad_wb.sheet_by_index(Excel_Sheetpage)

    sheet_nrows = sheet_0.nrows - 1  # got sheet's rows
    sheet_ncols = sheet_0.ncols  # got sheet's cols

    while(sheet_nrows >= 0):

        if sheet_0.cell_value(sheet_nrows, Find_Str_col) == keyword:
            ad_wb.release_resources()
            return sheet_0.cell_value(sheet_nrows, Replace_Str_cols)

        sheet_nrows = sheet_nrows - 1

    ad_wb.release_resources()
    return -1


def systime():
    today = time.strftime("%Y%m%d.%H%M%S")
    return "System time is " + today


# Fiter file name is *.docx then store to target_files
target_files = fnmatch.filter(os.listdir(os.getcwd()), '*.docx')

print "Total files :" + str(len(target_files))

for fileName in target_files:
    print "Start:" + fileName + ", " + systime()

    # add log (Log format: xxxx.docx.txt )
    logFile = open(fileName + ".txt", 'w')

    # Got doc full test
    fullText = []
    doc = docx.Document(fileName)
    paras = doc.paragraphs

    for p in paras:
        fullText.append(p.text)  # fullText is list object

    # Find keyword to store keyword_list that need conform to the "[xxxxxx]"
    # format  from docx file

    keyword_list = []
    for line in fullText:
        first_index = line.find('[', 0)

        while (first_index != -1):
            last_index = line.find(']', first_index)
            keyword_list.append(line[first_index + 1:last_index])
            first_index = line.find('[', first_index + 1)

    keyword_list = list(set(keyword_list))  # Remove duplicate string

    # Based on keyword_list to find execl chinese string then replace original
    # doxc file.
    for keyword in keyword_list:

        keyword_result = findkeyword(keyword)

        if keyword_result != -1:
            word_replace_all(fileName, "[" + keyword + "]", keyword_result)
        else:
            logFile.write("keyword not found," + keyword + "\n")

    logFile.close()
