import pandas as pd
import PyPDF2
import re
import sys
import collections
import os
import nltk
from nltk.corpus import stopwords
nltk.download("stopwords")
stop_words = set(stopwords.words("italian"))

def main():
    #read the file and get the data
    try:
        path = input("Please provide the absolute path for the PDF file: ") 
        pdf_file = open(path, mode="rb")
        p = input("Please provide from which PDF page start counting: ")

        #print(filename)

    except FileNotFoundError:
        print("File not found. Please check the inserted path.")
        main()

    except:
        print("An error occurred ", sys.exc_info()[0])
        exit()
    head, tail = os.path.split(path)
    filename = tail.split(".")[0] #splits the tail (eg. file.pdf) and get the name without the extension
    read_pdf = PyPDF2.PdfFileReader(pdf_file)
    pdfpages = read_pdf.getNumPages() 
    i = int(p) - 1
    if i < 0:
        i = 0
    text = ""
    for i in range(i, pdfpages):
        page = read_pdf.getPage(i)
        page_content = page.extractText()
        text += " " + page_content
        i += 1
    datahandling(text)
    newtext = datahandling(text)
    dataexport(newtext, filename)
 
def datahandling(text):
    #cleaning and tidying up
    text = " ".join([word for word in text.split() if word not in stop_words]) # removes stop words by adding only word not in the stop word list
    text = re.sub("https?:\/\/.*[\r\n]*", "", text) #remove urls
    text = re.sub("[\(\)\\\/\;\,\.\[\]\{\}\-\_\!\?\=\:\'\"]*", "", text) #removing punctuaction, parenthesis and stuff
    textlist = text.replace("\n", "").lower().split() #removes newline from string (\n) -> makes it all lowercase -> splits the string in a list
    #lower_case = [word.lower() for word in textlist]
    counts_text = collections.Counter(textlist)
    return counts_text

def dataexport(data, filename):
    #exporting the data
    try:
        writer = pd.ExcelWriter(filename + 'count.xlsx', engine='xlsxwriter')
        df = pd.DataFrame((list(data.items())), columns=['words', 'count'])
        df.to_excel(writer, index=False)
        writer.save()
        print("The report has been saved successfully.")
    except:
        print("An error occurred ", sys.exc_info()[0])
        exit()

if __name__ == '__main__':
    main()