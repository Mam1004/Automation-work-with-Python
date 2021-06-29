import pdfplumber
import re
import win32com.client
import os
import glob

File_path=r"M:\python\data_files\pdf"
os.chdir(File_path)


for FileList in glob.glob('*.pdf'):
  count=0
  with pdfplumber.open(File_path+"\\"+FileList) as pdf:
    pages = pdf.pages
    for i,pg in enumerate(pages):

     text=pages[i].extract_text()

    # print(text)

     for row in text.split('\n'):
         #print(row)
         if re.search('US Operations Training Team',row):
            count=count+1
   
      
  print(FileList+" --> no of occurnce :"+str(count))
   
