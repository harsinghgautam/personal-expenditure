import PyPDF2 as pp
import os
import pandas as pd
# pip install openpyxl to convert csv to excel

def pdf_to_text(ip_path,op_path,pwd):
    outputfile = open(op_path,'w')
    with open(ip_path,'rb') as inputfile:
            reader = pp.PdfReader(inputfile)
            reader.decrypt(pwd)
            for page in reader.pages:
                outputfile.write(page.extract_text())
    outputfile.close()

def text_to_excel(textfile,excelfilepath):
    xlfile = open(excelfilepath,'w')
    with open(textfile,'r') as myfile:
        newtext=myfile.read()
        str = newtext[newtext.find('Date'):newtext.find('TRANSACTION OVERVIEW')]
        str = str.replace(' ',',')
        str = str.replace(',-',',0')
        xlfile.write(str)
    xlfile.close()      

newfilename = ['jan_','feb_','march_','april_','may_','june_','july_','aug_','sept_','oct_','nov_','dec_']
# newfilename = ['april_24']
year = '22'
for i in range(len(newfilename)):
    print(i)
    inputpath = "D:\\mess\\bank statements\\sbi\\account statement_"+newfilename[i]+year+".pdf"
    path1 = "D:\\mess\\bank statements\\sbi\\text statements"
    if not os.path.exists(path1):
        os.makedirs(path1)
    outputpath = path1+"\\"+newfilename[i]+year+".txt"
    newoutputpath = "D:\\mess\\bank statements\\sbi\\text statements\\"+newfilename[i]+year+".csv"
    password = "46805220802"
    pdf_to_text(inputpath,outputpath,password)
    # text_to_excel(outputpath,newoutputpath)
    # csvfile = pd.read_csv(newoutputpath)
    # path2 = "D:\\mess\\bank statements\\sbi\\excel statements"
    # if not os.path.exists(path2):
    #     os.makedirs(path2)
    # csvfile.to_excel("D:\\mess\\bank statements\\sbi\\excel statements\\"+newfilename[i]+".xlsx",index=None,header = True)
print('done')                
    
# text_to_excel("D:\\mess\\bank statements\\sbi\\text statements\\aug_23.txt","D:\\mess\\bank statements\\sbi\\text statements\\april_23hue.csv")
# csvfile = pd.read_csv("D:\\mess\\bank statements\\sbi\\text statements\\april_23hue.csv")
# csvfile.to_excel("D:\\mess\\bank statements\\sbi\\text statements\\april_23hue.xlsx",index=None,header = True)
