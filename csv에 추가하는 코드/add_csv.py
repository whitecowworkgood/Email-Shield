import os
import sys
from zipfile import ZipFile
import csv
from csv import writer
import re
from win10toast import ToastNotifier
import time

path = os.getcwd()
path = path + "/Target/"   
toaster = ToastNotifier()
def add_make():#csv파일을 만들기 위해 파일의 정보를 저장하는 함수를 따로 만듬. 파일을 저장할 때에만 사용하기.
    Check_Format_Doc = re.compile("doc")
    Check_Format_ppt = re.compile("ppt")
    Check_Format_excel = re.compile("xls")
    for ftc in os.listdir(path):
        
        #doc 검사 부분
        if re.search( Check_Format_Doc, ftc.split(".")[1]):

            new_filename = os.path.splitext(ftc)[0] + '.zip'
            
            os.rename(path+ftc, path+new_filename)
            os.rename(path+new_filename, path+ftc)

            document=ZipFile(path+ftc)

            try:
                vbadata=document.read("word/vbaProject.bin")
            except:
                toaster.show_toast("검사", "{%s}는 매크로 파일이 아닙니다." % ftc)
                document.close()
            else:
                #print(str(vbadata).split('CMG="')[1].split('"')[0])

                str_vba = str(vbadata)
                CMG = str_vba.split('CMG="')[1].split('"')[0]
                DPB = str_vba.split('DPB="')[1].split('"')[0]
                GC = str_vba.split('GC="')[1].split('"')[0]
                
                docData=document.read("docProps/core.xml")
                doc_Data = docData.decode("utf-8")
                Creator = doc_Data.split("<dc:creator>")[1].split("</dc:creator>")[0]
                Create_Date = doc_Data.split('<dcterms:created xsi:type="dcterms:W3CDTF">')[1].split("</dcterms:created>")[0]
            
                document.close()
                
                data = [CMG, DPB, GC, Creator, Create_Date]

                f = open('./metafiles/Meta_Data.csv','a', newline='', encoding="euc-kr")
                writer_object = writer(f)
            
                writer_object.writerow(data)
                f.close()
                
                try:
                    External = document.read("word/_rels/settings.xml.rels")
                except:
                    toaster.show_toast("검사", "{%s}는 External 파일이 아닙니다." % ftc)
                    document.close()
                else:
                    
                    ex = "External"
                    ex_string = re.compile(ex)
                    if re.search(ex_string,  External.decode("utf-8")):
                    
                        f = open('./metafiles/Xml_Injection_Url.csv','a', newline='', encoding="euc-kr")
                        writer_object = writer(f)
                        data = [str(External.decode("utf-8").split('Type="')[1].split('" Target="External"')[0])]
                        writer_object.writerow(data)
                        f.close()
        

        #ppt 검사 부분
        elif re.search( Check_Format_ppt, ftc.split(".")[1]):
          
            new_filename = os.path.splitext(ftc)[0] + '.zip'
            
            os.rename(path+ftc, path+new_filename)
            os.rename(path+new_filename, path+ftc)

            document=ZipFile(path+ftc)

            try:
              
                vbaData=document.read("ppt/vbaProject.bin")
            except:
                toaster.show_toast("검사", "{%s}는 매크로 파일이 아닙니다." % ftc)
                document.close()
            else:
                
                str_vba = str(vbadata)
                CMG = str_vba.split('CMG="')[1].split('"')[0]
                DPB = str_vba.split('DPB="')[1].split('"')[0]
                GC = str_vba.split('GC="')[1].split('"')[0]
                
                docData=document.read("docProps/core.xml")
                doc_Data = docData.decode("utf-8")
                Creator = doc_Data.split("<dc:creator>")[1].split("</dc:creator>")[0]
                Create_Date = doc_Data.split('<dcterms:created xsi:type="dcterms:W3CDTF">')[1].split("</dcterms:created>")[0]
 
                document.close()
                
                data = [CMG, DPB, GC, Creator, Create_Date]

                f = open('./metafiles/Meta_Data.csv','a', newline='', encoding="euc-kr")
                writer_object = writer(f)
            
                writer_object.writerow(data)
                f.close()
                
                try:
                    External = document.read("word/_rels/settings.xml.rels")
                except:
                    toaster.show_toast("검사", "{%s}는 External 파일이 아닙니다." % ftc)
                    document.close()
                else:
                    
                    ex = "External"
                    ex_string = re.compile(ex)
                    if re.search(ex_string,  External.decode("utf-8")):
                    
                        f = open('./metafiles/Xml_Injection_Url.csv','a', newline='', encoding="euc-kr")
                        writer_object = writer(f)
                        data = [str(External.decode("utf-8").split('Type="')[1].split('" Target="External"')[0])]
                        writer_object.writerow(data)
                        f.close()
        
        
        
        #excel 검사 부분
        elif re.search( Check_Format_excel, ftc.split(".")[1]):
          
            new_filename = os.path.splitext(ftc)[0] + '.zip'
            
            os.rename(path+ftc, path+new_filename)
            os.rename(path+new_filename, path+ftc)

            document=ZipFile(path+ftc)
            
            try:

                vbaData=document.read("xl/vbaProject.bin") # 문서에서 정보를 추출, 추후 저장이나 사용하는 코드 추가
            except:
                toaster.show_toast("검사", "{%s}는 매크로 파일이 아닙니다." % ftc)#
                document.close()
            else:
                
                str_vba = str(vbadata)
                CMG = str_vba.split('CMG="')[1].split('"')[0]
                DPB = str_vba.split('DPB="')[1].split('"')[0]
                GC = str_vba.split('GC="')[1].split('"')[0]
                
                docData=document.read("docProps/core.xml")
                doc_Data = docData.decode("utf-8")
                Creator = doc_Data.split("<dc:creator>")[1].split("</dc:creator>")[0]
                Create_Date = doc_Data.split('<dcterms:created xsi:type="dcterms:W3CDTF">')[1].split("</dcterms:created>")[0]
                document.close()
                
                data = [CMG, DPB, GC, Creator, Create_Date]

                f = open('./metafiles/Meta_Data.csv','a', newline='', encoding="euc-kr")
                writer_object = writer(f)
            
                writer_object.writerow(data)
                f.close()
                
                try:
                    External = document.read("word/_rels/settings.xml.rels")
                except:
                    toaster.show_toast("검사", "{%s}는 External 파일이 아닙니다." % ftc)
                    document.close()
                else:
                    
                    ex = "External"
                    ex_string = re.compile(ex)
                    if re.search(ex_string,  External.decode("utf-8")):
                    
                        f = open('./metafiles/Xml_Injection_Url.csv','a', newline='', encoding="euc-kr")
                        writer_object = writer(f)
                        data = [str(External.decode("utf-8").split('Type="')[1].split('" Target="External"')[0])]
                        writer_object.writerow(data)
                        f.close()
        
        
        
        
        #기타 파일은 검사에서 제외
        else:
            continue
            
            
            
if __name__ == "__main__":
    while True:
        add_make()
        time.sleep(20)