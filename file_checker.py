from urllib import request
from urllib import parse
from urllib.request import urlopen
from win10toast import ToastNotifier
import hashlib
import time
import json
import os
import sys
from zipfile import ZipFile
import csv
from csv import writer
import re

api_key = '바이러스토털의 API키를 입력하세요'
toaster = ToastNotifier()

path = os.getcwd()+"/Attachments/"


def file_checker():

    virus_total_check = virus_total()
    make_file_check = make_check()

    if(virus_total_check == False or make_file_check == False):
        return False
        
    else:
        return True

        
def Find_External_Injection(in_file, document, ftc):
    try:
        External = document.read(in_file)
    except:
        toaster.show_toast("검사", "{%s}는 External 파일이 아닙니다." % ftc)
        return True
    else:
        ex = 'Target="External"'
        ex_string = re.compile(ex)
        if re.search(ex_string, External.decode("utf-8")):
                        
            data_to_file=[str(External.decode("utf-8").split('Type="')[1].split('" '+ex)[0])]
                       
            f = open('./metafiles/Xml_Injection_Url.csv','r', encoding="euc-kr")#폴더로 정리 및 파일 구별함.
            dr = csv.reader(f)
            data_to_csv=[]
            
            for line in dr:
                data_to_csv.append(line)
                            
            f.close()  
                
        for i in range(1, len(data_to_csv)):
            if data_to_csv[i][0] == data_to_file[0]:
                return False                 
            else:
                continue
    return True
        
#----------------- 파일 메타 검사 관련 코드------------------

def read_file_meta(in_file, document, ftc): #파일에서 메타 정보를 가져오는 코드, 코드 간편화 및 특수문자 제거작업 완료
    
    try:
        vbadata=document.read(in_file)
    except:
        toaster.show_toast("검사", "{%s}는 매크로 파일이 아닙니다." % ftc)#
        document.close()
        return True
    else:
        str_vba = str(vbadata)      
    
        CMG = str_vba.split('CMG="')[1].split('"')[0]
        DPB = str_vba.split('DPB="')[1].split('"')[0]
        GC = str_vba.split('GC="')[1].split('"')[0]
                
        xml_Data = document.read("docProps/core.xml").decode("utf-8")
        Creator = xml_Data.split("<dc:creator>")[1].split("</dc:creator>")[0]
        Create_Date = xml_Data.split('<dcterms:created xsi:type="dcterms:W3CDTF">')[1].split("</dcterms:created>")[0]
        
        data_to_file=[CMG, DPB, GC, Creator, Create_Date]
        
        f = open('./metafiles/Meta_Data.csv','r', encoding="euc-kr")
        dr = csv.reader(f)
        data_to_csv=[]
            
        for line in dr:
            data_to_csv.append(line)
            
        f.close()
    
        for i in range(1, len(data_to_csv)):
            for j in range(0, len(data_to_csv[i])):
                if data_to_csv[i][j] == data_to_file[j]:
                    return False                 
                else:
                    continue  
    return True
#------------------------------------------------------------  

  
def virus_total():# 바이러스토탈 키가 검사 4개가 한계
    
    for ftc in os.listdir(path):

        file_path = path + ftc # 경로 + 파일이름

        '''
        f = open(file_path, 'rb')
        file_hash = hashlib.md5(f.read()).hexdigest()
        #file_hash="221b9de416d42a979288cfa196912af4" 테스트를 위해 실제 문서형 악성코드의 해시값을 대입
        f.close()
        '''

        url = "https://www.virustotal.com/vtapi/v2/file/scan"
        report_url = "https://www.virustotal.com/vtapi/v2/file/report"
        api = api_key
       

        params = {'apikey': api, 'resource':ftc}
        
        data = parse.urlencode(params).encode('utf-8')
        req = request.Request(report_url, data)
        response = urlopen(req)
        data = response.read()
        time.sleep(24)
        
        data = json.loads(data.decode('utf-8'))
        md5 = data.get('md5', {})
        scan = data.get('scans', {})
        #print(scan)
        
        keys=scan.keys()
        #print(keys)
        
        for key in keys:
            if key == 'AhnLab-V3':
                if scan[key]['result'] != None:
                    return False
    return True
    
def make_check():
    Check_Format_doc = re.compile("doc")
    Check_Format_ppt = re.compile("ppt")
    Check_Format_excel = re.compile("xls")
    
    
    for ftc in os.listdir(path):
        # ---------------doc 검사 부분  ---------------
        if re.search( Check_Format_doc, ftc.split(".")[1]):
            

            new_filename = os.path.splitext(ftc)[0] + '.zip'
            os.rename(path+ftc, path+new_filename)
            os.rename(path+new_filename, path+ftc)
            
            document=ZipFile(path+ftc)
            
            
            # 첫번째 검사: External Injection- 매크로 파일이 없어도 되는 공격기법
            
            if Find_External_Injection("word"+"/_rels/settings.xml.rels", document, ftc) == False:
                document.close()
                return False
            
            #두번째 검사: vbaProject.bin을 가지고 하는 검사
            
            if read_file_meta("word"+"/vbaProject.bin", document, ftc) == False:
                document.close()
                return False

            document.close()
        
        #---------------ppt 검사 부분  ---------------      
        elif re.search( Check_Format_ppt, ftc.split(".")[1]):
            

            new_filename = os.path.splitext(ftc)[0] + '.zip'
            os.rename(path+ftc, path+new_filename)
            os.rename(path+new_filename, path+ftc)
            
            document=ZipFile(path+ftc)
            
            # 첫번째 검사: External Injection- 매크로 파일이 없어도 되는 공격기법
            
            if Find_External_Injection("ppt"+"/_rels/settings.xml.rels", document, ftc) == False:
                document.close()
                return False
                        
            #두번째 검사: vbaProject.bin을 가지고 하는 검사
            
            if read_file_meta("ppt"+"/vbaProject.bin", document, ftc) == False:
                document.close()
                return False

            document.close()  
        #---------------엑셀검사 부분---------------
        elif re.search( Check_Format_excel, ftc.split(".")[1]):
            

            new_filename = os.path.splitext(ftc)[0] + '.zip'
            os.rename(path+ftc, path+new_filename)
            os.rename(path+new_filename, path+ftc)
            
            document=ZipFile(path+ftc)
            
            # 첫번째 검사: External Injection- 매크로 파일이 없어도 되는 공격기법
            
            if Find_External_Injection("xl"+"/_rels/settings.xml.rels", document, ftc) == False:
                document.close()
                return False
           

            #두번째 검사: vbaProject.bin을 가지고 하는 검사
            
            if read_file_meta("xl"+"/vbaProject.bin", document, ftc)== False:
                document.close()
                return False
            document.close()
            
        #검사 대상이 아닌 파일들 (ex. txt, gif, ...)
        else: 
            toaster.show_toast("검사", "{%s}는 검사 대상이 아닙니다." % ftc)#
            continue
    
    
    return True

if __name__ == "__main__":
    
    file_checker()
