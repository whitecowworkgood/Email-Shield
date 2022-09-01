from win10toast import ToastNotifier
from file_checker import *
from email.header import decode_header, make_header
import imaplib
import email
import time
import os
import tkinter
from tkinter import ttk

toaster = ToastNotifier()
pass_ = False
window = tkinter.Tk()

def Main_Function(imap, itemlist, file_):
    recent = b'0'
    toaster.show_toast("메일", "새로운 메일이 왔습니다.")
    recent = itemlist[-1]
    
    
    result,email_data = imap.uid('fetch',recent,'(RFC822)')
    
    email_message = email.message_from_bytes(email_data[0][1])
    #print(email_message.get('subject'))
    subject = make_header(decode_header(email_message.get('subject')))
 

    path_dir = os.getcwd() + "/Attachments/"
    #print(path_dir) # 경로 확인용
    if(os.path.isdir(path_dir)):
        pass

    for item in email_message.walk():
        if(item.get_content_maintype() != 'multipart' and item.get('Content-Disposition') is not None ):
            
            file_name = make_header(decode_header(item.get_filename()))

            #첨부파일이 있는지 없는지 확인하는 분기문을 만들어야 할 듯
            if(file_name != None):
                file_ = True
                
                toaster.show_toast("메일", "메일 {%s}의 모든 첨부파일을 가져옵니다." % subject)#
                time.sleep(2)
                if(os.path.isdir(path_dir) != True):
                    #첨부파일이 여러개면 계속 생성해서 폴더가 없을때만 생성하게 만듬
                    os.system("mkdir Attachments")
                    
                if bool(file_name):
                    #print(file_name)
                    subpath = os.getcwd()
                    
                    subpath = subpath + "/Attachments/"
                    filepath = os.path.join(subpath,str(file_name))
                
                    if not os.path.isfile(filepath):
                        fpo = open(filepath,"wb")
                        fpo.write(item.get_payload(decode=True))
                        fpo.close()    
                        

    if(file_ == True):
        toaster.show_toast("메일", "{%s}메일의 첨부파일 검사를 진행합니다...." % subject)#
        time.sleep(1)
   
        check_result = file_checker()#file_checker.py의 함수

        if (check_result == False):
            toaster.show_toast("메일", "메일 {%s}는 위험한 메일인 것 같습니다. 주의해 주십시오." % subject)#
            #메일이름이 길면 자르는 기능도 생각해 볼 수 있음

        else:
            toaster.show_toast("메일", "메일 {%s}는 안전합니다."% subject)# 
            #메일이름이 길면 자르는 기능도 생각해 볼 수 있음
    else:
        toaster.show_toast("메일", "메일 {%s}은 첨부파일이 없습니다." % subject)#
            
    return file_,recent

def input_data():

    id, password, host = tkinter.StringVar(), tkinter.StringVar(), tkinter.StringVar()
    
    window.title("Login")
    window.geometry('300x200')
    window.resizable(False, False)
        
    ttk.Label(window, text = "Mailaddress : ").grid(row = 0, column = 0, padx = 10, pady = 10)
    ttk.Label(window, text = "Username : ").grid(row = 1, column = 0, padx = 10, pady = 10)
    ttk.Label(window, text = "Password : ").grid(row = 2, column = 0, padx = 10, pady = 10)
    ttk.Entry(window, textvariable = host).grid(row = 0, column = 1, padx = 10, pady = 10)
    ttk.Entry(window, textvariable = id).grid(row = 1, column = 1, padx = 10, pady = 10)
    ttk.Entry(window, textvariable = password, show="*").grid(row = 2, column = 1, padx = 10, pady = 10)
    ttk.Button(window, text = "Login", command = lambda: window.destroy()).grid(row = 3, column = 1, padx = 10, pady = 10)

    window.mainloop()

    return host.get(), id.get(), password.get()
def connect():
    mostrecent = b'0'
    file_ = False
    _connect = False
   
    _host, _id, _pw = input_data()
    while True:
        
        try :
            imap = imaplib.IMAP4_SSL(host=_host, port=imaplib.IMAP4_SSL_PORT)#imap 주소 입력
            if _connect == False:
                toaster.show_toast("이메일 연결 중...", "연결 중...")
            imap.login(_id,_pw)
            time.sleep(2)
        except:
            toaster.show_toast("로그인", "로그인에 실패하였습니다. 다시 로그인해 주세요.")
            return 0
            #continue
        else:
            if _connect == False:
                toaster.show_toast("이메일 연결", "연결 성공!!")
                toaster.show_toast("메일 점검", "메일 점검을 진행합니다.")
                _connect = True
            time.sleep(2)
        
            imap.select('INBOX')    
            result, data = imap.uid('search',None , 'ALL')   
            #위의 두 코드에서 오류가 발생, 일정 시간이 되면 연결을 끊어버리는 듯함, 그래서 연결 검사 연결을 시도하게 코드를 변경
            retcode,message = imap.uid('search',"UNSEEN")
    
            length = len(message[0].split())
    
            if(length != 0):
                itemlist = data[0].split()
                if(mostrecent == itemlist[-1]):
                    toaster.show_toast("메일", "새로운 메일이 없습니다.")
                    time.sleep(10)
                    continue
                        
                else:
                        
                    file_stat, mostrecent = Main_Function(imap, itemlist, file_)
                        
                    if(file_stat == True):
                        final_path = os.getcwd()
                        os.system("rmdir /s /q Attachments")
                        file_ == False # 다시 초기화
                        time.sleep(10)#너무 자주 출력되서 넣음
                       
                    else:
                        file_ == False
                        continue
            else:
                continue
                
            imap.logout()

        
 
if __name__ == "__main__":
    
    connect()
    
    
   