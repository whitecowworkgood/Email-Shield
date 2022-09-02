
<h3 align="center"><b>:skull: Use Language :skull:</b></h3>
</br>
<p align="center">

<img src="https://img.shields.io/badge/PYTHON-3776AB?style=for-the-badge&logo=PYTHON&logoColor=white">
<img src="https://img.shields.io/badge/json-000000?style=for-the-badge&logo=json&logoColor=white">

<h3 align="center"><b>🛠 Use Tools 🛠</b></h3>
</br>
<p align="center">

<img src="https://img.shields.io/badge/Notepad++-90E59A?style=for-the-badge&logo=Notepad++&logoColor=white">
<img src="https://img.shields.io/badge/VirusTotal-394EFF?style=for-the-badge&logo=VirusTotal&logoColor=white">

<h3 align="center"><b>:feet: Targets :feet:</b></h3>
</br>
<p align="center">

<img src="https://img.shields.io/badge/Gmail-EA4335?style=for-the-badge&logo=Gmail&logoColor=white">
<img src="https://img.shields.io/badge/Microsoft Word-2B579A?style=for-the-badge&logo=Microsoft Word&logoColor=white">
<img src="https://img.shields.io/badge/Microsoft Excel-217346?style=for-the-badge&logo=Microsoft Excel&logoColor=white">
<img src="https://img.shields.io/badge/Microsoft Office-D83B01?style=for-the-badge&logo=Microsoft Office&logoColor=white">


## emailshield
우리가 주로쓰는 이메일에서 제공하는 imap서비스를 파이썬의 imaplib 라이브러리를 활용하여 이메일 수신시 첨부파일을 검사하는 프로그램
해당 프로젝트는 목포대학교 정보보호학과 2022년 3학년 1학기 창의적공학설계 팀 프로젝트의 최종본임을 밝힙니다.
## 사전 작업
1. 프로그램과 연동할 메일의 imap의 서비스 포트를 오픈한다.
2. pip로 tkinter와 win10toast를 필수로 설치한다.

## emailshield 프리뷰
**실행하면 나오는 imap주소와 아이디 비밀번호 입력창**

![tkinter](https://user-images.githubusercontent.com/112620533/187892350-15b2cf85-dda4-4893-91dd-9da003bf7cea.png)

**win10toast를 활용한 사용자 알림 기능**

![win10tost](https://user-images.githubusercontent.com/112620533/187892344-b7094a54-9477-4d6b-8513-fae7d4cae3da.png)

## 원리

1. 우선 메일이 수신되면 첨부파일의 존재유무 확인
2. 첨부파일이 없으면 건너띄고, 있으면 첨부파일을 격리폴더로 다운로드
3. 다운로드된 첨부파일들을 바이러스 토탈에 업로드 하여 결과 확인
4. 최근 이슈가 되었던 XML Injection 공격 유무를 확인하기 위해 settings.xml.rels 파일의 존재 여부 확인
5. settings.xml.rels 파일이 존재하면 악성 url이 저장되어 있는 csv파일의 값과 확인

## 사용법
python imap_connect.py

## 추후 개선안
XML Injection의 경우 지금은 수동으로 csv에 업데이트를 해야 한다. 서버를 두어서 서버에서 xml 주소를 자동으로 가져오는 기능과, 서버 클라이언트 통신을 통해서 자동으로 csv를 업데이트를 할 수 있게 개선할 예정
