'''
동일 구조로 작성된 엑셀파일을 통합하는 Python code
NTIS 과제 정보, KIPRIS 특허 정보 파일들을 사용하여 디버깅 
특이사항 : 이전 경력 업무에서 작성한 각각의 파일 통합 코드(Google Colab에서 구동)를 갈무리하여 Python 실행파일로 작성
'''

#라이브러리 준비
import pandas as pd
from time import sleep
import platform
import os
import sys
import xlrd
import openpyxl
from openpyxl import load_workbook

#함수 작성-화면 지우기
runningSystem=platform.system() #구동환경 확인
def clear(time):
    sleep(time)
    if runningSystem=="Windows":
        os.system('cls') #windows os
    elif runningSystem=="Darwin":
        os.system('clear') #mac os

#함수 작성-디렉토리 생성
def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print ('Error: Creating directory. ' +  directory)

#함수 작성-계속 진행 확인
def continueNext():
    print("\n계속 진행하시겠습니까?")
    while True:
        command=input("입력(Y/N):")
        if command=='Y' or command=='y':
            break
        elif command=='N' or command=='n':
            sys.exit()
        else:
            print("올바른 입력이 아닙니다.")

#메인
clear(0.5)
print("엑셀 파일 통합 프로그램\n\n본 프로그램은 동일 구조(시트,프레임)의 엑셀 파일을 하나로 통합하는 프로그램입니다.\n시트의 순서가 다르거나 파일별 열이 다르게 저장된 파일의 경우 정상적인 동작이 불가능합니다.")
continueNext() #진행 확인
clear(0.5)

#디렉토리 내의 파일 리스트
path = "./input_file"
createFolder(path)
current_path = os.getcwd()
print(current_path)
print("위 경로에 통합대상 excel 파일을 이동하여 주세요.")
input("임의의 키를 입력하면 다음 단계로 진행합니다.")
clear(0)#화면 clear

#폴더 내 파일 목록
file_list = os.listdir(path)
file_list_xls = [file for file in file_list if file.endswith(".xls")]
#print ("file_list: {}".format(file_list_xls))
print("폴더 내 파일 수: ",len(file_list_xls),"개")
print("참고용 폴더 내 파일 목록(목록 상위 5개)")
for i in range(5):
    try:
        print(file_list_xls[i])
        #print(path+"/"+str(file_list_xls[i]))
    except:
        pass
continueNext() #진행 확인

#sample cheack #리스트 첫번째 엑셀파일을 사용하여 복수의 시트 확인
#sheetName=openpyxl.load_workbook(filename=path+"/"+str(file_list_xls[0])) #xlsx
sheetName=xlrd.open_workbook(path+"/"+str(file_list_xls[0])) #xls
snList=list(sheetName.sheet_names())
if len(snList)>1: #시트가 2개 이상일 경우 시트 확인
    print("\n통합 대상 파일에서 복수의 시트를 확인했습니다.")
    print(snList)
    print("통합할 시트의 번호를 입력하여 주세요(1부터).")
    targetSheet = int(input("입력: "))-1
    print("'"+snList[targetSheet]+"' 시트로 통합을 진행합니다.")
else:
    targetSheet=0
clear(2)

#sample cheack #1번 행 이외에 column이 작성된 경우를 고려한 행단위 조회
i=0
while True:
    clear(0)
    print("데이터프레임의 columns를 확인합니다.")
    print(pd.read_excel(path+"/"+str(file_list_xls[0]),sheet_name=targetSheet,header=i).columns)
    command = input('위 데이터가 맞습니까?(Y/N, 이전:B)')
    if command == "Y" or command=='y':
        head=i
        print("해당 columns를 지정하여 진행합니다.")
        break
    elif command == "N" or command=='n':
        i+=1
    elif command == "B" or command=='b':
        i-=1
    else:
        print('올바른 입력이 아닙니다.')

#디렉토리 파일명을 사용하여 파일 불러오기
df=[]
error_file=[]
for i in range(len(file_list_xls)):
    try:
        df.append(pd.read_excel(path+"/"+str(file_list_xls[i]),sheet_name=targetSheet,header=head))
    except:
        error_file.append(i)
    print('진행도: '+str(i)+'/'+str(len(file_list_xls)),end='\r')
print(str(len(df))+'건'+'불러오기 완료')

#에러가 발생하여 불러오지 못한 파일 명
if len(error_file)>0:
    print("불러오기 실패: "+str(len(error_file)),"개")
    for i in range(len(error_file)):
        print(file_list_xls[error_file[i]])
    continueNext() #진행 확인

#데이터 출처 확인을 위해 파일명을 사용하여 DataFrame 가장 앞에 입력
print("\n데이터 출처 구분을 위한 column을 추가합니다.")
#index_txt = str(input("입력 :"))
index_txt = "데이터 출처"
for i in range(len(df)):
    df[i].insert(0,index_txt,''.join(file_list_xls[i].split('.xls')))
print('column 삽입 완료')

#df 리스트 안에 있는 dataframe concat실행
concat_df=pd.concat(df)

#전체 데이터 확인
print("전체 데이터 수: ",len(concat_df))

#저장 확인
print("\n통합 파일 저장을 진행합니다.",end=" ")
continueNext() #진행 확인
clear(1)
print("저장할 파일 이름을 입력해 주세요.\n파일은 실행폴더내 output 폴더에 저장됩니다.\nxlsx확장자로 저장됩니다.")
fileName = str(input("확장자를 제외한 파일명 입력: "))

#concat_df 저장
path = "./output_file"
createFolder(path)
#concat_df.to_excel(path+'/'+fileName+'.xlsx',index=False, engine='xlsxwriter')
writer = pd.ExcelWriter(path+'/'+fileName+'.xlsx', engine='xlsxwriter',engine_kwargs={'options':{'strings_to_formulas': False}})
concat_df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()

clear(0)
print("통합 완료")
input("임의의 키를 입력하면 종료합니다.")
