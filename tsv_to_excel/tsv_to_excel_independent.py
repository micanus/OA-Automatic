import pandas as pd
import json
import datetime
import openpyxl
import tkinter as tk
from tkinter import messagebox

#pyinstaller 사용법
#명령 프롬프트에서 아래 명령어 실행
#cd C:\Users\~ #스크립트 파일이 있는 디렉토리로 이동
#pyinstaller --onefile --noconsole tsv_to_excel_independent.py

#처음 실행시 jinja2 관련 오류 발생시 --hidden-import=jinja2 옵션 추가
#pyinstaller --onefile tsv_to_excel.py --hidden-import=jinja2
#--onefile 옵션: 단일 실행 파일 생성
#--noconsole 옵션: 콘솔 창 없이 실행 (GUI 프로그램용)

#실행 파일은 dist 폴더에 생성됨
#실행 파일과 동일한 디렉토리에 'BookingsReportingData.tsv' 파일을 저장해야 함
#생성된 실행 파일은 'BookingsReportingData_result.xlsx' 파일을 생성/업데이트함
#기존 엑셀 파일이 있으면 새로운 데이터 추가, 없으면 새로 생성
#중복된 행 데이터 제거를 위해 특정 열('예약일시', '고객명', '사원번호', '자산태그') 기준으로 중복 제거
#필요시 컬럼명 변경
#필요시 불필요한 컬럼 제거, 새 컬럼 추가, 컬럼 순서 변경 등 추가 작업 가능
#실행 파일 생성 후, 명령 프롬프트에서 아래 명령어로 실행

def show_message(msg):
    root = tk.Tk()
    root.withdraw()  # 메인 윈도우 숨기기
    messagebox.showinfo("Information", msg)
    root.destroy()   

def main():
    try:
        df = pd.read_csv("BookingsReportingData.tsv", sep="\t", encoding="utf-8") #tsv_읽기
    except FileNotFoundError:
        show_message("'BookingsReportingData.tsv' 파일을 실행 파일과 같은 위치에 저장해주세요.")
        exit()

    #작업하기 용이하도록 컬럼명 변경
    df = df.rename(columns={'Date Time':"예약일시", 
                            'Customer Name':"고객명", 
                            'Customer Email':"고객 Email",
                            'Customer Phone':"고객 연락처",
                            'Customer Address':"고객 주소",
                            'Staff':"담당자 소속",
                            'Staff Name':"담당자 이름",
                            'Staff Email':"담당자 Email",
                            'Service':"서비스",
                            'Location':"장소", 
                            'Duration (mins.)':"소요시간(분)", 
                            'Pricing Type':"가격 유형", 
                            'Price':"가격", 
                            'Currency':"통화",
                            'Cc Attendees':'참석자(참조)',
                            'Signed Up Attendees Count':'신청 수',
                            'Text Notifications Enabled':'문자 알림 설정',
                            ' Custom Fields':'Custom Fields',
                            'Event Type':'이벤트 유형',
                            'Booking Id':'예약 ID',
                            'Tracking Data':'추적 데이터',
                            })

    customFields=pd.DataFrame()
    try:
        addData=[]
        for i in range(len(df)):
            try:
                data_dic=json.loads(df['Custom Fields'][i])
            except Exception as e:
                data_dic={} #Custom Fields 값이 비어있거나 잘못된 경우 빈 딕셔너리로 처리
            data_dic_new = {k.split(" (")[0]: v for k, v in data_dic.items()} #culumns 값이 될 key값 정제 #사원번호, 사용자연락처, 자산태그, 기기 사용(설치) 위치, 임시 대여PC 필요여부, 메모
            pd_data=pd.DataFrame.from_dict([data_dic_new])
            addData.append(pd_data)
        customFields=pd.concat(addData,ignore_index=True)
    except Exception as e:
        show_message("데이터 정리를 진행할 수 없습니다.\ntsv 파일이 손상되었거나 예약 내용이 없는 상태인지 확인이 필요합니다.")
        exit()

    new_df=pd.concat([df.reset_index(drop=True),customFields.reset_index(drop=True)],axis=1)
    new_df = new_df.dropna(axis=1, how="all") #전체 예약에서 각 항목 전부 공란=필요없음

    new_df = new_df.rename(columns={'사원번호':'사원번호', 
                                    '사용자연락처':'연락처', 
                                    '자산태그':'자산태그', 
                                    '기기 사용(설치) 위치':'기기 사용 위치', 
                                    '임시 대여PC 필요여부':'임시PC 필요여부', 
                                    '메모':'메모(추가요청사항)'})

    #불필요한 컬럼 제거
    columns_to_drop = ["고객 Email", "고객 연락처", "고객 주소", 
                       "담당자 소속", "담당자 이름", "담당자 Email",
                       "장소", "소요시간(분)", "가격 유형", "가격", "통화",
                       '참석자(참조)', '신청 수', '문자 알림 설정',
                       'Custom Fields', '이벤트 유형', '예약 ID', '추적 데이터'] #제거할 컬럼명 리스트
    new_df = new_df.drop(columns=columns_to_drop, errors='ignore')

    #컬럼 추가
    #new_df['New Column'] = 'Default Value' #새 컬럼 추가 예시
    #new_df["처리 날짜"] = datetime.datetime.today().strftime("%Y-%m-%d")
    #new_df['완료 여부'] = 'N'

    #컬럼 순서 변경 #원하는 컬럼 순서 리스트
    columns_order = ["서비스", "예약일시", "고객명",  #출처 원본 데이터
                     "사원번호", "연락처", "자산태그", "기기 사용 위치", "임시PC 필요여부", "메모(추가요청사항)"] #출처 Custom Fileds
    new_df = new_df.reindex(columns=columns_order) #reindex로 컬럼 순서 변경

    #중복 제거
    #new_df = new_df.drop_duplicates(subset=['예약일시', '고객명', '사원번호', '자산태그']) #특정 컬럼 기준 중복 제거 #필요시 컬럼명 변경

    '''
    #최초 실행시 생성되는 엑셀 파일 #디렉토리에 없다면 파일을 생성
    if not pd.io.common.file_exists("BookingsReportingData_result.xlsx"):
        try:
            new_df.to_excel("BookingsReportingData_result.xlsx", index=False)
        except PermissionError:
            show_message("'BookingsReportingData_result.xlsx' 파일을 생성할 수 없습니다.\n파일이 열려 있거나, 폴더에 쓰기 권한이 없는 경우입니다.\n엑셀을 종료하거나 권한을 확인한 후 다시 실행해 주세요.")
            exit()
        except FileNotFoundError:
            show_message("파일을 생성할 수 없습니다.\n폴더 경로가 잘못되었거나 존재하지 않습니다.\n경로를 확인한 후 다시 실행해 주세요.")
            exit()
        except Exception as e:
            show_message("파일을 생성할 수 없습니다.\n알 수 없는 오류가 발생했습니다.\n오류 내용: " + str(e))
            exit()
    else:
        #파일이 이미 존재하는 경우, 기존 파일을 열고 새로운 데이터를 추가 #특정 열('예약일시', '고객명')의 값이 동일한, 확정적으로 중복된 행 데이터 제거 필요시 추가 작업 필요
        show_message("기존 파일을 확인 했습니다.\n새 데이터를 추가 하겠습니다.")
        #기존 엑셀 파일 읽기 #사원번호 컬럼을 텍스트로 읽기 #사원번호가 0으로 시작하는 경우가 있어 str로 읽기
        try:
            existing_df = pd.read_excel("BookingsReportingData_result.xlsx", dtype={'사원번호': str})
        except PermissionError:
            show_message("'BookingsReportingData_result.xlsx' 파일이 열려 있어 읽을 수 없습니다.\n엑셀을 종료한 후 다시 실행해 주세요.")
            exit()
        except Exception as e:
            show_message("'BookingsReportingData_result.xlsx' 파일을 읽을 수 없습니다.\n파일이 손상되었거나 엑셀 형식이 아닐 수 있습니다.\n파일을 확인하거나 삭제 후 다시 실행해 주세요.\n\n오류 내용: " + str(e))
            exit()

        #데이터 병합 및 중복 제거
        combined_df = pd.concat([existing_df, new_df], ignore_index=True)
        combined_df = combined_df.drop_duplicates(subset=['예약일시', '고객명', '사원번호', '자산태그']) #특정 컬럼 기준 중복 제거 #필요시 컬럼명 변경
        
        #병합된 데이터를 다시 엑셀 파일로 저장
        try:
            combined_df.to_excel("BookingsReportingData_result.xlsx", index=False)
        except PermissionError:
            show_message("'BookingsReportingData_result.xlsx' 파일이 열려 있어 저장할 수 없습니다.\n엑셀을 종료한 후 다시 실행해 주세요.")
            exit()
    '''

    #new_df.to_excel("BookingsReportingData_result.xlsx", index=False) #excel_저장 #rawdata 작업 #날짜별 생성 필요없음 #일일이 작업하는 경우
    #1일 2회 오전, 오후 생성, 실행 파일명에 작업날짜 추가
    try:
        data_processed = str(datetime.datetime.today().strftime("%Y-%m-%d"))
        date_processed = data_processed + "_AM" if datetime.datetime.now().hour < 12 else data_processed + "_PM"
        new_df.to_excel("BookingsData_result_"+date_processed+".xlsx", index=False) #excel_저장 #작업날짜_파일명에 추가
        show_message("파일을 생성 했습니다.\n내용 확인 후 사용 바랍니다.")
    except PermissionError:
            show_message("'BookingsReportingData_result.xlsx' 파일을 생성할 수 없습니다.\n파일이 열려 있거나, 폴더에 쓰기 권한이 없는 경우입니다.\n엑셀을 종료하거나 권한을 확인한 후 다시 실행해 주세요.")
            exit()
    except FileNotFoundError:
            show_message("파일을 생성할 수 없습니다.\n폴더 경로가 잘못되었거나 존재하지 않습니다.\n경로를 확인한 후 다시 실행해 주세요.")
            exit()
    except Exception as e:
            show_message("파일을 생성할 수 없습니다.\n알 수 없는 오류가 발생했습니다.\n오류 내용: " + str(e))
            exit()

if __name__ == "__main__":
    main()