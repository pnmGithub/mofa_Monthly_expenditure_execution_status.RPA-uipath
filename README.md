## [외교부 월별세출집행현황]  ##
> 외교부 월별 세출집행현황(검색기간 설정) 데이터를 추출하여 결과 파일 작성
> 제공한 템플릿 양식을 Copy하여 결과파일 작성    
> REFramework로 구현 (QueueItem이 아닌 **DataTable.DataRow** 이용)     
> **TransactionData = 월별 집행액 정보 [항목], [연도], [월], [집행액]**


#### [작업순서] ####
> _이해를 돕기위한 작업가이드.xlsx 파일 참고
1. Template.xlsx 파일을 Output로 Copy하기
2. DataTable 생성. [항목], [연도], [월], [집행액] 으로 생성. Config에 설정한 날짜 기간 데이터 넣기
3. 브라우저 열어서 데이터 추출할 페이지로 이동
4. 웹 페이지에서 기간 세팅 후 [검색] 클릭. 데이터스크래핑 - Datatable 생성해서 저장, Datatable 초기화 하고 저장하기
5. 엑셀에 월별 데이터 저장하기 (집행액=0 삭제) - 기본 엑셀시트를 복사해서 사용하기(스타일 유지위해)
6. 월별 합계 구하는 엑셀함수(SUM())를 TransactionData의 [집행액] 컬럼에 쓰기
7. 엑셀 저장 - 합계 추가


#### [Config.xlsx 참고] ####
* webURL = 데이터 수집할 웹 페이지 주소
* saveFolderPath = 파일 저장할 폴더 경로
* saveFilePath = 저장할 파일명
* copyFilePath = Copy할 파일 경로(폴더 포함)
* myName = 본인이름
* searchStartYear = 검색 시작년
* searchStartMonth = 검색 시작월
* searchEndYear = 검색 종료년
* searchEndMonth = 검색 종료월


#### [How It Works] ####

1. **INITIALIZE PROCESS**
   + Config 파일 세팅
   + 브라우저 강제 종료
   + Template.xlsx 파일을 Output로 Copy하기
   + DataTable 생성. [항목], [연도], [월], [집행액] 으로 생성. Config에 설정한 날짜 기간 데이터 넣기
   + 브라우저 열어서 데이터 추출할 페이지로 이동

2. **GET TRANSACTION DATA**
   + Get transaction item 비활성
   + TransactionNumber가 DataTable Row 수만큼 반복을 위한 처리 -  TransactionData의 Row 갯수 체크해서 TransactionItem에 현재 Row 정보 할당

4. **PROCESS TRANSACTION**
   + SetTransactionStatus에서 필요없는 부분 비활성. 인수설정 ( TransactionData, TransactionNumber)
   + 웹 페이지에서 검색기간 세팅 후 [검색] 클릭
   + 데이터스크래핑 - Datatable 생성해서 저장, Datatable 초기화 하고 저장하기
   + 엑셀에 월별 데이터 저장하기 (집행액=0 삭제) - 기본 엑셀시트를 복사해서 사용하기(스타일 유지위해)
   + 월별 합계 구하는 엑셀함수(SUM())를 TransactionData의 [집행액] 컬럼에 쓰기

4. **END PROCESS**
   + 엑셀 저장 - 합계 추가
   + 브라우저 닫기

* * *
![mofa_Monthly_expenditure_execution_status_guide](https://github.com/pnmGithub/mofa_Monthly_expenditure_execution_status.RPA-uipath/assets/149296871/26172f88-9339-413b-9419-3feb29a03297)   
* * *

### REFrameWork Template ###
**Robotic Enterprise Framework**
