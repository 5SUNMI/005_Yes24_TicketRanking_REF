#[시험1] 005_Yes24_TicketRanking_REF

**[How It Works]**

1. **INITIALIZE PROCESS**
    - Config.xlsx 설정
    - 브라우저(Chrome, Excel) 강제종료
    - 폴더, 파일 관리(폴더 있으면 삭제 후 생성, 파일 있으면 삭제)
    - 예스24 티켓 - 랭킹 카테고리 가져오기 (dt_TransactionData에 저장)
        1. 랭킹 카테고리 가져오기 (Find Children )
        2. List에 입력
    - 카테고리별 엑셀파일 사전 준비 - 시트 순서대로 작성되도록
      
2. **GET TRANSACTION DATA**
    - Get transaction item 비활성
    - 처리할 TransactionItem 체크
      
3. **PROCESS TRANSACTION**
    - 카테고리 별 NavigateTo로 랭킹 페이지 이동
    - Merge용 데이터테이블 생성 - 스크래핑 결과 없을 경우 오류 방지
    - 베스트 영역 스크래핑 ([랭킹], [공연명], [공연기간 공연장소], [예매율])
    - 랭킹 영역 스크래핑 ( [랭킹], [공연명], [공연기간 공연장소], [예매율])
    - DataTable 합치고 컬럼 정리 ([랭킹], [공연명], [공연기간], [공연장소], [예매율])
    - 카테고리별로 엑셀에 쓰기
      
4. **END PROCESS**
    - 엑셀 레이아웃 정리
    - 브라우저 닫기
    - 브라우저 강제 종료
      
