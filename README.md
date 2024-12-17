# Excel_Growth_Analysis
Excel Processor with Growth Rapid growth Analysis openpyxl only.py
이 스크립트는 다음과 같은 기능을 제공합니다:

GUI 인터페이스로 파일 선택과 진행 상황을 시각적으로 보여줍니다
성장 상품 분석:

성장성 >= 0
검색량 >= 8000
쇼핑성키워드 = True
경쟁률 < 4


급성장 상품 분석:

성장성 >= 0.15
검색량 >= 10000
쇼핑성키워드 = True


결과 파일은 요청하신 형식대로 저장됩니다:

원본파일명_(성장/급성장)_날짜_시간.xlsx
주요 변경사항:

pandas 의존성 제거
openpyxl만 사용하여 모든 데이터 처리 구현
메모리 효율적인 데이터 처리 방식 적용
더 자세한 진행 상황 표시

사용 방법:

스크립트를 .py 파일로 저장
pip install openpyxl 실행
스크립트 실행
GUI에서 파일 선택 후 원하는 분석 옵션 선택
