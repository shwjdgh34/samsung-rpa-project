목표: 팀별 상위 3명의 연봉 평균값 구하기
1. 연도별 employee.xlsx 파일들을 xlwings를 이용하여 읽어들인다.
2. 읽어들인 데이터를 병합하여 하나의 데이터 프레임을 생성한다.
3. result.xlsx 파일을 생성한다.
4. result.xlsx에서 team별로 sheet를 만들어준다.
5. team별 sheet의 A1셀에 각 team 연봉 상위 3명의 연봉 평균을 적어 놓는다.

    5-1. team별로 dataframe을 만들어준다.

    5-2. team별로 생성된 dataframe에서 연봉 상위 3명의 연봉을 평균낸다. (일의자리까지 반올림. 예를들어 1524.7 => 1525)

    5-3. 값을 A1셀에 적어넣는다.

6. result.xlsx을 저장한다.


심화 목표: 연봉 실수령액(salary - 비용)을 내림차순으로 정렬하여 result sheet에 입력하기

(참고할점: 이름 중복 없음.)

1. 각 사원들의 구매목록중에서 support == False인 경우를 salary에서 제외 해서 연봉 실수령액을 구한다.
2. 모든 사원의 연봉 실수령액 정보를 담은 하나의 병합된 dataframe을 만든다.
3. 병합된 dataframe을 내림차순 정렬하여 result sheet에 입력한다.