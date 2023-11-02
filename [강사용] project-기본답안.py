import os, glob
import xlwings as xw
import pandas as pd

files = glob.glob('[!~]*_employee.xlsx')
app = xw.App(visible=False) #엑셀을 최소화 된 상태로 생성
df_list = []
for file in files:
    wb = app.books.open(file) #최소화 된 엑셀에서 해당하는 파일이름의 파일 오픈
    ws = wb.sheets[0]
    df = ws.range('A1').expand().options(pd.DataFrame).value
    df_list.append(df)
concated_df = pd.concat(df_list)
concated_df.reset_index(inplace=True)
wb.close()

wb = xw.Book()

teams = concated_df['team'].unique()
for team_name in teams:
    team_df = concated_df[ (concated_df['team'] == team_name)]
    soted_df = team_df.sort_values(by='salary', ascending=False, ignore_index=True)
    result = round( soted_df.head(3)['salary'].mean() )
    ws = wb.sheets.add(team_name)
    ws.range('A1').value = result

wb.save('result-ex.xlsx')
wb.close()
app.kill()