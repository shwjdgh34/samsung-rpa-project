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

#--- 심화 ------------------------------------------------
name_list = []
team_list = []
real_salary_list = []

#app = xw.App(visible=False) #엑셀을 최소화 된 상태로 생성

for name in concated_df['name']:
    filename = './empinfo/'+name+'.xlsx'
    wb = app.books.open(filename)
    ws = wb.sheets[0]
    df = ws.range('A1').expand().options(pd.DataFrame).value
    
    salary = (concated_df[concated_df['name'] == name]['salary'].values[0]) #해당 사원의 급여
    team = (concated_df[concated_df['name'] == name]['team'].values[0]) #해당 사원의 부서

    unsupport_item_price_sum = df[ df['support'] == "FALSE" ]['price'].sum() #sum으로 처리 해도 됨.
    salary -= unsupport_item_price_sum

    name_list.append(name)
    team_list.append(team)
    real_salary_list.append(salary)
    wb.close()

result = pd.DataFrame({
    'name' : name_list,
    'team' : team_list,
    'salary' : real_salary_list
})
result = result.sort_values(by='salary', ascending=False, ignore_index=True)

wb = app.books.open('result-ex.xlsx') #최소화 된 엑셀에서 해당하는 파일이름의 파일 오픈
ws = wb.sheets.add("result")
ws.range('A1').options(index=False).value = result
wb.save()
wb.close
app.kill()