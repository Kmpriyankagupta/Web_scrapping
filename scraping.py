import json
from requests_html import HTMLSession
import pandas as pd
session = HTMLSession()

resList=[]

for page in range(7):
    r = session.get('https://www.athens.edu/wp-json/upup-baseline/v1/search?s=&post_type=program&filters=%257B%257D&page={}&posts_per_page'.format(page+1))
    res = json.loads(r.text)
    res = res['results']
    for i in res:
        degreeTitle = i['post_title']
        k = i['terms']
        degreeString = ''
        try:
            degreetype = k['degree-type']
            for j in degreetype:
                degreeString += j['name'] + ', '
        except:
            try:
                degreetype = k['program-focus']
                for j in degreetype:
                    degreeString += j['name'] + ', '
            except:
                degreetype = k['department']
                for j in degreetype:
                    degreeString += j['name'] + ', '

        degreeString = degreeString[:-2]

        degreeModalityString = ''
        degreeModality = k['modality']
        for j in degreeModality:
            degreeModalityString += j['name'] + ', '
        degreeModalityString = degreeModalityString[:-2]

        degreeProgramFocus = k['program-focus'][0]['name']

        resList.append([degreeProgramFocus, degreeTitle, degreeString, degreeModalityString])



data = {'Program Focus': [i[0] for i in resList],
        'Program Title': [i[1] for i in resList],
        'Degree Type': [i[2] for i in resList],
        'Study Modality':[i[3] for i in resList]}

df = pd.DataFrame(data)

df1 = df.copy()
df1 = df1[(df1['Study Modality'].str.contains('On-Campus', regex=True))]

df2 = df.copy()
df2 = df2[(df2['Program Focus'].str.contains('Science', regex=True))]

df3 = df.copy()
df3 = df3[(df3['Degree Type'].str.contains('Master', regex=True))]

df4 = df.copy()
df4 = df4[(df4['Study Modality'].isin(['Online']))]


with pd.ExcelWriter('output.xlsx') as writer:
    df.to_excel(writer, sheet_name='Sheet_name_1')
    df1.to_excel(writer, sheet_name='Sheet_name_2')
    df2.to_excel(writer, sheet_name='Sheet_name_3')
    df3.to_excel(writer, sheet_name='Sheet_name_4')
    df4.to_excel(writer, sheet_name='Sheet_name_5')

