import pandas as pd


df=pd.read_excel("new_assign.xlsx")
print(df.head())

#to calculate Percentage
df["Percentage"] = ((df["sub1"] + df["sub2"] + df["sub3"] + df["sub4"])/400)*100

#for Result
df.loc[df['Percentage']<40, 'Result'] = 'failed'
df.loc[df['Percentage']>=40, 'Result'] = 'passed'

#for Grade
df.loc[df['Percentage']<40, 'Grade'] = 'Failed'
df.loc[df['Percentage']>70, 'Grade'] = 'Distinction'
df.loc[df['Percentage'].between(40,50), 'Grade'] = 'Pass'
df.loc[df['Percentage'].between(50,60), 'Grade'] = 'Second Class'
df.loc[df['Percentage'].between(60,70), 'Grade'] = 'First Class'
print(df.head())

#to save the file
writer = pd.ExcelWriter('new_assign1.xlsx')
df.to_excel(writer,'new_sheet')
writer.save()