import pandas as pd

df = pd.read_excel('/Applications/办公自动化/Excel/cs.xlsx', index_col=0, sheet_name='Sheet1')
df.loc[(df.name.notnull()), 'name'] = "Yes"
df.to_excel('/Applications/办公自动化/Excel/10.xlsx')
