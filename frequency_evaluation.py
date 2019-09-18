import xlrd, xlwt
import numpy as np
import pandas as pd

from nltk.stem import PorterStemmer

ps = PorterStemmer()
dataset = pd.read_excel('administration.xls')

rows = dataset.iloc[3:, 3:].values
n_rows = np.array(rows)

skill_set = set()
for row in n_rows:
    for cell in row:
        skill_set.add(cell)

freq_skill = dict()

for row in rows:
    for skill in skill_set:
        if skill in row:
            val = freq_skill.get(skill)
            if val is None:
                freq_skill.update({skill:  1})
            else:
                freq_skill.update({skill: val + 1})


wb = xlwt.Workbook()
ws = wb.add_sheet('social_care_frequency')

row, col = 0, 0
ws.write(row, col, 'Skill')
col = 1
ws.write(row, col, 'Frequency')
row, col = 1, 0

for key, val in freq_skill.items():
    ws.write(row, col, key)
    col = col + 1
    ws.write(row, col, val)
    row = row + 1
    col = 0

wb.save("administration_frequency.xls")