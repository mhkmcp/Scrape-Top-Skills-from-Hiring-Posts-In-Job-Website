import requests
import xlwt
from bs4 import BeautifulSoup
from Assessment.custom_functions import job_url, job_detail_url

keywords = ["Accountant", "Retail", "Sales & Marketing",
    "Administration", "Social Care"]

BASE_URL = "https://www.reed.co.uk/"

key = keywords[1]
key = key.lower()

cnt, col, row = 0, 0, 3

wb = xlwt.Workbook()
ws = wb.add_sheet("test", cell_overwrite_ok=True)

ws.write(row, col, 'Title')
col = col + 1
ws.write(row, col, 'ID')
col = col + 1
ws.write(row, col, 'Category')
col = col + 1
ws.write(row, col, 'Skill 1')
col = col + 1
ws.write(row, col, 'Skill 2')
col = col + 1
ws.write(row, col, 'Skill 3')
col = col + 1
ws.write(row, col, 'Skill 4')
col = col + 1
ws.write(row, col, 'Skill 5')
col = col + 1
ws.write(row, col, 'Skill 6')
col = col + 1
ws.write(row, col, 'Skill 7')
col = col + 1
ws.write(row, col, 'Skill 8')
col = col + 1
ws.write(row, col, 'Skill 9')
col = col + 1
ws.write(row, col, 'Skill 10')
col = col + 1
row = row + 1
col = 0


for num in range(1, 31):
    url = job_url(key, num)
    response = requests.get(url)

    soup = BeautifulSoup(response.content, "html.parser")
    articles = soup.select('article', {'class': 'job-result'})

    for job in articles:
        title = job.find('h3').get_text()[1:-1].lower()
        id = job.get('id')[10:]

        if title[-1] == ' ':
            title = title[:-1]

        ws.write(row, col, title)
        col = col + 1
        ws.write(row, col, id)
        col = col + 1
        ws.write(row, col, keywords[0])
        col = col + 1

        detail_url = job_detail_url(key, title, id)

        response2 = requests.get(detail_url)
        soup2 = BeautifulSoup(response2.content, "html.parser")
        skills = soup2.select('div.skills  ul li')

        if len(skills) >= 1:
            for skill in skills:
                skill_ = skill.get_text()

                ws.write(row, col, skill_)

            row = row + 1

        col = 0

wb.save('retail.xls')