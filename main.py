import requests
from xlsxwriter import Workbook
from bs4 import BeautifulSoup
from datetime import datetime

start_url = 'http://blueir.investproductions.com/investor-relations/press-releases/2017'
end_url = 'http://blueir.investproductions.com/investor-relations/press-releases/2021'
header = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36',
}

wb = Workbook('Releases_Template.xlsx')
sheet = wb.add_worksheet(name = 'Press Releases')
bold = wb.add_format({"bold":1})
sheet.write("A1","Date",bold)
sheet.write("B1","Press Releases",bold)
row = 1
col = 0


for i in range(5) :
    start_url = 'http://blueir.investproductions.com/investor-relations/press-releases/'+str(2021-i)
    with requests.session() as s:
        r = s.get(start_url, headers=header)

        soup = BeautifulSoup(r.content, 'html.parser') 
        press_releases = soup.find_all(class_="RowStyle")

        for press_release in press_releases :
            date = press_release.find(class_="date").get_text()
            release = press_release.find(class_="title-link").get_text()
            release = release.split('\n')[len(release.split('\n'))-2]
            sheet.write(row,col, date)
            sheet.write(row,col+1, release)
            row += 1
wb.close()
print("Successfully saved!")


# save_press_releases()