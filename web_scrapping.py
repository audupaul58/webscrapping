from bs4 import BeautifulSoup
import requests, openpyxl

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'MySheet1'
sheet.append(['Movie rank', 'Movie name', 'Release year', 'Movie rating'])





try:
    webpage = requests.get('https://www.imdb.com/chart/top/')
    webpage.raise_for_status()

    soup = BeautifulSoup(webpage.text, 'html.parser')
    movies = soup.find('tbody', class_='lister-list').find_all('tr')
    for movie in movies:
        name = movie.find('td', class_="titleColumn").a.text
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])

except Exception as e:
    print(e)

excel.save('Movie Rating.Xlsx')

