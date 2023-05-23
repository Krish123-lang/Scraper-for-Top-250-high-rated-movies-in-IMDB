import requests
import pandas as pd
from bs4 import BeautifulSoup  # importing BeautifulSoup from bs4

try:

    # getting the source code of the URL
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()  # This will raise an error if the URL is invalid

    # takes the html code in text format and parse it
    soup = BeautifulSoup(source.text, 'html.parser')

    # getting the tag and class and also getting the content inside of tr tag
    movies = soup.find('tbody', class_="lister-list").find_all('tr')

    posts = []

    for movie in movies:
        # getting the text written in "a" tag after entering into the "td" tag
        name = movie.find('td', class_="titleColumn").a.text

        # getting the rank in the form of list which is inside the "titleColumn" class
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]

        year = movie.find('td', class_="titleColumn").span.text.strip("()")  # fetching the year

        # fetching the rating of the movie
        rating = movie.find('td', class_="ratingColumn imdbRating").strong.text

        # print(rank, name, year, rating)

        imdb_data = {
            "Rank": rank,
            "Name": name,
            "Year": year,
            "Rating": rating,
            }
        posts.append(imdb_data)

    data = pd.DataFrame(posts)
    data.to_csv("IMDB_ratings.csv", index=None)

except Exception as e:
    print(e)
