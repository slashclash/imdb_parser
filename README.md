Скрипт парсит фильмографию актера по его IMDB ID с сайта [www.imdb.com/](https://www.imdb.com/) и сохраняет в json и xlsx файлы.
Для парсинга используется библиотека Beautiful Soup, для запросов - Requests, для работы с json и xlsx - json и openpyxl.


The script can scrape the actor's filmography by actor's ID from the [www.imdb.com/](https://www.imdb.com/) and save it in json and xlsx formats.
The script scrapes the title, year of release and the link of each movie in all categories.

Main libraries:
  - Beautiful Soup for parsing.
  - Requests for HTTP requests.
  - openpyxl for xlsx files.
  - json for json files.
