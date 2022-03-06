from win32com.client import Dispatch
from datetime import datetime
from random import choice
from time import sleep
import requests
import json


def read(what_to: str):
    speak = Dispatch("SAPI.SpVoice")
    print(what_to)
    speak.Speak(what_to.strip("-\n"))


# --------------------------- newsapi.org variables ------------------------------#

baseURL = "https://newsapi.org/v2/"
apiKEY = "1ddfaf43bd944293b9610545a5308f87"
endpoint = ["everything", "top-headlines"]
country = ["in", "us"]
category = ["general", "sports", "business", "entertainment", "health", "science", "technology"]
source = ""  # source
q = ""  # search query
pageSize = 20  # number of results per page
page = 1  # number of pages

example_url = "https://newsapi.org/v2/top-headlines?country=in&apiKey=1ddfaf43bd944293b9610545a5308f87"
headline_url = f"{baseURL}{endpoint[1]}?country={country[0]}&category={category[5]}&apiKey={apiKEY}"
news_details = json.loads(requests.get(headline_url).text)


# --------------------------- datetime variables ------------------------------#

week_day = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
month = [None, "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
         "November", "December"]
day = datetime.now().day
day_suffix = "st" if day in [1, 21, 31] else "nd" if day in [2, 22] else "rd" if day in [3, 23] else "th"
hour = datetime.now().hour
hour_suffix = "Morning" if 11 >= hour >= 4 else "Noon" if 14 >= hour >= 12 else \
    "Afternoon" if 17 >= hour >= 15 else "Evening" if 23 >= hour >= 18 else "Night"


# --------------------------- speaking variables ------------------------------#

news_intro = ["reading", "viewing", "approaching", "revealing", "uncovering", "it's", "this is", "it's", "here's",
              "there's", "listen", "see", "look at", "take a look at", "have a look at", "presenting", "opening"]
news_num = lambda i: "st" if i in [1, 21, 31, 41, 51, 61, 71, 81, 91] else \
    "nd" if i in [2, 22, 32, 42, 52, 62, 72, 82, 92] else "rd" if i in [3, 23, 33, 43, 53, 63, 73, 83, 93] else "th"
news_headline = ["headline", "highlight", "title"]
news_everything = []
# news_brief = ["a brief description", "news brief", "something about it", "something about the news", "news flash"]
news_source = ["from", "extracted from", "collected from", "news source", "courtesy", "published by"]


# --------------------------- main program ------------------------------#

article_num = 1

if news_details["status"] == "ok":
    read(f"Good {hour_suffix}. Today is {week_day[(datetime.now().weekday())]}, the {day}{day_suffix} day of"
         f" {month[datetime.now().month]} in {datetime.now().year} and it's now {hour}:{datetime.now().minute}.")
    read("I'm presenting the top news of the day from the city")
    print("---------------------------------------------------------------------\n\n")

    # sleep(1)
    for article in news_details["articles"]:
        print(f"{article_num} :", end="")
        read(f"{choice(news_intro)} the {article_num}{news_num(article_num)} {choice(news_headline)}")
        read(f"-- {article.get('title').split(' - ')[0]}")
        # read(f"- {choice(news_brief)}")
        # read(f"-- {article.get('description')}")
        print(f"-- For full content visit: {article.get('url')}")
        read(f"-{choice(news_source)}: {article.get('title').split(' - ')[1]}\n")

        article_num += 1
    else:
        read("\nThat's all for now. Come again after some hours to get the updates.")
else:
    read("Sorry, news cannot be fetched for some internet connection issue. Try again after some times.")
