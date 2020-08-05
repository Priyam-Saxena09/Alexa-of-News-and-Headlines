import requests
import json
from win32com.client import Dispatch
def speak(str):
    speak = Dispatch("SAPI.SpVoice")
    speak.Speak(str)
try:
    head = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=general&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline = json.loads(head.text)
    speak("Welcome to Alexa of Newspaper.")
    speak("Let's have some general headlines")
    for i in range(8):
        print(f"This is the news from {hdline['articles'][i]['source']['name']}")
        print(f"{hdline['articles'][i]['title']}")
        print(f"{hdline['articles'][i]['description']}", end="\n\n") if hdline['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline['articles'][i]['source']['name']}")
        speak(f"{hdline['articles'][i]['title']}")
        speak(f"{hdline['articles'][i]['description']}") if hdline['articles'][i]['description'] != None else print()

    head1 = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=health&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline1 = json.loads(head1.text)
    speak("Now let's have some health headlines")
    for i in range(8):
        print(f"This is the news from {hdline1['articles'][i]['source']['name']}")
        print(f"{hdline1['articles'][i]['title']}")
        print(f"{hdline1['articles'][i]['description']}", end="\n\n") if hdline1['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline1['articles'][i]['source']['name']}")
        speak(f"{hdline1['articles'][i]['title']}")
        speak(f"{hdline1['articles'][i]['description']}") if hdline1['articles'][i]['description'] != None else print()

    head2 = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=entertainment&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline2 = json.loads(head2.text)
    speak("Now let's have some entertainment headlines")
    for i in range(8):
        print(f"This is the news from {hdline2['articles'][i]['source']['name']}")
        print(f"{hdline2['articles'][i]['title']}")
        print(f"{hdline2['articles'][i]['description']}", end="\n\n") if hdline2['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline2['articles'][i]['source']['name']}")
        speak(f"{hdline2['articles'][i]['title']}")
        speak(f"{hdline2['articles'][i]['description']}") if hdline2['articles'][i]['description'] != None else print()

    head3 = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=business&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline3 = json.loads(head3.text)
    speak("Now let's have some business headlines")
    for i in range(8):
        print(f"This is the news from {hdline3['articles'][i]['source']['name']}")
        print(f"{hdline3['articles'][i]['title']}")
        print(f"{hdline3['articles'][i]['description']}", end="\n\n") if hdline3['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline3['articles'][i]['source']['name']}")
        speak(f"{hdline3['articles'][i]['title']}")
        speak(f"{hdline3['articles'][i]['description']}") if hdline3['articles'][i]['description'] != None else print()

    head4 = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=sports&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline4 = json.loads(head4.text)
    speak("Now let's have some sports headlines")
    for i in range(8):
        print(f"This is the news from {hdline4['articles'][i]['source']['name']}")
        print(f"{hdline4['articles'][i]['title']}")
        print(f"{hdline4['articles'][i]['description']}", end="\n\n") if hdline4['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline4['articles'][i]['source']['name']}")
        speak(f"{hdline4['articles'][i]['title']}")
        speak(f"{hdline4['articles'][i]['description']}") if hdline4['articles'][i]['description'] != None else print()

    head5 = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline5 = json.loads(head5.text)
    speak("Now let's have some science headlines")
    for i in range(8):
        print(f"This is the news from {hdline5['articles'][i]['source']['name']}")
        print(f"{hdline5['articles'][i]['title']}")
        print(f"{hdline5['articles'][i]['description']}", end="\n\n") if hdline5['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline5['articles'][i]['source']['name']}")
        speak(f"{hdline5['articles'][i]['title']}")
        speak(f"{hdline5['articles'][i]['description']}") if hdline5['articles'][i]['description'] != None else print()

    head6 = requests.get(
        "https://newsapi.org/v2/top-headlines?country=in&category=technology&apiKey=d6127f8fabb247d38514da47fbedbe86")
    hdline6 = json.loads(head6.text)
    speak("At last,let's have some technology headlines")
    for i in range(8):
        print(f"This is the news from {hdline6['articles'][i]['source']['name']}")
        print(f"{hdline6['articles'][i]['title']}")
        print(f"{hdline6['articles'][i]['description']}", end="\n\n") if hdline6['articles'][i]['description'] != None else print("--",end="\n\n")
        speak(f"This is a news from {hdline6['articles'][i]['source']['name']}")
        speak(f"{hdline6['articles'][i]['title']}")
        speak(f"{hdline6['articles'][i]['description']}") if hdline6['articles'][i]['description'] != None else print()
except Exception as e:
    print("There is Server connection Error.")
    speak("There is Server connection Error.")

