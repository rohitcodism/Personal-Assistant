import win32com.client
import speech_recognition as sr
import webbrowser
import openai
import os
import datetime
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import python_weather
import spotipy
import json
import pywhatkit

# OpenAI API #
# API Key: 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

speaker = win32com.client.Dispatch("SAPI.Spvoice")

# Twilio API #
#account_sid = ' AC4834f2764ff1b12135e6d433876fb6d7 '
#auth_token = '200dfddce5f7a137e315ad292a8d579d'
#userName = 'rohitpaulhhs04@gmail.com'

# Spotify API #
#client_id : b69c4958b8f24a4e9455062741a752a1
#client_secret : 3c90bf6ce0704ac493807f7d843dad01

# Spotify Music Functionality #
userName = 'i8umlmo3e5brs22s6t87nsngk'
clientID = 'b69c4958b8f24a4e9455062741a752a1'
clientSecret = '3c90bf6ce0704ac493807f7d843dad01'
redirect_uri = 'https://google.com'



# Weather Function #
async def getweather():
    # declare the client. the measuring unit used defaults to the metric system (celcius, km/h, etc.)
    async with python_weather.Client(unit=python_weather.IMPERIAL) as client:
        # fetch a weather forecast from a city
        weather = await client.get(f'{text}')

        # returns the current day's forecast temperature (int)
        speaker.Speak(f"temperature is {weather.current.temperature} degree Farhenite")

        # get the weather forecast for a few days
        for forecast in weather.forecasts:
            speaker.Speak(forecast)
            break


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.4)
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"You said: {query}")
            return query
        except Exception as e:
            return "Sorry can you repeat again. I'm listening..."


if __name__ == '__main__':
    speaker.Speak("Hello I'm Jarvis")
    speaker.Speak("Listening")
    print("Listening...")
    text = takeCommand()
    while True:
        sites = [["YouTube", "https://youtube.com"], ["Google", "https://google.com"],
                 ["Linkedin", "https://linkedin.com"]]
        apps = [["Spotify", "C:\\Users\\user\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Spotify.lnk"],
                ["Figma", "C:\\Users\\user\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Figma.lnk"],
                ["Notion", "C:\\Users\\user\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Notion.lnk"],
                ["Code", "C:\\Users\\user\\OneDrive\\Desktop\\Visual Studio Code.lnk"]]
        for site in sites:
            if f"Open {site[0]}".lower() in text.lower():
                webbrowser.open(site[1])
                speaker.Speak(f"Opening {site[0]}")
        for app in apps:
            if f"Open {app[0]}".lower() in text.lower():
                os.startfile(f"{app[1]}")
                speaker.Speak(f"Opening {app[0]}")
        if "the time" in text:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            speaker.Speak(f"Okay now its {strfTime}")
        elif "Search" in text:
            speaker.Speak("What do you want to search")
            search_text = takeCommand().lower()
            speaker.Speak(f"Searching {search_text}")
            chrome_options = Options()
            chrome_options.add_experimental_option("detach", True)
            browser = webdriver.Chrome(options=chrome_options)
            speaker.Speak(f"Opening {search_text}")
            browser.get('https://google.co.in/search?q=' + search_text)
        elif ("play song" or "Play Song") in text:
           browser = webdriver.Chrome(Options().add_experimental_option("detach", True))
           browser.fi
        elif "I want to watch YouTube" in text:
            speaker.Speak('Which youtube video do you want to watch ?')
            mkv = takeCommand()
            speaker.Speak(f"playing {mkv}")
            print(f"playing {mkv}...")
            pywhatkit.playonyt(mkv, open_video=True)
        else:
            exit(0)
        break
