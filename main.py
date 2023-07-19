import win32com.client
import speech_recognition as sr
import os

speaker = win32com.client.Dispatch("SAPI.Spvoice")

def takeCommand():
    r = sr.Recogniser()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        audio = r.listen(source)
        query = r.recoginze_google(audio, language="en-in")
        print(f"User said: {query}")
if __name__ == "__main__" :
    s = takeCommand();
    speaker.Speak(s)