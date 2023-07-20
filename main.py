import win32com.client
import speech_recognition as sr
import webbrowser
import openai

speaker = win32com.client.Dispatch("SAPI.Spvoice")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.4)
        r.pause_threshold = 0.5
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Sorry can you repeat again. I'm listening..."


if __name__ == '__main__':
    speaker.Speak("Hello I'm Jarvis AI")
    speaker.Speak("Listening")
    print("Listening...")
    text = takeCommand()
    while True:
        sites = [["YouTube", "https://youtube.com"], ["Google", "https://google.com"],
                 ["Linkedin", "https://linkedin.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in text.lower():
                webbrowser.open(site[1])
                speaker.Speak(f"Opening {site[0]}")
        break
