import win32com.client as winC1
import easygui
import datetime
import speech_recognition as sr
import webbrowser
import os
import wikipedia
import smtplib
from time import localtime, strftime

speak = winC1.Dispatch('SAPI.spvoice')

def speaks(audio):
    speak.speak(audio)


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speaks("Good Morning!")

    elif hour >= 12 and hour < 18:
        speaks("Good Afternoon!")

    else:
        speaks("Good Evening!")

    speaks("Suyash Sir. Please tell me how may I help you")


def takeCommand():
    # It takes microphone input from the user and returns string output

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        # print(e)
        print("Say that again please...")
        return "None"
    return query


if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()


        if 'wikipedia' in query:
            speaks('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speaks("According to Wikipedia")
            #print(results)
            speaks(results)

        elif 'the time' in query:
            strTime = [strftime("%X", localtime())]
            print(strTime)
        elif 'goodbye' in query:
            speaks('Goodbye Sir , Jarvis powering off in 3, 2, 1, 0')
            break
        elif 'jarvis' in query:
            speaks('Yes Sir?, What can I doo for you sir?')
        elif 'open google' in query:
            webbrowser.open("google.com")
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        elif 'open github' in query:
            webbrowser.open("github.com")
        elif 'open stackoverflow' in query:
            webbrowser.open("stackoverflow.com")
        elif 'play music' in query:
            music_dir = 'C:\\Users\\suyash\\Music'
            songs = os.listdir(music_dir)
            print(songs)
            os.startfile(os.path.join(music_dir, songs[1]))
        elif 'open profile' in query:
            profile_dir = '"S:\images\profile"'
            os.startfile(profile_dir)
            speaks(f"Sir, the time is {strTime}")
        elif 'your name' in query:
            speaks('My name is Jarvis, at your service sir')



