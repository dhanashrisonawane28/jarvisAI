import datetime
import os
import subprocess

import speech_recognition as sr
import win32com.client
import webbrowser
import openai
import subprocess
import sys

speaker = win32com.client.Dispatch("SAPI.SpVoice")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        print("Listening...")
        try:
            audio = r.listen(source)
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand what you said.")
            return ""
        except sr.RequestError as e:
            print(f"Sorry, there was an error with the request: {e}")
            return ""

# Say the initial greeting
s = "Hello, I am Jarvis. How can I assist you?"
speaker.Speak(s)

while True:
    text = takeCommand()
    # if text:
    #     speaker.Speak(text)
    sites = [["youtube","https://youtube.com"],["wikipedia","https://wikipedia.com"],["google","https://google.com"]]
    for site in sites:
        if f"Open {site[0]}".lower() in text.lower():
            speaker.Speak(f"Opening {site[0]} ")
            webbrowser.open(site[1])

    if "open music" in text:
        musicpath = r"C:\Users\dhanu\Downloads\cleardrum.mp3"
        os.system(f"start {musicpath}")
        subprocess.call(["start", musicpath], shell=True)


    # customize music
    # downloads_folder = r"C:\Users\dhanu\Downloads"
    # if "open music" in text:
    #     # Prompt the user to speak the filename
    #     print("Please speak the name of the song you want to open.")
    #
    #     spoken_filename = takeCommand()
    #
    #     if spoken_filename:
    #         # Search for the spoken filename in the "Downloads" folder
    #         for root, dirs, files in os.walk(downloads_folder):
    #             for file in files:
    #                 if spoken_filename in file.lower():
    #                     song_path = os.path.join(root, file)
    #                     os.system(f"start {song_path}")
    #                     break
    #             if "song_path" in locals():
    #                 break
    #         else:
    #             print(f"Song '{spoken_filename}' not found in Downloads folder.")
    if "the time" in text:
        hour = datetime.datetime.now().strftime("%H")
        minutes = datetime.datetime.now().strftime("%M")
        seconds = datetime.datetime.now().strftime("%S")
        speaker.Speak(f"the time is {hour} hours {minutes} minutes {seconds} seconds")

    # open app
    # if "open visual studio code".lower() in text.lower():
    #     app_path = r"C:/Users/dhanu/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Visual Studio Code.exe"
    #     os.system(f'"start {app_path} "')

