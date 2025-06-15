import speech_recognition as sr
import win32com.client
import webbrowser
import datetime
import subprocess
import os

speaker = win32com.client.Dispatch("SAPI.Spvoice")

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 1
        print("Listening...")
        audio = r.listen(source)
        try:
            print("Recognizing.....")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query.lower()
        except sr.UnknownValueError:
            print("Sorry, I couldn't understand.")
            return ""
        except sr.RequestError:
            print("Error with speech recognition service.")
            return ""

print("AI Jarvis")
speaker.speak("Hello, I am Jarvis AI")
speaker.speak("How can I help you?")
speaker.speak("Listening...")

sites = [
    ["youtube", "https://youtube.com"],
    ["chat_gpt", "https://copilot.microsoft.com"],
    ["google", "https://google.com"]
]
musics = [
    ["faded", r"E:\song\audio song\1 Alan walker Faded.mp3"],
    ["ram", r"E:\song\audio song\Ram Siya Ram - Adipurush 128 Kbps.mp3"],
    ["angel", r"E:\song\audio song\40 oh angel sent from up above.mp3"]
]
apps = [
    ["camera", "start microsoft.windows.camera:"],
    ["screenshot", "start snippingtool"],
    ["word", "start winword"],
]

while True:
    text = takecommand()

    if text:
        speaker.speak(text)

        # Open websites dynamically
        for site in sites:
            if f"open {site[0]}" in text:
                speaker.speak(f"Opening {site[0]}..")
                webbrowser.open(site[1])  #correct method for open URLs:
                break

        # Open music files properly
        for music in musics:
            if f"open music {music[0]}" in text:
                speaker.speak(f"Opening {music[0]} song")
                os.startfile(music[1])  # Correct method for opening local files

        # Open applications
        for app in apps:
            if f"open {app[0]}" in text:
                speaker.speak(f"Opening {app[0]} app")
                subprocess.run(app[1], shell=True) # used to run system commands and external programs from within a script.
                # it useful for launching applications, executing shell commands, and interacting with other programs.
                # the shell= True argument in subprocess.run() (or subprocess.Popen()) tells Python to execute the command through the system's shell (like cmd.exe on Windows or /bin/sh on Linux).
                break

        # Announce the current time
        if "the time" in text:
            current_time = datetime.datetime.now().strftime("%I:%M %p")
            speaker.speak(f"The current time is {current_time}")
            print(f"The current time is {current_time}")

        # Exit condition with flexibility for variations
        if text in ["good bye", "goodbye"]:
            speaker.speak("Goodbye!..sir")
            break
