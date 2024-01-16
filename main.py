import time

import subprocess

import speech_recognition as sr
import win32com.client
import webbrowser
import os


def say(text_to_speak):
    spk = win32com.client.Dispatch("SAPI.SpVoice")
    spk.Speak(text_to_speak)


def Command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.2)
        r.pause_threshold = 1
        r.energy_threshold = 300
        print("Listening...")
        audio = r.listen(source)
        try:
            print("recognizing")
            query = r.recognize_google(audio, language="en-in")
            query = query.lower()
            print(f"User said: {query}")
            return query
        except sr.UnknownValueError:
            print("Sorry, I did not get that. Please try again.")
            return ""


if __name__ == "__main__":
    say("hello, I am Jarvis")
    while True:
        query = Command()

        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.org"],
                 ["google", "https://www.google.com"]]

        for site in sites:
            if f"open {site[0]}".lower() in query:
                say(f"opening {site[0]} sir....")
                webbrowser.open(site[1])

        if "open music" in query:
            music = r"C:\Users\ratho\Downloads\glossy-168156.mp3"
            print("Music file path:", music)

            os.startfile(music)
        if "open spotify" in query:
            vscode_path = r"C:\Users\ratho\AppData\Roaming\Spotify\Spotify.exe"
            subprocess.Popen([vscode_path])
        # if "open whatsapp" in query:
        #     whatsapp

        if "open cmd" in query:
            cmd_path = os.path.expandvars(r"%windir%\system32\cmd.exe")
            subprocess.Popen([cmd_path])

        import subprocess

        if "open power" in query:
            power_bi_path = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
            try:
                subprocess.Popen([power_bi_path])
            except Exception as e:
                print(f"Error: {e}")

        if "open word" in query:
            word_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            subprocess.Popen([word_path])

        if "ppt" in query:
            power_point = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            subprocess.Popen([power_point])

        if "open excel" in query:
            path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            subprocess.Popen([path])

        if "open chrome" in query:
            chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            subprocess.Popen([chrome_path])


