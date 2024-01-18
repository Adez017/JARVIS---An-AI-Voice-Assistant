import time

import subprocess

import speech_recognition as sr
import win32com.client
import webbrowser
import os
import openai

#function to convert for speak as response
def say(text_to_speak):
    spk = win32com.client.Dispatch("SAPI.SpVoice")
    spk.Speak(text_to_speak)


openai.api_key = "your_api_key"
# Update the model name because the text da vinvci is suspended
OPENAI_MODEL = "gpt-3.5-turbo-instruct"

chat_str = ""

#this function will create a response to the command given as voice command
def chat(query):
    global chat_str
    print(chat_str)

    chat_str += f"User: {query}\n Jarvis: "
    response = openai.Completion.create(
        model=OPENAI_MODEL,
        prompt=chat_str,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    say(response["choices"][0]["text"])
    chat_str += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]

#the function will be going to display the output of the response generate by the ai openai
def ai(prompt):
    text = f"OpenAI response for Prompt: {prompt} \n *************************\n\n"
    response = openai.Completion.create(
        model=OPENAI_MODEL,
        prompt=prompt,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    text += response["choices"][0]["text"]
    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    text += response["choices"][0]["text"]
    os.makedirs("Openai", exist_ok=True)
    with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip() }.txt", "w") as f:
        f.write(text)



#the command function will be going to take the command as input 

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

       #for opening google,wikipedia and youtube 
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.org"],
                 ["google", "https://www.google.com"]]

        for site in sites:
            if f"open {site[0]}".lower() in query:
                say(f"opening {site[0]} sir....")
                webbrowser.open(site[1])
#for opening the glossy music on the local computer 
        if "open music" in query:
            music = r"C:\Users\ratho\Downloads\glossy-168156.mp3"
            print("Music file path:", music)

            os.startfile(music)
            #for opening the spotify on our local computer 
        elif "open spotify" in query:
            vscode_path = r"C:\Users\ratho\AppData\Roaming\Spotify\Spotify.exe"
            subprocess.Popen([vscode_path])
    #for opening command prompt

        elif "open cmd" in query:
            cmd_path = os.path.expandvars(r"%windir%\system32\cmd.exe")
            subprocess.Popen([cmd_path])

        
#for opening the powerBI presentation on our computer
        if "open power".lower() in query.lower():
            power_bi_path = r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe"
            try:
                subprocess.Popen([power_bi_path])
            except Exception as e:
                print(f"Error: {e}")
#for opening word on our computer (local)
        elif "open word".lower() in query.lower():
            word_path = r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            subprocess.Popen([word_path])
#for opening power point on our computer
        elif "ppt".lower() in query.lower():
            power_point = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
            subprocess.Popen([power_point])
#for opening the excel on our computer
        elif "open excel".lower() in query.lower():
            path = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            subprocess.Popen([path])
#for opening the chrome on our computer
        elif "open chrome".lower() in query.lower():
            chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            subprocess.Popen([chrome_path])
#this command will be using the the openai model for generating the result 
        elif "Using artificial intelligence".lower() in query.lower():
            ai(prompt=query)
#this will be going to quit the program
        elif "Jarvis Quit".lower() in query.lower():
            exit()
#this command will be reboot or reset the chat 
        elif "reset chat".lower() in query.lower():
            chatStr = ""

        else:
            print("Chatting...")
            chat(query)
