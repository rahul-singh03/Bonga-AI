import speech_recognition as sr
import os
import win32com.client
from config import apikey
import openai
import webbrowser
import datetime
speaker = win32com.client.Dispatch("SAPI.SpVoice")

def say(text):
    speaker.Speak(text)

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
        try:
            print("Recognizing....")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said: {query}")
            return query
        except Exception as e:
            return "Some error occurred while taking command."

def ai(prompt):
    openai.api_key = apikey
    text = f"OpenAI response for prompt: {prompt} \n *********************** \n\n"
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=prompt,
        temperature=1,
        max_tokens=500,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    try:
        # print(response["choices"][0]["text"])
        text += response["choices"][0]["text"]
        if not os.path.exists("Openai"):
            os.mkdir("Openai")
        with open(f"Openai/{''.join(prompt.split('intelligence')[1:]).strip()}.txt", "w") as f:
            f.write(text)
        print("Prompt saved successfully!")
    except Exception as e:
        print("Some Error Occurred!")
chat_text = ""
def chat(query):
    global chat_text
    print(chat_text)
    openai.api_key = apikey
    chat_text += f"Rahul: {query}\n Bonga: "
    response = openai.ChatCompletion.create(
        model="gpt-3.5-turbo",
        messages=[
            {
                "role": "user",
                "content": chat_text
            }
        ],
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    say(response["choices"][0]["message"]["content"])
    chat_text += f"{response['choices'][0]['message']['content']}\n"
    return response["choices"][0]["message"]["content"]

if __name__ == '__main__':
    print("Hello, I am Bonga A.I")
    say("Hello, I am Bonga A I")
    while True:
        print("Listening....")
        query = takeCommand()
        # list of websites
        sites = [["youtube", "https://youtube.com"], ["wikipedia", "https://wikipedia.com"],
                 ["google", "https://google.com"]]
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                say(f"Opening {site[0]} Maaleek")
                webbrowser.open(site[1])

        if "using artificial intelligence".lower() in query.lower():
            ai(prompt=query)
        elif "open music" in query:
            musicPath = r"C:\Users\RAHUL KUMAR SINGH\Downloads\real.wav"
            os.startfile(musicPath)
        elif "the time" in query:
            # strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            # say(f"Sir, The time is {strfTime}")
            hour = datetime.datetime.now().strftime("%H")
            min = datetime.datetime.now().strftime("%M")
            say(f"Maaleek, The time is {hour} bajke {min} minutes.")
        elif "exit".lower() in query.lower():
            exit()
        elif "reset chat".lower() in query.lower():
            chat_text = ""
        else:
            if query != "Some error occurred while taking command.":
                print("Chatting.....")
                chat(query)