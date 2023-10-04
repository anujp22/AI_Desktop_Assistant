import datetime
import speech_recognition as sr
import win32com.client
import os
import webbrowser
import openai
from config import apikey

import requests

ChatStr = ""


# chat function to interact with the chatbot

def chat(user_query):
    global ChatStr
    openai.api_key = apikey
    ChatStr += f"Your_Name: {user_query}\n Jarvis:"
    messages = [
        {"role": "system", "content": "You are helpful"},
        {"role": "user", "content": ChatStr}
    ]
    response = openai.ChatCompletion.create(
        messages=messages,
        model="gpt-3.5-turbo",
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    assistant_response = response["choices"][0]["message"]["content"]
    print(assistant_response)
    speaker.Speak(assistant_response)
    ChatStr += f"\nJarvis: {assistant_response}"
    return assistant_response

    # print(f"User Query: {user_query}")
    # print(f"Assistant Response: {assistant_response}")


# AI function to interact with the chatbot and save the conversation in a file

def ai(user_query):
    openai.api_key = apikey
    messages = [
        {"role": "user", "content": "You are helpful"},
        {"role": "user", "content": user_query}
    ]
    response = openai.ChatCompletion.create(
        messages=messages,
        model="gpt-3.5-turbo",
        temperature=1,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )

    assistant_response = response["choices"][0]["message"]["content"]

    # print(f"User Query: {user_query}")
    # print(f"Assistant Response: {assistant_response}")

    if not os.path.exists("Openai"):
        os.mkdir("Openai")

    with open(f"Openai/{''.join(user_query.split('intelligence')[1:]).strip()}.txt", "w") as f:
        f.write(f"User Query: {user_query}\n\nAssistant Response:\n{assistant_response}")


speaker = win32com.client.Dispatch("SAPI.SPVoice")


#  weather function to check the weather of a city

def weather_check(city_query):
    api = f"http://api.weatherapi.com/v1/current.json?key=072cad3a8f1d4fb2bfb185306232409&q={city_query}&aqi=no"
    data = requests.get(api).json()
    print(f"The temperature is {data['current']['temp_f']} F")
    print(f"The humidity is {data['current']['humidity']}")
    print(f"The wind speed is {data['current']['wind_mph']} mph")


# outlook calendar function to create a calendar event

def create_calendar_event(heading, starttime, endtime, body):
    while True:
        try:
            starttime = datetime.datetime.strptime(starttime, "%Y, %m, %d, %H, %M")
            endtime = datetime.datetime.strptime(endtime, "%Y, %m, %d, %H, %M")
            break
        except ValueError:
            print("Invalid date and time format. Please use 'YYYY, MM, DD, HH, MM' format.")
    outlook = win32com.client.Dispatch("Outlook.Application")
    calendar_item = outlook.CreateItem(0)
    calendar_item.StartTime = starttime
    calendar_item.EndTime = endtime
    calendar_item.Body = body
    calendar_item.Subject = heading
    calendar_item.Save()
    speaker.Speak(f"Event created successfully")

# speech recognition function to take user input


def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        # r.pause_threshold = 1
        audio = r.listen(source)
        try:
            query = r.recognize_google(audio, language="en-us")
            print(f"User said: {query}")
            return query
        except Exception:
            return "Some error occurred"


if __name__ == '__main__':
    print('PyCharm')
    # print("write you want to hear")
    # s = input()
    speaker.Speak(f"welcome to the program sir")
    while True:
        print("Listening.....")
        query = take_command()
        sites = [["youtube", "https://www.youtube.com"], ["wikipedia", "https://www.wikipedia.com"],
                 ["google", "https://www.google.com"], ["spotify", "https://www.spotify.com"],
                 ["code", "https://www.neetcode.io"], ["github", "https://github.com"],
                 ["linkedin", "https://www.linkedin.com"], ["stackoverflow", "https://stackoverflow.com"],
                 ["instagram", "https://www.instagram.com"], ["facebook", "https://www.facebook.com"],
                 ["twitter", "https://www.twitter.com"], ["googlemaps", "https://www.google.com/maps"]]
        # opening websites
        for site in sites:
            if f"Open {site[0]}".lower() in query.lower():
                speaker.Speak(f"Opening {site[0]} sir....")
                webbrowser.open(site[1])
        # printing time
        if "the time" in query:
            strfTime = datetime.datetime.now().strftime("%H:%M:%S")
            print(strfTime)
            speaker.Speak(f"Sir the time is {strfTime}")

        # todo: create list for os opening functions
        elif "Using artificial intelligence".lower() in query.lower():
            ai(user_query=query)
        # todo: openAI more creative functions
        # exiting the loop
        elif "Jarvis Quit".lower() in query.lower():
            speaker.Speak(f"thank you for using my services")
            exit()
        elif "what is weather" in query.lower():
            speaker.Speak("Please say the name of city \n NOTE: Just speak the name of the city no other words.")
            city = take_command()
            weather_check(city)
        elif "create a calendar event".lower() in query.lower():
            speaker.Speak("Please enter the title of the event")
            title = take_command()
            speaker.Speak("Please enter the start time of the event")
            start_time = input("Please type the start time of the event in 24 hour format YYYY, MM, DD, HH, MM")
            speaker.Speak("Please enter the end time of the event")
            end_time = input("Please type the end time of the event in 24 hour format YYYY, MM, DD, HH, MM")
            speaker.Speak("Please enter the description of the event")
            description = input("Please type the description of the event")
            create_calendar_event(title, start_time, end_time, description)
        else:
            print("You are talking A.I. Desktop Assistant developed by YourName")
            chat(query)
