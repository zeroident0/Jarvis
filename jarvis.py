import os
import subprocess
import speech_recognition as sr
import pyttsx3
import webbrowser
import datetime
import wikipedia
import pyautogui
import time
import requests
from bs4 import BeautifulSoup

# Initialize text-to-speech engine
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  # 0 for male, 1 for female voice


# Function to speak text
def speak(text):
    print(f"Assistant: {text}")
    engine.say(text)
    engine.runAndWait()


# Function to recognize speech
def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User: {query}\n")
    except Exception as e:
        print("Say that again please...")
        return "None"
    return query.lower()


# Function to open applications
def open_application(app_name):
    app_mapping = {
        'notepad': 'notepad.exe',
        'calculator': 'calc.exe',
        'paint': 'mspaint.exe',
        'word': 'winword.exe',
        'excel': 'excel.exe',
        'powerpoint': 'powerpnt.exe',
        'chrome': 'chrome.exe',
        'edge': 'msedge.exe',
        'visual studio code': 'code',
        'command prompt': 'cmd.exe',
        'file explorer': 'explorer.exe'
    }

    if app_name in app_mapping:
        try:
            subprocess.Popen(app_mapping[app_name])
            speak(f"Opening {app_name}")
        except Exception as e:
            speak(f"Sorry, I couldn't open {app_name}")
    else:
        speak("Application not recognized")


# Function to search files
def search_file(file_name):
    try:
        # This will search in the entire C drive (can be slow)
        # For better performance, specify your common directories
        for root, dirs, files in os.walk('C:\\'):
            if file_name in files:
                file_path = os.path.join(root, file_name)
                os.startfile(file_path)
                speak(f"Opening {file_name}")
                return
        speak("File not found")
    except Exception as e:
        speak("An error occurred while searching for the file")


# Function to control volume
def set_volume(level):
    try:
        if 'mute' in level:
            pyautogui.press('volumemute')
            speak("Volume muted")
        elif 'max' in level or 'maximum' in level:
            for _ in range(50):
                pyautogui.press('volumeup')
            speak("Volume set to maximum")
        elif 'min' in level or 'minimum' in level:
            for _ in range(50):
                pyautogui.press('volumedown')
            speak("Volume set to minimum")
        else:
            speak("Volume command not recognized")
    except Exception as e:
        speak("Could not control volume")


# Function to get weather
def get_weather(city="your city"):
    try:
        base_url = "http://api.openweathermap.org/data/2.5/weather?"
        api_key = "your_api_key_here"  # You need to get this from openweathermap.org
        complete_url = f"{base_url}appid={api_key}&q={city}"
        response = requests.get(complete_url)
        data = response.json()

        if data["cod"] != "404":
            main = data["main"]
            temperature = main["temp"] - 273.15  # Convert from Kelvin to Celsius
            humidity = main["humidity"]
            weather_desc = data["weather"][0]["description"]

            speak(f"Temperature in {city} is {temperature:.1f} degrees Celsius")
            speak(f"Humidity is {humidity} percent")
            speak(f"Weather description: {weather_desc}")
        else:
            speak("City not found")
    except Exception as e:
        speak("Could not fetch weather information")


# Main function
if __name__ == "__main__":
    speak("Hello, I am your AI assistant. How can I help you today?")

    while True:
        query = take_command().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            try:
                results = wikipedia.summary(query, sentences=2)
                speak("According to Wikipedia")
                speak(results)
            except:
                speak("No results found on Wikipedia")

        elif 'open' in query:
            app_name = query.replace('open', '').strip()
            open_application(app_name)

        elif 'search file' in query:
            file_name = query.replace('search file', '').strip()
            search_file(file_name)

        elif 'time' in query:
            str_time = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {str_time}")

        elif 'date' in query:
            str_date = datetime.datetime.now().strftime("%d %B %Y")
            speak(f"Today's date is {str_date}")

        elif 'volume' in query:
            set_volume(query)

        elif 'weather' in query:
            if 'in' in query:
                city = query.split('in')[-1].strip()
                get_weather(city)
            else:
                get_weather()

        elif 'search' in query:
            search_query = query.replace('search', '').strip()
            url = f"https://www.google.com/search?q={search_query}"
            webbrowser.open(url)
            speak(f"Here are the search results for {search_query}")

        elif 'type' in query:
            text_to_type = query.replace('type', '').strip()
            pyautogui.write(text_to_type)

        elif 'press' in query:
            key = query.replace('press', '').strip()
            pyautogui.press(key)
            speak(f"Pressed {key}")

        elif 'screenshot' in query:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"screenshot_{timestamp}.png"
            pyautogui.screenshot(filename)
            speak("Screenshot taken and saved")

        elif 'shutdown' in query:
            speak("Shutting down the computer in 10 seconds")
            time.sleep(10)
            os.system("shutdown /s /t 1")

        elif 'restart' in query:
            speak("Restarting the computer in 10 seconds")
            time.sleep(10)
            os.system("shutdown /r /t 1")

        elif 'sleep' in query:
            speak("Putting the computer to sleep")
            os.system("rundll32.exe powrprof.dll,SetSuspendState 0,1,0")

        elif 'exit' in query or 'quit' in query or 'goodbye' in query:
            speak("Goodbye! Have a nice day.")
            exit()

        else:
            speak("I didn't understand that command. Can you please repeat?")