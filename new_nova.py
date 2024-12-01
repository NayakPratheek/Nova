import subprocess
import speech_recognition as sr
import pyttsx3
import datetime
import webbrowser
import yt_dlp
import pyautogui
import os

# Initialize recognizer and text-to-speech engine
recognizer = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)
rate = engine.getProperty('rate')
engine.setProperty('rate', 150)

# Function to speak text aloud
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Function to listen for the wake word and user commands
def listen():
    with sr.Microphone() as source:
        print("Listening for wake word...")
        recognizer.adjust_for_ambient_noise(source, duration=0.5)  # Adjust faster
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=5)  # Limit listening time

        try:
            print("Recognizing wake word...")
            command = recognizer.recognize_google(audio).lower()
            print(f"Wake word detected: {command}")

            if "nova" in command:  # Wake word
                speak("How can I help you?")
                return get_command()

        except sr.UnknownValueError:
            pass
        except sr.RequestError:
            speak("Sorry, my speech service is down.")
        return ""

# Function to capture detailed commands after the wake word
def get_command():
    with sr.Microphone() as source:
        print("Listening for command...")
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)

        try:
            print("Recognizing command...")
            command = recognizer.recognize_google(audio).lower()
            print(f"User said: {command}")
            return command

        except sr.UnknownValueError:
            speak("Sorry, I did not understand that.")
        except sr.RequestError:
            speak("Sorry, my speech service is down.")
        return ""

# Function to search YouTube for a query
def search_in_yt(query):
    try:
        search_url = f"https://www.youtube.com/results?search_query={query}"
        speak(f"Here are your search results for {query} on YouTube.")
        webbrowser.open_new(search_url)
    except Exception as e:
        print("Error:", e)
        speak("Sorry, I couldn't search YouTube.")

# Function to play music on YouTube
def play_youtube_music(query):
    try:
        ydl_opts = {"format": "bestaudio/best", "quiet": True}
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            search_result = ydl.extract_info(f"ytsearch:{query}", download=False)
            if search_result['entries']:
                video_url = search_result["entries"][0]["webpage_url"]
                speak(f"Playing the first result on YouTube.")
                webbrowser.open(video_url)
            else:
                speak("Sorry, no results found.")
    except Exception as e:
        print("Error:", e)
        speak("I couldn't play the music.")

# Function to pause YouTube video
def pause_video():
    speak("Pausing the video.")
    pyautogui.press('k')

# Function to open Notepad and write what the user says
def type_in_notepad():
    speak("Opening Notepad.")
    notepad_path = r"C:\Windows\System32\notepad.exe"  # Default path for Notepad
    os.startfile(notepad_path)
    speak("Notepad is now open. You can start speaking, and I will type your words. Say 'save it' to stop.")

    while True:
        text_to_write = listen_for_command()
        if text_to_write:
            if "save it" in text_to_write:
                speak("Saving the document and stopping.")
                break
            else:
                pyautogui.typewrite(text_to_write, interval=0)
        else:
            speak("You didn't say anything.")

# Function to listen for user command input (continuous listening)
def listen_for_command():
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=0.5)
        print("Listening for command...")
        audio = recognizer.listen(source, timeout=5, phrase_time_limit=10)

        try:
            command = recognizer.recognize_google(audio).lower()
            print("Command:", command)
            return command
        except sr.UnknownValueError:
            speak("Sorry, I didn't understand.")
        except sr.RequestError:
            speak("Network error.")
    return ""

# Function to perform calculator operations using pyautogui
def perform_calculator_operation(command, operator_key):
    try:
        words = command.split()
        numbers = [word for word in words if word.replace('.', '', 1).isdigit()]
        if len(numbers) == 2:
            calculator_window = pyautogui.getWindowsWithTitle("Calculator")
            if calculator_window:
                calculator_window[0].activate()

            pyautogui.typewrite(numbers[0], interval=0)
            pyautogui.press(operator_key)
            pyautogui.typewrite(numbers[1], interval=0)
            pyautogui.press("enter")
            speak("Operation performed. Check the calculator for the result.")
        else:
            speak("I need exactly two numbers to perform this operation.")
    except Exception as e:
        print(f"Error in calculator operation: {e}")
        speak("Sorry, I couldn't perform the operation.")

# Function to open calculator and perform basic operations
def open_calculator():
    speak("Opening Calculator. Please tell me the operation you want to perform.")
    os.startfile("calc.exe")

    while True:
        command = get_command()
        if command:
            if "add" in command or "plus" in command:
                perform_calculator_operation(command, "+")
            elif "subtract" in command or "minus" in command:
                perform_calculator_operation(command, "-")
            elif "multiply" in command or "times" in command:
                perform_calculator_operation(command, "*")
            elif "divide" in command or "by" in command:
                perform_calculator_operation(command, "/")
            elif "close calculator" in command or "exit calculator" in command:
                speak("Closing calculator.")
            # Attempt to close calc.exe more robustly
                try:
                    # First try using taskkill
                    result = subprocess.run(['tasklist'], capture_output=True, text=True)
                    if 'calc.exe' in result.stdout:
                        subprocess.run(['taskkill', '/f', '/im', 'calc.exe'])
                        speak("Calculator closed.")
                    else:
                        speak("Calculator is not open.")
                except Exception as e:
                    print(f"Error: {e}")
                    speak("An error occurred while trying to close the calculator.")
                break
        else:
            speak("I didn't understand the operation. Please try again.")


def slide():
    slide_path = r"C:\Program Files\Microsoft Office\root\Office16\POWERPNT.EXE"
    os.startfile(slide_path)

# Function to process user commands
def process_command(command):
    if "what is the time" in command or "what's the time" in command:
        current_time = datetime.datetime.now().strftime("%H:%M")
        speak(f"The time is {current_time}")

    elif "can you search for" in command and "in youtube" in command:
        search = command.replace("can you search for", "").replace("in youtube", "").strip()
        search_in_yt(search)

    elif "play" in command or "can you play" in command:
        song = command.replace("play", "").strip()
        play_youtube_music(song)

    elif "open calculator" in command:
        open_calculator()

    elif "pause" in command:
        pause_video()

    elif "open notepad" in command:
        type_in_notepad()

    elif "open slides" in command:
        slide()

    elif "stop" in command or "exit" in command:
        speak("Goodbye!")
        return False

    else:
        speak("I'm sorry, I can't perform that task yet.")

    return True

# Main function to handle the assistant's workflow
def main():
    try:
        speak("Hello, I am Nova! Say 'Hey Nova!' to activate me.")
        while True:
            command = listen()
            if command:
                if not process_command(command):
                    break
    except KeyboardInterrupt:
        speak("Goodbye! Program interrupted.")
    except Exception as e:
        print(f"Error in main loop: {e}")
        speak("An unexpected error occurred. Exiting now.")
    finally:
        speak("Shutting down Nova.")

# Entry point for the program
if __name__ == "__main__":
    main()
