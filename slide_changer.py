import pyttsx3
import speech_recognition as sr
import win32com.client as win32
import tkinter as tk
from tkinter import filedialog
import pyautogui

# Initialize text-to-speech engine
engine = pyttsx3.init()


# Function to speak text
def speak(text):
    engine.say(text)
    engine.runAndWait()


# Function to listen for voice commands
def listen_for_command():
    recognizer = sr.Recognizer()
    microphone = sr.Microphone()

    with microphone as source:
        print("Listening for a command...")
        speak("I'm listening.")
        recognizer.adjust_for_ambient_noise(source, duration=1)  # Adjust noise for 1 second
        try:
            audio = recognizer.listen(source, timeout=10, phrase_time_limit=5)  # Longer listening times
            print("Processing your command...")
            command = recognizer.recognize_google(audio).lower()
            print(f"Command received: {command}")
            return command
        except sr.UnknownValueError:
            speak("Sorry, I didn't catch that. Please try again.")
        except sr.RequestError:
            speak("There was an error with the speech recognition service.")
        except Exception as e:
            print(f"Error: {e}")
            speak("An error occurred while processing your command.")
        return ""


# Function to control the PowerPoint presentation
def control_presentation(presentation, command):
    try:
        # Access the slides directly
        slides = presentation.Slides

        if "next slide" in command or "forward" in command:
            pyautogui.press('right')  # Simulate right arrow key press
            speak("Moving forward to the next slide.")

        elif "previous slide" in command or "back" in command:
            pyautogui.press('left')  # Simulate left arrow key press
            speak("Going back to the previous slide.")

        elif "start" in command or "presentation" in command:
            presentation.SlideShowSettings.Run()
            speak("Starting the slideshow.")

        elif "end presentation" in command or "stop" in command:
            presentation.SlideShowWindow.View.Exit()
            speak("Ending the slideshow.")

        elif "exit" in command or "quit" in command:
            speak("Exiting the program. Goodbye!")
            return "exit"

        else:
            speak("I didn't understand the command.")

    except Exception as e:
        print(f"Error controlling presentation: {e}")
        speak("An error occurred while controlling the presentation.")

    return "continue"


# Function to open a PowerPoint file
def open_presentation(file_path):
    try:
        # Initialize PowerPoint
        powerpoint = win32.Dispatch("PowerPoint.Application")
        presentation = powerpoint.Presentations.Open(file_path)
        powerpoint.Visible = True  # Make PowerPoint visible
        print(f"Presentation '{file_path}' opened successfully.")
        speak("Presentation opened. Ready to start.")
        return presentation
    except Exception as e:
        speak(f"Error opening the presentation: {str(e)}")
        print(f"Error: {str(e)}")
        return None


# Function to open a file picker dialog
def choose_file():
    root = tk.Tk()
    root.withdraw()  # Hide the main tkinter window
    file_path = filedialog.askopenfilename(
        title="Select a PowerPoint file",
        filetypes=[("PowerPoint files", "*.pptx")]
    )
    return file_path


# Main function
def main():
    # Allow the user to select a PowerPoint file
    file_path = choose_file()

    if not file_path:
        print("No file selected. Exiting.")
        speak("No file selected. Exiting.")
        return

    # Open the selected PowerPoint file
    presentation = open_presentation(file_path)
    if not presentation:
        return

    speak("You can now give voice commands to control the presentation.")

    # Start a loop to receive and execute commands
    while True:
        command = listen_for_command()
        if command:
            action = control_presentation(presentation, command)
            if action == "exit":
                break


if __name__ == "__main__":
    main()
