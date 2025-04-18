import ctypes
import json
import os
import threading
import time
from queue import Queue
import cv2
import mediapipe as mp
import pyautogui
import pyttsx3
import speech_recognition as sr

voice_typing_active = False
mic_lock = threading.Lock()


# Initialize speech engine, recognizer, and microphone
engine = pyttsx3.init()
recognizer = sr.Recognizer()
mic = sr.Microphone()

# Variables for profile handling
profiles = {}
current_user = None
profile_file = "profiles.json"
alpha = 1.5  # Sensitivity for mouse movement
voice_commands = [
    "create profile", "switch profile", "open word", "save document", "close word",
    "open excel", "create sheet", "close excel", "play youtube", "pause youtube",
    "next track spotify", "open file explorer", "new folder", "open folder", "click"
]
app_stack = []
command_queue = Queue()

# Initialize Mediapipe and OpenCV for face tracking
mp_face_mesh = mp.solutions.face_mesh
face_mesh = mp_face_mesh.FaceMesh(min_detection_confidence=0.5, min_tracking_confidence=0.5)
cap = cv2.VideoCapture(0)
screen_width, screen_height = pyautogui.size()
prev_x, prev_y = 0, 0
should_click = False
activate_keyboard = False
running = True
speech_lock = threading.Lock()  # Lock to manage concurrent speech synthesis


def load_profiles():
    global profiles
    try:
        with open(profile_file, "r") as file:
            profiles = json.load(file)
        print("Profiles loaded successfully:", profiles)
    except (FileNotFoundError, json.JSONDecodeError):
        profiles = {}
        print("No valid profile file found. Initializing new profiles.")
        save_profiles()


# Save profiles to the file
def save_profiles():
    with open(profile_file, "w") as file:
        json.dump(profiles, file, indent=4)
    print("Profiles saved successfully.")


# Create a new profile
def create_profile(user_name):
    global current_user
    if user_name not in profiles:
        profiles[user_name] = {"name": user_name, "preferences": {}, "sensitivity": alpha}
        save_profiles()
        print(f"Created profile: {user_name}")
    current_user = user_name


# Switch to an existing profile
def switch_profile(user_name):
    global current_user
    if user_name in profiles:
        current_user = user_name
        print(f"Switched to profile: {user_name}")
    else:
        print(f"Profile '{user_name}' not found. Creating a new one.")
        create_profile(user_name)


# Voice feedback function with thread safety
def say(text):
    with speech_lock:
        engine.say(text)
        engine.runAndWait()


# Function to open applications
def open_application(app_name, command):
    say(f"Opening {app_name}")
    os.system(command)


def close_last_application():
    if app_stack:
        last_app = app_stack.pop()  # Pop the last opened application
        say(f"Closing {last_app}.")
        pyautogui.hotkey("alt", "f4")  # Generic close command
    else:
        say("No applications to close.")


def add_command_to_queue(command):
    command_queue.put(command)


# Function to process the queue commands
def process_command_queue():
    while not command_queue.empty():
        command = command_queue.get()
        recognize_user_commands(command)
        command_queue.task_done()


def binary_search(arr, x):
    left, right = 0, len(arr) - 1
    while left <= right:
        mid = left + (right - left) // 2
        if arr[mid] == x:
            return mid
        elif arr[mid] < x:
            left = mid + 1
        else:
            right = mid - 1
    return -1


voice_commands.sort()
command_index = binary_search(voice_commands, "create profile")


### Word Commands ###
def word_create_document():
    open_application("Word", "start WINWORD")
    time.sleep(2)
    say("Creating a new document in Word.")


def word_save_document():
    pyautogui.hotkey('ctrl', 's')  # Save the document
    say("Saving the document.")


def word_close():
    pyautogui.hotkey('alt', 'f4')  # Close Wordf
    say("Closing Word.")


def word_open_document():
    say("Please say the name of the document to open.")
    file_name = get_voice_input()
    if file_name:
        pyautogui.hotkey('ctrl', 'o')  # Open file dialog
        pyautogui.write(file_name)
        pyautogui.press('enter')
        say(f"Opening document {file_name}.")


### Excel Commands ###
def excel_create_sheet():
    open_application("Excel", "start excel")
    time.sleep(2)
    say("Creating a new sheet in Excel.")


def excel_save_sheet():
    pyautogui.hotkey('ctrl', 's')  # Save the sheet
    say("Saving the sheet.")


def excel_close():
    pyautogui.hotkey('alt', 'f4')  # Close Excel
    say("Closing Excel.")


def excel_open_sheet():
    say("Please say the name of the spreadsheet to open.")
    file_name = get_voice_input()
    if file_name:
        pyautogui.hotkey('ctrl', 'o')  # Open file dialog
        pyautogui.write(file_name)
        pyautogui.press('enter')
        say(f"Opening spreadsheet {file_name}.")


    ### PowerPoint Commands ###
def powerpoint_create_presentation():
    open_application("PowerPoint", "start POWERPNT")
    time.sleep(2)
    say("Creating a new presentation in PowerPoint.")


def powerpoint_save_presentation():
    pyautogui.hotkey('ctrl', 's')  # Save the presentation
    say("Saving the presentation.")


def powerpoint_close():
    pyautogui.hotkey('alt', 'f4')  # Close PowerPoint
    say("Closing PowerPoint.")


def powerpoint_open_presentation():
    say("Please say the name of the presentation to open.")
    file_name = get_voice_input()
    if file_name:
        pyautogui.hotkey('ctrl', 'o')  # Open file dialog
        pyautogui.write(file_name)
        pyautogui.press('enter')
        say(f"Opening presentation {file_name}.")


def next_tab():
    pyautogui.hotkey('alt', 'tab')
    say("Switched to the next tab.")


def previous_tab():
    pyautogui.hotkey('alt', 'tab')
    say("Switched to the previous tab.")


def stop_typing():
    """Listens for 'Stop typing' command in a separate thread using its own mic instance."""
    local_mic = sr.Microphone()  # Create a separate microphone instance
    while True:
        try:
            with local_mic as source:
                recognizer.adjust_for_ambient_noise(source)
                print("Listening for stop command (voice typing)...")
                audio = recognizer.listen(source)
            command = recognizer.recognize_google(audio).lower()
            if "stop typing" in command:
                pyautogui.press('esc')
                say("Voice typing stopped.")
                break
        except Exception as e:
            print(f"Listening error in stop_typing thread: {e}")
            continue

def start_typing():
    """Starts Windows voice typing and listens for the stop command in a separate thread."""
    global voice_typing_active
    try:
        voice_typing_active = True  # Pause the main loop from using the mic
        pyautogui.click()
        time.sleep(2)
        pyautogui.hotkey('win', 'h')
        say("Voice typing started. Say 'Stop typing' to stop.")
        threading.Thread(target=stop_typing, daemon=True).start()
    except Exception as e:
        say(f"An error occurred during voice typing. Error: {e}")
        print(f"Error in start_typing: {e}")

def minimise_window():
    """Minimizes the active window."""
    pyautogui.hotkey("win", "down")
    say("Minimized the active window.")


def maximise_window():
    """Maximizes the active window."""
    pyautogui.hotkey("win", "up")
    say("Maximized the active window.")


def enter():
    """Attempts to click using multiple methods."""
    print("Attempting to click...")
    try:
        # Attempt pyautogui first
        pyautogui.click()
        print("Mouse click executed successfully with pyautogui.")
        say("Clicked at the current cursor position.")
    except Exception as e:
        print(f"pyautogui.click() failed: {e}. Trying ctypes...")
        try:
            # Fallback to ctypes
            ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # Left button down
            ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # Left button up
            print("Mouse click executed successfully with ctypes.")
            say("Clicked at the current cursor position.")
        except Exception as ctypes_error:
            print(f"ctypes click failed: {ctypes_error}")
            say(f"Failed to click. Error: {ctypes_error}")


### YouTube Commands ###
def open_youtube():
    try:
        youtube_url = "https://www.youtube.com"
        os.system(f"start {youtube_url}")
        say("Opening YouTube in your browser.")
    except Exception as e:
        say(f"Failed to open YouTube. Error: {e}")
        print(f"Error opening YouTube: {e}")


def youtube_play_pause():
    pyautogui.press('k')  # Play/pause video
    say("Toggling play/pause on YouTube.")


def youtube_volume_up():
    pyautogui.press('volumeup')
    say("Increasing YouTube volume.")


def youtube_volume_down():
    pyautogui.press('volumedown')
    say("Decreasing YouTube volume.")


def youtube_next_video():
    pyautogui.hotkey('shift', 'n')  # Shortcut for the next video
    say("Playing the next video on YouTube.")


def youtube_mute_unmute():
    pyautogui.press('m')  # Mute/unmute
    say("Toggling mute/unmute on YouTube.")


### Spotify Commands ###

def open_spotify():
    spotify_path = r"C:\Users\aarav\AppData\Roaming\Spotify\Spotify.exe"
    print(f"Attempting to open Spotify with path: {spotify_path}")
    try:
        os.startfile(spotify_path)
        say("Spotify is opening.")
    except Exception as e:
        print(f"Error in open_spotify(): {e}")
        say(f"Failed to open Spotify. Error: {e}")


def spotify_play_pause():
    pyautogui.press('playpause')  # Play/pause Spotify
    say("Toggling play/pause on Spotify.")


def spotify_next_track():
    pyautogui.press('nexttrack')  # Next track
    say("Skipping to the next track.")


def spotify_previous_track():
    pyautogui.press('prevtrack')  # Previous track
    say("Going back to the previous track.")


def spotify_volume_up():
    pyautogui.press('volumeup')  # Volume up
    say("Increasing Spotify volume.")


def spotify_volume_down():
    pyautogui.press('volumedown')  # Volume down
    say("Decreasing Spotify volume.")


def spotify_mute_unmute():
    pyautogui.press('volumemute')  # Mute/unmute
    say("Toggling mute/unmute on Spotify.")


### File Explorer Commands ###
def open_file_explorer():
    open_application("File Explorer", "explorer")
    say("Opening File Explorer.")


def create_new_folder():
    pyautogui.hotkey('ctrl', 'shift', 'n')  # Create a new folder
    say("Creating a new folder.")


def open_folder():
    say("Please say the name of the folder to open.")
    folder_name = get_voice_input()
    if folder_name:
        pyautogui.hotkey('ctrl', 'l')  # Focus on address bar
        pyautogui.write(folder_name)
        pyautogui.press('enter')
        say(f"Opening folder {folder_name}.")


def open_file():
    say("Please say the name of the file to open.")
    file_name = get_voice_input()
    if file_name:
        pyautogui.write(file_name)
        pyautogui.press('enter')
        say(f"Opening file {file_name}.")


# Voice command for taking file/folder name
def get_voice_input():
    with mic as source:
        recognizer.adjust_for_ambient_noise(source)
        audio = recognizer.listen(source)
        try:
            command = recognizer.recognize_google(audio).lower()
            return command
        except sr.UnknownValueError:
            say("I didn't catch that. Please try again.")
        except sr.RequestError as e:
            say("Sorry, I am having trouble connecting.")
        return None


# Function to track head movements and control the mouse
def track_head_for_mouse():
    global prev_x, prev_y, should_click, activate_keyboard, running
    while cap.isOpened() and running:
        success, frame = cap.read()

        # Validate the frame
        if not success or frame is None:
            print("Failed to capture frame from webcam.")
            continue

        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        results = face_mesh.process(frame_rgb)

        if results.multi_face_landmarks:
            for face_landmarks in results.multi_face_landmarks:
                # Get coordinates for the nose tip (use landmark 1 as an example)
                nose_tip = face_landmarks.landmark[1]
                x = int(nose_tip.x * screen_width)
                y = int(nose_tip.y * screen_height)

                # Smooth out movements by averaging previous position with new position
                new_x = prev_x + (x - prev_x) / alpha
                new_y = prev_y + (y - prev_y) / alpha

                # Move the mouse
                pyautogui.moveTo(new_x, new_y)

                # Store the new position
                prev_x, prev_y = new_x, new_y

                # Check for click activation
                if should_click:
                    pyautogui.click()
                    should_click = False

        # Break the loop if running is set to False
        if not running:
            break

    cap.release()


# Now include the function in the rest of your code

# Recognize and process voice commands
def recognize_user_commands():
    global should_click, activate_keyboard, running
    with mic as source:
        print("Adjusting for ambient noise...")
        recognizer.adjust_for_ambient_noise(source)
        print("Listening for command...")
        audio = recognizer.listen(source)

    try:
        command = recognizer.recognize_google(audio).lower()
        print(f"Recognized command: {command}")

        if "word" in command:
            if "create" in command:
                word_create_document()
            elif "save" in command:
                word_save_document()
            elif "close" in command:
                word_close()
            elif "open" in command:
                word_open_document()


        # Excel-related commands
        elif "excel" in command:
            if "create" in command:
                excel_create_sheet()
            elif "save" in command:
                excel_save_sheet()
            elif "close" in command:
                excel_close()
            elif "open" in command:
                excel_open_sheet()


        # PowerPoint-related commands
        elif "powerpoint" in command:
            if "create" in command:
                powerpoint_create_presentation()
            elif "save" in command:
                powerpoint_save_presentation()
            elif "close" in command:
                powerpoint_close()
            elif "open" in command:
                powerpoint_open_presentation()

        # YouTube commands
        elif "open youtube" in command:
            open_youtube()
        elif "youtube" in command:
            if "play" in command or "pause" in command:
                youtube_play_pause()
            elif "volume up" in command:
                youtube_volume_up()
            elif "volume down" in command:
                youtube_volume_down()
            elif "next" in command:
                youtube_next_video()
            elif "mute" in command or "unmute" in command:
                youtube_mute_unmute()

        # Spotify commands
        elif "spotify" in command:
            if "open" in command:
                open_spotify()
            elif "play" in command or "pause" in command:
                spotify_play_pause()
            elif "next" in command:
                spotify_next_track()
            elif "previous" in command:
                spotify_previous_track()
            elif "volume up" in command:
                spotify_volume_up()
            elif "volume down" in command:
                spotify_volume_down()
            elif "mute" in command or "unmute" in command:
                spotify_mute_unmute()

        # File Explorer commands
        elif "file explorer" in command:
            open_file_explorer()
        elif "new folder" in command:
            create_new_folder()
        elif "open folder" in command:
            open_folder()
        elif "open file" in command:
            open_file()

        #System Commands
        elif "next tab" in command:
            next_tab()
        elif "previous tab" in command:
            previous_tab()
        elif "maximise_window" in command:
            maximise_window()
        elif "minimise_window" in command:
            minimise_window()
        elif "enter" in command:
            enter()
        elif "start voice typing" in command:
            start_typing()
        elif "stop voice typing" in command:
            stop_typing()
        say(f"Executed command: {command}")

    except sr.UnknownValueError:
        say("I did not understand that. Could you please repeat?")
    except sr.RequestError as e:
        say("There was an issue with the recognition service.")


# Main function to start the assistant
def main():
    say("Hello, I am Jarvis. How can I assist you today?")
    load_profiles()
    user_name = None
    while user_name is None:
        say("Please say your name to create or switch profile.")
        with mic as source:
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)
            try:
                user_name = recognizer.recognize_google(audio).lower()
                switch_profile(user_name)
            except sr.UnknownValueError:
                say("I did not understand your name. Please try again.")
                user_name = None  # Keep looping until a name is recognized
            except sr.RequestError as e:
                say("There was an issue with the recognition service.")
                user_name = None

    # Start a thread for face tracking
    face_thread = threading.Thread(target=track_head_for_mouse)
    face_thread.daemon = True
    face_thread.start()

    # Voice command loop
    while running:
        recognize_user_commands()

    cap.release()


if __name__ == "__main__":
    main()
