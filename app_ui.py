import os
import sys
import threading
import cv2
import pyttsx3
import speech_recognition as sr
from PyQt5.QtCore import Qt, QTimer, QEventLoop
from PyQt5.QtGui import QPixmap, QImage
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QLabel, QStackedWidget, QSlider, QListWidget,
    QHBoxLayout, QTextEdit
)

import main  # Import your backend functions here


class VoiceAssistantUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.engine = pyttsx3.init()
        self.capture = cv2.VideoCapture(0)
        self.profile_name = ""  # Stores the current profile name

        # Start the profile selection process
        threading.Thread(target=self.askForProfile).start()
        threading.Thread(target=main.track_head_for_mouse).start()



    def initUI(self):
        # Setup main window properties
        self.setWindowTitle("Voice Assistant")
        self.setGeometry(50, 50, 600, 900)  # Increased size for better readability

        # Main container for stacked views (Profile selection and Main UI)
        self.stack = QStackedWidget(self)
        self.setCentralWidget(self.stack)

        # Profile selection view
        self.profileSelectionView = QWidget()
        profile_layout = QVBoxLayout(self.profileSelectionView)
        self.profile_label = QLabel("Please say your profile name", self)
        self.profile_label.setAlignment(Qt.AlignCenter)
        profile_layout.addWidget(self.profile_label)
        self.stack.addWidget(self.profileSelectionView)  # Index 0

        # Main UI view
        self.mainUIView = QWidget()
        main_layout = QVBoxLayout(self.mainUIView)

        # Profile and Mouse Sensitivity Section
        profile_info_layout = QHBoxLayout()
        self.profile_name_label = QLabel("Profile: ", self)
        self.mouse_sensitivity_label = QLabel("Mouse Sensitivity:", self)

        # Mouse sensitivity slider
        self.mouse_sensitivity_slider = QSlider(Qt.Horizontal)
        self.mouse_sensitivity_slider.setMinimum(1)
        self.mouse_sensitivity_slider.setMaximum(10)
        self.mouse_sensitivity_slider.setValue(5)  # Default sensitivity
        profile_info_layout.addWidget(self.profile_name_label)
        profile_info_layout.addWidget(self.mouse_sensitivity_label)
        profile_info_layout.addWidget(self.mouse_sensitivity_slider)
        main_layout.addLayout(profile_info_layout)

        # Microphone Control Icon (always listening)
        self.mic_label = QLabel("🎤 Listening...", self)
        self.mic_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.mic_label)

        # Webcam Display
        self.webcam_label = QLabel("Webcam Feed", self)
        self.webcam_label.setFixedSize(580, 300)  # Increased size for better visibility
        main_layout.addWidget(self.webcam_label)
        

        # Command Display
        self.command_label = QLabel("Status: Waiting for command...", self)
        self.command_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.command_label)

        # Command History List
        self.command_list = QListWidget()
        main_layout.addWidget(QLabel("Recent Commands:", self))
        main_layout.addWidget(self.command_list)

        # Commands Section: Available Commands
        self.commands_text = QTextEdit()
        self.commands_text.setReadOnly(True)
        self.commands_text.setPlainText(
            "Available Commands:\n"
            "- Create Word Document, Save Word Document, Close Word,  \n"
            "- Create Excel Sheet, Save Excel Sheet, Close Excel,  \n"
            "- Create PowerPoint Presentation, Save PowerPoint Presentation, Close PowerPoint, \n"
            "- Open youtube,Play/Pause YouTube, YouTube Volume Up/Down, Next YouTube Video, Mute/Unmute YouTube\n"
            "- Open spotify,Play/Pause Spotify, Next/Previous Spotify Track, Spotify Volume Up/Down, Mute/Unmute Spotify\n"
            "- Open File Explorer, Create New Folder, Open Folder, Open File\n"
            "- Switch to Next Tab, Switch to Previous Tab\n"
            "- enter to click\n"
            "- Start Voice Typing, Stop Voice Typing\n"

        )
        main_layout.addWidget(QLabel("Available Commands:", self))
        main_layout.addWidget(self.commands_text)

        # Add main UI to stack
        self.stack.addWidget(self.mainUIView)  # Index 1

        # Timer for webcam feed update
        self.timer = QTimer()
        self.timer.timeout.connect(self.updateWebcamFeed)
        self.timer.start(30)  # Update every 30ms

    def continuous_head_tracking(self):
        """Continuously tracks head movement for mouse control."""
        main.track_head_for_mouse()



    def askForProfile(self):
        """Ask for a profile when the program starts."""
        loop = QEventLoop()
        self.profile_label.setText("Please say your profile name.")
        self.speak("Please say your profile name.")

        recognizer = sr.Recognizer()
        with sr.Microphone() as source:
            audio = recognizer.listen(source)
            try:
                profile = recognizer.recognize_google(audio)
                self.profile_name = profile
                self.profile_name_label.setText(f"Profile: {profile}")
                self.speak(f"Welcome, {profile}")

                # Switch to main UI view after profile is confirmed
                self.switchToMainUI()

                # Start continuous listening in main UI
                threading.Thread(target=self.listenContinuously).start()
            except sr.UnknownValueError:
                self.profile_label.setText("Could not understand profile name")
                self.speak("Sorry, I didn't catch that.")
            except Exception as e:
                self.profile_label.setText(f"Error: {e}")

    def switchToMainUI(self):
        """Switch to the main UI after profile is set."""
        self.stack.setCurrentIndex(1)  # Switch to the main UI view

    def listenContinuously(self):
        """Continuously listens for voice commands."""
        recognizer = sr.Recognizer()
        while True:
            with sr.Microphone() as source:
                self.mic_label.setText("🎤 Listening...")
                audio = recognizer.listen(source)
                try:
                    command = recognizer.recognize_google(audio).lower()
                    self.command_label.setText(f"Command: {command}")
                    self.command_list.addItem(command)  # Add command to the list
                    self.speak(f"Executing {command}")
                    self.executeCommand(command)
                except sr.UnknownValueError:
                    self.command_label.setText("Could not understand audio")
                    self.speak("Sorry, I didn't catch that.")
                except Exception as e:
                    self.command_label.setText(f"Error: {e}")

    def open_application(app_name, command):
        """Open an application based on the app_name and command provided."""
        try:
            os.system(command)
            print(f"{app_name} opened successfully.")
        except Exception as e:
            print(f"Failed to open {app_name}: {e}")
    def executeCommand(self, command):
        # Profile Management
        if "create profile" in command:
            user_name = main.get_voice_input("Please say the profile name.")
            main.create_profile(user_name)
        elif "switch profile" in command:
            user_name = main.get_voice_input("Please say the profile name.")
            main.switch_profile(user_name)
        elif "load profiles" in command:
            main.load_profiles()
        elif "save profiles" in command:
            main.save_profiles()

        # Microsoft Word Commands
        elif "create word document" in command:
            main.word_create_document()
        elif "save word document" in command:
            main.word_save_document()
        elif "close word" in command:
            main.word_close()
        elif "open word document" in command:
            main.word_open_document()


        # Microsoft Excel Commands
        elif "create excel sheet" in command:
            main.excel_create_sheet()
        elif "save excel sheet" in command:
            main.excel_save_sheet()
        elif "close excel" in command:
            main.excel_close()
        elif "open excel sheet" in command:
            main.excel_open_sheet()


        # Microsoft PowerPoint Commands
        elif "create powerpoint presentation" in command:
            main.powerpoint_create_presentation()
        elif "save powerpoint presentation" in command:
            main.powerpoint_save_presentation()
        elif "close powerpoint" in command:
            main.powerpoint_close()
        elif "open powerpoint presentation" in command:
            main.powerpoint_open_presentation()


        # YouTube Commands
        elif "open youtube" in command:
            main.open_youtube()
        elif "play youtube" in command or "pause youtube" in command:
            main.youtube_play_pause()
        elif "youtube volume up" in command:
            main.youtube_volume_up()
        elif "youtube volume down" in command:
            main.youtube_volume_down()
        elif "next youtube video" in command:
            main.youtube_next_video()
        elif "mute youtube" in command or "unmute youtube" in command:
            main.youtube_mute_unmute()

            # Spotify commands
        elif "spotify" in command:
         if "open" in command:
                main.open_spotify()
         elif "play" in command or "pause" in command:
                main.spotify_play_pause()
         elif "next" in command:
                main.spotify_next_track()
         elif "previous" in command:
                main.spotify_previous_track()
         elif "volume up" in command:
                main.spotify_volume_up()
         elif "volume down" in command:
                main.spotify_volume_down()
         elif "mute" in command or "unmute" in command:
                main.spotify_mute_unmute()

        # File Explorer Commands
        elif "open file explorer" in command:
            main.open_file_explorer()
        elif "create new folder" in command:
            main.create_new_folder()
        elif "open folder" in command:
            main.open_folder()
        elif "open file" in command:
            main.open_file()

        # System Navigation Commands
        elif "switch to next tab" in command:
            main.next_tab()
        elif "start voice typing" in command:
            main.start_typing()
        elif "stop voice typing" in command:
            main.stop_typing()
        elif "switch to previous tab" in command:
            main.previous_tab()
        elif "minimize window" in command:
            main.minimise_window()
        elif "maximise window" in command:
            main.maximise_window()
        elif "enter" in command:
            main.enter()

        else:
            self.command_label.setText("Command not recognized")
            self.command_list.addItem("Command not recognized")


    def speak(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

    def updateWebcamFeed(self):
        """Update the webcam feed on the UI."""
        ret, frame = self.capture.read()
        if ret:
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            h, w, ch = frame.shape
            bytes_per_line = ch * w
            convert_to_qt_format = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)
            self.webcam_label.setPixmap(QPixmap.fromImage(convert_to_qt_format))

    def closeEvent(self, event):
        """Release resources when the UI is closed."""
        self.timer.stop()
        self.capture.release()


# Run the application
app = QApplication(sys.argv)
window = VoiceAssistantUI()
window.show()
sys.exit(app.exec_())
