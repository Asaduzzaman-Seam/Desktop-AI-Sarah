# Graphics Files
import sys
from PyQt5.QtCore import QSize
from PyQt5.QtGui import QMovie, QPainter, QPixmap
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QWidget, QLabel
# System Files
import speech_recognition as sr
import pyttsx3
import wheel
import pyaudio
import win32
import win32com.client
import win32api
import wikipedia
import os
import random
import webbrowser
import requests
# Import Files and Functions
from additional import func, life


# Graphical Interphase
class UIWindow(QWidget):
    def __init__(self, parent=None):
        super(UIWindow, self).__init__(parent)
        self.resize(QSize(600, 750))
        self.ToolsBTN = QPushButton('Start', self) 
        self.ToolsBTN.resize(100, 40)
        self.ToolsBTN.move(50, 50)
        self.ToolsBTN.clicked.connect(restart)

        self.CPS = QPushButton('Shut Down', self)
        self.CPS.resize(100, 40)
        self.CPS.move(450, 650)
        self.CPS.clicked.connect(close)

        self.label = QLabel("DAsarah2.1", self)
        self.label.resize(100, 40)
        self.label.move(50, 660)


class MainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(MainWindow, self).__init__(parent)
        self.setGeometry(50, 50, 600, 750)
        self.setFixedSize(600, 750)
        self.startUIWindow()

        self.movie = QMovie("img.gif")
        self.movie.frameChanged.connect(self.repaint)
        self.movie.start()

    def startUIWindow(self):
        self.Window = UIWindow(self)
        self.setWindowTitle("Dextop Assistant Sarah 2.1")
        self.show()

    def paintEvent(self, event):
        currentFrame = self.movie.currentPixmap()
        frameRect = currentFrame.rect()
        frameRect.moveCenter(self.rect().center())
        if frameRect.intersects(event.rect()):
            painter = QPainter(self)
            painter.drawPixmap(frameRect.left(), frameRect.top(), currentFrame)


# System File Start Here

# Program Start and Stop
def restart():
    speak('How may I help you , sir?')
    assistant(take_command())


def close():
    speak('I am going to shut down')
    speak('Have a nice day sir , see you , bye')
    sys.exit()

# Text to Voice
def speak(text):
    engine.say(text)
    engine.runAndWait()

# Take Command
def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print ("Listening...")
        speak('Your command sir...')
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print ("Recognizing...")
        query = r.recognize_google(audio)
        print ("TEXT : "+query)

    except (ValueError, Exception):
        print ("Say that again please")
        query = "None"
    return query

# Main Logic Function
def assistant(query):

    # Command Number
    num = 0

    # Assistant Life
    num = life.human(query, num)

    # Math Logic
    if 'mathematics' in query.lower():
        speak('Operation name...')
        op_name = take_command()
        num = func.math(op_name, num)

    # Logic
    if 'wikipedia' in query.lower():
        speak('Searching wikipedia ...')
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences=2)
        print (results)
        speak(results)
        num = num + 1

    if 'which is' in query.lower():
        if 'my favourite anime' in query.lower():
            speak('Your favourite anime is Naruto')
            num = num + 1
        elif 'my favourite color' in query.lower():
            speak('Your favourite color is White and Black')
            num = num + 1
        else:
            speak("I don't know Sir")
            num = num + 1

    if 'tell me' in query.lower():
        if 'about you' in query.lower():
            speak('I am just a programme of python. You give me my name sarah. I like to take your command')
            num = num + 1
        elif 'about myself' in query.lower():
            speak('Your name is Asaduzzaman Siam , You are truly a good person and kind to others...')
            num = num + 1
        elif 'about my mother' in query.lower():
            speak('She is the most important person in your life . I know that how much you love your mom')
            num = num + 1
        else:
            speak("I don't know Sir")
            num = num + 1

    if 'google' in query.lower():
        speak('Wait for Google results.....')
        query = query.replace('search on google', '')
        url = "http://www.google.com/?#q="
        chrome_path = 'C:/Program Files (x86)/Mozilla Firefox/firefox.exe %s'
        webbrowser.get(chrome_path).open(url + query)
        num = num + 1

    if 'open youtube' in query.lower():
        speak('Please wait for a moment sir')
        url = "youtube.com"
        firefox_path = 'C:/Program Files (x86)/Mozilla Firefox/firefox.exe %s'
        webbrowser.get(firefox_path).open(url)
        num = num + 1

    if 'open facebook' in query.lower():
        speak('Please wait for a moment sir')
        url = "www.facebook.com"
        firefox_path = 'C:/Program Files (x86)/Mozilla Firefox/firefox.exe %s'
        webbrowser.get(firefox_path).open(url)
        num = num + 1

    if 'close firefox' in query.lower():
        speak('Please wait for a moment sir')
        os.system("taskkill /f /im firefox.exe")
        speak('Done sir')
        num = num + 1

    if 'song' in query.lower():
        if 'hindi' in query.lower():
            speak('Please wait for a moment sir')
            songs_dir = "E:\\Video Song\\Hindi"
            songs = os.listdir(songs_dir)
            max_value = len(songs)
            os.startfile(os.path.join(songs_dir, songs[random.randint(0, max_value - 1)]))
            speak('Enjoy Sir...')
            num = num + 1
        elif 'english' in query.lower():
            speak('Please wait for a moment sir')
            songs_dir = "E:\\Video Song\\English"
            songs = os.listdir(songs_dir)
            max_value = len(songs)
            os.startfile(os.path.join(songs_dir, songs[random.randint(0, max_value - 1)]))
            speak('Enjoy Sir...')
            num = num + 1
        elif 'favourite' in query.lower():
            speak('Please wait for a moment sir')
            songs_dir = "E:\\Video Song\\English"
            songs = os.listdir(songs_dir)
            os.startfile(os.path.join(songs_dir, songs[35]))
            speak('Enjoy Sir...')
            num = num + 1
        elif 'audio' in query.lower():
            speak('Please wait for a moment sir')
            songs_dir = "E:\\Music\\2017"
            songs = os.listdir(songs_dir)
            max_value = len(songs)
            os.startfile(os.path.join(songs_dir, songs[random.randint(0, max_value - 1)]))
            speak('Enjoy Sir...')
            num = num + 1
        elif 'stop' in query.lower():
            speak('Please wait for a moment sir')
            os.system("taskkill /f /im vlc.exe")
            speak('Done sir')
            num = num + 1

    if 'what are you doing' in query.lower():
        speak('Just doing my thing')
        num = num + 1

    if 'joke' in query.lower():
        res = requests.get(
            'https://icanhazdadjoke.com/',
            headers={"Accept": "application/json"}
        )
        if res.status_code == requests.codes.ok:
            speak(str(res.json()['joke']))
            num = num + 1
        else:
            speak('oops!I ran out of jokes')
            num = num + 1
    if 'thank you' in query.lower():
        speak('Thank you too sir , I am very happy for you...')
        num = num + 1

    if 'pass' in query.lower():
        num = num + 1

    if 'shutdown' in query.lower():
        speak('I am going to shut down')
        speak('Have a nice day sir , see you , bye')
        sys.exit()

    if 'goodbye' in query.lower():
        speak('I am going to sleep Sir')
        print ('Sleep Mode Active')
        return 0

    if num == 0:
        num = life.auto_work(num)

    if num == 0:
        speak('I do not understand sir')

    print ('Stage : loop')
    talk_to_me(0)


# Active State Counter
def talk_to_me(counter):
    print (counter)
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)

    try:
        r.pause_threshold = 1
        query = r.recognize_google(audio)
        print ("Active : "+query)
        if 'goodbye' in query.lower():
            speak('I am going to sleep Sir')
            print ('Sleep Mode Active')
            return 0
        else:
            call = func.inner_func(query)
            if call == 0:
                assistant(take_command())
            else:
                assistant('pass')

    except (ValueError, Exception):
        counter = counter + 1
        if counter > 5:
            speak('I am going to sleep Sir')
            print ('Sleep Mode Active')
            return 0
        talk_to_me(counter)

# Main Function Loop
if __name__ == '__main__':

    # Text to Audio converter Module
    engine = pyttsx3.init()
    rate = engine.getProperty('rate')
    engine.setProperty('rate', 125)

    app = QApplication(sys.argv)
    w = MainWindow()
    sys.exit(app.exec_())