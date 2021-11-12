import pyttsx3
import speech_recognition as sr
import os
import webbrowser
import datetime
import wikipedia
import sys
import cv2
from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtCore import QTimer, QTime, Qt
from PyQt5.QtGui import QMovie
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUiType
from jarvisUi import Ui_jarvisUi


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice',voices[0].id)

def speak(audio):
    engine.say(audio)
    print(audio)
    engine.runAndWait()

def createEngines():
    speak("creating engines. running setup. waiting for engine to start. please wait sir i am starting engines. done. everything is ready to run")

def wish():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour <12:
        speak("good morning sir, i am jarvis sir so please tell me how may i help you")
    elif hour>=12 and hour<18:
        speak("good afternoon sir, i am jarvis sir so please tell me how may i help you")
    else:
        speak("good evening sir, i am jarvis sir so please tell me how may i help you")

class MainThread(QThread):
    def __init__(self):
        super(MainThread,self).__init__()

    def run(self):
        self.JARVIS()
        while True:
            self.permission = self.STT()
            if 'wake up' in self.permission:
                self.JARVIS()
            elif 'goodbye' in self.permission:
                speak("good bye sir, have a good day")
                sys.exit()


    def STT(self):
        r = sr.Recognizer()
        with sr.Microphone() as source:
            print("Listning...........")
            audio = r.listen(source)
        try:
            print("Recognizing......")
            text = r.recognize_google(audio,language='en-in')
            print(">> ",text)
        except Exception:
            return "None"
        text = text.lower()
        return text

    def JARVIS(self):
        createEngines()
        wish()
        while True:
            self.query = self.STT()
            if 'wikipedia' in self.query:
                speak("searching wikipedia....please wait")
                self.query = self.query.replace("wikipedia", "")
                results = wikipedia.summary(self.query, sentences=2)
                speak("according to wikipedia")
                speak(results)
            elif 'open camera' in self.query:
                cap = cv2.VideoCapture(0)
                while True:
                    ret, img = cap.read()
                    cv2.imshow('webcam', img)
                    k = cv2.waitKey(50)
                    if k==27:
                        break;
                cap.release()
                cap.destroyAllWindows
            elif 'goodbye' in self.query:
                speak("good bye it was nice to meet you")
                sys.exit()
            elif 'open google' in self.query:
                speak("what should i search on google?")
                search = self.STT()
                webbrowser.open(f"{search}")
                speak("searching in google")
            elif 'open youtube' in self.query:
                speak("what should i search on youtube")
                search2 = self.STT()
                web = search2.replace(" ", "+")
                webbrowser.open(f"https://www.youtube.com/results?search_self.query={web}")
                speak("searching in youtube")
            elif 'open bing' in self.query:
                speak("what should i search on bing")
                search3 = self.STT()
                web2 = search3.replace(" ", "+")
                webbrowser.open(f"https://www.bing.com/videos/search?q={web2}&go=Search&qs=ds&form=QBVDMH")
            elif 'open whatsapp' in self.query:
                webbrowser.open("https://web.whatsapp.com")
                speak("opening whatsapp")
            elif 'open instagram' in self.query:
                webbrowser.open("https://www.instagram.com")
                speak("opening instagram")
            elif 'open facebook' in self.query:
                webbrowser.open("https://www.facebook.com")
                speak("opening facebook")
            elif 'open spotify' in self.query:
                webbrowser.open("https://open.spotify.com")
                speak("opening spotify")
            elif 'open notepad' in self.query:
                npath = "C:\\windows\\system32\\notepad.exe"
                os.startfile(npath)
                speak("opening notepad")
            elif 'close notepad' in self.query:
                os.system("taskkill /f /im notepad.exe")
            elif 'open file explorer' in self.query:
                fepath = "C:\\Windows\\explorer.exe"
                os.startfile(fepath)
                speak("opening file explorer")
            elif 'close file explorer' in self.query:
                os.system("taskkill /f /im explorer.exe")
            elif 'open command prompt' in self.query:
                os.system("start cmd")
                speak("opening command prompt")
            elif 'open powerpoint' in self.query:
                pppath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"
                os.startfile(pppath)
                speak("opening powerpoint")
            elif 'close powerpoint' in self.query:
                os.system("taskkill /f /im POWERPNT.EXE")
            elif 'open word' in self.query:
                wpath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"
                os.startfile(wpath)
                speak("opening word")
            elif 'close word' in self.query:
                os.system("taskkill /f /im WINWORD.EXE")
            elif 'open excel' in self.query:
                epath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE"
                os.startfile(epath)
                speak("opening excel")
            elif 'close excel' in self.query:
                os.system("taskkill /f /im EXCEL.EXE")
            elif 'open paint' in self.query:
                ppath = ("C:\\windows\\system32\\mspaint.exe")
                os.startfile(ppath)
                speak("opening paint")
            elif 'close paint' in self.query:
                os.system("taskkill /f /im mspaint.exe")
            elif 'open designer' in self.query:
                qtpath = "C:\\Users\\ishag\\AppData\\Roaming\\Python\\Python310\\site-packages\\QtDesigner\\designer.exe"
                os.startfile(qtpath)
                speak("opening qt designer")
            elif 'close designer' in self.query:
                os.system("taskkill /f /im designer.exe")
            elif 'open vs code' in self.query:
                vspath = "C:\\Users\\ishag\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
                os.startfile(vspath)
                speak("opening vs code")
            elif 'close vs code' in self.query:
                os.system("taskkill /f /im Code.exe")
            elif "shutdown the system" in self.query:
                os.system("shutdown /s /t 5")
                sys.exit()
            elif "restart the system" in self.query:
                os.system("shutdown /r /t 5")
                sys.exit()
            elif "sleep the system" in self.query:
                os.system("rundll32.exe powrprof.dll, SetSuspendState 0,1,0")
                sys.exit()
            elif 'play music' in self.query or 'play song' in self.query or 'play songs' in self.query:
                webbrowser.open("https://www.youtube.com/watch?v=7mFvyrNHZRY&list=PL8ZaXDTUkC_snB1F4QK6oIQCWGGY6cWHX&index=1")
                speak("playing music")
                break
            elif 'sleep' in self.query:
                speak("okay i am going to sleep, you can call me anytime")
                break
            else:
                speak("say that again please")

startExecution = MainThread()

class Main(QMainWindow):
        def __init__(self):
            super().__init__()
            self.ui = Ui_jarvisUi()
            self.ui.setupUi(self)
            self.ui.pushButton.clicked.connect(self.startTask)
            self.ui.pushButton_2.clicked.connect(self.close)

        def startTask(self):
            self.ui.movie = QtGui.QMovie("C:/Users/ishag/OneDrive/Desktop/Projects/Main-Jarvis-with-ui/gif/7LP8.gif")
            self.ui.label.setMovie(self.ui.movie)
            self.ui.movie.start()
            self.ui.movie = QtGui.QMovie("C:/Users/ishag/OneDrive/Desktop/Projects/Main-Jarvis-with-ui/gif/T8bahf.gif")
            self.ui.label_2.setMovie(self.ui.movie)
            self.ui.movie.start()
            timer = QTimer(self)
            timer.timeout.connect(self.showTime)
            timer.start(1000)
            startExecution.start()
        
        def showTime(self):
            current_time = QTime.currentTime()
            current_date = QDate.currentDate()
            label_time = current_time.toString('hh:mm:ss')
            label_date = current_date.toString(Qt.ISODate)
            self.ui.textBrowser.setText(label_date)
            self.ui.textBrowser_2.setText(label_time)

app = QApplication(sys.argv)
jarvis = Main()
jarvis.show()
exit(app.exec_())