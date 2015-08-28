from __future__ import print_function
import json
import urlparse
import time
from datetime import datetime
from eyex.api import EyeXInterface, SampleGaze, SampleFixation
from pywinauto import Application
import Queue
import webbrowser
from collections import namedtuple
import pywinauto
from pywinauto import Application
import os.path
import win32api
import win32net
import signal, sys
from SimpleWebSocketServer import WebSocket, SimpleWebSocketServer
import easygui as eg
import threading
from threading import Timer, Thread

PageTurn = namedtuple('PageTurn', ['slug', 'page', 'choice'])
Picture = namedtuple('Picture', ['pt', 'pb', 'pl', 'pr'])
Text = namedtuple('Text', ['tt','tb','tl','tr'])
path = os.getcwd()
if path[-14:-1] == '\\program file':
    path = path[0:len(path)-14]
eyetrack_on = False # Determines if EyeTracker is recording SampleGaze and SampleFixation
data_saved = False # Determines if you've reached End of Book and have saved all content
t_length = 120.0 # How long the program waits until it does a timeout save on a page
t0 = 0 # Time when video started
ttime = 0 # Total length of the video

alldata = Queue.Queue() # Holds SampleGaze and SampleFixation data taken from Tobii EyeGo
funcQ = Queue.Queue() # Handles 'events' from the three threads marked in Main

# FUNCTION DEFINITION(5) ######################################################################################

## 1 ##
def OpenPrograms():
    """Starts Tobii Gaze Viewer and Internet Explorer with Tarheel Reader."""
    try: # Connect/start Tobii Gaze Viewer regardless if its already running or not
        app = Application().connect_(path="C:\\Program Files (x86)\\Tobii Dynavox\\Gaze Viewer\\Tobii.GazeViewer.Startup.exe")
    except pywinauto.application.ProcessNotFoundError:
        app = Application.start("C:\\Program Files (x86)\\Tobii Dynavox\\Gaze Viewer\\Tobii.GazeViewer.Startup.exe")
    tw = app.Window_().Wait('visible', timeout=60, retry_interval=0.1)
    tw.Minimize()
    webbrowser.open('http://gbserver3.cs.unc.edu/favorites/?voice=silent&pageColor=fff&textColor=000&fpage=1&favorites=94348,97375,94147,91140&eyetracker=1') # Start iexplorer
    n = 0
    while n == 0:
        n = 1
        try:
            iexplorer = Application().connect_(path='C:\Program Files\Internet Explorer\iexplore.exe')
        except pywinauto.application.ProcessNotFoundError:
            n = 0
    w_handle = pywinauto.findwindows.find_windows(title_re="Tar Heel")
    while not w_handle:
        time.sleep(0.1)
        w_handle = pywinauto.findwindows.find_windows(title_re="Tar Heel")
    w_handle = w_handle[0]
    window = iexplorer.window_(handle=w_handle)
    time.sleep(1)
    window.SetFocus()
    window.TypeKeys('{F11}')
    return [iexplorer, app, tw, window]

## 2 ##
def SaveData(datetime, timeout):
    """Empties contents of queue (SampleGaze and SampleFixation points, PageTurn, Picture and
	Text namedtuples) into a json file in the folder 'data'."""
    data = []
    suffix = ''
    while True:
        try:
            r = alldata.get(False)
        except Queue.Empty:
            break
        if isinstance(r, PageTurn):
            d = {
                'type': 'PageTurn',
                'slug': r.slug,
                'page': r.page
            }
            data.append(d)
        elif isinstance(r, SampleGaze):
            d = {
                'type': 'SampleGaze',
                'timestamp': r.timestamp,
                'x': r.x,
                'y': r.y
			}
            data.append(d)
        elif isinstance(r, SampleFixation):
            d = {
                'type': 'SampleFixation',
                'event_type': r.event_type,
                'timestamp': r.timestamp,
                'x': r.x,
                'y': r.y
			}
            data.append(d)
        elif isinstance(r, Picture):
            d = {
                'type': 'Picture',
                'pt': r.pt,
                'pb': r.pb,
                'pl': r.pl,
                'pr': r.pr
            }
            data.append(d)
        elif isinstance(r, Text):
            d = {
                'type': 'Text',
                'tt': r.tt,
                'tb': r.tb,
                'tl': r.tl,
                'tr': r.tr
            }
            data.append(d)
    if timeout:
        suffix = '_timeout'
    filename = path +"\\data\\eyegazedata_" + datetime.strftime('%Y-%m-%d-%H-%M-%S') + suffix + '.json'
    json.dump(data, file(filename, 'w'))

## 3 ##
def SaveVid(datetime, timeout):
    """Saves a video that includes a screen capture, heat map and gaze plot by interacting with
	Tobii Gaze Viewer.  Videos are stored as mp4 files in 'Videos' -> 'Gaze Viewer'.

    tw - the original GV window that allows you to start/stop recording.
    tw1 - window that allows you to edit/save the video you've recorded.
    tw2 - popup for naming video/file explorer."""
    global timer
    global addres_bar
    global ttime
    eg.msgbox(msg="The video of this book reading will now be saved. While saving, please refrain from moving the cursor.  Click the button below to continue.", title="TAR HEEL READER - START VIDEO SAVING", ok_button="Start")
    time.sleep(1)
    window.Minimize()
    time.sleep(0.5)
    w = app.Windows_()
    tw1 = tw
    while tw1 == tw:
        for i,d in enumerate(w):
            if d.IsVisible():
                tw1 = w[i]
        w = app.Windows_()
	tw1.Maximize()
    tw1.ClickInput(coords=(1700,65))
    time.sleep(1)
    w = app.Windows_()
    tw2 = tw1
    while tw2 == tw1:
        for i,d in enumerate(w):
            if d.IsVisible():
                tw2 = w[i]
                break
        w = app.Windows_()
    tw2.SetFocus()

    date = path + "\\videos\\" + 'eyex_' + datetime.strftime('%Y-%m-%d-%H-%M-%S')
    d = ""
    for char in date:
        d = d + "{" + char + "}"
    filename = "{BACK}" + d
    if timeout == True:
        filename = filename + "{_}{t}{i}{m}{e}{o}{u}{t}"
        date = date + "_timeout"
    time.sleep(0.1)
    tw2.TypeKeys(filename)
    tw2.TypeKeys("{ENTER}")
    time.sleep(1)
    tw1.SetFocus()
    filename = date + '.mp4'
    print('Before filename loop')
    while os.path.isfile(filename) == False: # Waits for video to exist before trying to exit
        time.sleep(0.1)
    time.sleep(5)
    tw1.ClickInput(coords=(1870, 65)) # Tries to exit, but even if video file exists GazeViewer might still not be done saving
    while tw1.IsVisible() and tw.IsVisible() == False:
        tw1.ClickInput(coords=(1870, 65)) # Exit Button
        time.sleep(1)
        if tw1.IsVisible() and tw.IsVisible() == False:
            tw1.ClickInput(coords=(1100, 570)) # Popup that appears if its not done saving
    tw.Minimize()
    window.Maximize()
    window.SetFocus()
    if timeout == True:
        eg.msgbox(msg="Saving complete upon timeout. To restart recording, start a new book.", title="TAR HEEL READER - SAVING COMPLETE TIMEOUT", ok_button="Read Another Book")
        window.TypeKeys('{F11}')
        time.sleep(0.1)
        webbrowser.open('http://gbserver3.cs.unc.edu/favorites/?voice=silent&pageColor=fff&textColor=000&fpage=1&favorites=94348,97375,94147,91140&eyetracker=1')
        time.sleep(1)
        window.TypeKeys('{F11}')
    else:
	    eg.msgbox(msg="Saving complete!", title="TAR HEEL READER - SAVING COMPLETE", ok_button="Continue")
    timer = threading.Timer(t_length, timeoutHandler)

## 4 ##
def timeoutHandler():
    funcQ.put([timeoutHandlerHelper])
## 5 ##
def timeoutHandlerHelper():
    """Function called when timer reaches t_length seconds"""
    global eyetrack_on
    tw.TypeKeys("{F7}")
    eyetrack_on = False
    data_saved = True
    date = datetime.now()
    SaveVid(date, True)
    SaveData(date, True)


# CLASS DEFINITION(1) ############################################################################

## 1 ##
class Logger(WebSocket):

    def handleMessage(self):
        """Function called to handle a "PageTurn" event."""

        query = json.loads(self.data)
        funcQ.put([self.handleMessageHelper, query])

    def handleMessageHelper(self, query):
        global eyetrack_on
        global data_saved
        global timer
        global t0
        global ttime
        print ('got', query)
        print ('choice is type ', query['choice'])
        print('')
        choice = query['choice']

        if int(query['page']) == 1: # Resets data_saved upon the start of a new book
            data_saved = False
            timer.start()
            t0 = time.time()
        if choice == True and data_saved == False:
            timer.cancel()
            ttime = time.time() - t0
            tw.TypeKeys("{F7}")
            eyetrack_on = False
            data_saved = True # Prevents rate pages still considered part of the book from causing trouble afterwards
            date = datetime.now()
            SaveVid(date, False)
            SaveData(date, False)
        elif choice == False and data_saved == False:
            pt = PageTurn(query['slug'], query['page'], query['choice'])
            pic = Picture(query['pt'], query['pb'], query['pl'], query['pr'])
            text = Text(query['tt'], query['tb'], query['tl'], query['tr'])
            alldata.put(pt)
            alldata.put(pic)
            alldata.put(text)
            if int(query['page']) != 1:
                timer.cancel()
                timer = Timer(t_length, timeoutHandler)
                timer.start()
            if eyetrack_on == False:
                tw.SetFocus()
                tw.TypeKeys("{F7}")

            eyetrack_on = True


    def handleConnected(self):
        """Function called when Web Socket connects to Tar Heel Reader"""
        print('connected')
        pass

    def handleClose(self):
        """Function called when Web Socket becomes disconnected from Tar Heel Reader"""
        print('disconnected')
        pass

## WebSocket Server
class Server(Thread):
    def run(self):
        """Starts the thread for the Web Socket Server"""
        host = ''
        port = 8008
        self.server = SimpleWebSocketServer(host, port, Logger)

        print ('serving')
        self.server.serveforever()

    def close_sig_handler(self, signal, frame):
        """Function called upon Ctrl+C that kills the program"""
        print ("closing")
        self.server.close()
        #app.kill_()
        #iexplorer.kill_()
        #sys.exit()


## MAIN ##################################################################

answer = eg.buttonbox(msg="Ready to Start?", choices=["Yes","No"])

if answer == "Yes":
    iexplorer, app, tw, window = OpenPrograms()
    def handle_data(data):
        """Function called to handle EyeX SampleGaze and SampleFixation events"""
        global eyetrack_on
        if eyetrack_on:
            alldata.put(data)

    eye_api = EyeXInterface() # Thread one
    eye_api_f = EyeXInterface(fixation=True)
    eye_api.on_event += [handle_data]

    serverThread = Server() # Thread two
    serverThread.start()
    signal.signal(signal.SIGTERM, serverThread.close_sig_handler)

    timer = threading.Timer(t_length, timeoutHandler) # Thread three

    while True:
        func = funcQ.get()
        print ('calling', func)
        func[0](*func[1:])