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

# NAMED TUPLES ###############################################################################################

PageTurn = namedtuple('PageTurn', ['slug', 'page', 'choice'])
Picture = namedtuple('Picture', ['pt', 'pb', 'pl', 'pr'])
Text = namedtuple('Text', ['tt','tb','tl','tr'])

# GLOBAL VARIABLES ###########################################################################################

path = os.getcwd()      # Return program start in path
bookshelf = open(path + '\\Bookshelf.txt', 'r').read() # Opens Bookshelf.txt
if path[-14:-1] == '\\program file': # path goes up a level from program files folder
    path = path[0:len(path)-14]

eyetrack_on = False     # True: Tobii EyeGo is recording data, False: Tobii EyeGo is not recording data
data_saved = False      # Determines if you've reached end of book and have saved all content (vdieo, data)
t_length = 120.0        # How long the program waits until it does a timeout save on a page

alldata = Queue.Queue() # Holds SampleGaze and SampleFixation data taken from Tobii EyeGo
funcQ = Queue.Queue()   # Handles 'events' from the three threads marked in Main

# F UNCTION DEFINITION(5) #####################################################################################

## 1 ##
def OpenPrograms():
    """Starts Tobii Dynavox Gaze Viewer and Internet Explorer with Tarheel Reader.
    
    app - the pywinauto Application associated with Tobii Dynavox Gaze Viewer
    tw - main window of app upon startup
    window - main window of IE upon startup"""

    try: # Connect/start Tobii Gaze Viewer regardless if its already running or not
        app = Application().connect_(path="C:\\Program Files (x86)\\Tobii Dynavox\\Gaze Viewer\\Tobii.GazeViewer.Startup.exe")
    except pywinauto.application.ProcessNotFoundError:
        app = Application.start("C:\\Program Files (x86)\\Tobii Dynavox\\Gaze Viewer\\Tobii.GazeViewer.Startup.exe")
    tw = app.Window_().Wait('visible', timeout=60, retry_interval=0.1)
    tw.Minimize()
    if bookshelf == 'A':
        webbrowser.open('http://test.tarheelreader.org/favorites/?collection=vap-study-bookshelf-a&eyetracker=1')
    else:
        webbrowser.open('http://test.tarheelreader.org/favorites/?collection=vap-study-bookshelf-b&eyetracker=1')    
    c = pywinauto.findwindows.find_windows(title_re="Tar Heel Reader|Favorites")
    while len(c) == 0:
        c = pywinauto.findwindows.find_windows(title_re="Tar Heel Reader|Favorites")
    window = pywinauto.controls.HwndWrapper.HwndWrapper(c[0])
    time.sleep(1)
    window.SetFocus()
    window.TypeKeys('{F11}')
    return [app, tw, window]

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
    if bookshelf == 'A':
        filename = path +"\\data\\Bookshelf A\\eyegazedata_" + datetime.strftime('%Y-%m-%d-%H-%M-%S') + suffix + '.json'
    else:
        filename = path +"\\data\\Bookshelf B\\eyegazedata_" + datetime.strftime('%Y-%m-%d-%H-%M-%S') + suffix + '.json'
    json.dump(data, file(filename, 'w'))

## 3 ##
def SaveVid(datetime, timeout):
    """Saves a video that includes a screen capture, heat map and gaze plot by interacting with
	Tobii Gaze Viewer.  Videos are stored as mp4 files in folder 'videos' in subfolder corresponding
    to the bookshelf read from.

    tw - the original GV window that allows you to start/stop recording.
    tw1 - window that allows you to edit/save the video you've recorded.
    tw2 - popup for naming video/file explorer."""

    global timer
    global addres_bar

    eg.msgbox(msg="The video of this book reading will now be saved. While saving, please refrain from moving the cursor.  Click the button below to continue.", title="TAR HEEL READER - START VIDEO SAVING", ok_button="Start")
    time.sleep(1)
    window.Minimize()
    time.sleep(0.5)
    w = app.Windows_()
    tw1 = tw
    while tw1 == tw: # Common way this program attempts to switch to a newly appeared window by waiting for it to be visible and enabled
        for i,d in enumerate(w):
            if d.IsVisible() and d.IsEnabled():
                tw1 = w[i]
        w = app.Windows_()
	tw1.Maximize()
    tw1.ClickInput(coords=(1700,65)) # Dependent on screen resolution
    time.sleep(1)
    w = app.Windows_()
    tw2 = tw1
    while tw2 == tw1:
        for i,d in enumerate(w):
            if d.IsVisible() and d.IsEnabled():
                tw2 = w[i]
                break
        w = app.Windows_()
    tw2.SetFocus()
    if bookshelf == 'A':
        date = path + "\\videos\\Bookshelf A\\" + 'eyex_' + datetime.strftime('%Y-%m-%d-%H-%M-%S')
    else:
        date = path + "\\videos\\Bookshelf B\\" + 'eyex_' + datetime.strftime('%Y-%m-%d-%H-%M-%S')
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
    tw1.ClickInput(coords=(1870, 65)) # Dependent on screen resolution; tries to exit, but even if video file exists GazeViewer might still not be done saving
    while tw1.IsVisible() and tw.IsVisible() == False:
        tw1.ClickInput(coords=(1870, 65)) # Dependent on screen resolution; Exit Button
        time.sleep(1)
        if tw1.IsVisible() and tw.IsVisible() == False:
            tw1.ClickInput(coords=(1100, 570)) # Depdendent on screen resolution; Popup that appears if its not done saving
    tw.Minimize()
    window.Maximize()
    window.SetFocus()
    if timeout == True:
        eg.msgbox(msg="Saving complete upon timeout. To restart recording, start a new book.", title="TAR HEEL READER - SAVING COMPLETE TIMEOUT", ok_button="Read Another Book")
        window.TypeKeys('{F11}')
        time.sleep(0.1)
        if bookshelf == 'A':
            webbrowser.open('http://test.tarheelreader.org/favorites/?collection=vap-study-bookshelf-a&eyetracker=1')
        else:
            webbrowser.open('http://test.tarheelreader.org/favorites/?collection=vap-study-bookshelf-b&eyetracker=1')
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
    """Function called when timer reaches t_length seconds. Force saves video and data."""
    global eyetrack_on

    tw.TypeKeys("{F7}")
    eyetrack_on = False
    data_saved = True
    date = datetime.now()
    SaveVid(date, True)
    SaveData(date, True)


# CLASS DEFINITION(2) ##########################################################################################

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

        # Optional print statements of query values from Tar Heel Reader
        #print ('got', query)
        #print ('choice is type ', query['choice'])
        #print('')

        choice = query['choice']

        if int(query['page']) == 1: # Resets data_saved upon the start of a new book, start timer
            data_saved = False
            timer.start()
        if choice == True and data_saved == False: # Executed on last page of a book
            timer.cancel()
            tw.TypeKeys("{F7}")
            eyetrack_on = False
            data_saved = True # Prevents rate pages still considered part of the book from causing trouble afterwards
            date = datetime.now()
            SaveVid(date, False) # See ## 2 ##
            SaveData(date, False) # See ## 3 ##
        elif choice == False and data_saved == False: # Executed on any page but the last page
            pt = PageTurn(query['slug'], query['page'], query['choice'])
            pic = Picture(query['pt'], query['pb'], query['pl'], query['pr'])
            text = Text(query['tt'], query['tb'], query['tl'], query['tr'])
            alldata.put(pt)
            alldata.put(pic)
            alldata.put(text)
            if int(query['page']) != 1: # Reset timer upon reaching new page
                timer.cancel()
                timer = Timer(t_length, timeoutHandler)
                timer.start()
            if eyetrack_on == False: # If recording video isn't on, turn it on
                tw.SetFocus()
                tw.TypeKeys("{F7}")
            eyetrack_on = True # Turn on data recording

    def handleConnected(self):
        """Function called when Web Socket connects to Tar Heel Reader"""
        print('connected')
        pass

    def handleClose(self):
        """Function called when Web Socket becomes disconnected from Tar Heel Reader"""
        print('disconnected')
        pass

## 2 ##
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
        #sys.exit()


## MAIN ###################################################################################

answer = eg.buttonbox(msg="Ready to Start?", choices=["Yes","No"])

if answer == "Yes":
    app, tw, window = OpenPrograms()
    
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
else:
    # Program ends and closes
    pass

