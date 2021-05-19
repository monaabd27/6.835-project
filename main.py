import Leap, sys, thread, time
import win32api, win32con
import speech_recognition as sr
# from pyautogui import press
import pyautogui
import keyboard
# from pynput.mouse import Controller
import PySimpleGUI27 as sg
from win32gui import GetForegroundWindow, SetForegroundWindow
import subprocess
#modified from the sample given in SDK

PREV_LOCATION = None
PREV_RUN_TIME = round(time.time() * 1000)
GRAB_COUNT = round(time.time() * 1000)
IS_GRABBING = False
INPUT_COUNT = round(time.time() * 1000)
INPUT_MODE = False
SMOOTH_FRAME = []
DRAGGING = False
USER_HAND = "right"

class SampleListener(Leap.Listener):
    finger_names = ['Thumb', 'Index', 'Middle', 'Ring', 'Pinky']
    bone_names = ['Metacarpal', 'Proximal', 'Intermediate', 'Distal']

    def on_init(self, controller):
        print "Initialized"

    def on_connect(self, controller):
        print "Connected"
        dom_hand()

    def on_disconnect(self, controller):
        # Note: not dispatched when running in a debugger.
        print "Disconnected"

    def on_exit(self, controller):
        print "Exited"

    def on_frame(self, controller):
        
        global PREV_LOCATION
        global SMOOTH_FRAME
        global DRAGGING
        global USER_HAND
        # Get the most recent frame and report some basic information
        frame = controller.frame()
        # for hand in frame.hands:
            # print(hand.stabilized_palm_position)
        if len(frame.hands) == 0:
            SMOOTH_FRAME = []
            PREV_LOCATION = None
        if len(frame.hands) == 1:
            hand = frame.hands[0]
            # move_cursor(hand.palm_position[0], hand.palm_position[1])
            # PREV_LOCATION = (hand.palm_position[0], (hand.palm_position[1]))
            move_cursor(hand.wrist_position[0], hand.wrist_position[1], hand)
            # PREV_LOCATION = (hand.wrist_position[0], (hand.wrist_position[1]))
            # print(hand.grab_strength, hand.palm_normal[0], hand.palm_normal[1], hand.palm_normal[2])
            if DRAGGING:
                click_and_drag(hand)
            else:
                click_handler(hand)
            if not IS_GRABBING:
                input_handler(hand)
        elif len(frame.hands) == 2:
            hand1 = frame.hands[0]
            hand2 = frame.hands[1]
            if USER_HAND == 'right':
                if hand1.is_right:
                    # move_cursor(hand1.palm_position[0], hand1.palm_position[1])
                    # PREV_LOCATION = (hand1.palm_position[0], (hand1.palm_position[1]))
                    move_cursor(hand1.wrist_position[0], hand1.wrist_position[1], hand1)
                    # PREV_LOCATION = (hand1.wrist_position[0], (hand1.wrist_position[1]))
                    if DRAGGING:
                        click_and_drag(hand2)
                    else:
                        click_handler(hand2)
                    if not IS_GRABBING:
                        input_handler(hand1)

                else:
                    # move_cursor(hand2.palm_position[0], hand2.palm_position[1])
                    # PREV_LOCATION = (hand2.palm_position[0], (hand2.palm_position[1]))
                    move_cursor(hand2.wrist_position[0], hand2.wrist_position[1], hand2)
                    # PREV_LOCATION = (hand2.wrist_position[0], (hand2.wrist_position[1]))
                    input_handler(hand2)
                    click_handler(hand1)
            else:
                if hand1.is_left:
                    # move_cursor(hand1.palm_position[0], hand1.palm_position[1])
                    # PREV_LOCATION = (hand1.palm_position[0], (hand1.palm_position[1]))
                    move_cursor(hand1.wrist_position[0], hand1.wrist_position[1], hand1)
                    # PREV_LOCATION = (hand1.wrist_position[0], (hand1.wrist_position[1]))
                    input_handler(hand1)
                    click_handler(hand2)
                else:
                    # move_cursor(hand2.palm_position[0], hand2.palm_position[1])
                    # PREV_LOCATION = (hand2.palm_position[0], (hand2.palm_position[1]))
                    move_cursor(hand2.wrist_position[0], hand2.wrist_position[1], hand2)
                    # PREV_LOCATION = (hand2.wrist_position[0], (hand2.wrist_position[1]))
                    input_handler(hand2)
                    click_handler(hand1)

        # print(PREV_LOCATION, SMOOTH_FRAME)

#code to navigate the desktop
def input_handler(hand):
    global INPUT_COUNT
    global INPUT_MODE
    normal = hand.palm_normal
    # if normal[0] < 0 and normal[1] > 0.8 and normal[2] > 0.2:
    if normal[1] > 0.7 and normal[2] > 0.1:

        if INPUT_MODE:
            wind = GetForegroundWindow()
            if round(time.time() * 1000) - INPUT_COUNT > 500:
                text = recognize_speech()
                SetForegroundWindow(wind)
                tl = text.lower()
                if text:
                    if tl == 'enter':
                        keyboard.send('enter')
                    if tl == 'backspace':
                        keyboard.send('backspace')
                    if tl == 'tab':
                        keyboard.send('tab')
                    elif tl == 'open file':
                        keyboard.send('ctrl+o')
                    elif tl == 'redo':
                        keyboard.send('ctrl+y')
                    elif tl == 'undo':
                        keyboard.send('ctrl+x')
                    elif tl == 'save file':
                        keyboard.send('ctrl+s')
                    # elif tl == 'navigate open windows':
                    #     keyboard.send('alt+tab')
                    elif tl == 'exit window':
                        keyboard.send('alt+F4')
                    elif tl == 'exit tab':
                        keyboard.send('ctrl+w')
                    elif tl == 'help':
                        help()
                    else:
                        # typewrite(text)
                        keyboard.write(text, delay=0.01)
                # else:
                #     keyboard.send('ctrl+windows+o')
        else:
            INPUT_MODE = True
            INPUT_COUNT = round(time.time() * 1000)
    else:
        INPUT_MODE= False

def click_handler(hand):
    # print(hand.pinch_strength)
        # if hand.pinch_strength > 0.95:
        #     print(hand.pinch_strength)
        #     click()
    # print(hand.grab_strength, "grab")
    # print(hand.pinch_strength, "pinch")
    # print(hand.grab_strength)
    if hand.grab_strength >= 0.96:# and hand.grab_strength >=0.85:
        global GRAB_COUNT
        global IS_GRABBING
        global DRAGGING
        if IS_GRABBING:
            if 4100 > round(time.time() * 1000) - GRAB_COUNT > 4000:
                print('double click')
                double_click()
            if 2100 > round(time.time() * 1000) - GRAB_COUNT > 2000:
                normal = hand.palm_normal
                if normal[1] > 0.7 and normal[2] > 0.1:
                    print('right click')
                    right_click()
                    IS_GRABBING = False
                    GRAB_COUNT = None
                    return
                else:
                    print('click')
                    click()
            if 6100 > round(time.time() * 1000) - GRAB_COUNT > 6000:
                print('click and drag')
                DRAGGING = True
                click()
                click_and_drag(hand)
            # if round(time.time() * 1000) - GRAB_COUNT > 2000:
            #     print('right click')
            #     right_click()
            #     GRAB_COUNT = None
            #     IS_GRABBING = False
            #     LAST_CLICK_TIME = round(time.time() * 1000)
        else:
            IS_GRABBING = True
            GRAB_COUNT = round(time.time() * 1000)
    else:
        IS_GRABBING = False
        GRAB_COUNT = None

def click_and_drag(hand):
    global DRAGGING
    global GRAB_COUNT
    global IS_GRABBING
    x, y = win32api.GetCursorPos()
    # pyautogui.mouseDown(x,y)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    print(hand.grab_strength)
    if hand.grab_strength < 0.55:
        # pyautogui.mouseUp(x,y)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
        DRAGGING = False
        GRAB_COUNT = None
        IS_GRABBING = False
        print('not dragging')

def right_click():
    x, y = win32api.GetCursorPos()
    pyautogui.rightClick(x,y)

def double_click():
    x, y = win32api.GetCursorPos()
    pyautogui.doubleClick(x,y)

def click():
    x, y = win32api.GetCursorPos()
    # pyautogui.mouseDown(x, y)
    # time.sleep(2)
    # pyautogui.mouseUp(x, y)
    pyautogui.click(x,y)
    # pyautogui.mouseUp(x,y)
    # pyautogui.mouseUp(1000,400)
    # mouse = Controller()
    # mouse.press(Button.left)
    # mouse.release(Button.left)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0, 0,0,0)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    # win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x, y,0,0)
def recognize_speech():
    # obtain audio from the microphone
    #code sampled from https://github.com/Uberi/speech_recognition/blob/master/examples/microphone_recognition.py
    r = sr.Recognizer()
    layout = [[sg.Image(filename=r'microphone.gif',
                enable_events=True,
                background_color='white',
                key='_IMAGE_',
                right_click_menu=['UNUSED', 'Exit'])],]
    window = sg.Window('My new window', no_titlebar=True, grab_anywhere=True, keep_on_top=True, background_color='white', alpha_channel=.7, margins=(0,0), location=(3250,1500)).Layout(layout)
    # window.read()
    try:
        with sr.Microphone() as source:
            
            print("Say something!")
            # time = round(time.time() * 1000)
            # while round(time.time() * 1000) - time < 100:
            #     event, values = window.Read()
            # sg.Window('My new window', no_titlebar=True, grab_anywhere=True, keep_on_top=True, background_color='white', alpha_channel=.7, margins=(0,0), location=(3250,1500)).Layout(layout)
            window.Read(timeout=26)
            print('check2')
            audio = r.listen(source, timeout= 5) #removed timeout
        
        text = r.recognize_google(audio)
        window.close()
        if text:
            print(text)
            return text
        # if text.lower() == 'enter':
        #     keyboard.send('enter')
    except:
        text = None
        window.close()

    print(text)
    # window.close()
    return text

def move_cursor(x, y, hand):
    global PREV_LOCATION
    global PREV_RUN_TIME
    global SMOOTH_FRAME
    if PREV_LOCATION == None:
        PREV_LOCATION = (hand.wrist_position[0], (hand.wrist_position[1]))
        return
    if round(time.time() * 1000) - PREV_RUN_TIME < 10:
        return
    if len(SMOOTH_FRAME) < 40:
        SMOOTH_FRAME.append((x,y))
    else:
        del SMOOTH_FRAME[0]
        # if correct_motion(x, y):
        #     return
        SMOOTH_FRAME.append((x,y))
        x, y= zip(*SMOOTH_FRAME)
        x = (sum(x)/len(x))
        y = (sum(y)/len(y))

    shift_x = x - PREV_LOCATION[0]
    shift_y = PREV_LOCATION[1] - y
    curr_x, curr_y = win32api.GetCursorPos()
    win32api.SetCursorPos((curr_x-int(shift_x), curr_y-int(shift_y)))
    # print((curr_x+shift_x*5, (curr_y+shift_y*10)))
    PREV_RUN_TIME = round(time.time() * 1000)
    PREV_LOCATION = (hand.wrist_position[0], (hand.wrist_position[1]))


# def predict_trajectory():
#     #calculate the trajectory of the hand based on previous movement but also the locaiton on the screen
#     #ie if we are 
#     #returns an (x,y) vector of the predicted trajectory
#     global SMOOTH_FRAME
#     x1, y1 = SMOOTH_FRAME[0]
#     x2, y2 = SMOOTH_FRAME[-1]
#     dx, dy = (x2-x1)/len(SMOOTH_FRAME), (y2-y1)/len(SMOOTH_FRAME)
#     return dx, dy

# add in the motion correction + prediction (research some parkinsons motion behavior)
# - correct motions by predicting the trajectory of the hand using the previous points, and if falling out of the smoothed trajectory predicted, then correct


# def correct_motion(x,y):
#     xc, yc = predict_trajectory()
#     dx, dy = abs(xc-x), abs(yc-y)
#     print(dx/x, dy/y)
#     if abs(dx/x) < 0.5 and abs(dy/y <0.5):
#         return True
#     return False

    #if the motion if way off the currect trajectory/happened too quickly, dont move the cursor

def main():
    # Create a sample listener and controller
    listener = SampleListener()
    controller = Leap.Controller()

    # Have the sample listener receive events from the controller
    controller.add_listener(listener)
    # Keep this process running until Enter is pressed
    print "Press Enter to quit..."
    try:
        sys.stdin.readline()
    except KeyboardInterrupt:
        pass
    finally:
        # Remove the sample listener when done
        controller.remove_listener(listener)

def dom_hand():
    global USER_HAND
    choices = ('right', 'left')

    layout = [  [sg.Text('Please choose a hand as your main cursor', font=('Helvetica', 12))],
                [sg.Drop(values=choices, auto_size_text=True)], 
                [sg.Text("To view the help document, click 'Help' below.")],
                [sg.Text("To get help later, face your palm up and say 'help'.")],
                [sg.Submit(), sg.Button('Help')]]

    window = sg.Window('Pick cursor hand', layout)

    while True:                  # the event loop
        event, values = window.read()
        if event == 'Help':
            help()
        elif event == 'Submit' or event == 'None':
            USER_HAND = values[0]
            print(USER_HAND) 
            window.close()
            break

def help():
    subprocess.call("help.pdf", shell=True)


if __name__ == "__main__":
    # dom_hand()
    main()
    # layout = [[sg.Image(filename=r'microphone.gif',
    #             enable_events=True,
    #             background_color='white',
    #             key='_IMAGE_',
    #             right_click_menu=['UNUSED', 'Exit'])],]
    # window = sg.Window('My new window', no_titlebar=True, grab_anywhere=True, keep_on_top=True, background_color='white', alpha_channel=.7, margins=(0,0), location=(3250,1500)).Layout(layout)
    # window.read()
    


   



