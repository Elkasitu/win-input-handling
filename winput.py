# This module handles both KB and Mouse input for Windows applications
#
# Most of this code is a combination of code found in the following links:
# http://stackoverflow.com/questions/2964051/python-win32-simulate-click
# http://stackoverflow.com/questions/14489013/simulate-python-keypresses-for-controlling-a-game
#
# DirectInput Keyboard Scan Codes:
# http://www.gamespp.com/directx/directInputKeyboardScanCodes.html
#
# I merely put the code together, cleaned it up a bit and added some functionality
# in order to get a working back-end for any future applications of mine that
# need Windows input emulation.

"""winput comes with the following high-level functions:

mouse_lclick() -- calls left mouse click
mouse_lhold() -- presses and holds left mouse button
mouse_lrelease() -- releases left mouse button

mouse_rclick() -- calls right mouse click
mouse_rhold() -- presses and holds right mouse button
mouse_rrelease() -- releases right mouse button

mouse_mclick() -- calls middle mouse click
mouse_mhold() -- presses and holds middle mouse button
mouse_mrelease() -- releases middle mouse button

move(x,y) -- moves mouse to x/y coordinates (in pixels)
getpos() -- returns mouse x/y coordinates (in pixels)
slide(x,y) -- slides mouse to x/y coodinates (in pixels)
              also supports optional speed='slow', speed='fast'

key_down(code) -- presses the key associated with the passed-in keycode
                  and holds it until key_up with the same code is called

key_up(code) -- releases the key associated with the passed-in keycode
"""

__all__ = ['mouse_lclick', 'mouse_lhold', 'mouse_lrelease', 'mouse_rclick',
           'mouse_rhold', 'mouse_rrelease', 'mouse_mclick',
           'mouse_mhold', 'mouse_mrelease', 'move', 'slide',
           'getpos', 'key_down', 'key_up']

__version__ = "0.1"

from ctypes import *
from ctypes.wintypes import *
from time import sleep
import win32ui


# Constants
# Keyboard constants
KEY_W = 0x11
KEY_A = 0x1E
KEY_S = 0x1F
KEY_D = 0x20

# Mouse constants
MIDDLE_DOWN = 0x00000020
MIDDLE_UP   = 0x00000040
MOVE       = 0x00000001
ABSOLUTE   = 0x00008000
RIGHT_DOWN  = 0x00000008
RIGHT_UP    = 0x00000010


# C struct redefinitions 
PUL = POINTER(c_ulong)

class KeyBdInput(Structure):
    _fields_ = [("wVk", c_ushort),
                ("wScan", c_ushort),
                ("dwFlags", c_ulong),
                ("time", c_ulong),
                ("dwExtraInfo", PUL)]

class HardwareInput(Structure):
    _fields_ = [("uMsg", c_ulong),
                ("wParamL", c_short),
                ("wParamH", c_ushort)]

class MouseInput(Structure):
    _fields_ = [("dx", c_long),
                ("dy", c_long),
                ("mouseData", c_ulong),
                ("dwFlags", c_ulong),
                ("time", c_ulong),
                ("dwExtraInfo", PUL)]

class Input_I(Union):
    _fields_ = [("ki", KeyBdInput),
                 ("mi", MouseInput),
                 ("hi", HardwareInput)]

class Input(Structure):
    _fields_ = [("type", c_ulong),
                ("ii", Input_I)]

class Point(Structure):
    _fields_ = [("x", c_ulong),
             ("y", c_ulong)]


# Set-up

SendInput = windll.user32.SendInput

FInputs = Input * 2
extra = c_ulong(0)

click = Input_I()
click.mi = MouseInput(0, 0, 0, 2, 0, pointer(extra))
release = Input_I()
release.mi = MouseInput(0, 0, 0, 4, 0, pointer(extra))

x = FInputs( (0, click), (0, release) )
#SendInput(2, pointer(x), sizeof(x[0])) CLICK & RELEASE

x2 = FInputs( (0, click) )
#SendInput(2, pointer(x2), sizeof(x2[0])) CLICK & HOLD

x3 = FInputs( (0, release) )
#SendInput(2, pointer(x3), sizeof(x3[0])) RELEASE HOLD


# Functions
# Mouse functions

def move(x, y):
    windll.user32.SetCursorPos(x,y)

def getpos():
    global pt
    pt = Point()
    windll.user32.GetCursorPos(byref(pt))
    return pt.x, pt.y

def slide(a, b, speed=0):
    while True:
        if speed == 'slow':
            sleep(0.005)
            Tspeed = 2
        if speed == 'fast':
            sleep(0.001)
            Tspeed = 5
        if speed == 0:
            sleep(0.001)
            Tspeed = 3

        x = getpos()[0]
        y = getpos()[1]
        if abs(x-a) < 5:
            if abs(y-b) < 5:
                break

        if a < x:
            x -= Tspeed
        if a > x:
            x += Tspeed
        if b < y:
            y -= Tspeed
        if b > y:
            y += Tspeed
        move(x,y)

def mouse_lclick():
    SendInput(2, pointer(x), sizeof(x[0]))

def mouse_lhold():
    SendInput(2, pointer(x2), sizeof(x2[0]))

def mouse_lrelease():
    SendInput(2, pointer(x3), sizeof(x3[0]))

def mouse_rclick():
    windll.user32.mouse_event(RIGHT_DOWN, 0, 0, 0, 0)
    windll.user32.mouse_event(RIGHT_UP, 0, 0, 0, 0)

def mouse_rhold():
    windll.user32.mouse_event(RIGHT_DOWN, 0, 0, 0, 0)

def mouse_rrelease():
    windll.user32.mouse_event(RIGHT_UP, 0, 0, 0, 0)

def mouse_mclick():
    windll.user32.mouse_event(MIDDLE_DOWN, 0, 0, 0, 0)
    windll.user32.mouse_event(MIDDLE_UP, 0, 0, 0, 0)

def mouse_mhold():
    windll.user32.mouse_event(MIDDLE_DOWN, 0, 0, 0, 0)

def mouse_mrelease():
    windll.user32.mouse_event(MIDDLE_UP, 0, 0, 0, 0)

# Keyboard functions

def key_down(code):
    extra = c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, code, 0x0008, 0, pointer(extra) )
    x = Input( c_ulong(1), ii_ )
    SendInput(1, pointer(x), sizeof(x))

def key_up(hexKeyCode):
    extra = c_ulong(0)
    ii_ = Input_I()
    ii_.ki = KeyBdInput( 0, code, 0x0008 | 0x0002, 0, pointer(extra) )
    x = Input( c_ulong(1), ii_ )
    SendInput(1, pointer(x), sizeof(x))
