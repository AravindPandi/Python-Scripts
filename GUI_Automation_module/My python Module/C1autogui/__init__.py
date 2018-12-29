import ctypes
from ctypes import *
import ctypes.wintypes
import collections
import sys
import time
import Kboard


MOUSEEVENTF_LEFTDOWN = 0x0002
MOUSEEVENTF_LEFTUP = 0x0004
MOUSEEVENTF_LEFTCLICK = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
MOUSEEVENTF_RIGHTDOWN = 0x0008
MOUSEEVENTF_RIGHTUP = 0x0010
MOUSEEVENTF_RIGHTCLICK = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
MOUSEEVENTF_MIDDLEDOWN = 0x0020
MOUSEEVENTF_MIDDLEUP = 0x0040
MOUSEEVENTF_MIDDLECLICK = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP

MOUSEEVENTF_WHEEL = 0x0800
MOUSEEVENTF_HWHEEL = 0x01000

FAILSAFE = True
PAUSE = 0.1
MINIMUM_DURATION = 0.1
MINIMUM_SLEEP = 0.05

KEY_NAMES = ['\t', '\n', '\r', ' ', '!', '"', '#', '$', '%', '&', "'", '(',
     ')', '*', '+', ',', '-', '.', '/', '0', '1', '2', '3', '4', '5', '6', '7',
     '8', '9', ':', ';', '<', '=', '>', '?', '@', '[', '\\', ']', '^', '_', '`',
     'a', 'b', 'c', 'd', 'e','f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o',
     'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '{', '|', '}', '~',
     'accept', 'add', 'alt', 'altleft', 'altright', 'apps', 'backspace',
     'browserback', 'browserfavorites', 'browserforward', 'browserhome',
     'browserrefresh', 'browsersearch', 'browserstop', 'capslock', 'clear',
     'convert', 'ctrl', 'ctrlleft', 'ctrlright', 'decimal', 'del', 'delete',
     'divide', 'down', 'end', 'enter', 'esc', 'escape', 'execute', 'f1', 'f10',
     'f11', 'f12', 'f13', 'f14', 'f15', 'f16', 'f17', 'f18', 'f19', 'f2', 'f20',
     'f21', 'f22', 'f23', 'f24', 'f3', 'f4', 'f5', 'f6', 'f7', 'f8', 'f9',
     'final', 'fn', 'hanguel', 'hangul', 'hanja', 'help', 'home', 'insert', 'junja',
     'kana', 'kanji', 'launchapp1', 'launchapp2', 'launchmail',
     'launchmediaselect', 'left', 'modechange', 'multiply', 'nexttrack',
     'nonconvert', 'num0', 'num1', 'num2', 'num3', 'num4', 'num5', 'num6',
     'num7', 'num8', 'num9', 'numlock', 'pagedown', 'pageup', 'pause', 'pgdn',
     'pgup', 'playpause', 'prevtrack', 'print', 'printscreen', 'prntscrn',
     'prtsc', 'prtscr', 'return', 'right', 'scrolllock', 'select', 'separator',
     'shift', 'shiftleft', 'shiftright', 'sleep', 'space', 'stop', 'subtract', 'tab',
     'up', 'volumedown', 'volumemute', 'volumeup', 'win', 'winleft', 'winright', 'yen',
     'command', 'option', 'optionleft', 'optionright']
KEYBOARD_KEYS = KEY_NAMES   # keeping old KEYBOARD_KEYS for backwards compatibility
keyboardMapping = dict([(key, None) for key in KEY_NAMES])

# Populate the basic printable ascii characters.
for c in range(32, 128):
    keyboardMapping[chr(c)] = ctypes.windll.user32.VkKeyScanA(ctypes.wintypes.WCHAR(chr(c)))



KEYEVENTF_KEYUP = 0x0002


INPUT_MOUSE = 0
INPUT_KEYBOARD = 1


####################################################################
#            GET CURRENT POSITION OF MOUSE POINTER                 #
####################################################################

class current_position():
    def __init__(self, cursor=None):
        self.cursor=POINT()
        windll.user32.GetCursorPos(ctypes.byref(self.cursor))
        self.x1=int(self.cursor.x)
        self.y1=int(self.cursor.y)

class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_ulong),
                ("y", ctypes.c_ulong)]


###################################################################
#           FUNCTIONS HELPS TO PERFORM MOVES & CLICKS             #
###################################################################

    
class FailSafeException(Exception):
    pass


def _failSafeCheck():
    if FAILSAFE and current_position() == (0, 0):
        raise FailSafeException('fail-safe triggered from mouse moving to upper-left corner. To disable this fail-safe, set c1_GUI.FAILSAFE to False.')

def linear(n):
    
    """
    Copied from get_tween module.
    """
    if not 0.0 <= n <= 1.0:
        raise ValueError('Argument must be between 0.0 and 1.0.')
    return n

def _autoPause(pause, _pause):
    if _pause:
        if pause is not None:
            time.sleep(pause)
        elif PAUSE != 0:
            time.sleep(PAUSE)

def getPointOnLine(x1, y1, x2, y2, n):
    x = ((x2 - x1) * n) + x1
    y = ((y2 - y1) * n) + y1
    return (x, y)


def check_cord(x, y):
    
    """Assign value of two co-orinates as separate x and y"""
    
    if isinstance(x, collections.Sequence):
        if len(x) == 2:
            if y is None:
                x, y = x
            else:
                raise ValueError('When passing a sequence at the x argument, the y argument must not be passed (received {0}).'.format(repr(y)))
        else:
            raise ValueError('The supplied sequence must have exactly 2 elements ({0} were received).'.format(len(x)))
    else:
        pass

    return x, y

def position():
    cursor = POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(cursor))
    x=int(cursor.x)
    y=int(cursor.y)
    return (x, y)

def _mouseMoveDrag(moveOrDrag, x, y, xOffset, yOffset, duration, tween=linear, button=None):
    
    """Handles the actual move event.

    """

    
    assert moveOrDrag in ('move', 'drag'), "moveOrDrag must be in ('move', 'drag'), not %s" % (moveOrDrag)

    if sys.platform != 'darwin':
        moveOrDrag = 'move' # For Windows it should be move.

    xOffset = int(xOffset) if xOffset is not None else 0
    yOffset = int(yOffset) if yOffset is not None else 0

    if x is None and y is None and xOffset == 0 and yOffset == 0:
        return  # Special case for no mouse movement at all.

    startx, starty = position()

    x = int(x) if x is not None else startx
    y = int(y) if y is not None else starty

    # x, y, xOffset, yOffset are now int.
    x += xOffset
    y += yOffset

    width, height = (ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))

    # Make sure x and y are within the screen boundries.
    
    x = max(0, min(x, width - 1))
    y = max(0, min(y, height - 1))

    # If the duration is small enough, just move the cursor there instantly.
    steps = [(x, y)]

    if duration > MINIMUM_DURATION:
        # Non-instant moving/dragging involves tweening:
        num_steps = max(width, height)
        sleep_amount = duration / num_steps
        if sleep_amount < MINIMUM_SLEEP:
            num_steps = int(duration / MINIMUM_SLEEP)
            sleep_amount = duration / num_steps

        steps = [
            getPointOnLine(startx, starty, x, y, tween(n / num_steps))
            for n in range(num_steps)
        ]
        # Making sure the last position is the actual destination.
        steps.append((x, y))

    for tweenX, tweenY in steps:
        if len(steps) > 1:
            # A single step does not require tweening.
            time.sleep(sleep_amount)

        _failSafeCheck()
        tweenX = int(round(tweenX))
        tweenY = int(round(tweenY))
        if moveOrDrag == 'move':
            ctypes.windll.user32.SetCursorPos(tweenX, tweenY)
        else:
            raise NotImplementedError('Unknown value of moveOrDrag: {0}'.format(moveOrDrag))

    _failSafeCheck()


###########################################################################
#                          MOUSE MOVE ACTION                              #
###########################################################################



def moveTo(x=None, y=None, duration=0.0, tween=linear, pause=None, _pause=True):
    
    """Moves the mouse cursor to a point on the screen."""
    
    x, y = check_cord(x, y)
    _failSafeCheck()
    _mouseMoveDrag('move', x, y, 0, 0, duration, tween)
    _autoPause(pause, _pause)


###########################################################################
#                          MOUSE CLICK ACTION                             #
###########################################################################


def click(x=None, y=None, clicks=1, interval=0.0, button=None, duration=0.0, tween=linear, pause=None, _pause=True):
    
    """Performs pressing a mouse button down and then immediately releasing it."""

    if button not in ('left', 'middle', 'right', 1, 2, 3):
        raise ValueError("button argument must be one of ('left', 'middle', 'right', 1, 2, 3)")

    _failSafeCheck()
    x, y = check_cord(x, y)
    _mouseMoveDrag('move', x, y, 0, 0, duration, tween)

    x, y = position()
    for i in range(clicks):
        _failSafeCheck()
        if button == 1 or str(button).lower() == 'left':
            _click(x, y, 'left')
        elif button == 2 or str(button).lower() == 'middle':
            _click(x, y, 'middle')
        elif button == 3 or str(button).lower() == 'right':
            _click(x, y, 'right')
        else:
            # These mouse buttons for hor. and vert. scrolling only apply to x11:
            _click(x, y, button)

        time.sleep(interval)

    _autoPause(pause, _pause)


def _click(x, y, button):
    
    """Send the mouse click event to Windows by calling the mouse_event() win32
    function."""
    
    if button == 'left':
        try:
            _sendMouseEvent(MOUSEEVENTF_LEFTCLICK, x, y)
        except (PermissionError, OSError): 
            pass
    elif button == 'middle':
        try:
            _sendMouseEvent(MOUSEEVENTF_MIDDLECLICK, x, y)
        except (PermissionError, OSError): 
            pass
    elif button == 'right':
        try:
            _sendMouseEvent(MOUSEEVENTF_RIGHTCLICK, x, y)
        except (PermissionError, OSError): 
            pass
    else:
        assert False, "button argument not in ('left', 'middle', 'right')"


def _sendMouseEvent(ev, x, y, dwData=0):

    assert x != None and y != None, 'x and y cannot be set to None'

    width, height = (ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))
    convertedX = 65536 * x // width + 1
    convertedY = 65536 * y // height + 1
    ctypes.windll.user32.mouse_event(ev, ctypes.c_long(convertedX), ctypes.c_long(convertedY), dwData, 0)




#######################################################################################################
#                                    MOUSE SCROLL ACTION                                              #
#######################################################################################################


##def scroll(clicks, x=None, y=None, pause=None, _pause=True):
##    
##    """Performs a scroll of the mouse scroll wheel."""
##    
##    _failSafeCheck()
##    if type(x) in (tuple, list):
##        x, y = x[0], x[1]
##    startx, starty = position()
##    width, height = (ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))
##
##    if x is None:
##        x = startx
##    else:
##        if x < 0:
##            x = 0
##        elif x >= width:
##            x = width - 1
##    if y is None:
##        y = starty
##    else:
##        if y < 0:
##            y = 0
##        elif y >= height:
##            y = height - 1
##
##    try:
##        _sendMouseEvent(MOUSEEVENTF_WHEEL, x, y, dwData=clicks)
##    except (PermissionError, OSError): 
##            pass
##
##    _autoPause(pause, _pause)



########################################################################################################
#                                "KEYBOARD FUNCTION STARTS"                                            #
########################################################################################################
    

def keyPress(key, pause=None, _pause=True):
    
    """
    Performs key press

    Args:
      key (str): The key to be pressed down. The valid names are listed in
      KEYBOARD_KEYS.
    """

    if len(key) > 1:
        key = key.lower()

    _failSafeCheck()
    Kboard._keyDown(key)

    _autoPause(pause, _pause)


def keyRelease(key, pause=None, _pause=True):
    
    """Performs a keyboard key release (without the press down beforehand)."""
    
    if len(key) > 1:
        key = key.lower()

    _failSafeCheck()
    Kboard._keyUp(key)

    _autoPause(pause, _pause)


def press(keys, presses=1, interval=0.0, pause=None, _pause=True):
    
    """Performs a keyboard key press down, followed by a release.

    """
    
    if type(keys) == str:
        keys = [keys] # put string in a list
    else:
        lowerKeys = []
        for s in keys:
            if len(s) > 1:
                lowerKeys.append(s.lower())
            else:
                lowerKeys.append(s)
    interval = float(interval)
    for i in range(presses):
        for k in keys:
            _failSafeCheck()
            Kboard._keyDown(k)
            Kboard._keyUp(k)
        time.sleep(interval)

    _autoPause(pause, _pause)

    
def passString(message, interval=0.0, pause=None, _pause=True):
    
    """Performs a keyboard key press down, followed by a release, for each of
    the characters in message.

    """
    
    interval = float(interval)

    _failSafeCheck()

    for c in message:
        if len(c) > 1:
            c = c.lower()
        press(c, _pause=False)
        time.sleep(interval)
        _failSafeCheck()

    _autoPause(pause, _pause)



def hotkeys(*args, **kwargs):
    
    """Performs key down presses on the arguments passed in order, then performs
    key releases in reverse order.

    """
    
    interval = float(kwargs.get('interval', 0.0))

    _failSafeCheck()

    for c in args:
        if len(c) > 1:
            c = c.lower()
        Kboard._keyDown(c)
        time.sleep(interval)
    for c in reversed(args):
        if len(c) > 1:
            c = c.lower()
        Kboard._keyUp(c)
        time.sleep(interval)

    _autoPause(kwargs.get('pause', None), kwargs.get('_pause', True))
