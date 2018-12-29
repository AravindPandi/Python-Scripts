import ctypes
import ctypes.wintypes
import sys




try:
   ctypes.windll.user32.SetProcessDPIAware()
except AttributeError:
    pass # Windows XP doesn't support this, so just do nothing.


"""
A lot of this code is probably repeated from win32 extensions module, but I didn't want to have that dependency.

Note: According to http://msdn.microsoft.com/en-us/library/windows/desktop/ms646260(v=vs.85).aspx
the ctypes.windll.user32.mouse_event() function has been superceded by SendInput.

SendInput() is documented here: http://msdn.microsoft.com/en-us/library/windows/desktop/ms646310(v=vs.85).aspx

UPDATE: SendInput() doesn't seem to be working for me. I've switched back to mouse_event()."""


# Event codes to be passed to the mouse_event() win32 function.
# Documented here: http://msdn.microsoft.com/en-us/library/windows/desktop/ms646273(v=vs.85).aspx
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

# Documented here: http://msdn.microsoft.com/en-us/library/windows/desktop/ms646304(v=vs.85).aspx
KEYEVENTF_KEYUP = 0x0002

# Documented here: http://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
INPUT_MOUSE = 0
INPUT_KEYBOARD = 1


# This ctypes structure is for a Win32 POINT structure,
# which is documented here: http://msdn.microsoft.com/en-us/library/windows/desktop/dd162805(v=vs.85).aspx
# The POINT structure is used by GetCursorPos().
class POINT(ctypes.Structure):
    _fields_ = [("x", ctypes.c_ulong),
                ("y", ctypes.c_ulong)]

# These ctypes structures are for Win32 INPUT, MOUSEINPUT, KEYBDINPUT, and HARDWAREINPUT structures,
# used by SendInput and documented here: http://msdn.microsoft.com/en-us/library/windows/desktop/ms646270(v=vs.85).aspx
# Thanks to BSH for this StackOverflow answer: https://stackoverflow.com/questions/18566289/how-would-you-recreate-this-windows-api-structure-with-ctypes
class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ('dx', ctypes.wintypes.LONG),
        ('dy', ctypes.wintypes.LONG),
        ('mouseData', ctypes.wintypes.DWORD),
        ('dwFlags', ctypes.wintypes.DWORD),
        ('time', ctypes.wintypes.DWORD),
        ('dwExtraInfo', ctypes.POINTER(ctypes.wintypes.ULONG)),
    ]

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ('wVk', ctypes.wintypes.WORD),
        ('wScan', ctypes.wintypes.WORD),
        ('dwFlags', ctypes.wintypes.DWORD),
        ('time', ctypes.wintypes.DWORD),
        ('dwExtraInfo', ctypes.POINTER(ctypes.wintypes.ULONG)),
    ]

class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ('uMsg', ctypes.wintypes.DWORD),
        ('wParamL', ctypes.wintypes.WORD),
        ('wParamH', ctypes.wintypes.DWORD)
    ]

class INPUT(ctypes.Structure):
    class _I(ctypes.Union):
        _fields_ = [
            ('mi', MOUSEINPUT),
            ('ki', KEYBDINPUT),
            ('hi', HARDWAREINPUT),
        ]

    _anonymous_ = ('i', )
    _fields_ = [
        ('type', ctypes.wintypes.DWORD),
        ('i', _I),
    ]
# End of the SendInput win32 data structures.

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



""" Keyboard key mapping for pyautogui:
Documented at http://msdn.microsoft.com/en-us/library/windows/desktop/dd375731(v=vs.85).aspx

The *KB dictionaries in pyautogui map a string that can be passed to keyDown(),
keyUp(), or press() into the code used for the OS-specific keyboard function.

They should always be lowercase, and the same keys should be used across all OSes."""
keyboardMapping = dict([(key, None) for key in KEY_NAMES])
keyboardMapping.update({
    'backspace': 0x08, # VK_BACK
    '\b': 0x08, # VK_BACK
    'super': 0x5B, #VK_LWIN
    'tab': 0x09, # VK_TAB
    '\t': 0x09, # VK_TAB
    'clear': 0x0c, # VK_CLEAR
    'enter': 0x0d, # VK_RETURN
    '\n': 0x0d, # VK_RETURN
    'return': 0x0d, # VK_RETURN
    'shift': 0x10, # VK_SHIFT
    'ctrl': 0x11, # VK_CONTROL
    'alt': 0x12, # VK_MENU
    'pause': 0x13, # VK_PAUSE
    'capslock': 0x14, # VK_CAPITAL
    'kana': 0x15, # VK_KANA
    'hanguel': 0x15, # VK_HANGUEL
    'hangul': 0x15, # VK_HANGUL
    'junja': 0x17, # VK_JUNJA
    'final': 0x18, # VK_FINAL
    'hanja': 0x19, # VK_HANJA
    'kanji': 0x19, # VK_KANJI
    'esc': 0x1b, # VK_ESCAPE
    'escape': 0x1b, # VK_ESCAPE
    'convert': 0x1c, # VK_CONVERT
    'nonconvert': 0x1d, # VK_NONCONVERT
    'accept': 0x1e, # VK_ACCEPT
    'modechange': 0x1f, # VK_MODECHANGE
    ' ': 0x20, # VK_SPACE
    'space': 0x20,
    'pgup': 0x21, # VK_PRIOR
    'pgdn': 0x22, # VK_NEXT
    'pageup': 0x21, # VK_PRIOR
    'pagedown': 0x22, # VK_NEXT
    'end': 0x23, # VK_END
    'home': 0x24, # VK_HOME
    'left': 0x25, # VK_LEFT
    'up': 0x26, # VK_UP
    'right': 0x27, # VK_RIGHT
    'down': 0x28, # VK_DOWN
    'select': 0x29, # VK_SELECT
    'print': 0x2a, # VK_PRINT
    'execute': 0x2b, # VK_EXECUTE
    'prtsc': 0x2c, # VK_SNAPSHOT
    'prtscr': 0x2c, # VK_SNAPSHOT
    'prntscrn': 0x2c, # VK_SNAPSHOT
    'printscreen': 0x2c, # VK_SNAPSHOT
    'insert': 0x2d, # VK_INSERT
    'del': 0x2e, # VK_DELETE
    'delete': 0x2e, # VK_DELETE
    'help': 0x2f, # VK_HELP
    'win': 0x5b, # VK_LWIN
    'winleft': 0x5b, # VK_LWIN
    'winright': 0x5c, # VK_RWIN
    'apps': 0x5d, # VK_APPS
    'sleep': 0x5f, # VK_SLEEP
    'num0': 0x60, # VK_NUMPAD0
    'num1': 0x61, # VK_NUMPAD1
    'num2': 0x62, # VK_NUMPAD2
    'num3': 0x63, # VK_NUMPAD3
    'num4': 0x64, # VK_NUMPAD4
    'num5': 0x65, # VK_NUMPAD5
    'num6': 0x66, # VK_NUMPAD6
    'num7': 0x67, # VK_NUMPAD7
    'num8': 0x68, # VK_NUMPAD8
    'num9': 0x69, # VK_NUMPAD9
    'multiply': 0x6a, # VK_MULTIPLY  ??? Is this the numpad *?
    'add': 0x6b, # VK_ADD  ??? Is this the numpad +?
    'separator': 0x6c, # VK_SEPARATOR  ??? Is this the numpad enter?
    'subtract': 0x6d, # VK_SUBTRACT  ??? Is this the numpad -?
    'decimal': 0x6e, # VK_DECIMAL
    'divide': 0x6f, # VK_DIVIDE
    'f1': 0x70, # VK_F1
    'f2': 0x71, # VK_F2
    'f3': 0x72, # VK_F3
    'f4': 0x73, # VK_F4
    'f5': 0x74, # VK_F5
    'f6': 0x75, # VK_F6
    'f7': 0x76, # VK_F7
    'f8': 0x77, # VK_F8
    'f9': 0x78, # VK_F9
    'f10': 0x79, # VK_F10
    'f11': 0x7a, # VK_F11
    'f12': 0x7b, # VK_F12
    'f13': 0x7c, # VK_F13
    'f14': 0x7d, # VK_F14
    'f15': 0x7e, # VK_F15
    'f16': 0x7f, # VK_F16
    'f17': 0x80, # VK_F17
    'f18': 0x81, # VK_F18
    'f19': 0x82, # VK_F19
    'f20': 0x83, # VK_F20
    'f21': 0x84, # VK_F21
    'f22': 0x85, # VK_F22
    'f23': 0x86, # VK_F23
    'f24': 0x87, # VK_F24
    'numlock': 0x90, # VK_NUMLOCK
    'scrolllock': 0x91, # VK_SCROLL
    'shiftleft': 0xa0, # VK_LSHIFT
    'shiftright': 0xa1, # VK_RSHIFT
    'ctrlleft': 0xa2, # VK_LCONTROL
    'ctrlright': 0xa3, # VK_RCONTROL
    'altleft': 0xa4, # VK_LMENU
    'altright': 0xa5, # VK_RMENU
    'browserback': 0xa6, # VK_BROWSER_BACK
    'browserforward': 0xa7, # VK_BROWSER_FORWARD
    'browserrefresh': 0xa8, # VK_BROWSER_REFRESH
    'browserstop': 0xa9, # VK_BROWSER_STOP
    'browsersearch': 0xaa, # VK_BROWSER_SEARCH
    'browserfavorites': 0xab, # VK_BROWSER_FAVORITES
    'browserhome': 0xac, # VK_BROWSER_HOME
    'volumemute': 0xad, # VK_VOLUME_MUTE
    'volumedown': 0xae, # VK_VOLUME_DOWN
    'volumeup': 0xaf, # VK_VOLUME_UP
    'nexttrack': 0xb0, # VK_MEDIA_NEXT_TRACK
    'prevtrack': 0xb1, # VK_MEDIA_PREV_TRACK
    'stop': 0xb2, # VK_MEDIA_STOP
    'playpause': 0xb3, # VK_MEDIA_PLAY_PAUSE
    'launchmail': 0xb4, # VK_LAUNCH_MAIL
    'launchmediaselect': 0xb5, # VK_LAUNCH_MEDIA_SELECT
    'launchapp1': 0xb6, # VK_LAUNCH_APP1
    'launchapp2': 0xb7, # VK_LAUNCH_APP2
    #';': 0xba, # VK_OEM_1
    #'+': 0xbb, # VK_OEM_PLUS
    #',': 0xbc, # VK_OEM_COMMA
    #'-': 0xbd, # VK_OEM_MINUS
    #'.': 0xbe, # VK_OEM_PERIOD
    #'/': 0xbf, # VK_OEM_2
    #'~': 0xc0, # VK_OEM_3
    #'[': 0xdb, # VK_OEM_4
    #'|': 0xdc, # VK_OEM_5
    #']': 0xdd, # VK_OEM_6
    #"'": 0xde, # VK_OEM_7
    #'': 0xdf, # VK_OEM_8
    #'': 0xe7, # VK_PACKET
    #'': 0xf6, # VK_ATTN
    #'': 0xf7, # VK_CRSEL
    #'': 0xf8, # VK_EXSEL
    #'': 0xf9, # VK_EREOF
    #'': 0xfa, # VK_PLAY
    #'': 0xfb, # VK_ZOOM
    #'': 0xfc, # VK_NONAME
    #'': 0xfd, # VK_PA1
    #'': 0xfe, # VK_OEM_CLEAR
})

# Populate the basic printable ascii characters.
for c in range(32, 128):
    keyboardMapping[chr(c)] = ctypes.windll.user32.VkKeyScanA(ctypes.wintypes.WCHAR(chr(c)))

def isShiftCharacter(character):
    """Returns True if the key character is uppercase or shifted."""
    return character.isupper() or character in '~!@#$%^&*()_+{}|:"<>?'


def _keyDown(key):
    """Performs a keyboard key press without the release. This will put that
    key in a held down state.

    NOTE: For some reason, this does not seem to cause key repeats like would
    happen if a keyboard key was held down on a text field.

    Args:
      key (str): The key to be pressed down. The valid names are listed in
      pyautogui.KEY_NAMES.

    Returns:
      None
    """
    if key not in keyboardMapping or keyboardMapping[key] is None:
        return

    needsShift = isShiftCharacter(key)

    """
    # OLD CODE: The new code relies on having all keys be loaded in keyboardMapping from the start.
    if key in keyboardMapping.keys():
        vkCode = keyboardMapping[key]
    elif len(key) == 1:
        # note: I could use this case to update keyboardMapping to cache the VkKeyScan results, but I've decided not to just to make any possible bugs easier to reproduce.
        vkCode = ctypes.windll.user32.VkKeyScanW(ctypes.wintypes.WCHAR(key))
        if vkCode == -1:
            raise ValueError('There is no VK code for key "%s"' % (key))
        if vkCode > 0x100: # the vk code will be > 0x100 if it needs shift
            vkCode -= 0x100
            needsShift = True
    """
    mods, vkCode = divmod(keyboardMapping[key], 0x100)

    for apply_mod, vk_mod in [(mods & 4, 0x12), (mods & 2, 0x11),
        (mods & 1 or needsShift, 0x10)]: #HANKAKU not suported! mods & 8
        if apply_mod:
            ctypes.windll.user32.keybd_event(vk_mod, 0, 0, 0) #
    ctypes.windll.user32.keybd_event(vkCode, 0, 0, 0)
    for apply_mod, vk_mod in [(mods & 1 or needsShift, 0x10), (mods & 2, 0x11),
        (mods & 4, 0x12)]: #HANKAKU not suported! mods & 8
        if apply_mod:
            ctypes.windll.user32.keybd_event(vk_mod, 0, KEYEVENTF_KEYUP, 0) #


def _keyUp(key):
    """Performs a keyboard key release (without the press down beforehand).

    Args:
      key (str): The key to be released up. The valid names are listed in
      pyautogui.KEY_NAMES.

    Returns:
      None
    """
    if key not in keyboardMapping or keyboardMapping[key] is None:
        return

    needsShift = isShiftCharacter(key)
    """
    # OLD CODE: The new code relies on having all keys be loaded in keyboardMapping from the start.
    if key in keyboardMapping.keys():
        vkCode = keyboardMapping[key]
    elif len(key) == 1:
        # note: I could use this case to update keyboardMapping to cache the VkKeyScan results, but I've decided not to just to make any possible bugs easier to reproduce.
        vkCode = ctypes.windll.user32.VkKeyScanW(ctypes.wintypes.WCHAR(key))
        if vkCode == -1:
            raise ValueError('There is no VK code for key "%s"' % (key))
        if vkCode > 0x100: # the vk code will be > 0x100 if it needs shift
            vkCode -= 0x100
            needsShift = True
    """
    mods, vkCode = divmod(keyboardMapping[key], 0x100)

    for apply_mod, vk_mod in [(mods & 4, 0x12), (mods & 2, 0x11),
        (mods & 1 or needsShift, 0x10)]: #HANKAKU not suported! mods & 8
        if apply_mod:
            ctypes.windll.user32.keybd_event(vk_mod, 0, 0, 0) #
    ctypes.windll.user32.keybd_event(vkCode, 0, KEYEVENTF_KEYUP, 0)
    for apply_mod, vk_mod in [(mods & 1 or needsShift, 0x10), (mods & 2, 0x11),
        (mods & 4, 0x12)]: #HANKAKU not suported! mods & 8
        if apply_mod:
            ctypes.windll.user32.keybd_event(vk_mod, 0, KEYEVENTF_KEYUP, 0) #

