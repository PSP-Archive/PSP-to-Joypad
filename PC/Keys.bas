Attribute VB_Name = "Keys"
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As GENERALINPUT, ByVal cbSize As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Const VK_SHIFT = &H10
Const VK_CONTROL = &H11
Const VK_MENU = &H12
Const KEYEVENTF_KEYUP = &H2
Const INPUT_KEYBOARD = 1

Private Type KEYBDINPUT
  wVk As Integer
  wScan As Integer
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Private Type MOUSEINPUT
  dx As Long
  dy As Long
  mouseData As Long
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Public Type HARDWAREINPUT
  uMsg As Long
  value As Long
  'wParamL As Integer
  'wParamH As Integer
End Type

Private Type GENERALINPUT
  dwType As Long
  ki(0 To 23) As Byte
End Type

Private Type GENERALINPUT2
  dwType As Long
  'mi(0 To 23) As Byte
  ki(0 To 23) As Byte
  'hi(0 To 23) As Byte
End Type

Enum CONST_DIKEYFLAGS
    DIK_0 = 11
    DIK_1 = 2
    DIK_2 = 3
    DIK_3 = 4
    DIK_4 = 5
    DIK_5 = 6
    DIK_6 = 7
    DIK_7 = 8
    DIK_8 = 9
    DIK_9 = 10
    DIK_A = 30             '(&H1E)
    DIK_ABNT_C1 = 115      '(&H73)
    DIK_ABNT_C2 = 126      '(&H7E)
    DIK_ADD = 78           '(&H4E)
    DIK_APOSTROPHE = 40    '(&H28)
    DIK_APPS = 221         '(&HDD)
    DIK_AT = 145           '(&H91)
    DIK_AX = 150           '(&H96)
    DIK_B = 48             '(&H30)
    DIK_BACK = 14
    DIK_BACKSLASH = 43     '(&H2B)
    DIK_BACKSPACE = 14
    DIK_C = 46             '(&H2E)
    DIK_CALCULATOR = 161   '(&HA1)
    DIK_CAPITAL = 58       '(&H3A)
    DIK_CAPSLOCK = 58      '(&H3A)
    DIK_CIRCUMFLEX = 144   '(&H90)
    DIK_COLON = 146        '(&H92)
    DIK_COMMA = 51         '(&H33)
    DIK_CONVERT = 121      '(&H79)
    DIK_D = 32             '(&H20)
    DIK_DECIMAL = 83       '(&H53)
    DIK_DELETE = 211       '(&HD3)
    DIK_DIVIDE = 181       '(&HB5)
    DIK_DOWN = 208         '(&HD0)
    DIK_DOWNARROW = 208    '(&HD0)
    DIK_E = 18             '(&H12)
    DIK_END = 207          '(&HCF)
    DIK_EQUALS = 13
    DIK_ESCAPE = 1
    DIK_F = 33             '(&H21)
    DIK_F1 = 59            '(&H3B)
    DIK_F2 = 60            '(&H3C)
    DIK_F3 = 61            '(&H3D)
    DIK_F4 = 62            '(&H3E)
    DIK_F5 = 63            '(&H3F)
    DIK_F6 = 64            '(&H40)
    DIK_F7 = 65            '(&H41)
    DIK_F8 = 66            '(&H42)
    DIK_F9 = 67            '(&H43)
    DIK_F10 = 68           '(&H44)
    DIK_F11 = 87           '(&H57)
    DIK_F12 = 88           '(&H58)
    DIK_F13 = 100          '(&H64)
    DIK_F14 = 101          '(&H65)
    DIK_F15 = 102          '(&H66)
    DIK_G = 34             '(&H22)
    DIK_GRAVE = 41         '(&H29)
    DIK_H = 35             '(&H23)
    DIK_HOME = 199         '(&HC7)
    DIK_I = 23             '(&H17)
    DIK_INSERT = 210       '(&HD2)
    DIK_J = 36             '(&H24)
    DIK_K = 37             '(&H25)
    DIK_KANA = 112         '(&H70)
    DIK_KANJI = 148        '(&H94)
    DIK_L = 38             '(&H26)
    DIK_LALT = 56          '(&H38)
    DIK_LBRACKET = 26      '(&H1A)
    DIK_LCONTROL = 29      '(&H1D)
    DIK_LEFT = 203         '(&HCB)
    DIK_LEFTARROW = 203    '(&HCB)
    DIK_LMENU = 56         '(&H38)
    DIK_LSHIFT = 42        '(&H2A)
    DIK_LWIN = 219         '(&HDB)
    DIK_M = 50             '(&H32)
    DIK_MAIL = 236         '(&HEC)
    DIK_MEDIASELECT = 237  '(&HED)
    DIK_MEDIASTOP = 164    '(&HA4)
    DIK_MINUS = 12
    DIK_MULTIPLY = 55      '(&H37)
    DIK_MUTE = 160         '(&HA0)
    DIK_MYCOMPUTER = 235   '(&HEB)
    DIK_N = 49             '(&H31)
    DIK_NEXT = 209         '(&HD1)
    DIK_NEXTTRACK = 153    '(&H99)
    DIK_NOCONVERT = 123    '(&H7B)
    DIK_NUMLOCK = 69       '(&H45)
    DIK_NUMPAD0 = 82       '(&H52)
    DIK_NUMPAD1 = 79       '(&H4F)
    DIK_NUMPAD2 = 80       '(&H50)
    DIK_NUMPAD3 = 81       '(&H51)
    DIK_NUMPAD4 = 75       '(&H4B)
    DIK_NUMPAD5 = 76       '(&H4C)
    DIK_NUMPAD6 = 77       '(&H4D)
    DIK_NUMPAD7 = 71       '(&H47)
    DIK_NUMPAD8 = 72       '(&H48)
    DIK_NUMPAD9 = 73       '(&H49)
    DIK_NUMPADCOMMA = 179  '(&HB3)
    DIK_NUMPADENTER = 156  '(&H9C)
    DIK_NUMPADEQUALS = 141 '(&H8D)
    DIK_NUMPADMINUS = 74   '(&H4A)
    DIK_NUMPADPERIOD = 83  '(&H53)
    DIK_NUMPADPLUS = 78    '(&H4E)
    DIK_NUMPADSLASH = 181  '(&HB5)
    DIK_NUMPADSTAR = 55    '(&H37)
    DIK_O = 24             '(&H18)
    DIK_OEM_102 = 86       '(&H56)
    DIK_P = 25             '(&H19)
    DIK_PAUSE = 197        '(&HC5)
    DIK_PERIOD = 52        '(&H34)
    DIK_PGDN = 209         '(&HD1)
    DIK_PGUP = 201         '(&HC9)
    DIK_PLAYPAUSE = 162    '(&HA2)
    DIK_POWER = 222        '(&HDE)
    DIK_PREVTRACK = 144    '(&H90)
    DIK_PRIOR = 201        '(&HC9)
    DIK_Q = 16             '(&H10)
    DIK_R = 19             '(&H13)
    DIK_RALT = 184         '(&HB8)
    DIK_RBRACKET = 27      '(&H1B)
    DIK_RCONTROL = 157     '(&H9D)
    DIK_RETURN = 28        '(&H1C)
    DIK_RIGHT = 205        '(&HCD)
    DIK_RIGHTARROW = 205   '(&HCD)
    DIK_RMENU = 184        '(&HB8)
    DIK_RSHIFT = 54        '(&H36)
    DIK_RWIN = 220         '(&HDC)
    DIK_S = 31             '(&H1F)
    DIK_SCROLL = 70        '(&H46)
    DIK_SEMICOLON = 39     '(&H27)
    DIK_SLASH = 53         '(&H35)
    DIK_SLEEP = 223        '(&HDF)
    DIK_SPACE = 57         '(&H39)
    DIK_STOP = 149         '(&H95)
    DIK_SUBTRACT = 74      '(&H4A)
    DIK_SYSRQ = 183        '(&HB7)
    DIK_T = 20             '(&H14)
    DIK_TAB = 15
    DIK_U = 22             '(&H16)
    DIK_UNDERLINE = 147    '(&H93)
    DIK_UNLABELED = 151    '(&H97)
    DIK_UP = 200           '(&HC8)
    DIK_UPARROW = 200      '(&HC8)
    DIK_V = 47             '(&H2F)
    DIK_VOLUMEDOWN = 174   '(&HAE)
    DIK_VOLUMEUP = 176     '(&HB0)
    DIK_W = 17             '(&H11)
    DIK_WAKE = 227         '(&HE3)
    DIK_WEBBACK = 234      '(&HEA)
    DIK_WEBFAVORITES = 230 '(&HE6)
    DIK_WEBFORWARD = 233   '(&HE9)
    DIK_WEBHOME = 178      '(&HB2)
    DIK_WEBREFRESH = 231   '(&HE7)
    DIK_WEBSEARCH = 229    '(&HE5)
    DIK_WEBSTOP = 232      '(&HE8)
    DIK_X = 45             '(&H2D)
    DIK_Y = 21             '(&H15)
    DIK_YEN = 125          '(&H7D)
    DIK_Z = 44             '(&H2C)
End Enum



Sub Main()
Dim bKey As Byte
If ListVar2 = "shift" Then bKey = VK_SHIFT
If ListVar2 = "control" Then bKey = VK_CONTROL
If ListVar2 = "alt" Then bKey = VK_MENU

Select Case ListVar1
Case "hold down", "press", "hold"
PressKey bKey
Case "release", "lift", "reset"
ReleaseKey bKey
End Select
End Sub

Public Sub PressKey(bKey As Byte)
Dim GInput(0) As GENERALINPUT
Dim KInput As KEYBDINPUT
KInput.wScan = bKey 'the key we're going to press
KInput.dwFlags = 0 'press the key
'copy the structure into the input array's buffer.
GInput(0).dwType = INPUT_KEYBOARD ' keyboard input
CopyMemory GInput(0).ki(0), KInput, Len(KInput)
'send the input now
Call SendInput(1, GInput(0), Len(GInput(0)))
End Sub

Public Sub ReleaseKey(bKey As Byte)
Dim GInput(0) As GENERALINPUT
Dim KInput As KEYBDINPUT
KInput.wScan = bKey ' the key we're going to realease
KInput.dwFlags = KEYEVENTF_KEYUP ' release the key
GInput(0).dwType = INPUT_KEYBOARD ' keyboard input
CopyMemory GInput(0).ki(0), KInput, Len(KInput)
'send the input now
Call SendInput(1, GInput(0), Len(GInput(0)))
End Sub

