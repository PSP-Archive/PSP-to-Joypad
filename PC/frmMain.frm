VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PSP to JOYPAD"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2370
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   2370
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkEcho 
      Caption         =   "Enable ECHO"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   2040
      Width           =   2175
   End
   Begin VB.CommandButton btnKillJoypad 
      Caption         =   "Kill Joypad"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   240
      Width           =   915
   End
   Begin VB.Timer yMove 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1
      Left            =   1140
      Top             =   2820
   End
   Begin VB.Timer xMove 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   1
      Left            =   1560
      Top             =   2820
   End
   Begin MSWinsockLib.Winsock wskJoypad 
      Index           =   1
      Left            =   480
      Top             =   2820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8081
   End
   Begin MSWinsockLib.Winsock wskJoypad 
      Index           =   0
      Left            =   60
      Top             =   2820
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8081
   End
   Begin MSWinsockLib.Winsock wskServer 
      Left            =   60
      Top             =   3240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   8081
   End
   Begin VB.ListBox lstControllers 
      Height          =   450
      ItemData        =   "frmMain.frx":08CA
      Left            =   60
      List            =   "frmMain.frx":08CC
      TabIndex        =   6
      Top             =   2520
      Width           =   2235
   End
   Begin VB.HScrollBar ySensibility 
      Height          =   135
      LargeChange     =   10
      Left            =   60
      Max             =   10
      Min             =   1
      TabIndex        =   3
      Top             =   1800
      Value           =   5
      Width           =   2235
   End
   Begin VB.HScrollBar xSensibility 
      Height          =   135
      LargeChange     =   10
      Left            =   60
      Max             =   10
      Min             =   1
      TabIndex        =   2
      Top             =   1080
      Value           =   5
      Width           =   2235
   End
   Begin VB.CheckBox chkAnalogY 
      Caption         =   "Enable Analog Y Stick"
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   1380
      Width           =   2235
   End
   Begin VB.CheckBox chkAnalogX 
      Caption         =   "Enable Analog X Stick"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   660
      Width           =   2235
   End
   Begin VB.Timer xMove 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   1560
      Top             =   3240
   End
   Begin VB.Timer yMove 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   1140
      Top             =   3240
   End
   Begin VB.Timer tmrWinsock 
      Interval        =   200
      Left            =   480
      Top             =   3240
   End
   Begin VB.Shape stateJoypad 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Joypad 2"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   1620
      TabIndex        =   9
      Top             =   0
      Width           =   735
   End
   Begin VB.Shape stateJoypad 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   180
      Shape           =   4  'Rounded Rectangle
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Joypad 1"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Controllers"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   60
      TabIndex        =   7
      Top             =   2340
      Width           =   2235
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Y Sensibility"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   1620
      Width           =   2235
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "X Sensibility"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   900
      Width           =   2235
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim xMoving(2) As Integer
Dim yMoving(2) As Integer


Private Sub btnKillJoypad_Click()
Dim i As Integer
  For i = 0 To wskJoypad.Count - 1
    wskJoypad(i).Close
  Next
End Sub

Private Sub chkAnalogX_Click()
  xSensibility.Enabled = chkAnalogX.value = 1
End Sub

Private Sub chkAnalogY_Click()
  ySensibility.Enabled = chkAnalogY.value = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub

Private Sub tmrWinsock_Timer()
Dim i As Integer
  If wskServer.State = 0 Or _
  wskServer.State = 3 Or _
  wskServer.State = 8 Or _
  wskServer.State = 9 Then
    wskServer.Close
    wskServer.Listen
  End If
  
  For i = 0 To wskJoypad.Count - 1
    If wskJoypad(i).State = 7 Then
      stateJoypad(i).BackColor = &HC000&
    Else
      stateJoypad(i).BackColor = vbYellow
      If lstControllers.List(i) <> "" Then
        lstControllers.RemoveItem i
      End If
    End If
  Next
End Sub

Private Sub wskJoypad_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim strCommand() As String
Dim i As Integer

  stateJoypad(Index).BackColor = vbGreen
  wskJoypad(Index).GetData strData
  If Left(strData, 2) = "IP" Then
    lstControllers.AddItem (Index + 1) & ") " & Mid(strData, 3), Index
  Else
    strCommand = Split(strData, ",")
    For i = 0 To UBound(strCommand) - 1
      ProcessaDato strCommand(i), Index
    Next
  End If
End Sub

Private Sub wskServer_ConnectionRequest(ByVal requestID As Long)
  If wskJoypad(0).State <> 7 Then
    wskJoypad(0).Close
    wskJoypad(0).Accept requestID
  Else
    If wskJoypad(1).State <> 7 Then
      wskJoypad(1).Close
      wskJoypad(1).Accept requestID
    Else
      MsgBox "All controllers is connected! Please disconnect one controller and reconnect.", vbExclamation, "PSP to JOYPAD"
    End If
  End If
End Sub

Private Sub ProcessaDato(strData As String, nJoyPad As Integer)
Dim strParam() As String
Dim i As Integer

  strData = Trim(strData)
  If Len(strData) <= 3 Then
    Select Case strData
      Case "up1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_UP
          If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_H
          If chkEcho.value = 1 Then keybd_event vbKeyH, 0, 0, 0
        End If
      Case "up0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_UP
          If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_H
          If chkEcho.value = 1 Then keybd_event vbKeyH, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "dw1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_DOWN
          If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_B
          If chkEcho.value = 1 Then keybd_event vbKeyB, 0, 0, 0
        End If
      Case "dw0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_DOWN
          If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_B
          If chkEcho.value = 1 Then keybd_event vbKeyB, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "lf1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_LEFT
          If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_N
          If chkEcho.value = 1 Then keybd_event vbKeyN, 0, 0, 0
        End If
      Case "lf0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_LEFT
          If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_N
          If chkEcho.value = 1 Then keybd_event vbKeyN, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "rg1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_RIGHT
          If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_M
          If chkEcho.value = 1 Then keybd_event vbKeyM, 0, 0, 0
        End If
      Case "rg0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_RIGHT
          If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_M
          If chkEcho.value = 1 Then keybd_event vbKeyM, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "tr1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_W
          If chkEcho.value = 1 Then keybd_event vbKeyW, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_I
          If chkEcho.value = 1 Then keybd_event vbKeyI, 0, 0, 0
        End If
      Case "tr0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_W
          If chkEcho.value = 1 Then keybd_event vbKeyW, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_I
          If chkEcho.value = 1 Then keybd_event vbKeyI, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "cr1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_S
          If chkEcho.value = 1 Then keybd_event vbKeyS, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_K
          If chkEcho.value = 1 Then keybd_event vbKeyK, 0, 0, 0
        End If
      Case "cr0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_S
          If chkEcho.value = 1 Then keybd_event vbKeyS, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_K
          If chkEcho.value = 1 Then keybd_event vbKeyK, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "sq1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_A
          If chkEcho.value = 1 Then keybd_event vbKeyA, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_J
          If chkEcho.value = 1 Then keybd_event vbKeyJ, 0, 0, 0
        End If
      Case "sq0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_A
          If chkEcho.value = 1 Then keybd_event vbKeyA, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_J
          If chkEcho.value = 1 Then keybd_event vbKeyJ, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "ci1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_D
          If chkEcho.value = 1 Then keybd_event vbKeyD, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_L
          If chkEcho.value = 1 Then keybd_event vbKeyL, 0, 0, 0
        End If
      Case "ci0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_D
          If chkEcho.value = 1 Then keybd_event vbKeyD, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_L
          If chkEcho.value = 1 Then keybd_event vbKeyL, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "l11":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_Q
          If chkEcho.value = 1 Then keybd_event vbKeyQ, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_U
          If chkEcho.value = 1 Then keybd_event vbKeyU, 0, 0, 0
        End If
      Case "l10":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_Q
          If chkEcho.value = 1 Then keybd_event vbKeyQ, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_U
          If chkEcho.value = 1 Then keybd_event vbKeyU, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "l21":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_1
          If chkEcho.value = 1 Then keybd_event vbKey1, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_7
          If chkEcho.value = 1 Then keybd_event vbKey7, 0, 0, 0
        End If
      Case "l20":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_1
          If chkEcho.value = 1 Then keybd_event vbKey1, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_7
          If chkEcho.value = 1 Then keybd_event vbKey7, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "r11":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_E
          If chkEcho.value = 1 Then keybd_event vbKeyE, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_O
          If chkEcho.value = 1 Then keybd_event vbKeyO, 0, 0, 0
        End If
      Case "r10":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_E
          If chkEcho.value = 1 Then keybd_event vbKeyE, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_O
          If chkEcho.value = 1 Then keybd_event vbKeyO, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "r21":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_3
          If chkEcho.value = 1 Then keybd_event vbKey3, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_9
          If chkEcho.value = 1 Then keybd_event vbKey9, 0, 0, 0
        End If
      Case "r20":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_3
          If chkEcho.value = 1 Then keybd_event vbKey3, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_9
          If chkEcho.value = 1 Then keybd_event vbKey9, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "st1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_F
          If chkEcho.value = 1 Then keybd_event vbKeyF, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_C
          If chkEcho.value = 1 Then keybd_event vbKeyC, 0, 0, 0
        End If
      Case "st0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_F
          If chkEcho.value = 1 Then keybd_event vbKeyF, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_C
          If chkEcho.value = 1 Then keybd_event vbKeyC, 0, KEYEVENTF_KEYUP, 0
        End If
      Case "se1":
        If nJoyPad = 0 Then
          PressKey CONST_DIKEYFLAGS.DIK_G
          If chkEcho.value = 1 Then keybd_event vbKeyG, 0, 0, 0
        ElseIf nJoyPad = 1 Then
          PressKey CONST_DIKEYFLAGS.DIK_V
          If chkEcho.value = 1 Then keybd_event vbKeyV, 0, 0, 0
        End If
      Case "se0":
        If nJoyPad = 0 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_G
          If chkEcho.value = 1 Then keybd_event vbKeyG, 0, KEYEVENTF_KEYUP, 0
        ElseIf nJoyPad = 1 Then
          ReleaseKey CONST_DIKEYFLAGS.DIK_V
          If chkEcho.value = 1 Then keybd_event vbKeyV, 0, KEYEVENTF_KEYUP, 0
        End If
    End Select
  Else
    strParam = Split(strData, "|")
    Select Case strParam(0)
      Case "aX":
        If chkAnalogX.value = 1 Then
          If xSensibility = xSensibility.Max Then
            If CInt(strParam(1)) > 0 Then
              If nJoyPad = 0 Then
                PressKey CONST_DIKEYFLAGS.DIK_RIGHT
                If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, 0, 0
              ElseIf nJoyPad = 1 Then
                PressKey CONST_DIKEYFLAGS.DIK_M
                If chkEcho.value = 1 Then keybd_event vbKeyM, 0, 0, 0
              End If
            ElseIf CInt(strParam(1)) < 0 Then
              If nJoyPad = 0 Then
                PressKey CONST_DIKEYFLAGS.DIK_LEFT
                If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, 0, 0
              ElseIf nJoyPad = 1 Then
                PressKey CONST_DIKEYFLAGS.DIK_N
                If chkEcho.value = 1 Then keybd_event vbKeyN, 0, 0, 0
              End If
            Else
              If nJoyPad = 0 Then
                ReleaseKey CONST_DIKEYFLAGS.DIK_RIGHT
                If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, KEYEVENTF_KEYUP, 0
                ReleaseKey CONST_DIKEYFLAGS.DIK_LEFT
                If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, KEYEVENTF_KEYUP, 0
              ElseIf nJoyPad = 1 Then
                ReleaseKey CONST_DIKEYFLAGS.DIK_M
                If chkEcho.value = 1 Then keybd_event vbKeyM, 0, KEYEVENTF_KEYUP, 0
                ReleaseKey CONST_DIKEYFLAGS.DIK_N
                If chkEcho.value = 1 Then keybd_event vbKeyN, 0, KEYEVENTF_KEYUP, 0
              End If
            End If
          Else
            xMoving(nJoyPad) = CInt(strParam(1))
          End If
        End If
      Case "aY":
        If chkAnalogY.value = 1 Then
          If ySensibility = ySensibility.Max Then
            If CInt(strParam(1)) > 0 Then
              If nJoyPad = 0 Then
                PressKey CONST_DIKEYFLAGS.DIK_DOWN
                If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, 0, 0
              ElseIf nJoyPad = 1 Then
                PressKey CONST_DIKEYFLAGS.DIK_B
                If chkEcho.value = 1 Then keybd_event vbKeyB, 0, 0, 0
              End If
            ElseIf CInt(strParam(1)) < 0 Then
              If nJoyPad = 0 Then
                PressKey CONST_DIKEYFLAGS.DIK_UP
                If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, 0, 0
              ElseIf nJoyPad = 1 Then
                PressKey CONST_DIKEYFLAGS.DIK_H
                If chkEcho.value = 1 Then keybd_event vbKeyH, 0, 0, 0
              End If
            Else
              If nJoyPad = 0 Then
                ReleaseKey CONST_DIKEYFLAGS.DIK_UP
                If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, KEYEVENTF_KEYUP, 0
                ReleaseKey CONST_DIKEYFLAGS.DIK_DOWN
                If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, KEYEVENTF_KEYUP, 0
              ElseIf nJoyPad = 1 Then
                ReleaseKey CONST_DIKEYFLAGS.DIK_H
                If chkEcho.value = 1 Then keybd_event vbKeyH, 0, KEYEVENTF_KEYUP, 0
                ReleaseKey CONST_DIKEYFLAGS.DIK_B
                If chkEcho.value = 1 Then keybd_event vbKeyB, 0, KEYEVENTF_KEYUP, 0
              End If
            End If
          Else
            yMoving(nJoyPad) = CInt(strParam(1))
          End If
        End If
    End Select
    MoveAnalog xMoving(nJoyPad), yMoving(nJoyPad), nJoyPad
  End If
End Sub

Private Sub MoveAnalog(xM As Integer, yM As Integer, Index As Integer)
  If chkAnalogX.value = 1 And xSensibility <> xSensibility.Max Then
    If xM < 0 Then
      If Index = 0 Then
        PressKey CONST_DIKEYFLAGS.DIK_LEFT
        If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, 0, 0
      ElseIf Index = 1 Then
        PressKey CONST_DIKEYFLAGS.DIK_N
        If chkEcho.value = 1 Then keybd_event vbKeyN, 0, 0, 0
      End If
      If xSensibility.value < 100 Then
        xMove(Index).Interval = Abs(xM) * xSensibility.value
        xMove(Index).Enabled = True
      End If
    ElseIf xM > 0 Then
      If Index = 0 Then
        PressKey CONST_DIKEYFLAGS.DIK_RIGHT
        If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, 0, 0
      ElseIf Index = 1 Then
        PressKey CONST_DIKEYFLAGS.DIK_M
        If chkEcho.value = 1 Then keybd_event vbKeyM, 0, 0, 0
      End If
      If xSensibility.value < 100 Then
        xMove(Index).Interval = Abs(xM) * xSensibility.value
        xMove(Index).Enabled = True
      End If
    Else
      xMove_Timer Index
    End If
  End If

  If chkAnalogY.value = 1 And ySensibility <> ySensibility.Max Then
    If yM < 0 Then
      If Index = 0 Then
        PressKey CONST_DIKEYFLAGS.DIK_UP
        If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, 0, 0
      ElseIf Index = 1 Then
        PressKey CONST_DIKEYFLAGS.DIK_H
        If chkEcho.value = 1 Then keybd_event vbKeyH, 0, 0, 0
      End If
      If ySensibility.value < 100 Then
        yMove(Index).Interval = Abs(yM) * ySensibility.value
        yMove(Index).Enabled = True
      End If
    ElseIf yM > 0 Then
      If Index = 0 Then
        PressKey CONST_DIKEYFLAGS.DIK_DOWN
        If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, 0, 0
      ElseIf Index = 1 Then
        PressKey CONST_DIKEYFLAGS.DIK_B
        If chkEcho.value = 1 Then keybd_event vbKeyB, 0, 0, 0
      End If
      If ySensibility.value < 100 Then
        yMove(Index).Interval = Abs(yM) * ySensibility.value
        yMove(Index).Enabled = True
      End If
    Else
      yMove_Timer Index
    End If
  End If
End Sub

Private Sub xMove_Timer(Index As Integer)
  xMove(Index).Enabled = False
  If xMoving(Index) < 0 Then
    If Index = 0 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_LEFT
      If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 1 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_N
      If chkEcho.value = 1 Then keybd_event vbKeyN, 0, KEYEVENTF_KEYUP, 0
    End If
  ElseIf xMoving(Index) > 0 Then
    If Index = 0 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_RIGHT
      If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 1 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_M
      If chkEcho.value = 1 Then keybd_event vbKeyM, 0, KEYEVENTF_KEYUP, 0
    End If
  Else
    If Index = 0 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_LEFT
      If chkEcho.value = 1 Then keybd_event vbKeyLeft, 0, KEYEVENTF_KEYUP, 0
      ReleaseKey CONST_DIKEYFLAGS.DIK_RIGHT
      If chkEcho.value = 1 Then keybd_event vbKeyRight, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 1 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_N
      If chkEcho.value = 1 Then keybd_event vbKeyN, 0, KEYEVENTF_KEYUP, 0
      ReleaseKey CONST_DIKEYFLAGS.DIK_M
      If chkEcho.value = 1 Then keybd_event vbKeyM, 0, KEYEVENTF_KEYUP, 0
    End If
  End If
End Sub

Private Sub yMove_Timer(Index As Integer)
  yMove(Index).Enabled = False
  If yMoving(Index) < 0 Then
    If Index = 0 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_UP
      If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 1 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_H
      If chkEcho.value = 1 Then keybd_event vbKeyH, 0, KEYEVENTF_KEYUP, 0
    End If
  ElseIf yMoving(Index) > 0 Then
    If Index = 0 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_DOWN
      If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 1 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_B
      If chkEcho.value = 1 Then keybd_event vbKeyB, 0, KEYEVENTF_KEYUP, 0
    End If
  Else
    If Index = 0 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_UP
      If chkEcho.value = 1 Then keybd_event vbKeyUp, 0, KEYEVENTF_KEYUP, 0
      ReleaseKey CONST_DIKEYFLAGS.DIK_DOWN
      If chkEcho.value = 1 Then keybd_event vbKeyDown, 0, KEYEVENTF_KEYUP, 0
    ElseIf Index = 1 Then
      ReleaseKey CONST_DIKEYFLAGS.DIK_H
      If chkEcho.value = 1 Then keybd_event vbKeyH, 0, KEYEVENTF_KEYUP, 0
      ReleaseKey CONST_DIKEYFLAGS.DIK_B
      If chkEcho.value = 1 Then keybd_event vbKeyB, 0, KEYEVENTF_KEYUP, 0
    End If
  End If
End Sub



