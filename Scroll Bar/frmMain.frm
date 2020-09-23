VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Scroll bar..."
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "SET"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2535
      TabIndex        =   3
      Top             =   450
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   1335
      TabIndex        =   2
      Text            =   "25"
      Top             =   465
      Width           =   1080
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   105
      ScaleHeight     =   165
      ScaleWidth      =   3510
      TabIndex        =   0
      Top             =   105
      Width           =   3540
      Begin VB.PictureBox picBar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   180
         TabIndex        =   1
         Top             =   -15
         Width           =   180
      End
   End
   Begin VB.Label lblValue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   450
      Width           =   300
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'|------------------------------------------------|'
'|Scroll Bar[cool!!]                              |'
'|------------------------------------------------|'
'|Written by Pallab Mahmud                        |'
'|© Copyright 2001 by Pallab Mahmud               |'
'|email: pallmahmud@yahoo.com                     |'
'|                                                |'
'|This sample code is a FREEWARE. Use it in your  |'
'|own project as it fits You but do not re-sale   |'
'|this code or destroy the original authors name. |'
'|                                                |'
'|Warning: No warranty is provided with this set  |'
'|of code so use it in your own risk. The author  |'
'|is not responsible for the Damage caused by     |'
'|this code.                                      |'
'|------------------------------------------------|'
'|------------------------------------------------|'
'Comments:This is a cool scroll bar.You can change
'it base and bar picture whatever you want.It uses
'only one api call.I think it is great.What do you think?
'Hey,listen i am new in programing and i am 14 years old
'So,don't mind and Please please........vote for me
'--------------------------------------------------'
Option Explicit
Dim hTl, iXY, slPos
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Sub cmdSet_Click()
    slPos = FormatNumber(((Text1) / 100) * (picMain.ScaleWidth - picBar.Width), 0, , , vbFalse)
    If slPos > (picMain.ScaleWidth - picBar.Width) Then
        slPos = (picMain.ScaleWidth - picBar.Width)
        Text1 = 100
    ElseIf slPos < 0 Then
        slPos = 0
        Text1 = 0
    End If
    picBar.Left = slPos
    lblValue = FormatNumber((picBar.Left / (picMain.ScaleWidth - picBar.Width)) * 100, 0, , , vbFalse) & "%"
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdSet_Click
    End If
End Sub
Private Sub picbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iXY = Y
    hTl = picBar.Top
End Sub
Private Sub picbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        slPos = hTl + (X - iXY)
        If slPos < 0 Then slPos = 0
            If slPos > (picMain.ScaleWidth - picBar.Width) Then slPos = (picMain.ScaleWidth - picBar.Width)
                hTl = slPos
                picBar.Left = slPos
    End If

    lblValue = FormatNumber((picBar.Left / (picMain.ScaleWidth - picBar.Width)) * 100, 0, , , vbFalse) & "%"
End Sub
Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    iXY = picBar.Width / 2
    slPos = X
    If slPos > picMain.ScaleWidth - picBar.Width Then slPos = picMain.ScaleWidth - picBar.Width
    picBar.Left = slPos
    SetCapture picBar.hwnd
End Sub
