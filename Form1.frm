VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   Caption         =   "PASTE ANY PICTURE TO RICHTEXTBOX....BY KAYHAN TANRISEVEN...THE BENCHMARKER®"
   ClientHeight    =   4035
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4035
   ScaleWidth      =   9540
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4515
      Top             =   3975
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1710
      Left            =   15
      TabIndex        =   3
      Top             =   2310
      Width           =   4230
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   $"Form1.frx":0000
         Height          =   1485
         Left            =   240
         TabIndex        =   4
         Top             =   1635
         Width           =   3915
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Paste ==>"
      Height          =   1350
      Left            =   3465
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2085
      Left            =   120
      Picture         =   "Form1.frx":014B
      ScaleHeight     =   2025
      ScaleWidth      =   3090
      TabIndex        =   0
      Top             =   120
      Width           =   3150
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3855
      Left            =   4920
      TabIndex        =   2
      Top             =   120
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6800
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":733D
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_PASTE = &H302
Private Sub Command1_Click()
    ' Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData Picture1.Picture
    
    ' Paste the picture into the RichTextBox.
    SendMessage RichTextBox1.hwnd, WM_PASTE, 0, 0
End Sub

Private Sub Timer1_Timer()
If Label1.Top = 165 Then
Timer1.Enabled = False
End If
Label1.Top = Label1.Top - 15
End Sub
