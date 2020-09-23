VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBlank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   7380
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   14
      Top             =   660
      Visible         =   0   'False
      Width           =   6030
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calendar Properties"
      Height          =   4335
      Left            =   120
      TabIndex        =   13
      Top             =   60
      Width           =   4335
      Begin VB.Frame Frame3 
         Caption         =   "Month Font"
         Height          =   1515
         Left            =   180
         TabIndex        =   20
         Top             =   1440
         Width           =   3975
         Begin VB.ComboBox cboMFont 
            Height          =   315
            Left            =   780
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   300
            Width           =   2895
         End
         Begin VB.Label Label2 
            Caption         =   "Font"
            Height          =   195
            Left            =   240
            TabIndex        =   22
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Date"
         Height          =   1155
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   3975
         Begin VB.OptionButton optDate 
            Caption         =   "Current Month/Year"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optDate 
            Caption         =   "Set Month/Year"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   18
            Top             =   660
            Width           =   1515
         End
         Begin VB.ComboBox cboMmm 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1740
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox cboYYYY 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2760
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   600
            Width           =   975
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808000&
      Height          =   975
      Left            =   660
      ScaleHeight     =   915
      ScaleWidth      =   1155
      TabIndex        =   2
      Top             =   5460
      Width           =   1215
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   9
         Left            =   60
         TabIndex        =   11
         Top             =   60
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   10
         Top             =   60
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   2
         Left            =   900
         TabIndex        =   9
         Top             =   60
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   3
         Left            =   900
         TabIndex        =   8
         Top             =   360
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   4
         Left            =   900
         TabIndex        =   7
         Top             =   660
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   5
         Left            =   480
         TabIndex        =   6
         Top             =   660
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   5
         Top             =   660
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   4
         Top             =   360
         Width           =   195
      End
      Begin VB.CommandButton Command1 
         Height          =   195
         Index           =   8
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   195
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         Height          =   315
         Left            =   420
         Top             =   300
         Visible         =   0   'False
         Width           =   315
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Draw Calendar"
      Height          =   495
      Left            =   8220
      TabIndex        =   1
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox picBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   5340
      Picture         =   "Form1.frx":57522
      ScaleHeight     =   298
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   0
      Top             =   240
      Width           =   6030
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Calendar Position"
      Height          =   195
      Left            =   660
      TabIndex        =   12
      Top             =   5220
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
    Shape1.Move Command1(Index).Left - 30, Command1(Index).Top - 30, Command1(Index).Width + 50, Command1(Index).Height + 50
    If Shape1.Visible = False Then Shape1.Visible = True
End Sub

Private Sub Command2_Click()
Dim Stamp As New clsCalendarStamp
With Stamp
    .Background = vbBlack
    .BackgroundTrim = border
    
    .TrimDepth = 1
    
    .Left = 10
    .Top = 10
    
    .CalendarMonth = Month(Now)
    .CalendarYear = Year(Now)
    
    .TargetImage = picBuffer
    
    .DayBold = True
    .DayColor = RGB(220, 220, 220)
    .DayFont = "Arial"
    .DayFontSize = 12
    
    .LabelBold = True
    .LabelColor = RGB(255, 255, 220)
    .LabelFont = "Arial"
    .LabelFontSize = 12
    
    .TitleBold = True
    .TitleColor = RGB(230, 230, 255)
    .TitleFont = "Arial"
    .TitleFontSize = 24
    
    .TodayColor = RGB(255, 130, 155)
    
    .DrawCalendar
End With

End Sub




Private Sub Form_Load()
Dim x As Integer
    For x = 1 To 12
        cboMmm.AddItem Format(DateSerial(2003, x, 1), "Mmm")
    Next
    For x = 1960 To 2029
        cboYYYY.AddItem x
    Next
    
End Sub

Private Sub optDate_Click(Index As Integer)
    If Index = 1 Then
        cboMmm.Enabled = True
        cboYYYY.Enabled = True
    Else
        cboMmm.Enabled = False
        cboYYYY.Enabled = False
    End If
End Sub
