VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TIME OUT"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":030A
   ScaleHeight     =   4260
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "        Remaining Time"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   210
      TabIndex        =   9
      Top             =   2820
      Width           =   4305
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "This computer will"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   150
         TabIndex        =   10
         Top             =   450
         Visible         =   0   'False
         Width           =   2085
      End
   End
   Begin VB.CommandButton cmdabort 
      Caption         =   "ABORT"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2580
      Picture         =   "Form1.frx":1192B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1590
      Width           =   1455
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3180
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Restart"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1650
      TabIndex        =   5
      Top             =   240
      Width           =   1305
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Log Off"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton cmdexecute 
      Caption         =   "EXECUTE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   750
      Picture         =   "Form1.frx":11C35
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1590
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   990
      Left            =   4110
      Top             =   1830
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   3450
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   930
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   1830
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   930
      Width           =   855
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   930
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hour(s)            Minute(s)         Second(s)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   180
      TabIndex        =   0
      Top             =   630
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim c As String
Dim h, m, s As Integer

Private Sub enable()
Option1.Enabled = Not Option1.Enabled
Option2.Enabled = Not Option2.Enabled
Option3.Enabled = Not Option3.Enabled
Combo1.Enabled = Not Combo1.Enabled
Combo2.Enabled = Not Combo2.Enabled
Combo3.Enabled = Not Combo3.Enabled
cmdexecute.Enabled = Not cmdexecute.Enabled
cmdabort.Enabled = Not cmdabort.Enabled
Label2.Visible = Not Label2.Visible
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub cmdabort_Click()
If MsgBox("Are you sure ?", vbYesNo, "System Message") = vbYes Then
enable
Me.Caption = "TIME OUT"
End If
End Sub

Private Sub cmdexecute_Click()
If (Combo1.Text <> "" And Combo2.Text <> "" And Combo3.Text <> "") And (Option1.Value Or Option2.Value Or Option3.Value) Then
If Option1.Value Then 'same as of If Option1.Value=True Then
c = "log off"
ElseIf Option2.Value Then 'same as of ElseIf Option2.Value=True Then
c = "restart"
Else
c = "shutdown"
End If
h = Val(Combo1.Text)
m = Val(Combo2.Text)
s = Val(Combo3.Text)
enable
Else
MsgBox "Please fill up the following : Hour(s), Minute(s), Second(s) and also choose any of the following : Log Off, Restart and Shutdown.", vbExclamation, "System Message"
End If
End Sub

Private Sub Form_Load()
For h = 0 To 59
Combo1.AddItem h
Combo2.AddItem h
Combo3.AddItem h
Next h
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Option1.ForeColor = vbBlack
Option2.ForeColor = vbBlack
Option3.ForeColor = vbBlack
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("Are you sure ?", vbYesNo, "System Message") <> vbYes Then
Cancel = 1
Me.WindowState = 1
End If
End Sub

Private Sub Option1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Option1.ForeColor = vbBlue
End Sub

Private Sub Option2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Option2.ForeColor = vbBlue
End Sub

Private Sub Option3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Option3.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
Label2.Caption = "This computer will " & c & Chr(13) & "  itself after : " & Format(h, "00:") & Format(m, "00:") & Format(s, "00")
Me.Caption = Right(Label2.Caption, 8)
If h = 0 And m = 0 And s = 0 Then
If c = "log off" Then
Shell "shutdown -l" 'This line will log off your computer
ElseIf c = "restart" Then
Shell "shutdown -r -t 00" 'This line will restart your computer
Else
Shell "shutdown -s -t 00" 'This line will shutdown your computer
End If
End 'I write End because sometimes the shell shutdown command is being delayed
End If

If s = 0 Then
If m = 0 And h <> 0 Then
h = h - 1
m = 60
End If
m = m - 1
s = 60
End If
s = s - 1
End Sub
